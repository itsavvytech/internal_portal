from datetime import datetime
from flask import render_template, redirect, url_for, request, current_app, send_from_directory, flash
from flask_wtf import FlaskForm
from flask_wtf.file import FileAllowed, FileRequired
from wtforms import FileField, SubmitField
from werkzeug.utils import secure_filename
from . import payroll
import os, xlrd, xlwt, json
from collections import OrderedDict

class UploadForm(FlaskForm):
    file = FileField('', validators=[FileAllowed(['xlsx', 'xls'], 'Expecting Microsoft Excel files'), FileRequired()])
    submit = SubmitField('Upload')

@payroll.route('/timecard', methods=['GET', 'POST'])
def index():
    form = UploadForm()
    form.file.description = 'Step 1: Please upload raw time card here.'
    if form.validate_on_submit():
        file = request.files['file']
        file_name = secure_filename(file.filename)
        file_path = current_app.config.get('UPLOAD_FOLDER')
        file.save(os.path.join(file_path, file_name))
        file_process(file_path, file_name)
        return send_from_directory(directory=current_app.config.get('UPLOAD_FOLDER'),
                                   filename=current_app.config.get('ATTENDANCE_NAME') + '.xls', as_attachment=True)
    return render_template('payroll/index.html', form=form)

@payroll.route('/commission', methods=['GET', 'POST'])
def commission_1():
    form = UploadForm()
    form.file.description = 'Step 2 a): Please upload good time card here.'
    if form.validate_on_submit():
        file = request.files['file']
        file_path = current_app.config.get('UPLOAD_FOLDER')
        file.save(os.path.join(file_path, current_app.config.get('ATTENDANCE_NAME') + '.xls'))
        flash('Time card was successfully uploaded!')
        return redirect(url_for('payroll.commission_2'))
    return render_template('payroll/index.html', form=form)

@payroll.route('/commission#', methods=['GET', 'POST'])
def commission_2():
    form = UploadForm()
    form.file.description = 'Step 2 b): Please upload commission file here.'
    if form.validate_on_submit():
        file = request.files['file']
        file_path = current_app.config.get('UPLOAD_FOLDER')
        file.save(os.path.join(file_path, current_app.config.get('COMMISSION_NAME') + '.xls'))
        return redirect(url_for('payroll.results'))
    return render_template('payroll/index.html', form=form)

@payroll.route('/results')
def results():
    from datetime import datetime
    return render_template('index.html', current_time=datetime.utcnow())

def compute_time(time_list):
    if len(time_list) < 4:
        return None
    period_1 = 60*int(time_list[1][:2])+int(time_list[1][3:]) - 60*int(time_list[0][:2])-int(time_list[0][3:])
    period_2 = 60*int(time_list[3][:2])+int(time_list[3][3:]) - 60*int(time_list[2][:2])-int(time_list[2][3:])
    return round((period_1 + period_2) / 60, 2)

def file_process(path, file_name):
    book = xlrd.open_workbook(os.path.join(path, file_name))
    sh = book.sheet_by_index(0)

    records_dict = {}
    for row_num in range(1, sh.nrows):
        row = sh.row(row_num)
        person_id = int(row[2].value)
        person_name = row[3].value
        person_name = '-'.join([person_name, str(person_id)])
        date = row[5].value
        date, time = str(xlrd.xldate.xldate_as_datetime(date, book.datemode)).split()
        time = time[:-3] # sort out seconds
        if person_name not in records_dict:
            records_dict[person_name] = OrderedDict()
            records_dict[person_name][date] = [time]
        else:
            if date not in records_dict[person_name]:
                records_dict[person_name][date] = [time]
            else:
                if len(records_dict[person_name][date]) < 4:
                    records_dict[person_name][date].append(time)

    book = xlwt.Workbook()
    for name, records in records_dict.items():
        sh = book.add_sheet(name)
        index = 0
        total_hours = []
        extra_hours = []
        reg_hours = []
        for date, time_list in records.items():
            row = sh.row(index)
            row.write(0, date)
            for i, time in enumerate(time_list):
                row.write(i+1, time)
            index += 1
            total_hour = compute_time(time_list)
            if total_hour:
                total_hours.append(total_hour)
                row.write(i+2, total_hours[-1])
                extra_hours.append(0 if total_hour <= 8 else total_hour-8)
                row.write(i+3, extra_hours[-1])
                reg_hours.append(8 if total_hour > 8 else total_hour)
                row.write(i+4, reg_hours[-1])
        row = sh.row(index)
        row.write(5, sum(total_hours))
        row.write(6, sum(extra_hours))
        row.write(7, sum(reg_hours))
    book.save(os.path.join(path, current_app.config.get('ATTENDANCE_NAME')+'.xls'))

    with open(os.path.join(path, current_app.config.get('ATTENDANCE_NAME')+'.json'), 'w') as fp:
        json.dump(records_dict, fp)

    os.remove(os.path.join(path, file_name))
