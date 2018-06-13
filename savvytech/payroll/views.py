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
        timecard_preprocess(file_path, file_name)
        return send_from_directory(directory=current_app.config.get('UPLOAD_FOLDER'),
                                   filename=current_app.config.get('TIMECARD_NAME') + '.xls', as_attachment=True)
    flash('Please upload your raw timecard file! Then complete the downloaded timecard and upload it in the commission tab!')
    return render_template('payroll/index.html', form=form)

@payroll.route('/commission', methods=['GET', 'POST'])
def commission_1():
    form = UploadForm()
    form.file.description = 'Step 2 a): Please upload completed timecard here.'
    if form.validate_on_submit():
        file = request.files['file']
        file_name = secure_filename(file.filename)
        file_path = current_app.config.get('UPLOAD_FOLDER')
        file.save(os.path.join(file_path, file_name))
        timecard_postprocess(file_path, file_name)
        return redirect(url_for('payroll.commission_2'))
    flash('Please upload your completed timecard file!')
    return render_template('payroll/index.html', form=form)

@payroll.route('/commission#', methods=['GET', 'POST'])
def commission_2():
    form = UploadForm()
    form.file.description = 'Step 2 b): Please upload raw commission file here.'
    if form.validate_on_submit():
        file = request.files['file']
        file_name = secure_filename(file.filename)
        file_path = current_app.config.get('UPLOAD_FOLDER')
        file.save(os.path.join(file_path, file_name))
        commission_preprocess(file_path, file_name)
        return send_from_directory(directory=current_app.config.get('UPLOAD_FOLDER'),
                                   filename=current_app.config.get('COMMISSION_NAME') + '.xls', as_attachment=True)
    flash('Please upload your raw commission file! Then complete the downloaded commission file and upload it in the results tab!')
    return render_template('payroll/index.html', form=form)

@payroll.route('/results', methods=['GET', 'POST'])
def results_1():
    form = UploadForm()
    form.file.description = 'Step 3): Please upload completed commission file here.'
    if form.validate_on_submit():
        file = request.files['file']
        file_name = secure_filename(file.filename)
        file_path = current_app.config.get('UPLOAD_FOLDER')
        file.save(os.path.join(file_path, file_name))
        commission_postprocess(file_path, file_name)
        return redirect(url_for('payroll.results_2'))
    flash('Please upload your completed commission file!')
    return render_template('payroll/index.html', form=form)

@payroll.route('/results#')
def results_2():
    for file in os.listdir(current_app.config.get('UPLOAD_FOLDER')):
        if file.endswith(".xls") or file.endswith(".xlsx"):
            os.remove(os.path.join(current_app.config.get('UPLOAD_FOLDER'), file))

    with open(os.path.join(current_app.config.get('UPLOAD_FOLDER'), current_app.config.get('TIMECARD_NAME') + '.json')) as fp:
        timecard_set = json.load(fp)
        with open(os.path.join(current_app.config.get('UPLOAD_FOLDER'), current_app.config.get('COMMISSION_NAME') + '.json')) as fp:
            commission_set = json.load(fp)
            return render_template('payroll/results.html', timecard_set=timecard_set, commission_set=commission_set)

def compute_time(time_list):
    if len(time_list) < 4:
        return None
    period_1 = 60*int(time_list[1][:2])+int(time_list[1][3:]) - 60*int(time_list[0][:2])-int(time_list[0][3:])
    period_2 = 60*int(time_list[3][:2])+int(time_list[3][3:]) - 60*int(time_list[2][:2])-int(time_list[2][3:])
    return round((period_1 + period_2) / 60, 2)

def timecard_preprocess(path, file_name):
    book = xlrd.open_workbook(os.path.join(path, file_name))
    sh = book.sheet_by_index(0)

    records_dict = {}
    for row_num in range(1, sh.nrows):
        row = sh.row(row_num)
        person_id = int(row[2].value)
        person_name = row[3].value
        person_name = '-'.join([person_name.upper(), str(person_id)])
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
        row = sh.row(0)
        row.write(0, 'Date')
        row.write(1, 'Checkin')
        row.write(2, 'Checkout')
        row.write(3, 'Checkin')
        row.write(4, 'Checkout')
        row.write(5, 'Working Hours')
        row.write(6, 'Overtime Hours')
        row.write(7, 'Regular Hours')

        index = 1
        total_hours = []
        overtime_hours = []
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
                overtime_hours.append(0 if total_hour <= 8 else total_hour-8)
                row.write(i+3, overtime_hours[-1])
                reg_hours.append(8 if total_hour > 8 else total_hour)
                row.write(i+4, reg_hours[-1])
        row = sh.row(index)
        row.write(4, 'Total')
        row.write(5, sum(total_hours))
        row.write(6, sum(overtime_hours))
        row.write(7, sum(reg_hours))

        row = sh.row(index + 1)
        row.write(0, 'Sick Hours')  # sick hour
        row.write(1, 'Vacation Hours')  # vacation hour
        row.write(2, 'Holiday Hours')  # holiday hour
        row = sh.row(index + 2)
        row.write(0, 0) # sick hour
        row.write(1, 0) # vacation hour
        row.write(2, 0) # holiday hour
    book.save(os.path.join(path, current_app.config.get('TIMECARD_NAME')+'.xls'))

    os.remove(os.path.join(path, file_name))

def timecard_postprocess(path, file_name):
    book = xlrd.open_workbook(os.path.join(path, file_name))

    records_dict = {}
    for sheet in book.sheets():
        person_name = sheet.name
        total_hour = 0
        overtime_hour = 0
        reg_hour = 0

        row_num = sheet.nrows
        for i in range(1, row_num - 3):
            current_hour = compute_time([sheet.cell(i, 1).value, sheet.cell(i, 2).value, sheet.cell(i, 3).value, sheet.cell(i, 4).value])
            total_hour += current_hour
            overtime_hour += 0 if current_hour <= 8 else current_hour-8
            reg_hour += 8 if current_hour > 8 else current_hour
        additional_hour = float(sheet.cell(row_num-1, 0).value) + float(sheet.cell(row_num-1, 1).value) + float(sheet.cell(row_num-1, 2).value)
        total_hour += additional_hour
        reg_hour += additional_hour
        records_dict[person_name] = [round(total_hour, 2), round(overtime_hour, 2), round(reg_hour, 2), round(additional_hour, 2)]

    with open(os.path.join(path, current_app.config.get('TIMECARD_NAME') + '.json'), 'w') as fp:
        json.dump(records_dict, fp)

    os.remove(os.path.join(path, file_name))

def compute_bonus_commission(sales, bucket_list):
    bonus, commission = 0, 0
    for index, bucket in enumerate(bucket_list):
        if sales > bucket[0]:
            commission += (bucket[0] - bucket_list[index-1][0]) * bucket[1] if index > 0 else bucket[0] * bucket[1]
            bonus = bucket[2]
        else:
            commission += (sales - bucket_list[index-1][0]) * bucket[1] if index > 0 else sales * bucket[1]
            break
    if sales > bucket_list[-1][0]:
        commission += (sales - bucket_list[-1][0]) * bucket_list[-1][1]
    return round(bonus, 2), round(commission, 2)

def commission_preprocess(path, file_name):
    book = xlrd.open_workbook(os.path.join(path, file_name))
    sh = book.sheet_by_index(0)

    records_dict = {}
    for row_num in range(1, sh.nrows):
        row = sh.row(row_num)
        person_id = int(row[0].value)
        person_name = '-'.join([row[1].value.upper(), str(person_id)])
        final_sale = float(row[2].value)

        bonus, commission = compute_bonus_commission(final_sale, current_app.config.get('SALES_PARAMS')[person_name])
        records_dict[person_name] = [bonus, commission, 0, 0, bonus + commission, round(bonus/2, 2), round(commission/2, 2), round(bonus/2 + commission/2, 2)]

    book = xlwt.Workbook()
    for name, records in records_dict.items():
        sh = book.add_sheet(name)
        row = sh.row(0)
        row.write(0, 'Bonus')
        row.write(1, 'Commission')
        row.write(2, 'KPI Bonus')
        row.write(3, 'Extra Bonus')
        row.write(4, 'Total')
        row.write(5, '1/2 Bonus')
        row.write(6, '1/2 Commission')
        row.write(7, 'Period Paid')

        row = sh.row(1)
        for i, record in enumerate(records):
            row.write(i, record)
    book.save(os.path.join(path, current_app.config.get('COMMISSION_NAME')+'.xls'))

def commission_postprocess(path, file_name):
    book = xlrd.open_workbook(os.path.join(path, file_name))

    records_dict = {}
    for sh in book.sheets():
        person_name = sh.name

        bonus = float(sh.cell(1, 0).value)
        commission = float(sh.cell(1, 1).value)
        kpi = float(sh.cell(1, 2).value)
        extra = float(sh.cell(1, 3).value)
        records_dict[person_name] = [round(commission, 2), round(bonus, 2), round(kpi, 2), round(extra, 2)]

    with open(os.path.join(path, current_app.config.get('COMMISSION_NAME') + '.json'), 'w') as fp:
        json.dump(records_dict, fp)

    os.remove(os.path.join(path, file_name))