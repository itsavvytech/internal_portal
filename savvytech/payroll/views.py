from datetime import datetime
from flask import render_template, redirect, url_for, request, current_app
from flask_wtf import FlaskForm
from flask_wtf.file import FileAllowed, FileRequired
from wtforms import FileField, SubmitField
from werkzeug.utils import secure_filename
from . import payroll
import os, xlrd, json
from collections import OrderedDict

class UploadForm(FlaskForm):
    file = FileField('', validators=[FileAllowed(['xlsx', 'xls'], 'Expecting Microsoft Excel files'), FileRequired()])
    submit = SubmitField('Upload')

@payroll.route('/', methods=['GET', 'POST'])
def index():
    form = UploadForm()
    form.file.description = 'Step1: Please upload the attendance file here.'
    if form.validate_on_submit():
        file = request.files['file']
        file_name = secure_filename(file.filename)
        file_path = current_app.config.get('UPLOAD_FOLDER')
        file.save(os.path.join(file_path, file_name))
        file_process(file_path, file_name)
        return redirect(url_for('payroll.attendance'))
    return render_template('payroll/index.html', form=form)

@payroll.route('/attendance', methods=['GET', 'POST'])
def attendance():
    with open(os.path.join(current_app.config.get('UPLOAD_FOLDER'), current_app.config.get('ATTENDANCE_NAME'))) as fp:
        data_set = json.load(fp)
        return render_template('payroll/form.html', data_set=data_set)


def file_process(path, file_name):
    book = xlrd.open_workbook(os.path.join(path, file_name))
    sh = book.sheet_by_index(0)

    records_dict = {}
    for row_num in range(1, sh.nrows):
        row = sh.row(row_num)
        person_name = row[3].value
        date = row[5].value
        date, time = str(xlrd.xldate.xldate_as_datetime(date, book.datemode)).split()
        if person_name not in records_dict:
            records_dict[person_name] = OrderedDict()
            records_dict[person_name][date] = [time]
        else:
            if date not in records_dict[person_name]:
                records_dict[person_name][date] = [time]
            else:
                records_dict[person_name][date].append(time)

    with open(os.path.join(path, current_app.config.get('ATTENDANCE_NAME')), 'w') as fp:
        json.dump(records_dict, fp)

    os.remove(os.path.join(path, file_name))
