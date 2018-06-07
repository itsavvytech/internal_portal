from flask import Blueprint

payroll = Blueprint('payroll', __name__, url_prefix='/payroll')
from . import views
