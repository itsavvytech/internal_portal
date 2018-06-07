from flask import Flask, Blueprint
from flask_bootstrap import Bootstrap
from flask_moment import Moment
from flask_mongoengine import MongoEngine
from config import config

bootstrap = Bootstrap()
moment = Moment()
db = MongoEngine()

main = Blueprint('main', __name__)
from . import views

def create_app(config_name):
    app = Flask(__name__)
    app.config.from_object(config[config_name])
    config[config_name].init_app(app)

    bootstrap.init_app(app)
    moment.init_app(app)
    db.init_app(app)

    app.register_blueprint(main)
    from .payroll import payroll
    app.register_blueprint(payroll)

    return app
