import os

base_dir = os.path.abspath(os.path.dirname(__file__))

class Config:
    SECRET_KEY = os.environ.get('SECRET_KEY', 'savvytech-it')
    SAVVYTECH_ADMIN = os.environ.get('SAVVYTECH_ADMIN', 'admin')
    UPLOAD_FOLDER = '/home/vincentni/project/savvytech/savvytech/upload'
    ATTENDANCE_NAME = 'attendance'

    @staticmethod
    def init_app(app):
        pass

class ProductionConfig(Config):
    DEBUG = True
    MONGODB_DB = 'savvytech'
    MONGODB_CONNECT = False
    MONGODB_USERNAME = ''
    MONGODB_PASSWORD = ''

config = {
    'production': ProductionConfig,
    'default': ProductionConfig
}
