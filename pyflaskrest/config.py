import os
import logging

class BaseConfig(object):
    DEBUG = False
    TESTING = False
    LOGGING_FORMAT = '%(asctime)s - %(name)s - %(levelname)s - %(message)s'
    LOGGING_LOCATION = 'pyflaskrest.log'
    LOGGING_LEVEL = logging.DEBUG
    CACHE_TYPE = 'simple'
    SUPPORTED_LANGUAGES = {'en': 'English'}
    BABEL_DEFAULT_LOCALE = 'en'
    BABEL_DEFAULT_TIMEZONE = 'UTC'

class DevelopmentConfig(BaseConfig):
    DEBUG = True
    TESTING = False
    ENV = 'dev'

class StagingConfig(BaseConfig):
    DEBUG = False
    TESTING = True
    ENV = 'staging'

class ProductionConfig(BaseConfig):
    DEBUG = False
    TESTING = False
    ENV = 'prod'

config = {
    "dev": "pyflaskrest.config.DevelopmentConfig",
    "staging": "pyflaskrest.config.StagingConfig",
    "prod": "pyflaskrest.config.ProductionConfig",
    "default": "pyflaskrest.config.DevelopmentConfig"
}

def configure_app(app):
    config_name = os.getenv('FLASK_CONFIGURATION', 'default')
    app.config.from_object(config[config_name])
    app.config.from_pyfile('config.cfg', silent=True)
    handler = logging.FileHandler(app.config['LOGGING_LOCATION'])
    handler.setLevel(app.config['LOGGING_LEVEL'])
    formatter = logging.Formatter(app.config['LOGGING_FORMAT'])
    handler.setFormatter(formatter)
    app.logger.addHandler(handler)