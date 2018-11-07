from flask import Flask
from pyflaskrest.main.controller import main
from pyflaskrest.config import configure_app
import os

app = Flask(__name__)
configure_app(app)


app.register_blueprint(main, url_prefix = '/')