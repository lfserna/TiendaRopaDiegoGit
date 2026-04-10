from flask import Flask

from app.config import Config
from app.routes.admin_routes import admin_bp
from app.routes.user_routes import user_bp
from app.services.excel_service import excel_service


def create_app() -> Flask:
    app = Flask(__name__)
    app.config.from_object(Config)

    with app.app_context():
        excel_service.ensure_files()

    app.register_blueprint(user_bp)
    app.register_blueprint(admin_bp)

    return app
