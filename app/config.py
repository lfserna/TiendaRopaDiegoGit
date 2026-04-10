import os


class Config:
    SECRET_KEY = os.getenv("SECRET_KEY", "cambiar-esta-clave-en-produccion")
    ADMIN_USER = os.getenv("ADMIN_USER", "Diego")
    ADMIN_PASSWORD = os.getenv("ADMIN_PASSWORD", "Diego")

    BASE_DIR = os.path.abspath(os.path.join(os.path.dirname(__file__), ".."))
    DATA_DIR = os.path.join(BASE_DIR, "data")
    REGISTROS_DIR = os.path.join(DATA_DIR, "registros")
    CONFIG_DIR = os.path.join(DATA_DIR, "config")

    PEDIDOS_FILE = os.path.join(REGISTROS_DIR, "pedidos.xlsx")
    CODIGOS_FILE = os.path.join(REGISTROS_DIR, "codigos.xlsx")
    CONFIG_FILE = os.path.join(CONFIG_DIR, "configuracion.xlsx")
