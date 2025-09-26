import os
import pyodbc
from contextlib import contextmanager

_DEFAULTS = {
    # coloque a porta se for a padrão pública
    "server":     os.getenv("SQLSERVER", "45.235.240.135,1433"),
    "database":   os.getenv("SQLDATABASE", "Stik"),
    "username":   os.getenv("SQLUSER", "ti"),
    "password":   os.getenv("SQLPASSWORD", "Stik0123"),
    "driver":     os.getenv("ODBC_DRIVER", "ODBC Driver 17 for SQL Server"),
    "trust_cert": os.getenv("TRUST_CERT", "yes"),
}

class DBConfig:
    def __init__(self, server=None, database=None, username=None, password=None,
                 driver=None, trust_cert=None):
        self.server = server or _DEFAULTS["server"]
        self.database = database or _DEFAULTS["database"]
        self.username = username or _DEFAULTS["username"]
        self.password = password or _DEFAULTS["password"]
        self.driver = driver or _DEFAULTS["driver"]
        self.trust_cert = trust_cert or _DEFAULTS["trust_cert"]

    def connection_string(self) -> str:
        parts = [
            f"DRIVER={{{self.driver}}}",
            f"SERVER={self.server}",
            f"DATABASE={self.database}",
            f"UID={self.username}",
            f"PWD={self.password}",
            "Encrypt=yes",
            f"TrustServerCertificate={'yes' if str(self.trust_cert).lower() in ['1','true','yes','y'] else 'no'}",
        ]
        return ";".join(parts)

    def connect(self):
        return pyodbc.connect(self.connection_string(), timeout=30)

@contextmanager
def get_conn(cfg: DBConfig):
    conn = cfg.connect()
    try:
        yield conn
    finally:
        conn.close()