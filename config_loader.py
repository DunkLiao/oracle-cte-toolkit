import json
import os
import sys
from pathlib import Path


def get_base_dir(reference_file=None):
    if getattr(sys, "frozen", False):
        return Path(sys.executable).resolve().parent
    if reference_file:
        return Path(reference_file).resolve().parent
    return Path(__file__).resolve().parent


def get_config_path(reference_file=None):
    return str(get_base_dir(reference_file) / "config.json")


def parse_dsn(dsn):
    host = ""
    port = "1521"
    service = ""

    try:
        if "/" in dsn:
            left, service = dsn.split("/", 1)
        else:
            left = dsn

        if ":" in left:
            host, port = left.split(":", 1)
        else:
            host = left
    except Exception:
        pass

    return host.strip(), port.strip(), service.strip()


def build_dsn(host, port, service):
    return f"{host.strip()}:{port.strip()}/{service.strip()}"


def build_config(
    username,
    password,
    dsn,
    sql_folder_path,
    output_excel_path,
):
    return {
        "database": {
            "username": username,
            "password": password,
            "dsn": dsn,
        },
        "sql_folder_path": sql_folder_path,
        "output_excel_path": output_excel_path,
    }


def load_config(config_path):
    with open(config_path, "r", encoding="utf-8") as f:
        return json.load(f)


def save_config(config_path, config):
    path = Path(config_path)
    parent = path.parent
    if str(parent):
        os.makedirs(parent, exist_ok=True)

    with open(path, "w", encoding="utf-8") as f:
        json.dump(config, f, indent=4, ensure_ascii=False)


def default_config(config_path):
    base_dir = Path(config_path).resolve().parent
    return build_config(
        username="your_username",
        password="your_password",
        dsn="your_host:1521/your_service",
        sql_folder_path=str(base_dir / "sql_files"),
        output_excel_path=str(base_dir / "output.xlsx"),
    )
