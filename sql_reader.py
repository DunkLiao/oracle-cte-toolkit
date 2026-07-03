import os
from pathlib import Path

import chardet


DEFAULT_EXTENSIONS = (".sql", ".txt")


def detect_encoding(file_path):
    try:
        with open(file_path, "rb") as f:
            raw = f.read()
    except Exception:
        return "utf-8"

    result = chardet.detect(raw)
    return result.get("encoding") or "utf-8"


def read_sql_file(file_path):
    path = Path(file_path)
    detected = detect_encoding(path)
    encodings = []

    for encoding in ("utf-8-sig", "utf-8", "cp950", "big5", detected):
        if encoding and encoding not in encodings:
            encodings.append(encoding)

    last_error = None
    for encoding in encodings:
        try:
            with open(path, "r", encoding=encoding) as f:
                return f.read()
        except UnicodeDecodeError as exc:
            last_error = exc

    try:
        with open(path, "r", encoding=detected or "utf-8", errors="ignore") as f:
            return f.read()
    except Exception:
        if last_error:
            raise last_error
        raise


def read_sql_files(folder, extensions=DEFAULT_EXTENSIONS):
    folder_path = Path(folder)
    sql_files = {}

    for item in sorted(os.listdir(folder_path)):
        if not item.lower().endswith(extensions):
            continue

        path = folder_path / item
        sql_files[path.stem] = read_sql_file(path)

    return sql_files
