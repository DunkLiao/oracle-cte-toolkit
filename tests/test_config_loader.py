import json
import tempfile
import unittest
from pathlib import Path

from config_loader import (
    build_config,
    build_dsn,
    get_config_path,
    load_config,
    parse_dsn,
    save_config,
)


class ConfigLoaderTests(unittest.TestCase):
    def test_parse_dsn_splits_host_port_and_service(self):
        self.assertEqual(
            parse_dsn("db.example.com:1522/ORCLPDB1"),
            ("db.example.com", "1522", "ORCLPDB1"),
        )

    def test_parse_dsn_defaults_port_when_missing(self):
        self.assertEqual(
            parse_dsn("db.example.com/ORCL"),
            ("db.example.com", "1521", "ORCL"),
        )

    def test_build_dsn_trims_parts(self):
        self.assertEqual(
            build_dsn(" db.example.com ", " 1521 ", " ORCL "),
            "db.example.com:1521/ORCL",
        )

    def test_save_and_load_config_round_trip(self):
        with tempfile.TemporaryDirectory() as tmp:
            path = Path(tmp) / "config.json"
            config = build_config(
                username="user1",
                password="secret",
                dsn="db:1521/service",
                sql_folder_path=str(Path(tmp) / "sql"),
                output_excel_path=str(Path(tmp) / "out.xlsx"),
            )

            save_config(path, config)
            loaded = load_config(path)

        self.assertEqual(loaded, config)

    def test_get_config_path_uses_reference_file_directory(self):
        with tempfile.TemporaryDirectory() as tmp:
            ref = Path(tmp) / "tool.py"
            ref.write_text("pass", encoding="utf-8")

            self.assertEqual(
                get_config_path(ref),
                str(Path(tmp) / "config.json"),
            )


if __name__ == "__main__":
    unittest.main()
