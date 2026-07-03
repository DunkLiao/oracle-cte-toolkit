import tempfile
import unittest
from pathlib import Path

from sql_reader import detect_encoding, read_sql_file, read_sql_files


class SqlReaderTests(unittest.TestCase):
    def test_reads_utf8_sql_file(self):
        with tempfile.TemporaryDirectory() as tmp:
            path = Path(tmp) / "a.sql"
            path.write_text("select '中文' as name from dual", encoding="utf-8")

            self.assertIn("中文", read_sql_file(path))

    def test_reads_cp950_sql_file(self):
        with tempfile.TemporaryDirectory() as tmp:
            path = Path(tmp) / "a.sql"
            path.write_bytes("select '中文' as name from dual".encode("cp950"))

            self.assertIn("中文", read_sql_file(path))

    def test_read_sql_files_returns_sorted_sql_and_txt_by_stem(self):
        with tempfile.TemporaryDirectory() as tmp:
            folder = Path(tmp)
            (folder / "b.txt").write_text("select 2 from dual", encoding="utf-8")
            (folder / "a.sql").write_text("select 1 from dual", encoding="utf-8")
            (folder / "ignore.md").write_text("ignored", encoding="utf-8")

            result = read_sql_files(folder)

        self.assertEqual(list(result.keys()), ["a", "b"])
        self.assertEqual(result["a"], "select 1 from dual")
        self.assertEqual(result["b"], "select 2 from dual")

    def test_detect_encoding_returns_a_string(self):
        with tempfile.TemporaryDirectory() as tmp:
            path = Path(tmp) / "a.sql"
            path.write_text("select 1 from dual", encoding="utf-8")

            self.assertIsInstance(detect_encoding(path), str)


if __name__ == "__main__":
    unittest.main()
