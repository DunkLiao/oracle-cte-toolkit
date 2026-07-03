import unittest
from unittest.mock import patch

from oracle_client import connect, oracle_type_from_description, wrap_metadata_sql


class OracleClientTests(unittest.TestCase):
    @patch("oracle_client.oracledb.connect")
    def test_connect_passes_standard_fields_to_oracledb(self, fake_connect):
        connect("user1", "secret", "db:1521/service")

        fake_connect.assert_called_once_with(
            user="user1",
            password="secret",
            dsn="db:1521/service",
        )

    def test_wrap_metadata_sql_strips_trailing_semicolon(self):
        wrapped = wrap_metadata_sql("select * from dual;")

        self.assertIn("select * from dual", wrapped)
        self.assertIn("WHERE 1=0", wrapped)
        self.assertNotIn("dual;", wrapped)

    def test_oracle_type_from_description_formats_number_precision_scale(self):
        col = ("AMOUNT", "NUMBER", None, None, 10, 2, None)

        self.assertEqual(oracle_type_from_description(col), "NUMBER(10,2)")


if __name__ == "__main__":
    unittest.main()
