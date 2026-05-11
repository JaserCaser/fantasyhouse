import os
import tempfile
import unittest
from pathlib import Path

import main


class SecurityBasicsTest(unittest.TestCase):
    def setUp(self):
        self._old_db_file = main.DB_FILE
        self._old_admin_init_password = os.environ.get("ADMIN_INIT_PASSWORD")
        self._old_cors_origins = os.environ.get("CORS_ORIGINS")

    def tearDown(self):
        main.DB_FILE = self._old_db_file
        if self._old_admin_init_password is None:
            os.environ.pop("ADMIN_INIT_PASSWORD", None)
        else:
            os.environ["ADMIN_INIT_PASSWORD"] = self._old_admin_init_password
        if self._old_cors_origins is None:
            os.environ.pop("CORS_ORIGINS", None)
        else:
            os.environ["CORS_ORIGINS"] = self._old_cors_origins

    def test_parse_cors_origins_from_env(self):
        os.environ["CORS_ORIGINS"] = "https://a.example.com, https://b.example.com"
        origins = main._parse_cors_origins()
        self.assertEqual(origins, ["https://a.example.com", "https://b.example.com"])

    def test_init_default_admin_uses_env_password(self):
        with tempfile.TemporaryDirectory() as tmpdir:
            main.DB_FILE = Path(tmpdir) / "test_kb.db"
            main.init_db()
            os.environ["ADMIN_INIT_PASSWORD"] = "StrongPass123"
            main.init_default_admin()

            conn = main.get_db()
            row = conn.execute(
                "SELECT username, password_hash FROM users WHERE username = ?",
                ("admin",),
            ).fetchone()
            conn.close()

            self.assertIsNotNone(row)
            self.assertEqual(row["username"], "admin")
            self.assertEqual(row["password_hash"], main.hash_pw("StrongPass123"))


if __name__ == "__main__":
    unittest.main()
