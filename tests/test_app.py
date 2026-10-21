import asyncio
import json
import tempfile
import unittest
from pathlib import Path
from unittest.mock import patch

import app


class FakeRequest:
    def __init__(self, payload):
        self.payload = payload

    async def json(self):
        return self.payload


class AppTests(unittest.TestCase):
    def test_create_and_download(self):
        sample = Path(__file__).resolve().parents[1] / "examples" / "sample_invoice.json"
        payload = json.loads(sample.read_text(encoding="utf-8"))
        with tempfile.TemporaryDirectory() as directory:
            root = Path(directory)
            with patch.object(app, "DB_FILE", root / "numbers.sqlite3"), \
                 patch.object(app, "OUTPUT_DIR", root / "output"), \
                 patch.object(app, "SPEC_DIR", root / "specs"):
                result = asyncio.run(app.create_invoice(FakeRequest(payload)))
                self.assertEqual(result["number"], "PFI-2026-0001")
                self.assertEqual(result["vat_included"], "156.98")
                self.assertEqual(Path(app.download("PFI-2026-0001.pdf").path),
                                 root / "output" / "PFI-2026-0001.pdf")
                with self.assertRaises(Exception):
                    app.download("../secret.pdf")


if __name__ == "__main__":
    unittest.main()

