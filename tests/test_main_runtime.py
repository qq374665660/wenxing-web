import io
import unittest

from fastapi.testclient import TestClient
from openpyxl import Workbook

from api.main_runtime import app, sanitize_upload_filename


def build_workbook_bytes(rows):
    workbook = Workbook()
    worksheet = workbook.active
    worksheet.title = "1.6地层信息"
    worksheet.append(["层号", "预留1", "预留2", "预留3", "预留4", "预留5", "预留6", "土层名称"])
    for row in rows:
        worksheet.append(row)

    stream = io.BytesIO()
    workbook.save(stream)
    stream.seek(0)
    return stream.getvalue()


class MainRuntimeTestCase(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        cls.client = TestClient(app)

    def test_health_endpoint_returns_ok(self):
        response = self.client.get("/api/health")

        self.assertEqual(response.status_code, 200)
        self.assertEqual(response.json()["status"], "ok")

    def test_root_returns_formal_frontend(self):
        response = self.client.get("/")

        self.assertEqual(response.status_code, 200)
        self.assertIn("地基基础分析", response.text)

    def test_precheck_detects_silt_and_silty_clay(self):
        workbook_bytes = build_workbook_bytes(
            [
                ["1-1", "", "", "", "", "", "", "粉土"],
                ["1-2", "", "", "", "", "", "", "粉质黏土"],
            ]
        )

        response = self.client.post(
            "/api/precheck",
            files={
                "file": (
                    "template.xlsx",
                    workbook_bytes,
                    "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                )
            },
        )

        self.assertEqual(response.status_code, 200)
        self.assertEqual(response.json(), {"has_silt": True, "has_silty_clay": True})

    def test_sanitize_upload_filename_removes_invalid_characters(self):
        sanitized = sanitize_upload_filename("bad<name>.xlsx")
        fallback_name = sanitize_upload_filename('<>:"/\\\\|?*.xlsx')

        self.assertEqual(sanitized, "bad_name.xlsx")
        self.assertEqual(fallback_name, "upload.xlsx")


if __name__ == "__main__":
    unittest.main()
