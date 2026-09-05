import tempfile
import unittest
from pathlib import Path
from unittest.mock import patch, MagicMock
from datetime import date

from openpyxl import load_workbook
from pypdf import PdfWriter
from extrair_recebiveis_gemini import (
    coerce_value, load_extraction_json, build_workbook, validate_pdf,
    extract_with_gemini, UserFacingError,
)

ROOT = Path(__file__).resolve().parents[1]


class SecurityTests(unittest.TestCase):
    def test_all_text_surfaces_and_empty_table(self):
        extraction = load_extraction_json(ROOT / 'exemplo_extracao.json')
        for table in extraction.tables:
            table.title = '=1+1'
            table.sheet_name = 'Mesmo/nome'
            for column in table.columns:
                column.name = '=1+1'
        extraction.tables[-1].rows = []
        with tempfile.TemporaryDirectory() as directory:
            path = Path(directory) / 'test.xlsx'
            build_workbook(extraction, path)
            wb = load_workbook(path)
            self.assertEqual(len(set(wb.sheetnames)), 3)
            for ws in wb:
                self.assertTrue(ws.tables)
                headers = [c.value for c in ws[4]]
                self.assertEqual(len(headers), len(set(headers)))
                self.assertFalse(any(c.data_type == 'f' for row in ws for c in row))
            wb.close()

    def test_types(self):
        self.assertEqual(coerce_value('31/08/2026', 'date'), date(2026, 8, 31))
        self.assertEqual(coerce_value('-12,30', 'currency'), -12.3)
        for prefix in '=+-@':
            self.assertTrue(coerce_value(prefix + 'abc', 'text').startswith("'"))

    def test_pdf_validation(self):
        with tempfile.TemporaryDirectory() as directory:
            path = Path(directory) / 'test.pdf'
            path.write_bytes(b'not pdf')
            with self.assertRaises(ValueError):
                validate_pdf(path, 10, 10)
            writer = PdfWriter()
            for _ in range(11):
                writer.add_blank_page(width=100, height=100)
            writer.write(path)
            with self.assertRaises(ValueError):
                validate_pdf(path, 10, 10)
            self.assertEqual(validate_pdf(path, 11, 10), 11)

    def test_sdk_contract_offline(self):
        from google import genai
        client = MagicMock()
        client.interactions.create.return_value.output_text = (ROOT / 'exemplo_extracao.json').read_text(encoding='utf-8')
        with tempfile.TemporaryDirectory() as directory, patch.object(genai, 'Client', return_value=client):
            path = Path(directory) / 'test.pdf'
            path.write_bytes(b'%PDF-synthetic')
            result = extract_with_gemini(path, 'test-model', 3, 1, api_key='synthetic-test-only')
            self.assertEqual(len(result.tables), 3)
            args = client.interactions.create.call_args.kwargs
            self.assertFalse(args['store'])
            self.assertEqual(args['response_format']['mime_type'], 'application/json')
            client.close.assert_called_once()

    def test_sdk_error_is_sanitized(self):
        from google import genai
        client = MagicMock()
        error = RuntimeError('SENSITIVE_MARKER')
        error.status_code = 404
        client.interactions.create.side_effect = error
        with tempfile.TemporaryDirectory() as directory, patch.object(genai, 'Client', return_value=client):
            path = Path(directory) / 'test.pdf'
            path.write_bytes(b'%PDF-synthetic')
            with self.assertRaises(UserFacingError) as caught:
                extract_with_gemini(path, 'missing', 1, 2, api_key='synthetic-test-only')
            self.assertNotIn('SENSITIVE_MARKER', str(caught.exception))
            self.assertIn('GEMINI_MODEL', str(caught.exception))
            self.assertEqual(client.interactions.create.call_count, 1)


if __name__ == '__main__':
    unittest.main()
