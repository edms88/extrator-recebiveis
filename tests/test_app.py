import io
import tempfile
import unittest
from pathlib import Path
from unittest.mock import patch

from pypdf import PdfWriter
from streamlit.testing.v1 import AppTest
import app
from extrair_recebiveis_gemini import load_extraction_json

ROOT = Path(__file__).resolve().parents[1]


class AppTests(unittest.TestCase):
    def test_initial_screen(self):
        at = AppTest.from_file(str(ROOT / 'app.py')).run()
        self.assertFalse(at.exception)
        self.assertEqual(at.info[0].value, 'Selecione um PDF para iniciar.')

    def test_process_and_cleanup_offline(self):
        writer = PdfWriter()
        writer.add_blank_page(width=595, height=842)
        buffer = io.BytesIO()
        writer.write(buffer)
        extraction = load_extraction_json(ROOT / 'exemplo_extracao.json')
        state = {}
        with tempfile.TemporaryDirectory() as temp_root:
            with patch.object(app, 'configure_api_key', return_value='synthetic-test-only'), patch.object(app, 'configured_model', return_value='test-model'), patch.object(app, 'extract_with_gemini', return_value=extraction), patch.object(app.st, 'session_state', state), patch.object(tempfile, 'tempdir', temp_root):
                app.process_pdf('anonimizado.pdf', buffer.getvalue())
                self.assertEqual(state['summary']['tables'], 3)
                self.assertTrue(state['xlsx_bytes'].startswith(b'PK'))
                self.assertTrue(state['json_bytes'])
                self.assertEqual(list(Path(temp_root).iterdir()), [])

    def test_upload_rejections(self):
        for name, content in [('file.txt', b'%PDF-'), ('file.pdf', b'invalid'), ('file.pdf', b'%PDF-' + b'0' * (10 * 1024 * 1024))]:
            with self.subTest(name=name, size=len(content)), self.assertRaises(ValueError):
                app.validate_upload(name, content)

    def test_stale_result_is_cleared(self):
        state = {'file_digest': 'old', 'xlsx_bytes': b'old', 'summary': {}}
        with patch.object(app.st, 'session_state', state):
            app.clear_previous_result('new')
        self.assertNotIn('xlsx_bytes', state)
