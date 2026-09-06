import json
import tempfile
import unittest
from pathlib import Path
from unittest.mock import patch
import httpx
from openai import OpenAI
from extracao_openai import extract_openai
from extracao_local import extract_local
from extrair_recebiveis_gemini import UserFacingError

ROOT=Path(__file__).resolve().parents[1]


class EngineTests(unittest.TestCase):
    def test_openai_real_sdk_with_mock_http(self):
        payload=json.loads((ROOT/'exemplo_extracao.json').read_text(encoding='utf-8'))
        captured=[]
        def handler(request):
            body=json.loads(request.content)
            captured.append(body)
            self.assertEqual(request.url.host,'api.openai.com')
            return httpx.Response(200,json={'id':'resp_test','object':'response','created_at':1,'status':'completed','model':'test-model','output':[{'id':'msg_test','type':'message','status':'completed','role':'assistant','content':[{'type':'output_text','text':json.dumps(payload),'annotations':[]}]}]})
        client=OpenAI(api_key='synthetic-test-only',http_client=httpx.Client(transport=httpx.MockTransport(handler)))
        with tempfile.TemporaryDirectory() as directory, patch('openai.OpenAI',return_value=client):
            pdf=Path(directory)/'test.pdf'; pdf.write_bytes(b'%PDF-synthetic')
            result=extract_openai(pdf,'test-model',1,'synthetic-test-only')
        self.assertEqual(len(result.tables),3)
        self.assertFalse(captured[0]['store'])
        self.assertEqual(captured[0]['text']['format']['type'],'json_schema')
        self.assertTrue(captured[0]['input'][1]['content'][0]['file_data'].startswith('data:application/pdf;base64,'))

    def test_openai_error_never_echoes_payload(self):
        client=OpenAI(api_key='synthetic-test-only',http_client=httpx.Client(transport=httpx.MockTransport(lambda request:httpx.Response(401,json={'error':{'message':'SECRET_PAYLOAD','type':'invalid_request_error'}}))))
        with tempfile.TemporaryDirectory() as directory, patch('openai.OpenAI',return_value=client):
            pdf=Path(directory)/'test.pdf'; pdf.write_bytes(b'%PDF-synthetic')
            with self.assertRaises(UserFacingError) as error:
                extract_openai(pdf,'test-model',1,'synthetic-test-only')
        self.assertIn('OPENAI_API_KEY',str(error.exception))
        self.assertNotIn('SECRET_PAYLOAD',str(error.exception))

    def test_local_rejects_unknown_without_network(self):
        from pypdf import PdfWriter
        with tempfile.TemporaryDirectory() as directory, patch('socket.socket',side_effect=AssertionError('network forbidden')):
            pdf=Path(directory)/'test.pdf'; w=PdfWriter(); w.add_blank_page(width=100,height=100); w.write(pdf)
            with self.assertRaises(UserFacingError):
                extract_local(pdf,1)

    def test_no_legacy_key_fallback(self):
        import app
        with patch.dict('os.environ',{'GEMINI_API_KEY':'synthetic-test-only','OPENAI_API_KEY':''}),patch.object(app.st,'secrets',{}):
            with self.assertRaises(UserFacingError):
                app.configure_api_key()
