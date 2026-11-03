from io import BytesIO
import unicodedata

import pytest
from fastapi.testclient import TestClient
from openpyxl import Workbook

from app.classifier import classify
from app.main import app
from app.normalize import normalize_names
from app.config import Settings


def _name_table():
    wb = Workbook()
    ws = wb.active
    ws.title = '지배구조'
    ws.append(['HLB글로벌(주) 특수관계자 지배구조'])
    ws.append([])
    ws.append(['정규(대표) 법인명', '구분', 'HLB글로벌과의 관계', '분개장 이표기(별칭)'])
    ws.append(['에이치엘비오션테크', '기타특수관계자', '특수관계자', '바다중공업,에이치엘비오션테크'])
    return wb


@pytest.mark.parametrize('filename', ['지배구조.xlsx', unicodedata.normalize('NFD', '4)_특관자상호정리_v2.xlsx')])
def test_governance_name_table_upload_is_assigned_to_names(filename):
    buffer = BytesIO()
    _name_table().save(buffer)
    client = TestClient(app)
    sid = client.post('/api/session').json()['sid']
    response = client.post('/api/upload', data={'sid':sid},
                           files={'file':(filename,buffer.getvalue(),'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')})
    assert response.status_code == 200
    data = response.json()
    assert data['suggested_slot'] == 'name_pivot'
    assert data['assign']['name_pivot'] == data['file_id']
    client.delete(f"/api/file/{sid}/{data['file_id']}")


def test_filename_classifier_handles_decomposed_korean(tmp_path):
    filename = unicodedata.normalize('NFD', '4)_특관자상호정리_v2.xlsx')
    result = classify(tmp_path/'upload.xlsx', filename, sheets={})
    assert result.slot == 'name_pivot'


def test_oceantech_and_explicit_old_name_are_matched_without_ai(tmp_path):
    path = tmp_path/'names.xlsx'
    _name_table().save(path)
    result = normalize_names(['바다중공업','에이치엘비오션테크'],path,
                             Settings(openai_api_key=None,openai_model='m'))
    assert result.canonical == {'에이치엘비오션테크'}
    assert result.mapping['바다중공업'] == '에이치엘비오션테크'
    assert result.mapping['에이치엘비오션테크'] == '에이치엘비오션테크'
    assert not result.unmatched
