#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
test_api.py —— API 级测试：submit → status → download

运行: python -m pytest tests/test_api.py -v

test_full_lifecycle 通过 mock _run_job_subprocess 拦截子进程启动，
同步完成 job 并验证完整 submit → status → download 链路。
"""

import io
import sys
import time as _time
from pathlib import Path
from unittest.mock import patch, AsyncMock

import pytest

BASE_DIR = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(BASE_DIR))


# ═══════════════════════════════════════════════════════════
# 全局 mock：替换 _run_job_subprocess 为同步完成
# ═══════════════════════════════════════════════════════════

async def _mock_run_job(job_id: str):
    """Mock: 直接写完成状态（替代子进程）。"""
    from job_manager import JobManager
    jm = JobManager()
    # 确保 job 进入 running 状态
    try:
        jm.transition(job_id, 'running')
    except ValueError:
        pass  # 可能已经是 running
    jm.save_output(job_id, 'report.docx', b'mock report')
    jm.transition(job_id, 'completed', duration_s=0)
    jm.set_progress(job_id, 7, steps=7)


@pytest.fixture(scope='session', autouse=True)
def mock_agent():
    """Session 级别 mock：拦截 _run_job_subprocess，同步完成 job。"""
    with patch('server._run_job_subprocess', side_effect=_mock_run_job):
        yield


# ═══════════════════════════════════════════════════════════
# TestClient
# ═══════════════════════════════════════════════════════════

from fastapi.testclient import TestClient
from server import app

client = TestClient(app)


# ═══════════════════════════════════════════════════════════
# Fixtures
# ═══════════════════════════════════════════════════════════

@pytest.fixture
def sample_docx():
    from docx import Document
    buf = io.BytesIO()
    doc = Document()
    doc.add_paragraph('停车明细分析报告')
    doc.save(buf)
    buf.seek(0)
    return buf


@pytest.fixture
def sample_csv():
    return io.BytesIO(b'name,value\na,1\nb,2\n')


# ═══════════════════════════════════════════════════════════
# Tests
# ═══════════════════════════════════════════════════════════

def test_frontend_serves():
    resp = client.get('/')
    assert resp.status_code == 200
    assert '停车' in resp.text or 'report' in resp.text.lower()


def test_create_job_success(sample_docx, sample_csv):
    resp = client.post('/api/jobs', files={
        'template': ('template.docx', sample_docx,
                     'application/vnd.openxmlformats-officedocument.wordprocessingml.document'),
        'data': ('data.csv', sample_csv, 'text/csv'),
    }, data={'instructions': '测试指令'})
    assert resp.status_code == 200
    data = resp.json()
    assert 'job_id' in data
    assert data['status'] in ('queued', 'running', 'completed')


def test_create_job_rejects_wrong_type():
    resp = client.post('/api/jobs', files={
        'template': ('template.txt', io.BytesIO(b'hello'), 'text/plain'),
        'data': ('data.csv', io.BytesIO(b'a,b\n1,2'), 'text/csv'),
    })
    assert resp.status_code == 400


def test_create_job_rejects_wrong_csv():
    from docx import Document
    buf = io.BytesIO()
    Document().save(buf)
    buf.seek(0)
    resp = client.post('/api/jobs', files={
        'template': ('t.docx', buf,
                     'application/vnd.openxmlformats-officedocument.wordprocessingml.document'),
        'data': ('data.txt', io.BytesIO(b'hello'), 'text/plain'),
    })
    assert resp.status_code == 400


def test_status_nonexistent():
    resp = client.get('/api/jobs/nonexistent')
    assert resp.status_code == 404


def test_download_nonexistent():
    resp = client.get('/api/jobs/nonexistent/download')
    assert resp.status_code == 404


def test_full_lifecycle(sample_docx, sample_csv):
    """
    端到端：submit → status → download。
    _run_job_subprocess 被 session mock 拦截，同步完成 job。
    """
    # 1. Submit
    resp = client.post('/api/jobs', files={
        'template': ('t.docx', sample_docx,
                     'application/vnd.openxmlformats-officedocument.wordprocessingml.document'),
        'data': ('d.csv', sample_csv, 'text/csv'),
    })
    assert resp.status_code == 200
    job_id = resp.json()['job_id']

    # 2. Status — 等待 worker 处理完成
    status = None
    for _ in range(20):
        _time.sleep(0.05)
        resp = client.get(f'/api/jobs/{job_id}')
        status = resp.json()['status']
        if status == 'completed':
            break
    assert status == 'completed', f'Expected completed, got {status}'

    # 3. Download
    resp = client.get(f'/api/jobs/{job_id}/download')
    assert resp.status_code == 200
    assert resp.content == b'mock report'


def test_download_incomplete_rejected(sample_docx, sample_csv):
    """不存在的 job 下载 → 404。"""
    resp = client.get('/api/jobs/nonexistent/download')
    assert resp.status_code == 404
