#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
server.py —— 停车报告 Agent 系统 Web 服务

API:
  POST /api/jobs                 上传文件，创建生成任务
  GET  /api/jobs/{job_id}        查询任务状态
  GET  /api/jobs/{job_id}/download  下载生成的 .docx

前端: GET / → templates/index.html
"""

import os
import sys
import json
import threading
import time as _time
from pathlib import Path
from datetime import datetime

from fastapi import FastAPI, File, UploadFile, Form, HTTPException
from fastapi.responses import FileResponse, HTMLResponse
from fastapi.staticfiles import StaticFiles
import uvicorn

# ── 本地模块 ──────────────────────────────────────────────
BASE_DIR = Path(__file__).resolve().parent
sys.path.insert(0, str(BASE_DIR))

from job_manager import JobManager
from logging_utils import log_request

# ═══════════════════════════════════════════════════════════
app = FastAPI(title='Parking Report Agent')
jobs = JobManager()

# 并发控制：信号量限制同时运行的 Agent 数量
MAX_CONCURRENT = int(os.environ.get('MAX_CONCURRENT_JOBS', '3'))
_semaphore = threading.BoundedSemaphore(MAX_CONCURRENT + 20)  # 3 并发 + 20 排队


# ═══════════════════════════════════════════════════════════
# 前端
# ═══════════════════════════════════════════════════════════

@app.get('/', response_class=HTMLResponse)
def index():
    html_path = BASE_DIR / 'templates' / 'index.html'
    if html_path.exists():
        return html_path.read_text('utf-8')
    return '<h1>Parking Report Agent</h1><p>templates/index.html not found</p>'


# ═══════════════════════════════════════════════════════════
# API
# ═══════════════════════════════════════════════════════════

@app.post('/api/jobs')
def create_job(
    template: UploadFile = File(...),
    data: UploadFile = File(...),
    instructions: str = Form(''),
):
    """上传 template.docx + data.csv，创建异步生成任务。"""
    # 校验文件类型
    if not template.filename.endswith('.docx'):
        raise HTTPException(400, 'template 必须是 .docx 文件')
    if not data.filename.endswith('.csv'):
        raise HTTPException(400, 'data 必须是 .csv 文件')

    # 创建 job
    job_id = jobs.create(instructions)
    jobs.save_input(job_id, 'template.docx', template.file.read())
    jobs.save_input(job_id, 'data.csv', data.file.read())

    log_request('job_created', job_id,
                input_files=['template.docx', 'data.csv'],
                instructions=instructions)

    # 并发控制（超过 3+20 上限则拒绝）
    acquired = _semaphore.acquire(blocking=False)
    if not acquired:
        raise HTTPException(503, '系统繁忙，请稍后重试')

    def _run_with_release():
        try:
            _run_generation(job_id)
        finally:
            _semaphore.release()

    threading.Thread(target=_run_with_release, daemon=True).start()

    return {'job_id': job_id, 'status': 'created'}


@app.get('/api/jobs/{job_id}')
def get_job_status(job_id: str):
    """查询任务状态。"""
    state = jobs.get_state(job_id)
    if not state:
        raise HTTPException(404, '任务不存在')
    return {
        'job_id': job_id,
        'status': state['status'],
        'progress': state.get('progress', 0),
        'steps': state.get('steps', 0),
        'created_at': state.get('created_at'),
        'updated_at': state.get('updated_at'),
        'error': state.get('error'),
    }


@app.get('/api/jobs/{job_id}/download')
def download_report(job_id: str):
    """下载生成的 .docx 报告。"""
    state = jobs.get_state(job_id)
    if not state:
        raise HTTPException(404, '任务不存在')
    if state['status'] != 'completed':
        raise HTTPException(400, f'任务尚未完成，当前状态: {state["status"]}')
    if not state.get('output_file'):
        raise HTTPException(400, '报告文件不存在')

    path = jobs.get_output_dir(job_id) / state['output_file']
    if not path.exists():
        raise HTTPException(404, '报告文件丢失')

    log_request('job_downloaded', job_id)
    return FileResponse(str(path), filename=state['output_file'],
                        media_type='application/vnd.openxmlformats-officedocument.wordprocessingml.document')


# ═══════════════════════════════════════════════════════════
# Agent 执行（后台线程）
# ═══════════════════════════════════════════════════════════

def _run_generation(job_id: str):
    """在后台线程中运行 Agent 生成报告。"""
    start_time = _time.time()

    try:
        jobs.transition(job_id, 'running')
        jobs.set_progress(job_id, 0, steps=7)

        # ── Step 1: 转换 template.docx → JSON ──
        _update_progress(job_id, 1, 'converting template')
        _convert_template(job_id)

        # ── Step 2-7: 运行 Agent ──
        _update_progress(job_id, 2, 'running agent')
        _run_agent_for_job(job_id)

        # ── 完成：注册输出文件（如果 Agent 未调用 restore_to_docx，兜底还原）──
        output_path = jobs.get_output_dir(job_id) / 'report.docx'
        filled_json = jobs.get_output_dir(job_id) / 'filled_template.json'
        if not output_path.exists() and filled_json.exists():
            print(f'[server] Agent 未生成 docx，从 filled_template.json 兜底还原')
            from json_to_docx import restore_docx
            restore_docx(str(filled_json), str(output_path))
        if not output_path.exists():
            raise RuntimeError('Agent 未生成输出文件')

        # 将输出文件注册到 job state（下载端点依赖 output_file 字段）
        jobs._set_output_file(job_id, 'report.docx')

        elapsed = round(_time.time() - start_time, 1)
        jobs.transition(job_id, 'completed', duration_s=elapsed)
        jobs.set_progress(job_id, 7, steps=7)
        log_request('job_completed', job_id,
                    output='report.docx', duration_s=elapsed,
                    steps=_get_agent_steps(job_id))

    except Exception as e:
        elapsed = round(_time.time() - start_time, 1)
        error_msg = f'{type(e).__name__}: {e}'
        jobs.set_error(job_id, error_msg)
        log_request('job_failed', job_id, error=error_msg, duration_s=elapsed)
        import traceback
        traceback.print_exc()


def _update_progress(job_id: str, step: int, detail: str = ''):
    jobs.set_progress(job_id, step)
    log_request('job_progress', job_id, step=step, detail=detail)


def _convert_template(job_id: str):
    """将用户上传的 template.docx 转为 template_full.json。"""
    from docx2json import docx_to_json

    docx_path = jobs.get_input_path(job_id, 'template.docx')
    json_path = jobs.get_input_path(job_id, 'template_full.json')

    full = docx_to_json(str(docx_path))
    json_path.write_text(json.dumps(full, ensure_ascii=False, indent=2), 'utf-8')

    log_request('job_progress', job_id, step=1,
                detail=f'converted: {len(full["blocks"])} blocks, {len(full["styles"])} styles')


def _run_agent_for_job(job_id: str):
    """设置环境变量后运行 core_agent。"""
    # 传递 job 根目录（一个变量替代所有路径，避免并发竞态覆盖）
    job_root = jobs.root / job_id
    os.environ['AGENT_JOB_DIR'] = str(job_root)
    os.environ['TEMPLATE_JSON'] = 'template_full.json'
    os.environ['TEMPLATE_CSV'] = 'data.csv'

    # 清除缓存的模块以获取新路径
    for key in list(sys.modules.keys()):
        if 'core_agent' in key:
            del sys.modules[key]

    # 动态导入并打补丁 — 注入 logger
    import core_agent as ca

    # 加载数据
    ca._load_template()
    ca._load_csv()

    csv_rows = ca._csv_rows
    blocks = len(ca._template_data['blocks'])

    log_request('job_started', job_id, csv_rows=len(csv_rows) if csv_rows else 0, blocks=blocks)

    # 打补丁：拦截 LLM 调用以记录日志
    _patch_agent_for_logging(ca, job_id)

    # 注入用户指令（通过环境变量传入，避免并发修改全局变量）
    state = jobs.get_state(job_id)
    instructions = (state or {}).get('instructions', '').strip()
    if instructions:
        os.environ['AGENT_USER_INSTRUCTIONS'] = instructions
    else:
        os.environ.pop('AGENT_USER_INSTRUCTIONS', None)

    # 运行 Agent
    agent = ca.TemplateFillerAgent()
    agent.run()


def _patch_agent_for_logging(ca_module, job_id: str):
    """通过环境变量传递 job_id，Agent 内部自行记录 LLM 日志。"""
    os.environ['AGENT_LOG_JOB_ID'] = job_id


def _get_agent_steps(job_id: str) -> int:
    """从 LLM 日志推断调用次数。"""
    log_path = jobs.root / job_id / 'llm_calls.jsonl'
    if log_path.exists():
        return sum(1 for _ in open(log_path, 'r', encoding='utf-8'))
    return 0

    AgentClass.setup = _setup_with_logging


def _get_agent_steps(job_id: str) -> int:
    """从 LLM 日志推断 agent 执行了多少步。"""
    log_path = jobs.root / job_id / 'llm_calls.jsonl'
    if log_path.exists():
        return sum(1 for _ in open(log_path, 'r', encoding='utf-8'))
    return 0


# ═══════════════════════════════════════════════════════════
# 入口
# ═══════════════════════════════════════════════════════════

if __name__ == '__main__':
    print('=' * 60)
    print('  Parking Report Agent — Server')
    print('=' * 60)
    print(f'  http://localhost:8000')
    print('=' * 60)
    uvicorn.run(app, host='0.0.0.0', port=8000)
