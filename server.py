#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
server.py —— 停车报告 Agent 系统 Web 服务（asyncio + subprocess）

架构:
  FastAPI (async) → asyncio.Semaphore → subprocess: agent_worker.py
  每个 Job 在独立子进程中运行，物理隔离，消除并发竞态。

并发:
  通过 asyncio.Semaphore 限制同时运行的子进程数（匹配 CPU 核数）。
  通过 _pending_count 限制排队数量，超限返回 503。

API:
  POST /api/jobs                 上传文件，创建生成任务
  GET  /api/jobs/{job_id}        查询任务状态
  GET  /api/jobs/{job_id}/download  下载生成的 .docx

前端: GET / → templates/index.html
"""

import asyncio
import os
import sys
from pathlib import Path

from fastapi import FastAPI, File, UploadFile, Form, HTTPException
from fastapi.responses import FileResponse, HTMLResponse
from contextlib import asynccontextmanager
import uvicorn

# ── 本地模块 ──────────────────────────────────────────────
BASE_DIR = Path(__file__).resolve().parent
sys.path.insert(0, str(BASE_DIR))

from job_manager import JobManager
from logging_utils import log_request

# ═══════════════════════════════════════════════════════════
# 配置（2 核 / 4 GB 服务器优化）
# ═══════════════════════════════════════════════════════════

MAX_CONCURRENT = int(os.environ.get('MAX_CONCURRENT_JOBS', '2'))
MAX_QUEUE_DEPTH = int(os.environ.get('MAX_QUEUE_DEPTH', '6'))
JOB_TIMEOUT = int(os.environ.get('JOB_TIMEOUT', '600'))
SHUTDOWN_GRACE = int(os.environ.get('SHUTDOWN_GRACE_SECONDS', '30'))

# ═══════════════════════════════════════════════════════════
# 全局状态
# ═══════════════════════════════════════════════════════════

jobs = JobManager()
_semaphore = asyncio.Semaphore(MAX_CONCURRENT)
_pending_count = 0                    # 等待中（含排队 + 已获取信号量但未完成）的 job 数
_running_procs: dict[str, asyncio.subprocess.Process] = {}
_shutting_down = False


# ═══════════════════════════════════════════════════════════
# 生命周期
# ═══════════════════════════════════════════════════════════

@asynccontextmanager
async def lifespan(app: FastAPI):
    """启动诊断；关闭时优雅终止所有子进程。"""
    global _shutting_down

    # ── Startup 诊断 ──
    print(f'[server] Python: {sys.executable}')
    print(f'[server] Event loop: {type(asyncio.get_event_loop()).__name__}')
    _check_subprocess_support()

    yield  # ← App 运行

    # ── Shutdown ──
    print('[server] 正在关闭...')
    _shutting_down = True

    # 终止所有运行中的子进程
    procs = list(_running_procs.items())
    if procs:
        print(f'[server] 终止 {len(procs)} 个运行中的 job...')
        for job_id, proc in procs:
            try:
                proc.terminate()
            except ProcessLookupError:
                continue

        for job_id, proc in procs:
            try:
                await asyncio.wait_for(proc.wait(), timeout=SHUTDOWN_GRACE)
            except asyncio.TimeoutError:
                try:
                    proc.kill()
                except ProcessLookupError:
                    pass
                await proc.wait()
                print(f'[server] Job {job_id} 被强制终止')
        print('[server] 所有子进程已终止')


app = FastAPI(title='Parking Report Agent', lifespan=lifespan)


# ═══════════════════════════════════════════════════════════
# 前端
# ═══════════════════════════════════════════════════════════

@app.get('/', response_class=HTMLResponse)
async def index():
    html_path = BASE_DIR / 'templates' / 'index.html'
    if html_path.exists():
        return html_path.read_text('utf-8')
    return '<h1>Parking Report Agent</h1><p>templates/index.html not found</p>'


# ═══════════════════════════════════════════════════════════
# API
# ═══════════════════════════════════════════════════════════

@app.post('/api/jobs')
async def create_job(
    template: UploadFile = File(...),
    data: UploadFile = File(...),
    instructions: str = Form(''),
):
    """上传 template.docx + data.csv，创建异步生成任务。"""
    global _pending_count

    if _shutting_down:
        raise HTTPException(503, '服务正在关闭')

    # 校验文件类型
    if not template.filename.endswith('.docx'):
        raise HTTPException(400, 'template 必须是 .docx 文件')
    if not data.filename.endswith('.csv'):
        raise HTTPException(400, 'data 必须是 .csv 文件')

    # 排队上限检查
    if _pending_count >= MAX_CONCURRENT + MAX_QUEUE_DEPTH:
        raise HTTPException(503, '系统繁忙，请稍后重试')

    # 创建 job（状态: queued）
    job_id = jobs.create(instructions)
    jobs.save_input(job_id, 'template.docx', await template.read())
    jobs.save_input(job_id, 'data.csv', await data.read())

    log_request('job_created', job_id,
                input_files=['template.docx', 'data.csv'],
                instructions=instructions)

    # 增加排队计数，启动后台任务
    _pending_count += 1
    asyncio.create_task(_run_with_semaphore(job_id))

    log_request('job_queued', job_id)
    return {'job_id': job_id, 'status': 'queued'}


@app.get('/api/jobs/{job_id}')
async def get_job_status(job_id: str):
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
async def download_report(job_id: str):
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
    return FileResponse(
        str(path),
        filename=state['output_file'],
        media_type='application/vnd.openxmlformats-officedocument.wordprocessingml.document',
    )


# ═══════════════════════════════════════════════════════════
# Job 执行
# ═══════════════════════════════════════════════════════════

def _safe_print(msg: str):
    """安全打印，避免 Windows GBK 终端编码错误。"""
    try:
        print(msg, flush=True)
    except UnicodeEncodeError:
        print(msg.encode('utf-8', errors='replace').decode('utf-8', errors='replace'), flush=True)


def _check_subprocess_support():
    """启动时检测子进程机制是否可用。"""
    import subprocess as _sp
    worker = BASE_DIR / 'agent_worker.py'
    if not worker.exists():
        print(f'[server] ⚠️  Worker 脚本不存在: {worker}')
        return

    # 测试同步 subprocess（兜底方案）
    try:
        r = _sp.run([sys.executable, str(worker), '--help'],
                    capture_output=True, timeout=5)
        print(f'[server] ✅ subprocess 可用 (sync)')
    except Exception as e:
        print(f'[server] ⚠️  subprocess 不可用: {e}')

    # 测试 asyncio subprocess
    try:
        loop = asyncio.get_event_loop()
        # 只是检查方法存在，不实际启动
        if hasattr(loop, 'subprocess_exec'):
            print(f'[server] ✅ asyncio subprocess 可用')
        else:
            print(f'[server] ⚠️  asyncio subprocess 不可用，将使用同步 fallback')
    except Exception as e:
        print(f'[server] ⚠️  asyncio 检查失败: {e}')


async def _run_with_semaphore(job_id: str):
    """获取信号量后执行子进程，完成后释放。"""
    global _pending_count
    try:
        async with _semaphore:
            await _run_job_subprocess(job_id)
    except Exception as e:
        _safe_print(f'[server] Job {job_id} 异常: {type(e).__name__}: {e}')
        # 确保 job 被标记为失败（worker 进程可能根本没启动）
        state = jobs.get_state(job_id)
        if state and state['status'] not in ('completed', 'failed'):
            jobs.set_error(job_id, f'执行异常: {type(e).__name__}: {e}')
    finally:
        _pending_count -= 1


async def _run_job_subprocess(job_id: str):
    """启动 agent_worker.py 子进程并等待完成。

    优先使用 asyncio.create_subprocess_exec；
    若抛出 NotImplementedError 则回退到同步 subprocess（线程池）。
    """
    worker_script = BASE_DIR / 'agent_worker.py'

    if not worker_script.exists():
        raise FileNotFoundError(f'Worker 脚本不存在: {worker_script}')

    try:
        await _run_via_async_subprocess(job_id, worker_script)
    except NotImplementedError:
        print(f'[server] asyncio subprocess 不可用，回退到同步模式')
        await _run_via_sync_subprocess(job_id, worker_script)


async def _run_via_async_subprocess(job_id: str, worker_script):
    """asyncio 原生子进程（高性能，首选）。"""
    env = {**os.environ, 'PYTHONIOENCODING': 'utf-8'}
    proc = await asyncio.create_subprocess_exec(
        sys.executable, str(worker_script),
        '--job-id', job_id,
        '--timeout', str(JOB_TIMEOUT),
        stdout=asyncio.subprocess.PIPE,
        stderr=asyncio.subprocess.PIPE,
        env=env,
    )
    _running_procs[job_id] = proc

    try:
        await asyncio.wait_for(proc.wait(), timeout=JOB_TIMEOUT + 10)

        if proc.stdout:
            output = await proc.stdout.read()
            if output:
                _safe_print(output.decode('utf-8', errors='replace').strip())
        if proc.stderr:
            err = await proc.stderr.read()
            if err:
                _safe_print('[worker stderr] ' + err.decode('utf-8', errors='replace').strip())

        if proc.returncode != 0:
            _safe_print(f'[server] Job {job_id} 退出码 {proc.returncode}')

    except asyncio.TimeoutError:
        _safe_print(f'[server] Job {job_id} 超时，强制终止')
        try:
            proc.kill()
        except ProcessLookupError:
            pass
        await proc.wait()
        jobs.set_error(job_id, f'任务执行超时（{JOB_TIMEOUT}s）')
        log_request('job_timeout', job_id)

    finally:
        _running_procs.pop(job_id, None)


async def _run_via_sync_subprocess(job_id: str, worker_script):
    """同步 subprocess（回退方案，在默认线程池中运行）。"""
    import subprocess as _sp

    loop = asyncio.get_event_loop()

    def _run():
        return _sp.run(
            [sys.executable, str(worker_script),
             '--job-id', job_id, '--timeout', str(JOB_TIMEOUT)],
            capture_output=True,
            encoding='utf-8',        # 强制 UTF-8，避免 Windows GBK 乱码
            errors='replace',
            timeout=JOB_TIMEOUT + 10,
            env={**os.environ, 'PYTHONIOENCODING': 'utf-8'},
        )

    try:
        result = await loop.run_in_executor(None, _run)
        if result.stdout:
            _safe_print(result.stdout.strip())
        if result.stderr:
            _safe_print('[worker stderr] ' + result.stderr.strip())
        if result.returncode != 0:
            _safe_print(f'[server] Job {job_id} 退出码 {result.returncode}')
    except _sp.TimeoutExpired:
        _safe_print(f'[server] Job {job_id} 超时')
        jobs.set_error(job_id, f'任务执行超时（{JOB_TIMEOUT}s）')
        log_request('job_timeout', job_id)
    except Exception as e:
        _safe_print(f'[server] Job {job_id} 同步执行异常: {type(e).__name__}: {e}')
        raise


# ═══════════════════════════════════════════════════════════
# 入口
# ═══════════════════════════════════════════════════════════

if __name__ == '__main__':
    print('=' * 60)
    print('  Parking Report Agent — Server')
    print(f'  Max concurrent: {MAX_CONCURRENT} | Queue depth: {MAX_QUEUE_DEPTH}')
    print(f'  Timeout: {JOB_TIMEOUT}s | Grace: {SHUTDOWN_GRACE}s')
    print(f'  http://localhost:8000')
    print('=' * 60)
    uvicorn.run(app, host='0.0.0.0', port=8000)
