#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
agent_worker.py —— 单个 Job 的独立执行进程

由 server.py 通过 subprocess 启动，每个 job 一个独立进程。
物理隔离全局变量，彻底消除并发竞态。

用法:
  python agent_worker.py --job-id <job_id> [--timeout <seconds>]

退出码:
  0 — 成功完成
  1 — 执行失败
  2 — 参数错误
"""

import os
import sys
import json
import argparse
import threading
import time as _time
import traceback
from pathlib import Path

# 确保项目根目录在 sys.path
BASE_DIR = Path(__file__).resolve().parent
sys.path.insert(0, str(BASE_DIR))

from job_manager import JobManager
from logging_utils import log_request

# ═══════════════════════════════════════════════════════════
# 超时处理
# ═══════════════════════════════════════════════════════════

_DEFAULT_TIMEOUT = int(os.environ.get('JOB_TIMEOUT', '600'))


def _start_timeout_timer(job_id: str, timeout: int, jobs):
    """启动超时定时器（threading.Timer，跨平台）。"""
    def _on_timeout():
        print(f'[worker] 任务超时 ({timeout}s)，标记失败并退出', flush=True)
        try:
            jobs.set_error(job_id, f'任务执行超时（{timeout}s）')
        except Exception:
            pass
        os._exit(1)

    timer = threading.Timer(timeout, _on_timeout)
    timer.daemon = True
    timer.start()
    return timer


# ═══════════════════════════════════════════════════════════
# 主逻辑
# ═══════════════════════════════════════════════════════════

def run_job(job_id: str, timeout: int = _DEFAULT_TIMEOUT):
    """执行单个 Job 的完整生成流程，返回 exit code。"""
    jobs = JobManager()
    start_time = _time.time()

    # 设置超时定时器
    timer = None
    if timeout > 0:
        timer = _start_timeout_timer(job_id, timeout, jobs)

    try:
        # ── 状态检查 ──
        state = jobs.get_state(job_id)
        if not state:
            print(f'[worker] Job {job_id} not found', flush=True)
            return 2
        if state['status'] not in ('queued', 'created'):
            print(f'[worker] Job {job_id} 状态异常: {state["status"]}', flush=True)
            return 1

        jobs.transition(job_id, 'running')
        jobs.set_progress(job_id, 0, steps=7)

        # ── Step 1: 转换 template.docx → template_full.json ──
        _update_progress(jobs, job_id, 1, 'converting template')
        _convert_template(jobs, job_id)

        # ── Step 2-7: 运行 Agent ──
        _update_progress(jobs, job_id, 2, 'running agent')
        _run_agent(jobs, job_id)

        # ── 兜底还原：Agent 未生成 docx 但有 filled_template.json ──
        output_path = jobs.get_output_dir(job_id) / 'report.docx'
        filled_json = jobs.get_output_dir(job_id) / 'filled_template.json'
        if not output_path.exists() and filled_json.exists():
            print(f'[worker] Agent 未生成 docx，从 filled_template.json 兜底还原', flush=True)
            from json_to_docx import restore_docx
            restore_docx(str(filled_json), str(output_path))
        if not output_path.exists():
            raise RuntimeError('Agent 未生成输出文件')

        # 注册输出文件到 state.json
        jobs._set_output_file(job_id, 'report.docx')

        elapsed = round(_time.time() - start_time, 1)
        jobs.transition(job_id, 'completed', duration_s=elapsed)
        jobs.set_progress(job_id, 7, steps=7)
        log_request('job_completed', job_id,
                    output='report.docx', duration_s=elapsed,
                    steps=_count_llm_calls(jobs, job_id))
        print(f'[worker] Job {job_id} completed in {elapsed}s', flush=True)
        return 0

    except Exception as e:
        elapsed = round(_time.time() - start_time, 1)
        error_msg = f'{type(e).__name__}: {e}'
        try:
            jobs.set_error(job_id, error_msg)
        except Exception:
            pass
        log_request('job_failed', job_id, error=error_msg, duration_s=elapsed)
        traceback.print_exc()
        print(f'[worker] Job {job_id} failed: {error_msg}', flush=True)
        return 1

    finally:
        if timer:
            timer.cancel()  # 取消超时定时器


# ═══════════════════════════════════════════════════════════
# Step 1: 模板转换
# ═══════════════════════════════════════════════════════════

def _convert_template(jobs: JobManager, job_id: str):
    """将用户上传的 template.docx 转为 template_full.json。"""
    from docx2json import docx_to_json

    docx_path = jobs.get_input_path(job_id, 'template.docx')
    json_path = jobs.get_input_path(job_id, 'template_full.json')

    full = docx_to_json(str(docx_path))
    json_path.write_text(json.dumps(full, ensure_ascii=False, indent=2), 'utf-8')

    log_request('job_progress', job_id, step=1,
                detail=f'converted: {len(full["blocks"])} blocks, {len(full["styles"])} styles')


# ═══════════════════════════════════════════════════════════
# Step 2-7: Agent 执行
# ═══════════════════════════════════════════════════════════

def _run_agent(jobs: JobManager, job_id: str):
    """在独立进程空间内运行 core_agent（无并发竞态）。"""
    # ⚠️ 必须在 import core_agent 之前设置环境变量！
    # core_agent 在模块加载时执行 _get_output_paths() 读取这些变量。
    job_root = jobs.root / job_id
    os.environ['AGENT_JOB_DIR'] = str(job_root)
    os.environ['TEMPLATE_JSON'] = 'template_full.json'
    os.environ['TEMPLATE_CSV'] = 'data.csv'
    os.environ['AGENT_LOG_JOB_ID'] = job_id

    import core_agent as ca

    # 加载模板和 CSV 数据到模块全局变量
    ca._load_template()
    ca._load_csv()

    csv_rows = ca._csv_rows
    blocks = len(ca._template_data['blocks'])

    log_request('job_started', job_id,
                csv_rows=len(csv_rows) if csv_rows else 0,
                blocks=blocks)

    # 注入用户指令
    state = jobs.get_state(job_id)
    instructions = (state or {}).get('instructions', '').strip()
    if instructions:
        os.environ['AGENT_USER_INSTRUCTIONS'] = instructions
    else:
        os.environ.pop('AGENT_USER_INSTRUCTIONS', None)

    # 运行 Agent
    agent = ca.TemplateFillerAgent()
    agent.run()


# ═══════════════════════════════════════════════════════════
# 辅助
# ═══════════════════════════════════════════════════════════

def _update_progress(jobs: JobManager, job_id: str, step: int, detail: str = ''):
    jobs.set_progress(job_id, step)
    log_request('job_progress', job_id, step=step, detail=detail)


def _count_llm_calls(jobs: JobManager, job_id: str) -> int:
    log_path = jobs.root / job_id / 'llm_calls.jsonl'
    if log_path.exists():
        return sum(1 for _ in open(log_path, 'r', encoding='utf-8'))
    return 0


# ═══════════════════════════════════════════════════════════
# 入口
# ═══════════════════════════════════════════════════════════

if __name__ == '__main__':
    parser = argparse.ArgumentParser(description='Agent Worker — 单 Job 执行进程')
    parser.add_argument('--job-id', required=True, help='Job ID')
    parser.add_argument('--timeout', type=int, default=_DEFAULT_TIMEOUT,
                        help=f'超时秒数（默认 {_DEFAULT_TIMEOUT}）')
    args = parser.parse_args()

    print(f'[worker] Starting job {args.job_id} (timeout={args.timeout}s)', flush=True)
    exit_code = run_job(args.job_id, args.timeout)
    sys.exit(exit_code)
