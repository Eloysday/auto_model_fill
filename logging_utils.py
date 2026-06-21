#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
logging_utils.py —— 结构化 JSON 日志

两种日志：
  1. Request log (requests.jsonl) — 用户请求生命周期，按行追加到全局文件
  2. LLM call log (per-job llm_calls.jsonl) — 由 job_manager 管理
"""

import json
import sys
from pathlib import Path
from datetime import datetime, timezone
from typing import Optional

# 全局请求日志路径（默认项目根目录）
REQUEST_LOG_PATH = Path(__file__).resolve().parent / 'requests.jsonl'


def _now() -> str:
    return datetime.now(timezone.utc).strftime('%Y-%m-%dT%H:%M:%SZ')


def log_request(event: str, job_id: str, **extra):
    """
    记录一次请求生命周期事件。

    使用方式：
      log_request('job_created',    job_id, input_files=['t.docx','d.csv'])
      log_request('job_started',    job_id, csv_rows=3674, blocks=106)
      log_request('job_completed',  job_id, output='report.docx', duration_s=133)
      log_request('job_downloaded', job_id)
      log_request('job_failed',     job_id, error='API timeout')
    """
    entry = {
        'ts': _now(),
        'event': event,
        'job_id': job_id,
        **extra,
    }
    _append_jsonl(REQUEST_LOG_PATH, entry)
    # 同时输出到 stdout（docker logs 可见）
    print(json.dumps(entry, ensure_ascii=False), flush=True)


def _append_jsonl(path: Path, entry: dict):
    path.parent.mkdir(parents=True, exist_ok=True)
    with open(path, 'a', encoding='utf-8') as f:
        f.write(json.dumps(entry, ensure_ascii=False) + '\n')
