#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
job_manager.py —— 文件级 Job 持久化与状态机

每个 Job 对应一个目录，状态完全通过文件系统反映：
  jobs/{job_id}/
  ├── state.json          # 状态快照
  ├── input/              # 用户上传的原始文件
  ├── output/             # 生成的 docx
  ├── charts/             # 图表 PNG
  ├── llm_calls.jsonl     # LLM 调用日志
  └── messages.jsonl      # Agent 对话历史（可选）

零依赖，单进程场景无需锁。
"""

import os
import json
import uuid
import shutil
from pathlib import Path
from datetime import datetime, timezone
from typing import Optional


JOBS_ROOT = Path(__file__).resolve().parent / 'jobs'

# ── 状态机 ────────────────────────────────────────────────
VALID_TRANSITIONS = {
    'created':   ['running'],
    'running':   ['completed', 'failed'],
    'completed': [],           # 终态
    'failed':    ['running'],  # 允许重试
}


class JobManager:
    """管理 Job 的创建、状态转换、文件存取。"""

    def __init__(self, root: Path = None):
        self.root = root or JOBS_ROOT
        self.root.mkdir(parents=True, exist_ok=True)

    # ── 创建 ────────────────────────────────────────────
    def create(self, instructions: str = '') -> str:
        """创建新 Job，返回 job_id。"""
        job_id = uuid.uuid4().hex[:12]
        job_dir = self.root / job_id
        (job_dir / 'input').mkdir(parents=True)
        (job_dir / 'output').mkdir(parents=True)
        (job_dir / 'charts').mkdir(parents=True)

        state = {
            'status': 'created',
            'progress': 0,
            'instructions': instructions,
            'created_at': _now(),
            'updated_at': _now(),
            'input_files': [],
            'output_file': None,
            'error': None,
            'steps': 0,
        }
        _write_json(job_dir / 'state.json', state)
        return job_id

    # ── 文件存取 ────────────────────────────────────────
    def save_input(self, job_id: str, filename: str, content: bytes):
        """保存用户上传的原始文件。"""
        dst = self.root / job_id / 'input' / filename
        dst.write_bytes(content)
        self._append_input_file(job_id, filename)

    def get_input_path(self, job_id: str, filename: str) -> Path:
        return self.root / job_id / 'input' / filename

    def get_charts_dir(self, job_id: str) -> Path:
        return self.root / job_id / 'charts'

    def get_output_dir(self, job_id: str) -> Path:
        return self.root / job_id / 'output'

    def save_output(self, job_id: str, filename: str, content: bytes):
        dst = self.root / job_id / 'output' / filename
        dst.write_bytes(content)
        self._set_output_file(job_id, filename)

    def read_output(self, job_id: str) -> Optional[bytes]:
        state = self.get_state(job_id)
        if state and state.get('output_file'):
            path = self.root / job_id / 'output' / state['output_file']
            if path.exists():
                return path.read_bytes()
        return None

    # ── 状态 ────────────────────────────────────────────
    def get_state(self, job_id: str) -> Optional[dict]:
        path = self.root / job_id / 'state.json'
        if not path.exists():
            return None
        return json.loads(path.read_text('utf-8'))

    def transition(self, job_id: str, new_status: str, **extra):
        """状态转移，不合法则抛异常。"""
        state = self.get_state(job_id)
        if not state:
            raise ValueError(f'Job {job_id} not found')
        old = state['status']
        if new_status not in VALID_TRANSITIONS.get(old, []):
            raise ValueError(f'Invalid transition: {old} → {new_status}')
        state['status'] = new_status
        state['updated_at'] = _now()
        state.update(extra)
        _write_json(self.root / job_id / 'state.json', state)

    def set_progress(self, job_id: str, progress: int, steps: int = None):
        state = self.get_state(job_id)
        if state:
            state['progress'] = progress
            if steps is not None:
                state['steps'] = steps
            state['updated_at'] = _now()
            _write_json(self.root / job_id / 'state.json', state)

    def set_error(self, job_id: str, error: str):
        state = self.get_state(job_id)
        if state:
            state['error'] = error
            state['status'] = 'failed'
            state['updated_at'] = _now()
            _write_json(self.root / job_id / 'state.json', state)

    # ── 日志追加 ────────────────────────────────────────
    def append_llm_log(self, job_id: str, entry: dict):
        entry['ts'] = _now()
        _append_jsonl(self.root / job_id / 'llm_calls.jsonl', entry)

    def append_message(self, job_id: str, entry: dict):
        entry['ts'] = _now()
        _append_jsonl(self.root / job_id / 'messages.jsonl', entry)

    # ── 内部 ────────────────────────────────────────────
    def _append_input_file(self, job_id: str, filename: str):
        state = self.get_state(job_id)
        if state is not None:
            files = state.get('input_files', [])
            if filename not in files:
                files.append(filename)
            state['input_files'] = files
            state['updated_at'] = _now()
            _write_json(self.root / job_id / 'state.json', state)

    def _set_output_file(self, job_id: str, filename: str):
        state = self.get_state(job_id)
        if state is not None:
            state['output_file'] = filename
            state['updated_at'] = _now()
            _write_json(self.root / job_id / 'state.json', state)


# ═══════════════════════════════════════════════════════════
# 工具
# ═══════════════════════════════════════════════════════════

def _now() -> str:
    return datetime.now(timezone.utc).strftime('%Y-%m-%dT%H:%M:%SZ')


def _write_json(path: Path, data: dict):
    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_text(json.dumps(data, ensure_ascii=False, indent=2), 'utf-8')


def _append_jsonl(path: Path, entry: dict):
    path.parent.mkdir(parents=True, exist_ok=True)
    with open(path, 'a', encoding='utf-8') as f:
        f.write(json.dumps(entry, ensure_ascii=False) + '\n')
