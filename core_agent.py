#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
core_agent.py —— 停车明细分析报告 智能填充 Agent
=================================================
功能：读取 template_full.json 模板 + data.csv 数据，
      让 LLM 驱动工具调用，自动计算指标、生成图表、填充模板，
      最终输出完整 JSON，再还原为 .docx 文档。

架构：
  - 工具层：Python 函数，LLM 可通过 function-calling 调用
  - 引擎层：OpenAI 兼容 API（function-calling loop）
  - 输出层：填充后的 JSON → json_to_docx 还原 .docx

用法：
  python core_agent.py
"""

import os
import sys
import json
import csv
import math
import hashlib
import time as _time
from datetime import datetime, timedelta
from collections import Counter, defaultdict
from pathlib import Path
import traceback

# ── 可选依赖 ──────────────────────────────────────────────
try:
    from openai import OpenAI
    HAS_OPENAI = True
except ImportError:
    HAS_OPENAI = False

try:
    import matplotlib
    matplotlib.use('Agg')
    import matplotlib.pyplot as plt
    import matplotlib.font_manager as fm
    HAS_MPL = True
except ImportError:
    HAS_MPL = False

try:
    from dotenv import load_dotenv
    load_dotenv()
except ImportError:
    pass

# ═══════════════════════════════════════════════════════════
# 0. 配置 & 常量（零硬编码路径 — 自动推断 + 环境变量覆盖）
# ═══════════════════════════════════════════════════════════

def _resolve_base_dir() -> Path:
    """推断项目根目录。优先级：环境变量 > 脚本所在目录"""
    env_dir = os.environ.get('TEMPLATE_PROJECT_DIR', '').strip()
    if env_dir:
        return Path(env_dir)
    return Path(__file__).resolve().parent

BASE_DIR = _resolve_base_dir()

def _get_output_paths():
    """根据 AGENT_JOB_DIR 环境变量动态解析路径（避免并发竞态）。"""
    job_dir = os.environ.get('AGENT_JOB_DIR', '')
    if job_dir:
        jd = Path(job_dir)
        return {
            'template': jd / 'input' / os.environ.get('TEMPLATE_JSON', 'template_full.json'),
            'csv': jd / 'input' / os.environ.get('TEMPLATE_CSV', 'data.csv'),
            'output_json': jd / 'output' / 'filled_template.json',
            'output_docx': jd / 'output' / 'report.docx',
            'charts': jd / 'charts',
        }
    # fallback: 旧方式
    return {
        'template': BASE_DIR / os.environ.get('TEMPLATE_JSON', 'template_full.json'),
        'csv': BASE_DIR / os.environ.get('TEMPLATE_CSV', 'data.csv'),
        'output_json': BASE_DIR / os.environ.get('OUTPUT_JSON', 'filled_template.json'),
        'output_docx': BASE_DIR / os.environ.get('OUTPUT_DOCX', 'filled_report.docx'),
        'charts': BASE_DIR / os.environ.get('CHARTS_DIR', 'charts'),
    }

_paths = _get_output_paths()
TEMPLATE_PATH = _paths['template']
CSV_PATH      = _paths['csv']
OUTPUT_JSON   = _paths['output_json']
OUTPUT_DOCX   = _paths['output_docx']
CHARTS_DIR    = _paths['charts']

# 确保输出目录存在
CHARTS_DIR.mkdir(parents=True, exist_ok=True)

# ── 占位符识别规则（可配置，支持环境变量覆盖） ──
PLACEHOLDER_COLOR = os.environ.get('PLACEHOLDER_COLOR', 'C00000')  # 占位符特征颜色（红）
FILLED_COLOR      = os.environ.get('FILLED_COLOR', '000000')       # 填充后黑色
# 占位符样式要求：italic / bold / underline 任意一个为 True 即匹配（None=不要求）
PLACEHOLDER_REQUIRE_ITALIC = os.environ.get('PLACEHOLDER_REQUIRE_ITALIC', '1') == '1'
PLACEHOLDER_REQUIRE_BOLD   = os.environ.get('PLACEHOLDER_REQUIRE_BOLD', '0') == '1'
# 内容模式：匹配这些 bracket 模式的 run 也被视为占位符（与颜色检测互补）
PLACEHOLDER_CONTENT_PATTERNS = [
    '【', '】',           # 中文方括号
    '[', ']',             # 英文方括号（含 [ ___ ] [笔数] 等）
    '{{', '}}',           # 双花括号
    '<<', '>>',           # 尖括号
    '___', '……',          # 下划线 / 省略号占位
]

# ── 中文字体（matplotlib） ──
CN_FONT_CANDIDATES = [
    'Noto Sans CJK SC', 'Noto Sans SC', 'WenQuanYi Micro Hei',
    'SimHei', 'Microsoft YaHei', 'PingFang SC',
    'Arial Unicode MS', 'DejaVu Sans'
]

# ── 当前状态 ──
_template_data: dict = None
_csv_rows: list = None
_csv_headers: list = None
_chart_files: list = []         # 已生成的图表文件路径
_stats_cache: dict = {}         # 统计缓存


# ═══════════════════════════════════════════════════════════
# 0.5. ColumnMapper — CSV 列名自动检测（零硬编码）
# ═══════════════════════════════════════════════════════════

class ColumnMapper:
    """根据 CSV headers 自动检测语义列的映射关系。

    通过关键词模糊匹配将实际列名映射到语义角色。
    LLM 也可从 load_csv_summary 获取列名后手动传入覆盖。
    """

    # 语义角色 → 关键词列表（按优先级排列，命中即止）
    SEMANTIC_PATTERNS = {
        'amount_ying':            ['应收金额', 'ying', 'receivable'],
        'amount_shi':             ['实收金额', 'shi shou', 'received'],
        'amount_free':            ['免费金额', 'free', 'mian fei'],
        'amount_dikou':           ['抵扣金额', 'dikou jin'],
        'amount_actual_dikou':    ['实际抵扣额', 'actual dikou'],
        'dikou_hours':            ['抵扣时长', 'dikou shi chang', 'hours'],
        'recharge_card':          ['充值卡扣费', '充值卡'],
        'payment_method':         ['支付方式', 'method', 'fang shi'],
        'payment_channel':        ['支付渠道', 'channel', 'qu dao'],
        'charge_time':            ['收费时间', 'charge', 'shou fei shi jian'],
        'entry_time':             ['进车时间', 'entry', 'jin che'],
    }

    def __init__(self, headers: list = None):
        self.mapping: dict[str, str] = {}
        if headers:
            self.detect(headers)

    def detect(self, headers: list):
        """从 CSV 表头自动检测列名映射。"""
        for role, patterns in self.SEMANTIC_PATTERNS.items():
            for h in headers:
                h_lower = h.lower().replace(' ', '')
                for p in patterns:
                    if p.lower().replace(' ', '') in h_lower:
                        self.mapping[role] = h
                        break
                if role in self.mapping:
                    break

    def get(self, role: str, default: str = None) -> str:
        """获取语义列名，可带默认值。"""
        return self.mapping.get(role, default)

    def set(self, role: str, column_name: str):
        """手动覆盖某个语义列的映射。"""
        self.mapping[role] = column_name

    @property
    def col_ying(self):           return self.mapping.get('amount_ying', '应收金额')
    @property
    def col_shi(self):            return self.mapping.get('amount_shi', '实收金额(元)')
    @property
    def col_free(self):           return self.mapping.get('amount_free', '免费金额(元)')
    @property
    def col_dikou(self):          return self.mapping.get('amount_actual_dikou', '实际抵扣额(元)')
    @property
    def col_dikou_hours(self):    return self.mapping.get('dikou_hours', '抵扣时长(小时)')
    @property
    def col_pay_method(self):     return self.mapping.get('payment_method', '支付方式')
    @property
    def col_pay_channel(self):    return self.mapping.get('payment_channel', '支付渠道')
    @property
    def col_charge_time(self):    return self.mapping.get('charge_time', '收费时间')
    @property
    def col_entry_time(self):     return self.mapping.get('entry_time', '进车时间')

    @property
    def all_numeric_columns(self) -> list:
        """返回可能为数值类型列的映射列表。"""
        return [self.col_ying, self.col_shi, self.col_free,
                self.col_dikou, self.col_dikou_hours]

    @property
    def all_time_columns(self) -> list:
        """返回时间列的映射列表。"""
        return [self.col_charge_time, self.col_entry_time]


# 全局列映射器（_load_csv 时自动初始化）
_column_mapper: ColumnMapper = ColumnMapper()


# ═══════════════════════════════════════════════════════════
# 1. 内部工具函数
# ═══════════════════════════════════════════════════════════

def _load_template(path: str = None) -> dict:
    """加载 JSON 模板到全局缓存"""
    global _template_data
    p = Path(path) if path else TEMPLATE_PATH
    with open(p, 'r', encoding='utf-8') as f:
        _template_data = json.load(f)
    return _template_data


def _load_csv(path: str = None) -> tuple:
    """加载 CSV 数据到全局缓存，自动检测列名映射，返回 (headers, rows)"""
    global _csv_rows, _csv_headers, _column_mapper
    p = Path(path) if path else CSV_PATH
    with open(p, 'r', encoding='utf-8-sig') as f:
        reader = csv.DictReader(f)
        _csv_headers = list(reader.fieldnames)
        _csv_rows = list(reader)
    # 自动检测列名映射
    _column_mapper = ColumnMapper(_csv_headers)
    return _csv_headers, _csv_rows


def _ensure_data_loaded():
    if _template_data is None:
        _load_template()
    if _csv_rows is None:
        _load_csv()


def _detect_time_format(rows: list, col: str) -> str:
    """从 CSV 样本自动检测时间格式。

    支持常见格式：%Y-%m-%d %H:%M:%S / %Y/%m/%d %H:%M:%S /
    %Y-%m-%d / %Y/%m/%d / %d-%m-%Y %H:%M:%S / %m/%d/%Y %H:%M:%S
    """
    candidates = [
        '%Y-%m-%d %H:%M:%S',
        '%Y/%m/%d %H:%M:%S',
        '%Y-%m-%d %H:%M',
        '%Y/%m/%d %H:%M',
        '%Y-%m-%d',
        '%Y/%m/%d',
        '%d-%m-%Y %H:%M:%S',
        '%d/%m/%Y %H:%M:%S',
        '%m/%d/%Y %H:%M:%S',
        '%m-%d-%Y %H:%M:%S',
        '%Y%m%d%H%M%S',
    ]
    sample_size = min(20, len(rows))
    for fmt in candidates:
        ok = 0
        for r in rows[:sample_size]:
            try:
                datetime.strptime(r.get(col, '').strip(), fmt)
                ok += 1
            except (ValueError, KeyError):
                pass
        if ok >= sample_size * 0.8:  # 80% 命中即认为正确
            return fmt
    return '%Y-%m-%d %H:%M:%S'  # 回退默认值


def _setup_matplotlib_font():
    """配置 matplotlib 中文字体"""
    if not HAS_MPL:
        return False
    for name in CN_FONT_CANDIDATES:
        for f in fm.fontManager.ttflist:
            if name.lower() in f.name.lower():
                plt.rcParams['font.family'] = f.name
                plt.rcParams['axes.unicode_minus'] = False
                return True
    # Fallback: 直接用 sans-serif
    plt.rcParams['font.sans-serif'] = CN_FONT_CANDIDATES + ['sans-serif']
    plt.rcParams['axes.unicode_minus'] = False
    return True


# ═══════════════════════════════════════════════════════════
# 2. 通用辅助
# ═══════════════════════════════════════════════════════════

def _reindex_blocks():
    """重新为所有 block 分配连续的 index（按数组顺序 0..N-1）"""
    _ensure_data_loaded()
    for i, b in enumerate(_template_data['blocks']):
        b['index'] = i


def _find_block_by_index(idx: int) -> dict | None:
    """在 blocks 数组中查找指定 index 的 block"""
    for b in _template_data['blocks']:
        if b.get('index') == idx:
            return b
    return None


def _find_block_pos(idx: int) -> int:
    """返回指定 index 的 block 在数组中的位置，-1 表示未找到"""
    for i, b in enumerate(_template_data['blocks']):
        if b.get('index') == idx:
            return i
    return -1


# ═══════════════════════════════════════════════════════════
# 2.5. 占位符检测（多层：格式 + 内容模式）
# ═══════════════════════════════════════════════════════════

def _is_placeholder_run(run: dict) -> bool:
    """判断单个 run 是否为占位符（两层检测，任一命中即为占位）。

    1. 格式检测：color 匹配 PLACEHOLDER_COLOR 且满足样式要求
    2. 内容检测：文本含 bracket 模式（如 [ ___ ]、【……】、{{name}}）
    """
    text = run.get('text', '')
    fmt = run.get('fmt', {})

    # ── 1. 格式检测 ──
    format_match = False
    if fmt.get('color') == PLACEHOLDER_COLOR:
        if not PLACEHOLDER_REQUIRE_ITALIC and not PLACEHOLDER_REQUIRE_BOLD:
            format_match = True
        elif PLACEHOLDER_REQUIRE_ITALIC and fmt.get('italic'):
            format_match = True
        elif PLACEHOLDER_REQUIRE_BOLD and fmt.get('bold'):
            format_match = True
        elif PLACEHOLDER_REQUIRE_ITALIC and PLACEHOLDER_REQUIRE_BOLD:
            if fmt.get('italic') or fmt.get('bold'):
                format_match = True
    if format_match:
        return True

    # ── 2. 内容模式检测（fallback，与格式检测独立） ──
    if text.strip():
        for pat in PLACEHOLDER_CONTENT_PATTERNS:
            if pat in text:
                stripped = text.strip()
                if len(stripped) <= 1:
                    continue
                return True

    return False


def _is_format_placeholder(run: dict) -> bool:
    """仅通过格式检测判断（不含内容 fallback）。用于区分「指令块」和「填充块」。"""
    fmt = run.get('fmt', {})
    if fmt.get('color') != PLACEHOLDER_COLOR:
        return False
    if not PLACEHOLDER_REQUIRE_ITALIC and not PLACEHOLDER_REQUIRE_BOLD:
        return True
    if PLACEHOLDER_REQUIRE_ITALIC and fmt.get('italic'):
        return True
    if PLACEHOLDER_REQUIRE_BOLD and fmt.get('bold'):
        return True
    if PLACEHOLDER_REQUIRE_ITALIC and PLACEHOLDER_REQUIRE_BOLD:
        if fmt.get('italic') or fmt.get('bold'):
            return True
    return False


def _is_instruction_block(block: dict) -> bool:
    """判断整个 block 是否为纯指令块（应删除，不出现在最终报告中）。

    关键区分：
      - 指令块（应删除）：所有 run 均通过 **格式检测**（有特殊颜色/样式）
      - 填充块（应填充）：run 含 bracket 模式但无特殊格式（如 "数据周期：【起始日期】"）
    只有格式占位块才是给 LLM 看的元指令，内容 bracket 块是模板的填空标记。
    """
    if block.get('type') != 'paragraph':
        return False
    runs = block.get('runs', [])
    if not runs:
        return False
    format_placeholder_count = sum(1 for r in runs if _is_format_placeholder(r))
    return format_placeholder_count == len(runs)


# ═══════════════════════════════════════════════════════════
# 3. 工具实现（供 LLM 调用的函数）
# ═══════════════════════════════════════════════════════════

def tool_get_template_structure() -> str:
    """获取模板的完整结构摘要。

    对每个 block 提供原始事实：
      - index / type / style
      - 文本预览（前 120 字）
      - fmt 摘要：has_color=C00000（红色） / has_italic / has_bracket（含【】[]等）
      - 图片标记

    LLM 需根据文本语义自行判断哪些是占位符、哪些是指令块、哪些是固定内容。
    代码不提供「🔴占位 / ✅已填充」等预分类标签。
    """
    _ensure_data_loaded()
    lines = []
    for b in _template_data['blocks']:
        idx = b.get('index')
        btype = b.get('type')
        style = b.get('style') or '—'
        texts = []
        has_red = False
        has_italic = False
        has_bracket = False
        if btype == 'paragraph':
            for r in b.get('runs', []):
                t = r.get('text', '')
                texts.append(t)
                fmt = r.get('fmt', {})
                if fmt.get('color') == 'C00000':
                    has_red = True
                if fmt.get('italic'):
                    has_italic = True
                if t and any(p in t for p in ('【', '】', '[', ']')):
                    has_bracket = True
        elif btype == 'table':
            for row in b.get('rows', []):
                for cell in row.get('cells', []):
                    for p in cell.get('paragraphs', []):
                        for r in p.get('runs', []):
                            t = r.get('text', '')
                            if t:
                                texts.append(t)
                            if r.get('fmt', {}).get('color') == 'C00000':
                                has_red = True

        text_preview = ''.join(texts)[:120]

        # ── 格式摘要（纯事实，不预判） ──
        fmt_parts = []
        if has_red:
            fmt_parts.append('RED')
        if has_italic:
            fmt_parts.append('italic')
        if has_bracket:
            fmt_parts.append('bracket')
        if b.get('image'):
            fmt_parts.append('IMG')
        fmt_str = ','.join(fmt_parts) if fmt_parts else '—'

        lines.append(f"[{idx}] {btype:11s} style={style:5s} fmt={fmt_str:20s} | {text_preview}")
    return '\n'.join(lines)


def tool_get_placeholder_blocks() -> str:
    """列出模板中所有具有非常规格式的 run（红色文字 或 含 bracket 模式）。

    返回每个 run 的：
      - block_index / run_index / 行号列号（表格时）
      - 当前文本
      - 格式标记：color / italic / bracket

    LLM 需根据文本语义自行决定每个 run 的处理方式（填充 / 删除 / 保留）。
    代码不预判——只提供原始数据。
    """
    _ensure_data_loaded()
    results = []
    for b in _template_data['blocks']:
        idx = b.get('index')
        btype = b.get('type')
        if btype == 'paragraph':
            for ri, r in enumerate(b.get('runs', [])):
                fmt = r.get('fmt', {})
                if fmt.get('color') == 'C00000' or fmt.get('italic') or \
                   any(p in r.get('text', '') for p in ('【', '】', '[', ']')):
                    results.append({
                        'block_index': idx,
                        'run_index': ri,
                        'current_text': r.get('text', ''),
                        'color': fmt.get('color'),
                        'italic': fmt.get('italic', False),
                    })
        elif btype == 'table':
            for ri_row, row in enumerate(b.get('rows', [])):
                for ci, cell in enumerate(row.get('cells', [])):
                    for pi, p in enumerate(cell.get('paragraphs', [])):
                        for ri_run, r in enumerate(p.get('runs', [])):
                            fmt = r.get('fmt', {})
                            if fmt.get('color') == 'C00000' or \
                               any(p in r.get('text', '') for p in ('【', '】', '[', ']')):
                                results.append({
                                    'block_index': idx,
                                    'row_index': ri_row,
                                    'col_index': ci,
                                    'run_index': ri_run,
                                    'current_text': r.get('text', ''),
                                    'color': fmt.get('color'),
                                })
    return json.dumps(results, ensure_ascii=False, indent=2)


def tool_get_placeholder_rules() -> str:
    """返回当前占位符检测规则（供 LLM 了解检测逻辑，必要时建议调整环境变量）。"""
    return json.dumps({
        'format_detection': {
            'color': PLACEHOLDER_COLOR,
            'require_italic': PLACEHOLDER_REQUIRE_ITALIC,
            'require_bold': PLACEHOLDER_REQUIRE_BOLD,
            'filled_color': FILLED_COLOR,
        },
        'content_detection': {
            'patterns': PLACEHOLDER_CONTENT_PATTERNS,
            'description': 'run 文本含任一 pattern 即视为占位符（排除单字符纯标点）',
        },
        'instruction_block_rule': '整段所有 run 均为占位 run → 标记为 🔴指令，应删除',
        'env_overrides': 'PLACEHOLDER_COLOR / PLACEHOLDER_REQUIRE_ITALIC / PLACEHOLDER_REQUIRE_BOLD / FILLED_COLOR',
    }, ensure_ascii=False, indent=2)


def tool_load_csv_summary() -> str:
    """读取 data.csv 并返回概览：总行数、列名、每列的数据类型/示例值/统计概要"""
    _load_csv()
    headers = _csv_headers
    rows = _csv_rows
    n = len(rows)
    info = {'total_rows': n, 'columns': {}}
    for h in headers:
        vals = [r[h] for r in rows if r.get(h, '').strip() != '']
        sample = vals[:5]
        # 推断类型
        numeric_count = 0
        for v in vals[:100]:
            try:
                float(v)
                numeric_count += 1
            except ValueError:
                pass
        is_numeric = numeric_count > len(vals[:100]) * 0.8
        info['columns'][h] = {
            'type': 'numeric' if is_numeric else 'text',
            'sample_values': sample,
            'non_empty_count': len(vals),
        }
    return json.dumps(info, ensure_ascii=False, indent=2)


def tool_compute_statistics(column: str, operation: str) -> str:
    """
    对 data.csv 的指定列执行统计运算。
    
    参数:
      column: 列名（如 "应收金额"、"实收金额(元)"、"实际抵扣额(元)"）
      operation: 运算名 —— sum / avg / count / max / min / median / mode / distinct
    """
    _ensure_data_loaded()
    rows = _csv_rows
    vals = []
    texts = []
    for r in rows:
        v = r.get(column, '').strip()
        if v:
            try:
                vals.append(float(v))
                texts.append(v)
            except ValueError:
                texts.append(v)
    
    try:
        if operation == 'sum':
            result = sum(vals)
        elif operation == 'avg':
            result = sum(vals) / len(vals) if vals else 0
        elif operation == 'count':
            result = len(vals)
        elif operation == 'max':
            result = max(vals) if vals else 0
        elif operation == 'min':
            result = min(vals) if vals else 0
        elif operation == 'median':
            sv = sorted(vals)
            mid = len(sv) // 2
            result = sv[mid] if len(sv) % 2 else (sv[mid - 1] + sv[mid]) / 2
        elif operation == 'mode':
            counts = Counter(texts)
            top = counts.most_common(1)
            result = top[0][0] if top else None
        elif operation == 'distinct':
            result = len(set(texts))
        else:
            return json.dumps({'error': f'不支持的运算：{operation}。支持：sum, avg, count, max, min, median, mode, distinct'})
        
        result = round(result, 2) if isinstance(result, float) else result
        return json.dumps({
            'column': column,
            'operation': operation,
            'result': result,
            'raw_count': len(vals),
            'total_rows': len(rows),
        }, ensure_ascii=False)
    except Exception as e:
        return json.dumps({'error': str(e), 'traceback': traceback.format_exc()})


def tool_compute_all_kpis() -> str:
    """
    一键计算模板关键指标表格所需的全部 6 项 KPI：
      1. 总交易笔数
      2. 应收总金额
      3. 实收总金额
      4. 实际抵扣总额
      5. 实收率
      6. 主要支付方式（按笔数 TOP1）
    返回完整结果 JSON。
    """
    _ensure_data_loaded()
    rows = _csv_rows
    cm = _column_mapper
    n = len(rows)

    col_ying = cm.col_ying
    col_shi = cm.col_shi
    col_dikou = cm.col_dikou
    col_free = cm.col_free
    col_dikou_hrs = cm.col_dikou_hours
    col_pay = cm.col_pay_method

    total_ying = sum(float(r[col_ying]) for r in rows if r.get(col_ying, '').strip())
    total_shi = sum(float(r[col_shi]) for r in rows if r.get(col_shi, '').strip())
    total_dikou = sum(float(r[col_dikou]) for r in rows if r.get(col_dikou, '').strip())
    shishou_rate = (total_shi / total_ying * 100) if total_ying > 0 else 0

    # 主要支付方式
    pay_counter = Counter(r.get(col_pay, '') for r in rows if r.get(col_pay, '').strip())
    top_pay = pay_counter.most_common(3)

    # 免费金额合计
    total_free = sum(float(r.get(col_free, '0') or '0') for r in rows)
    # 抵扣时长合计
    total_dikou_hours = sum(float(r.get(col_dikou_hrs, '0') or '0') for r in rows)
    
    result = {
        '总交易笔数': n,
        '应收总金额(元)': round(total_ying, 2),
        '实收总金额(元)': round(total_shi, 2),
        '实际抵扣总额(元)': round(total_dikou, 2),
        '实收率(%)': round(shishou_rate, 1),
        '主要支付方式_TOP1': {'方式': top_pay[0][0], '笔数': top_pay[0][1]} if top_pay else None,
        '主要支付方式_TOP3': [{'方式': p, '笔数': c} for p, c in top_pay],
        '免费总金额(元)': round(total_free, 2),
        '总抵扣时长(小时)': round(total_dikou_hours, 2),
    }
    return json.dumps(result, ensure_ascii=False, indent=2)


def tool_analyze_payment_distribution() -> str:
    """
    分析支付方式与支付渠道的完整分布情况。
    返回：各支付方式的笔数/金额统计、各渠道的笔数/金额统计、交叉分布。
    """
    _ensure_data_loaded()
    rows = _csv_rows
    cm = _column_mapper
    col_pay = cm.col_pay_method
    col_chan = cm.col_pay_channel
    col_ying = cm.col_ying
    col_shi = cm.col_shi

    # 支付方式统计
    by_method = defaultdict(lambda: {'笔数': 0, '应收金额合计': 0.0, '实收金额合计': 0.0})
    for r in rows:
        m = r.get(col_pay, '').strip()
        if m:
            by_method[m]['笔数'] += 1
            by_method[m]['应收金额合计'] += float(r.get(col_ying, 0) or 0)
            by_method[m]['实收金额合计'] += float(r.get(col_shi, 0) or 0)

    # 支付渠道统计
    by_channel = defaultdict(lambda: {'笔数': 0, '应收金额合计': 0.0, '实收金额合计': 0.0})
    for r in rows:
        c = r.get(col_chan, '').strip()
        if c:
            by_channel[c]['笔数'] += 1
            by_channel[c]['应收金额合计'] += float(r.get(col_ying, 0) or 0)
            by_channel[c]['实收金额合计'] += float(r.get(col_shi, 0) or 0)

    # 交叉分布
    cross = defaultdict(int)
    for r in rows:
        key = f"{r.get(col_pay,'')} / {r.get(col_chan,'')}"
        cross[key] += 1

    # 排序
    sorted_method = sorted(by_method.items(), key=lambda x: x[1]['笔数'], reverse=True)
    sorted_channel = sorted(by_channel.items(), key=lambda x: x[1]['笔数'], reverse=True)
    sorted_cross = sorted(cross.items(), key=lambda x: x[1], reverse=True)

    return json.dumps({
        '支付方式分布': [
            {'方式': k, **{kk: round(vv, 2) if isinstance(vv, float) else vv for kk, vv in v.items()}}
            for k, v in sorted_method
        ],
        '支付渠道分布': [
            {'渠道': k, **{kk: round(vv, 2) if isinstance(vv, float) else vv for kk, vv in v.items()}}
            for k, v in sorted_channel
        ],
        '交叉分布_TOP10': [
            {'组合': k, '笔数': v} for k, v in sorted_cross[:10]
        ],
    }, ensure_ascii=False, indent=2)


def tool_analyze_parking_duration() -> str:
    """
    分析停车时长：计算收费时间与进车时间的差值，
    返回平均时长、分布（每 2 小时的桶）、超过 12 小时的长时停车明细。
    """
    _ensure_data_loaded()
    rows = _csv_rows
    cm = _column_mapper
    col_charge = cm.col_charge_time
    col_entry = cm.col_entry_time
    col_ying = cm.col_ying
    time_fmt = _detect_time_format(rows, col_charge)

    durations_hours = []
    long_stays = []  # > 12h
    entry_hours = []
    charge_hours = []

    for r in rows:
        try:
            t_charge = datetime.strptime(r[col_charge].strip(), time_fmt)
            t_entry = datetime.strptime(r[col_entry].strip(), time_fmt)
            dh = (t_charge - t_entry).total_seconds() / 3600.0
            durations_hours.append(dh)
            entry_hours.append(t_entry.hour + t_entry.minute / 60.0)
            charge_hours.append(t_charge.hour + t_charge.minute / 60.0)
            if dh > 12:
                long_stays.append({
                    col_entry: r[col_entry],
                    col_charge: r[col_charge],
                    '时长(小时)': round(dh, 1),
                    col_ying: r[col_ying],
                })
        except (ValueError, KeyError):
            continue

    # 时长桶 (0-12h, 每 2 小时)
    buckets = defaultdict(int)
    for dh in durations_hours:
        bucket_idx = min(int(dh // 2), 6)  # 0-1, 2-3, ..., 12+
        bucket_label = f"{bucket_idx*2}-{bucket_idx*2+1}h" if bucket_idx < 6 else "12h+"
        buckets[bucket_label] += 1

    avg_duration = sum(durations_hours) / len(durations_hours) if durations_hours else 0
    median_duration = sorted(durations_hours)[len(durations_hours)//2] if durations_hours else 0
    max_duration = max(durations_hours) if durations_hours else 0
    min_duration = min(durations_hours) if durations_hours else 0

    # 入场/收费时间分布
    entry_dist = defaultdict(int)
    charge_dist = defaultdict(int)
    for h in entry_hours:
        if 6 <= h < 20:
            entry_dist[int(h)] += 1
    for h in charge_hours:
        if 6 <= h < 20:
            charge_dist[int(h)] += 1

    return json.dumps({
        '有效记录数': len(durations_hours),
        '平均停车时长(小时)': round(avg_duration, 2),
        '中位停车时长(小时)': round(median_duration, 2),
        '最长停车时长(小时)': round(max_duration, 2),
        '最短停车时长(小时)': round(min_duration, 2),
        '时长分布(每2小时)': {k: v for k, v in sorted(buckets.items())},
        '长时停车(>12小时)_数量': len(long_stays),
        '长时停车_TOP10': long_stays[:10],
        '入场时间分布(6-20点_每小时)': {f"{h}:00": entry_dist.get(h, 0) for h in range(6, 21)},
        '收费时间分布(6-20点_每小时)': {f"{h}:00": charge_dist.get(h, 0) for h in range(6, 21)},
    }, ensure_ascii=False, indent=2)


def tool_get_date_range() -> str:
    """获取数据的时间范围（起始日期、结束日期、跨越天数）"""
    _ensure_data_loaded()
    rows = _csv_rows
    cm = _column_mapper
    col_charge = cm.col_charge_time
    time_fmt = _detect_time_format(rows, col_charge)

    times = []
    for r in rows:
        try:
            t = datetime.strptime(r[col_charge].strip(), time_fmt)
            times.append(t)
        except (ValueError, KeyError):
            pass
    if not times:
        return json.dumps({'error': '无法解析时间'})
    times.sort()
    return json.dumps({
        '起始日期': times[0].strftime('%Y-%m-%d'),
        '结束日期': times[-1].strftime('%Y-%m-%d'),
        '起始时间戳': times[0].strftime('%Y-%m-%d %H:%M:%S'),
        '结束时间戳': times[-1].strftime('%Y-%m-%d %H:%M:%S'),
        '跨越天数': (times[-1] - times[0]).days + 1,
    }, ensure_ascii=False)


def tool_get_current_time() -> str:
    """获取当前系统时间。用于填充「生成时间」等字段。"""
    now = datetime.now()
    return json.dumps({
        '当前日期': now.strftime('%Y-%m-%d'),
        '当前时间戳': now.strftime('%Y-%m-%d %H:%M:%S'),
        'ISO8601': now.isoformat(),
    }, ensure_ascii=False)


def tool_generate_chart(chart_type: str, title: str, data_json: str, output_filename: str = '') -> str:
    """
    使用 matplotlib 生成图表并保存为 PNG。
    
    参数:
      chart_type: 'pie' 饼图 | 'bar' 柱形图 | 'horizontal_bar' 横向柱形图
      title: 图表标题
      data_json: JSON 字符串，格式根据 chart_type 变化:
                 pie/bar → {"labels": [...], "values": [...]}
                 horizontal_bar → {"labels": [...], "values": [...]}
      output_filename: 输出文件名（不含路径），留空则自动生成
    返回：生成的 PNG 文件完整路径
    """
    if not HAS_MPL:
        return json.dumps({'error': 'matplotlib 未安装，无法生成图表。请执行: pip install matplotlib'})
    
    _setup_matplotlib_font()
    CHARTS_DIR.mkdir(parents=True, exist_ok=True)
    
    data = json.loads(data_json) if isinstance(data_json, str) else data_json
    labels = data.get('labels', [])
    values = data.get('values', [])
    
    if not labels or not values:
        return json.dumps({'error': 'labels 或 values 为空'})
    
    fig, ax = plt.subplots(figsize=(7, 3.8))
    
    # 配色方案（专业蓝灰系，柔和过渡）
    colors = ['#2E5E8C', '#4A80B5', '#72A0D0', '#9ABFE0', '#C0D6ED',
              '#5C8A6E', '#7BA68C', '#A3C4AD', '#C5DDCD', '#E0EFE3']
    
    if chart_type == 'pie':
        wedges, texts, autotexts = ax.pie(
            values, labels=None, autopct='%1.1f%%',
            colors=colors[:len(labels)], pctdistance=0.78,
            startangle=140, wedgeprops={'linewidth': 0.5, 'edgecolor': 'white'}
        )
        for t in autotexts:
            t.set_fontsize(8)
        ax.legend(wedges, [f'{l}  {v}' for l, v in zip(labels, values)],
                  loc='lower center', bbox_to_anchor=(0.5, -0.15),
                  ncol=min(3, len(labels)), fontsize=8, frameon=False)
        ax.set_title(title, fontsize=12, fontweight='bold', pad=10)
        
    elif chart_type == 'bar':
        bars = ax.bar(range(len(labels)), values, color=colors[:len(labels)],
                       edgecolor='white', linewidth=0.5, width=0.65)
        ax.set_xticks(range(len(labels)))
        ax.set_xticklabels(labels, rotation=0, ha='center', fontsize=8)
        ax.set_title(title, fontsize=12, fontweight='bold', pad=8)
        ax.tick_params(axis='y', labelsize=8)
        for bar, v in zip(bars, values):
            ax.text(bar.get_x() + bar.get_width()/2, bar.get_height() + max(values)*0.02,
                    str(v), ha='center', va='bottom', fontsize=8, fontweight='bold', color='#333')
        ax.spines['top'].set_visible(False)
        ax.spines['right'].set_visible(False)
        ax.spines['left'].set_color('#cccccc')
        ax.spines['bottom'].set_color('#cccccc')
        ax.yaxis.grid(True, linestyle='--', alpha=0.3, color='#aaaaaa')
        ax.set_axisbelow(True)
        
    elif chart_type == 'horizontal_bar':
        bars = ax.barh(range(len(labels)), values, color=colors[:len(labels)],
                        edgecolor='white', linewidth=0.5, height=0.6)
        ax.set_yticks(range(len(labels)))
        ax.set_yticklabels(labels, fontsize=8)
        ax.set_title(title, fontsize=12, fontweight='bold', pad=8)
        ax.tick_params(axis='x', labelsize=8)
        for bar, v in zip(bars, values):
            ax.text(bar.get_width() + max(values)*0.02, bar.get_y() + bar.get_height()/2,
                    str(v), ha='left', va='center', fontsize=8, fontweight='bold', color='#333')
        ax.spines['top'].set_visible(False)
        ax.spines['right'].set_visible(False)
        ax.spines['left'].set_color('#cccccc')
        ax.spines['bottom'].set_color('#cccccc')
        ax.xaxis.grid(True, linestyle='--', alpha=0.3, color='#aaaaaa')
        ax.set_axisbelow(True)
        
    else:
        plt.close()
        return json.dumps({'error': f'不支持的图表类型：{chart_type}。支持：pie, bar, horizontal_bar'})
    
    # 生成文件名
    if not output_filename:
        safe_title = hashlib.md5(title.encode()).hexdigest()[:8]
        output_filename = f"{chart_type}_{safe_title}.png"
    output_path = CHARTS_DIR / output_filename
    
    fig.tight_layout(pad=1.2)
    fig.savefig(output_path, dpi=140, bbox_inches='tight', facecolor='white')
    plt.close(fig)
    
    global _chart_files
    _chart_files.append(str(output_path))
    
    return json.dumps({
        'chart_path': str(output_path),
        'chart_type': chart_type,
        'title': title,
        'filename': output_filename,
    }, ensure_ascii=False)


def tool_generate_payment_chart(column: str = None) -> str:
    """生成支付方式分布饼图。可指定列名，默认自动检测。"""
    _ensure_data_loaded()
    col = column or _column_mapper.col_pay_method
    counter = Counter(r.get(col, '') for r in _csv_rows if r.get(col, '').strip())
    labels = [k for k, v in counter.most_common()]
    values = [v for k, v in counter.most_common()]
    data_json = json.dumps({'labels': labels, 'values': values})
    return tool_generate_chart('pie', '支付方式分布', data_json, 'payment_method_pie.png')


def tool_generate_duration_chart(charge_col: str = None, entry_col: str = None,
                                  bucket_hours: int = 2, max_buckets: int = 7) -> str:
    """生成停车时长分布柱形图。

    参数:
      charge_col: 收费时间列名（默认自动检测）
      entry_col:  进车时间列名（默认自动检测）
      bucket_hours: 每桶小时数（默认 2）
      max_buckets:  桶数量（默认 7，超出部分归入最后桶）
    """
    _ensure_data_loaded()
    cm = _column_mapper
    col_c = charge_col or cm.col_charge_time
    col_e = entry_col or cm.col_entry_time
    time_fmt = _detect_time_format(_csv_rows, col_c)

    durations = []
    for r in _csv_rows:
        try:
            tc = datetime.strptime(r[col_c].strip(), time_fmt)
            te = datetime.strptime(r[col_e].strip(), time_fmt)
            durations.append((tc - te).total_seconds() / 3600.0)
        except (ValueError, KeyError):
            pass

    buckets = [0] * max_buckets
    bucket_labels = []
    for i in range(max_buckets - 1):
        lo = i * bucket_hours
        hi = lo + bucket_hours
        bucket_labels.append(f'{lo}-{hi}h')
    bucket_labels.append(f'{(max_buckets - 1) * bucket_hours}h+')

    for dh in durations:
        idx = min(int(dh // bucket_hours), max_buckets - 1)
        buckets[idx] += 1

    data_json = json.dumps({'labels': bucket_labels, 'values': buckets})
    return tool_generate_chart('bar', f'停车时长分布（每{bucket_hours}小时）', data_json, 'duration_bar.png')


def tool_generate_entry_hour_chart(column: str = None, hour_start: int = 6,
                                     hour_end: int = 20) -> str:
    """生成每日入场时间分布柱形图。

    参数:
      column: 入场时间列名（默认自动检测）
      hour_start / hour_end: 统计的小时范围（默认 6-20）
    """
    _ensure_data_loaded()
    col = column or _column_mapper.col_entry_time
    time_fmt = _detect_time_format(_csv_rows, col)
    dist = defaultdict(int)
    for r in _csv_rows:
        try:
            te = datetime.strptime(r[col].strip(), time_fmt)
            h = te.hour
            if hour_start <= h < hour_end:
                dist[h] += 1
        except (ValueError, KeyError):
            pass
    labels = [f'{h}:00' for h in range(hour_start, hour_end + 1)]
    values = [dist.get(h, 0) for h in range(hour_start, hour_end + 1)]
    data_json = json.dumps({'labels': labels, 'values': values})
    return tool_generate_chart('bar',
        f'每日入场时间分布（{hour_start}:00-{hour_end}:00）', data_json, 'entry_hour_bar.png')


def tool_generate_charge_hour_chart(column: str = None, hour_start: int = 6,
                                      hour_end: int = 20) -> str:
    """生成每日收费时间分布柱形图。

    参数:
      column: 收费时间列名（默认自动检测）
      hour_start / hour_end: 统计的小时范围（默认 6-20）
    """
    _ensure_data_loaded()
    col = column or _column_mapper.col_charge_time
    time_fmt = _detect_time_format(_csv_rows, col)
    dist = defaultdict(int)
    for r in _csv_rows:
        try:
            tc = datetime.strptime(r[col].strip(), time_fmt)
            h = tc.hour
            if hour_start <= h < hour_end:
                dist[h] += 1
        except (ValueError, KeyError):
            pass
    labels = [f'{h}:00' for h in range(hour_start, hour_end + 1)]
    values = [dist.get(h, 0) for h in range(hour_start, hour_end + 1)]
    data_json = json.dumps({'labels': labels, 'values': values})
    return tool_generate_chart('bar',
        f'每日收费时间分布（{hour_start}:00-{hour_end}:00）', data_json, 'charge_hour_bar.png')


def tool_fill_paragraph_run(block_index: int, run_index: int, new_text: str) -> str:
    """
    修改指定段落 block 中某个 run 的文本内容。
    自动将占位符红色（C00000）替换为黑色（000000），并移除斜体标记。
    
    参数:
      block_index: block 的 index 号
      run_index: run 在 runs 数组中的序号
      new_text: 新的文本内容
    """
    _ensure_data_loaded()
    blocks = _template_data['blocks']
    target = None
    for b in blocks:
        if b.get('index') == block_index:
            target = b
            break
    if target is None:
        return json.dumps({'error': f'找不到 block index={block_index}'})
    if target['type'] != 'paragraph':
        return json.dumps({'error': f'block {block_index} 不是 paragraph 类型，是 {target["type"]}'})
    
    runs = target.get('runs', [])
    if run_index >= len(runs):
        return json.dumps({'error': f'run_index {run_index} 超出范围（共 {len(runs)} 个 runs）'})
    
    run = runs[run_index]
    old_text = run.get('text', '')
    run['text'] = new_text
    
    # 自动修正格式：红字→黑字，斜体→正常
    fmt = run.get('fmt', {})
    if fmt.get('color') == PLACEHOLDER_COLOR:
        fmt['color'] = FILLED_COLOR
    fmt.pop('italic', None)
    fmt.pop('italic_cjk', None)
    
    return json.dumps({
        'success': True,
        'block_index': block_index,
        'run_index': run_index,
        'old_text': old_text,
        'new_text': new_text,
    }, ensure_ascii=False)


def tool_fill_table_cell(
    block_index: int, row_index: int, col_index: int,
    run_index: int, new_text: str
) -> str:
    """
    修改表格中某个单元格内某个 run 的文本。
    自动将占位符红色替换为黑色，移除斜体。
    
    参数:
      block_index: 表格 block 的 index
      row_index: 行号（从 0 开始）
      col_index: 列号（从 0 开始）
      run_index: run 在段落 runs 数组中的序号（通常为 0）
      new_text: 新文本
    """
    _ensure_data_loaded()
    blocks = _template_data['blocks']
    target = None
    for b in blocks:
        if b.get('index') == block_index:
            target = b
            break
    if target is None:
        return json.dumps({'error': f'找不到 block index={block_index}'})
    if target['type'] != 'table':
        return json.dumps({'error': f'block {block_index} 不是 table 类型'})
    
    rows = target.get('rows', [])
    if row_index >= len(rows):
        return json.dumps({'error': f'row_index {row_index} 超出范围（共 {len(rows)} 行）'})
    
    cells = rows[row_index].get('cells', [])
    if col_index >= len(cells):
        return json.dumps({'error': f'col_index {col_index} 超出范围（共 {len(cells)} 列）'})
    
    paragraphs = cells[col_index].get('paragraphs', [])
    # 默认对第一个 paragraph 操作
    if not paragraphs:
        return json.dumps({'error': '单元格内无段落'})
    
    old_text = ''.join(r.get('text', '') for r in paragraphs[0].get('runs', []))
    
    # 整个单元格替换为单个 run（彻底清除 \"N ]\" 等旧占位残留）
    paragraphs[0]['runs'] = [{'text': new_text, 'fmt': {'color': FILLED_COLOR}}]
    
    return json.dumps({
        'success': True,
        'block_index': block_index,
        'cell': f'[{row_index},{col_index}]',
        'old_text': old_text,
        'new_text': new_text,
    }, ensure_ascii=False)


def tool_replace_paragraph_text(block_index: int, new_text: str) -> str:
    """
    完全替换一个段落 block 的全部文本内容（合并所有 runs 为一个 run）。
    保留原段落格式（paraFmt, style），新文本为默认黑色。
    
    参数:
      block_index: block 的 index
      new_text: 新文本（支持 \n 自动拆分为多个段落 runs）
    """
    _ensure_data_loaded()
    blocks = _template_data['blocks']
    target = None
    for b in blocks:
        if b.get('index') == block_index:
            target = b
            break
    if target is None:
        return json.dumps({'error': f'找不到 block index={block_index}'})
    if target['type'] != 'paragraph':
        return json.dumps({'error': f'block {block_index} 不是 paragraph 类型'})
    
    old_text = ''.join(r.get('text', '') for r in target.get('runs', []))
    old_style = target.get('style')

    # ── 护栏：如果目标 block 是标题样式且新文本是长文正文 → 警告 ──
    style_warning = None
    if old_style and old_style not in ('', 'None', 'none', None):
        try:
            sid = int(old_style)
            if 2 <= sid <= 7 and len(new_text) > 30:
                style_warning = (
                    f'⚠️ 目标 block {block_index} 当前样式为 heading {sid - 1}（标题），'
                    f'但你正在填入 {len(new_text)} 字的长文正文。'
                    f'填入后请调用 set_block_style({block_index}, None) 将样式改为正文。'
                )
        except (ValueError, TypeError):
            pass

    # 替换为新的 runs（黑色，正常字体）
    target['runs'] = [{'text': new_text, 'fmt': {'color': FILLED_COLOR}}]
    
    # 清除段落级默认红色斜体格式
    if target.get('paraFmt', {}).get('defaultRunFmt'):
        drf = target['paraFmt']['defaultRunFmt']
        if drf.get('color') == PLACEHOLDER_COLOR:
            drf['color'] = FILLED_COLOR
        drf.pop('italic', None)
        drf.pop('italic_cjk', None)
    
    result = {
        'success': True,
        'block_index': block_index,
        'old_text': old_text[:200],
        'new_text': new_text[:200],
    }
    if style_warning:
        result['warning'] = style_warning
        result['suggested_fix'] = f"set_block_style({block_index}, None)"
    return json.dumps(result, ensure_ascii=False)


def tool_set_block_style(block_index: int, style_id: str) -> str:
    """
    修改指定段落 block 的样式 ID。

    用于 LLM 在不确定样式时做合理调整，例如：
      - 正文内容被填入了 heading 块 → set_block_style(idx, None) 改为正文
      - 需要提升某段为标题 → set_block_style(idx, '2') 改为 heading 1

    参数:
      block_index: 目标 block 的 index
      style_id: 样式 ID（'1'=Normal, '2'=heading 1, '3'=heading 2, None=无样式/正文）
    """
    _ensure_data_loaded()
    blocks = _template_data['blocks']
    target = None
    for b in blocks:
        if b.get('index') == block_index:
            target = b
            break
    if target is None:
        return json.dumps({'error': f'找不到 block index={block_index}'})
    if target['type'] != 'paragraph':
        return json.dumps({'error': f'block {block_index} 不是 paragraph 类型'})

    old_style = target.get('style')
    target['style'] = style_id if style_id and style_id.lower() != 'none' else None

    return json.dumps({
        'success': True,
        'block_index': block_index,
        'old_style': old_style,
        'new_style': style_id,
    }, ensure_ascii=False)


def tool_delete_blocks(block_indices: list) -> str:
    """
    从模板中删除指定 index 的 block。
    用于删除说明/指令性的占位块（如"红色斜体为需由智能体填写的占位内容"等）。
    
    参数:
      block_indices: 要删除的 block index 列表
    """
    _ensure_data_loaded()
    blocks = _template_data['blocks']
    idx_set = set(block_indices)
    removed = []
    kept = []
    for b in blocks:
        if b.get('index') in idx_set:
            removed.append({'index': b['index'], 'text_preview': ''.join(
                r.get('text', '') for r in b.get('runs', []))[:100]})
        else:
            kept.append(b)
    _template_data['blocks'] = kept
    _reindex_blocks()
    return json.dumps({
        'success': True,
        'removed_count': len(removed),
        'removed': removed,
        'remaining_blocks': len(kept),
    }, ensure_ascii=False)


def tool_insert_chart_block(
    after_block_index: int,
    image_path: str,
    title: str,
    caption: str = ''
) -> str:
    """
    在指定 block 之后插入图表段落（标题 + 图片）。
    如果紧邻的下一个 block 是红色占位符，自动替换之（先删后插）。
    JSON 还原为 docx 时由 json_to_docx 处理图片嵌入。

    参数:
      after_block_index: 在其后插入（填入该 block 的 index 号）
      image_path: 图表 PNG 文件的绝对路径
      title: 图表小标题（如"图1：支付方式分布"）
      caption: 图表下方的说明文字（可选）
    """
    _ensure_data_loaded()
    blocks = _template_data['blocks']

    insert_pos = _find_block_pos(after_block_index)
    if insert_pos < 0:
        return json.dumps({'error': f'找不到 after_block_index={after_block_index}'})
    insert_pos += 1  # 在目标 block 之后

    # ├─ 如果紧邻的下一个 block 是指令块/纯占位块 → 替换它
    replaced = None
    if insert_pos < len(blocks):
        nb = blocks[insert_pos]
        if _is_instruction_block(nb):
            replaced = blocks.pop(insert_pos)

    # 构建新 block
    new_blocks = []

    # 标题段落
    new_blocks.append({
        'type': 'paragraph',
        'style': '4',  # heading 3
        'runs': [{'text': title, 'fmt': {'color': '1F4D78'}}],
        '_generated': True,
    })

    # 图片段落
    new_blocks.append({
        'type': 'paragraph',
        'style': None,
        'runs': [],
        'paraFmt': {'alignment': 'center', 'spacing': {'before': 80, 'after': 80}},
        'image': str(image_path),
        '_generated': True,
    })

    # 说明文字
    if caption:
        new_blocks.append({
            'type': 'paragraph',
            'style': None,
            'runs': [{'text': caption, 'fmt': {'color': '595959', 'sz_hp': 18, 'szCs_hp': 18}}],
            'paraFmt': {'alignment': 'center', 'spacing': {'after': 160}},
            '_generated': True,
        })

    # 插入
    for nb in reversed(new_blocks):
        blocks.insert(insert_pos, nb)

    _reindex_blocks()

    return json.dumps({
        'success': True,
        'inserted_after': after_block_index,
        'new_block_count': len(new_blocks),
        'replaced_placeholder': bool(replaced),
        'image_path': str(image_path),
    }, ensure_ascii=False)


def _find_kpi_table_index() -> int:
    """在 blocks 中查找「关键指标」表格的 index。通过识别表头「指标/填写值/示例预期值」定位。"""
    _ensure_data_loaded()
    for b in _template_data['blocks']:
        if b.get('type') != 'table':
            continue
        rows = b.get('rows', [])
        if not rows:
            continue
        # 检查第一行第一列是否含 "指标"
        first_cell = rows[0].get('cells', [])
        if first_cell:
            first_text = ''.join(
                r.get('text', '') for p in first_cell[0].get('paragraphs', [])
                for r in p.get('runs', [])
            )
            if '指标' in first_text:
                return b.get('index')
    return -1


# ── KPI 标签语义模式（用于自动匹配表格行 → 指标类型）──
KPI_LABEL_PATTERNS = [
    ('count',       ['总交易笔数', '交易笔数', '总笔数', '笔数']),
    ('receivable',  ['应收总金额', '应收金额', '应收合计']),
    ('received',    ['实收总金额', '实收金额', '实收合计']),
    ('deduction',   ['实际抵扣总额', '抵扣总额', '实际抵扣额', '抵扣合计']),
    ('rate',        ['实收率', '收费率']),
    ('main_pay',    ['主要支付方式', '支付方式', '主力支付']),
]


def _match_kpi_label(label_text: str) -> str | None:
    """根据标签文字匹配 KPI 类型。返回类型名或 None。"""
    label_lower = label_text.replace(' ', '').lower()
    for kpi_type, patterns in KPI_LABEL_PATTERNS:
        for p in patterns:
            if p.replace(' ', '').lower() in label_lower:
                return kpi_type
    return None


def _build_kpi_value(kpi_type: str, kpis: dict) -> str:
    """根据 KPI 类型和计算结果构建显示文本。"""
    if kpi_type == 'count':
        return f'{kpis["总交易笔数"]} 笔'
    elif kpi_type == 'receivable':
        return f'¥{kpis["应收总金额(元)"]:,.2f}'
    elif kpi_type == 'received':
        return f'¥{kpis["实收总金额(元)"]:,.2f}'
    elif kpi_type == 'deduction':
        return f'¥{kpis["实际抵扣总额(元)"]:,.2f}'
    elif kpi_type == 'rate':
        return f'{kpis["实收率(%)"]}%'
    elif kpi_type == 'main_pay':
        top = kpis.get("主要支付方式_TOP1")
        return f'{top["方式"]} ×{top["笔数"]}' if top else '(无)'
    return ''


def tool_fill_kpi_table_all() -> str:
    """
    一键填充关键指标表格中所有占位单元格。

    通过表头「指标」定位表格，遍历每行的红色占位单元格，
    读取同行第 0 列的标签文字，通过语义匹配确定要填入的 KPI 值。
    不依赖固定的行号或列序 —— 模板增减行/换序后仍然正确。
    """
    kpis = json.loads(tool_compute_all_kpis())
    table_idx = _find_kpi_table_index()
    if table_idx < 0:
        return json.dumps({'error': '找不到 KPI 指标表格（表头含"指标"的 table）'})

    _ensure_data_loaded()
    table = _find_block_by_index(table_idx)
    if not table:
        return json.dumps({'error': f'找不到 table block index={table_idx}'})

    rows = table.get('rows', [])
    results = []

    for row_idx, row in enumerate(rows):
        cells = row.get('cells', [])
        if len(cells) < 2:
            continue

        # 读取第 0 列的标签文字
        label_text = ''.join(
            r.get('text', '')
            for p in cells[0].get('paragraphs', [])
            for r in p.get('runs', [])
        )
        kpi_type = _match_kpi_label(label_text)
        if not kpi_type:
            continue

        # 只在"填写值"列（第 1 列）找占位符 run，跳过"示例预期值"等参考列
        for col_idx in range(1, len(cells)):
            # 先检查该列是否已经有非占位内容（已填过）
            all_text = ''.join(
                r.get('text', '') for p in cells[col_idx].get('paragraphs', [])
                for r in p.get('runs', []))
            has_placeholder = any(
                _is_placeholder_run(r) for p in cells[col_idx].get('paragraphs', [])
                for r in p.get('runs', []))
            if not has_placeholder:
                continue  # 该列已填完，跳过

            for pi, p in enumerate(cells[col_idx].get('paragraphs', [])):
                for ri, run in enumerate(p.get('runs', [])):
                    if _is_placeholder_run(run):
                        value_text = _build_kpi_value(kpi_type, kpis)
                        if value_text:
                            res = json.loads(tool_fill_table_cell(
                                table_idx, row_idx, col_idx, ri, value_text))
                            results.append(res)
                        break
                if results and results[-1].get('success'):
                    break
            if results and results[-1].get('success'):
                break

    return json.dumps({
        'success': True,
        'table_index': table_idx,
        'filled_cells': len(results),
        'results': results,
    }, ensure_ascii=False, indent=2)


def tool_save_filled_json(output_path: str = None) -> str:
    """保存当前填充后的模板为 JSON 文件"""
    _ensure_data_loaded()
    p = Path(output_path) if output_path else OUTPUT_JSON
    with open(p, 'w', encoding='utf-8') as f:
        json.dump(_template_data, f, ensure_ascii=False, indent=2)
    size_kb = os.path.getsize(p) / 1024
    return json.dumps({
        'success': True,
        'output_path': str(p),
        'size_kb': round(size_kb, 1),
        'block_count': len(_template_data['blocks']),
    }, ensure_ascii=False)


def tool_restore_to_docx(json_path: str = None, docx_path: str = None) -> str:
    """
    将填充后的 JSON 还原为 .docx 文件。
    调用 json_to_docx 的 restore_docx 函数。
    """
    jp = json_path or str(OUTPUT_JSON)
    dp = docx_path or str(OUTPUT_DOCX)
    
    # 动态导入
    sys.path.insert(0, str(BASE_DIR))
    from json_to_docx import restore_docx
    restore_docx(jp, dp)
    
    size_kb = os.path.getsize(dp) / 1024 if os.path.exists(dp) else 0
    return json.dumps({
        'success': True,
        'docx_path': dp,
        'size_kb': round(size_kb, 1),
    }, ensure_ascii=False)


def tool_verify_report() -> str:
    """
    自检当前已填充的模板，返回所有残留问题：
      - 仍为红色 (C00000) 的占位 run
      - 正文内容跑到了标题样式 block（style=2~7 但文本不像标题）
      - block 总数和占位统计

    必须在 save_filled_json / restore_to_docx 之前调用。
    如果发现问题，应修复后再保存。
    """
    _ensure_data_loaded()
    blocks = _template_data['blocks']
    issues = []

    for b in blocks:
        idx = b['index']
        btype = b['type']
        style = b.get('style')

        if btype != 'paragraph':
            continue

        text = ''.join(r.get('text', '') for r in b.get('runs', []))

        # 1. 检测红色占位残留
        for ri, r in enumerate(b.get('runs', [])):
            if r.get('fmt', {}).get('color') == PLACEHOLDER_COLOR and r.get('text', '').strip():
                issues.append({
                    'type': 'red_placeholder_remaining',
                    'block_index': idx,
                    'run_index': ri,
                    'text': r['text'][:80],
                    'fix': f"fill_paragraph_run({idx}, {ri}, '...') 或 replace_paragraph_text({idx}, '...') 或 delete_blocks([{idx}])",
                })

        # 2. 检测正文内容跑到了标题样式
        if style and style not in ('', 'None', 'none', None) and text.strip():
            # 判断是否为标题样式（兼容整数 ID "2" 和字符串名 "Heading1"）
            is_heading = False
            style_lower = str(style).lower()
            if style_lower in ('heading1', 'heading2', 'heading3', 'heading4', 'heading5', 'heading6',
                               'heading 1', 'heading 2', 'heading 3', 'heading 4', 'heading 5', 'heading 6',
                               '2', '3', '4', '5', '6', '7'):
                is_heading = True
            if is_heading and (len(text) > 30 or text.startswith('1.') or text.startswith('①')):
                issues.append({
                    'type': 'body_text_in_heading_style',
                    'block_index': idx,
                    'style': style,
                    'text_preview': text[:80],
                    'fix': f"set_block_style({idx}, None)  ← 将样式改为正文",
                })

        # 3. 检测非标题内容顶着标题样式名
        if b.get('_generated') and style == '4':
            if len(text) > 20 and not text.startswith('图'):
                issues.append({
                    'type': 'generated_block_wrong_content',
                    'block_index': idx,
                    'text_preview': text[:80],
                    'fix': '可能是图表标题被其他内容覆盖。请检查并修复。',
                })

    # 4. 检测孤立标题（heading 后紧跟另一个 heading 或文件末尾，无正文内容）
    for i, b in enumerate(blocks):
        if b.get('type') != 'paragraph':
            continue
        style = b.get('style')
        if not style:
            continue
        style_lower = str(style).lower()
        if style_lower not in ('heading1', 'heading2', 'heading3', 'heading4', 'heading5', 'heading6',
                               'heading 1', 'heading 2', 'heading 3', 'heading 4', 'heading 5', 'heading 6',
                               '2', '3', '4', '5', '6', '7'):
            continue

        # 检查下一个 block
        heading_text = ''.join(r.get('text', '') for r in b.get('runs', []))[:40]
        if i + 1 >= len(blocks):
            issues.append({
                'type': 'orphan_heading_at_end',
                'block_index': b['index'],
                'heading_text': heading_text,
                'fix': 'Section 标题后无内容，可能缺少正文段落。',
            })
            continue

        next_b = blocks[i + 1]
        next_style = str(next_b.get('style', '')).lower()
        if next_style in ('heading1', 'heading2', 'heading3', 'heading4', 'heading5', 'heading6',
                          'heading 1', 'heading 2', 'heading 3', 'heading 4', 'heading 5', 'heading 6',
                          '2', '3', '4', '5', '6', '7'):
            issues.append({
                'type': 'consecutive_headings',
                'block_index': b['index'],
                'heading_text': heading_text,
                'next_heading': ''.join(r.get('text', '') for r in next_b.get('runs', []))[:40],
                'fix': '两个标题之间缺少正文内容。检查是否需要调整顺序或补充段落。',
            })

    return json.dumps({
        'total_blocks': len(blocks),
        'issue_count': len(issues),
        'issues': issues,
        'all_clear': len(issues) == 0,
    }, ensure_ascii=False, indent=2)


# ═══════════════════════════════════════════════════════════
# 3. 工具注册（OpenAI Function Calling 格式）
# ═══════════════════════════════════════════════════════════

TOOLS_FUNCTIONS = [
    {
        "name": "get_template_structure",
        "description": "获取模板的完整结构摘要：列出所有 block 的 index、类型、样式、文本预览，并标注哪些是占位符（红色斜体=待填充），哪些已有内容。这是理解模板的第一步。",
        "parameters": {"type": "object", "properties": {}, "required": []},
    },
    {
        "name": "get_placeholder_blocks",
        "description": "列出所有需要填充的占位 run（两层检测：格式颜色 + 内容 bracket 模式）。返回每个占位的 block_index、run_index、当前文本、检测方式（color/content）。用于精确定位填充目标。",
        "parameters": {"type": "object", "properties": {}, "required": []},
    },
    {
        "name": "get_placeholder_rules",
        "description": "返回当前占位符检测规则的详细说明（颜色阈值、样式要求、内容模式等），供 LLM 理解检测逻辑。如果模板使用非标准占位格式，可据此建议环境变量调整。",
        "parameters": {"type": "object", "properties": {}, "required": []},
    },
    {
        "name": "load_csv_summary",
        "description": "读取 data.csv 并返回概览：总行数、每列的字段名、数据类型（数值/文本）、示例值。",
        "parameters": {"type": "object", "properties": {}, "required": []},
    },
    {
        "name": "compute_all_kpis",
        "description": "一键计算模板「关键指标」表格所需的全部 KPI：总交易笔数、应收总金额、实收总金额、实际抵扣总额、实收率、主要支付方式 TOP1/TOP3。返回完整计算结果 JSON。",
        "parameters": {"type": "object", "properties": {}, "required": []},
    },
    {
        "name": "compute_statistics",
        "description": "对 data.csv 的指定列执行任意统计运算。可用于灵活探索数据。",
        "parameters": {
            "type": "object",
            "properties": {
                "column": {"type": "string", "description": "列名，如 '应收金额'、'实收金额(元)'、'支付方式'"},
                "operation": {"type": "string", "enum": ["sum", "avg", "count", "max", "min", "median", "mode", "distinct"], "description": "运算类型"},
            },
            "required": ["column", "operation"],
        },
    },
    {
        "name": "analyze_payment_distribution",
        "description": "分析支付方式与支付渠道的完整分布：各支付方式的笔数/金额、各渠道的笔数/金额、交叉分布 TOP10。用于撰写「二、支付方式与渠道」的分析叙述。",
        "parameters": {"type": "object", "properties": {}, "required": []},
    },
    {
        "name": "analyze_parking_duration",
        "description": "分析停车时长：平均/中位/最长/最短时长、每 2 小时的分布桶、超过 12 小时的长时停车明细、入场时间/收费时间的小时分布。用于撰写「三、停车时长分析」。",
        "parameters": {"type": "object", "properties": {}, "required": []},
    },
    {
        "name": "get_date_range",
        "description": "获取数据的时间范围：起始日期、结束日期、跨越天数。用于填充「数据周期」字段。",
        "parameters": {"type": "object", "properties": {}, "required": []},
    },
    {
        "name": "get_current_time",
        "description": "获取当前系统时间（精确到秒）。用于填充「生成时间」字段。LLM 不应使用内部时钟，必须调用此工具获取准确时间。",
        "parameters": {"type": "object", "properties": {}, "required": []},
    },
    {
        "name": "generate_chart",
        "description": "使用 matplotlib 生成图表并保存为 PNG。支持饼图(pie)、柱形图(bar)、横向柱形图(horizontal_bar)。返回 PNG 文件路径。",
        "parameters": {
            "type": "object",
            "properties": {
                "chart_type": {"type": "string", "enum": ["pie", "bar", "horizontal_bar"], "description": "图表类型"},
                "title": {"type": "string", "description": "图表标题"},
                "data_json": {"type": "string", "description": "JSON 字符串: {\"labels\": [...], \"values\": [...]}"},
                "output_filename": {"type": "string", "description": "输出文件名（不含路径），如 'chart1.png'，留空自动生成"},
            },
            "required": ["chart_type", "title", "data_json"],
        },
    },
    {
        "name": "generate_payment_chart",
        "description": "快捷生成支付方式分布饼图（一次性完成分析+绘图），返回 PNG 文件路径。",
        "parameters": {"type": "object", "properties": {}, "required": []},
    },
    {
        "name": "generate_duration_chart",
        "description": "快捷生成停车时长分布柱形图（每 2 小时为桶，0-12h+），返回 PNG 文件路径。",
        "parameters": {"type": "object", "properties": {}, "required": []},
    },
    {
        "name": "generate_entry_hour_chart",
        "description": "快捷生成每日入场时间分布柱形图（6:00-20:00，每小时），返回 PNG 文件路径。",
        "parameters": {"type": "object", "properties": {}, "required": []},
    },
    {
        "name": "generate_charge_hour_chart",
        "description": "快捷生成每日收费时间分布柱形图（6:00-20:00，每小时），返回 PNG 文件路径。",
        "parameters": {"type": "object", "properties": {}, "required": []},
    },
    {
        "name": "fill_paragraph_run",
        "description": "修改段落 block 中指定 run 的文本。自动将占位红色(C00000)改为黑色(000000)、移除斜体。用于精确替换段落中的占位符片段。",
        "parameters": {
            "type": "object",
            "properties": {
                "block_index": {"type": "integer", "description": "block 的 index 号"},
                "run_index": {"type": "integer", "description": "run 在 runs 数组中的序号（通常 0）"},
                "new_text": {"type": "string", "description": "新文本内容"},
            },
            "required": ["block_index", "run_index", "new_text"],
        },
    },
    {
        "name": "fill_table_cell",
        "description": "修改表格中指定单元格的文本。自动将占位红色改为黑色。用于精确填充表格中的占位值。",
        "parameters": {
            "type": "object",
            "properties": {
                "block_index": {"type": "integer", "description": "表格 block 的 index"},
                "row_index": {"type": "integer", "description": "行号（0 起始，表头=0）"},
                "col_index": {"type": "integer", "description": "列号（0 起始）"},
                "run_index": {"type": "integer", "description": "run 序号（通常 0）"},
                "new_text": {"type": "string", "description": "新文本"},
            },
            "required": ["block_index", "row_index", "col_index", "run_index", "new_text"],
        },
    },
    {
        "name": "replace_paragraph_text",
        "description": "完全替换一个段落 block 的全部文本（占位红色→黑色正文）。替换后 block 保持原有样式。如果发现填入后样式不对（如正文跑到标题块里），用 set_block_style 修正。",
        "parameters": {
            "type": "object",
            "properties": {
                "block_index": {"type": "integer", "description": "block 的 index"},
                "new_text": {"type": "string", "description": "新文本（自然语言叙述）"},
            },
            "required": ["block_index", "new_text"],
        },
    },
    {
        "name": "set_block_style",
        "description": "修改段落 block 的样式 ID。当填入的内容与当前样式不匹配时使用（如正文被填入标题块 → 改为 style=None；需要提升为标题 → style='2'）。style_id 为 None 表示正文。",
        "parameters": {
            "type": "object",
            "properties": {
                "block_index": {"type": "integer"},
                "style_id": {"type": "string", "description": "样式 ID：'1'=Normal, '2'=heading 1, '3'=heading 2, '4'=heading 3, None/null=正文"},
            },
            "required": ["block_index", "style_id"],
        },
    },
    {
        "name": "delete_blocks",
        "description": "从模板中删除指定的 block。⚠️ 调用后所有 block index 会重排，必须立即重新调用 get_template_structure 获取新 index！",
        "parameters": {
            "type": "object",
            "properties": {
                "block_indices": {"type": "array", "items": {"type": "integer"}, "description": "要删除的 block index 列表"},
            },
            "required": ["block_indices"],
        },
    },
    {
        "name": "insert_chart_block",
        "description": "在指定 block 后插入图表（标题+图片）。⚠️ 调用后所有 block index 会重排，必须立即重新调用 get_template_structure！",
        "parameters": {
            "type": "object",
            "properties": {
                "after_block_index": {"type": "integer", "description": "在其后插入的 block index"},
                "image_path": {"type": "string", "description": "图表 PNG 文件的完整路径"},
                "title": {"type": "string", "description": "图表小标题（如'图1：支付方式分布'）"},
                "caption": {"type": "string", "description": "图表下方说明（可选）"},
            },
            "required": ["after_block_index", "image_path", "title"],
        },
    },
    {
        "name": "fill_kpi_table_all",
        "description": "一键填充关键指标表格的所有占位单元格。自动查找表头含「指标」的表格，通过同行标签文字语义匹配确定填入值。不依赖固定行号或列序，模板增减行/换序后仍正确。",
        "parameters": {"type": "object", "properties": {}, "required": []},
    },
    {
        "name": "save_filled_json",
        "description": "将当前已填充的模板保存为 JSON 文件（默认 saved_filled_template.json）。在所有填充操作完成后调用。",
        "parameters": {
            "type": "object",
            "properties": {
                "output_path": {"type": "string", "description": "输出文件路径，默认 D:\\...\\filled_template.json"},
            },
            "required": [],
        },
    },
    {
        "name": "verify_report",
        "description": "自检当前已填充的模板：扫描残留的红色占位符、正文跑到标题样式、图表标题被覆盖等问题。必须在 save_filled_json 之前调用。返回问题列表和修复建议。",
        "parameters": {"type": "object", "properties": {}, "required": []},
    },
    {
        "name": "restore_to_docx",
        "description": "将填充后的 JSON 还原为 .docx Word 文档。在所有内容填充完毕、图表插入完成、且 verify_report 通过后调用。",
        "parameters": {
            "type": "object",
            "properties": {
                "json_path": {"type": "string", "description": "已填充 JSON 路径，默认 filled_template.json"},
                "docx_path": {"type": "string", "description": "输出 .docx 路径，默认 filled_report.docx"},
            },
            "required": [],
        },
    },
]

# 工具名 → 函数映射
TOOL_MAP = {
    'get_template_structure': tool_get_template_structure,
    'get_placeholder_blocks': tool_get_placeholder_blocks,
    'get_placeholder_rules': tool_get_placeholder_rules,
    'load_csv_summary': tool_load_csv_summary,
    'compute_all_kpis': tool_compute_all_kpis,
    'compute_statistics': tool_compute_statistics,
    'analyze_payment_distribution': tool_analyze_payment_distribution,
    'analyze_parking_duration': tool_analyze_parking_duration,
    'get_date_range': tool_get_date_range,
    'get_current_time': tool_get_current_time,
    'generate_chart': tool_generate_chart,
    'generate_payment_chart': tool_generate_payment_chart,
    'generate_duration_chart': tool_generate_duration_chart,
    'generate_entry_hour_chart': tool_generate_entry_hour_chart,
    'generate_charge_hour_chart': tool_generate_charge_hour_chart,
    'fill_paragraph_run': tool_fill_paragraph_run,
    'fill_table_cell': tool_fill_table_cell,
    'replace_paragraph_text': tool_replace_paragraph_text,
    'set_block_style': tool_set_block_style,
    'delete_blocks': tool_delete_blocks,
    'insert_chart_block': tool_insert_chart_block,
    'fill_kpi_table_all': tool_fill_kpi_table_all,
    'save_filled_json': tool_save_filled_json,
    'restore_to_docx': tool_restore_to_docx,
    'verify_report': tool_verify_report,
}


def execute_tool(name: str, args: dict) -> str:
    """执行工具调用，返回结果字符串"""
    func = TOOL_MAP.get(name)
    if not func:
        return json.dumps({'error': f'未知工具: {name}'})
    try:
        if args:
            return func(**args)
        else:
            return func()
    except Exception as e:
        return json.dumps({'error': str(e), 'tool': name, 'traceback': traceback.format_exc()})


# ═══════════════════════════════════════════════════════════
# 4. System Prompt
# ═══════════════════════════════════════════════════════════

SYSTEM_PROMPT = """你是一个报告模板自动填表 Agent。

## 核心原则：模板是给人看的，用你的语言理解能力去读它

模板本身是一份人类可读的文档。它的标题、说明文字、占位符标记——都是用人话写的。
**用你的语义理解能力去判断每个 block 的角色**，就像一个人拿到这份模板后会做的事：

- 读到「报告模板 —— 红色斜体为需由智能体填写的占位内容」→ 这是给制作者的说明，最终报告里不该出现 → **删除**
- 读到「【起始日期 – 结束日期】」→ 这是填空标记，需要替换为实际数据 → **填充**
- 读到「一、关键指标」→ 这是报告的固定标题，保持不变 → **不动**
- 读到「数据周期：2026-04-01 – 2026-04-30」→ 已经被填过了 → **不动**

代码只提供原始事实（`get_template_structure` 的格式标记 + 文本预览），**不做预分类**。
判断全由你根据文本含义 + 上下文做出。

## JSON 模板机制
- 模板由 `blocks` 数组组成，每个 block 有唯一的 `index`。类型：`paragraph` 或 `table`。
- `get_template_structure` 输出格式：`[index] type style=... fmt=RED,italic,bracket | 文本预览`
- `get_placeholder_blocks` 列出所有特殊格式的 run，供参考。
- 工具列表详见系统注册的函数，核心工具能力概述见下方。

## 工具能力概述

| 类别 | 工具 | 何时用 |
|------|------|--------|
| 探索 | `get_template_structure` | 第一步，了解全貌 |
| 探索 | `get_placeholder_blocks` | 定位需要处理的 run |
| 数据 | `load_csv_summary` / `compute_all_kpis` / `compute_statistics` | 理解 CSV 并计算指标 |
| 分析 | `analyze_payment_distribution` / `analyze_parking_duration` | 需要分类/时长分布时 |
| 时间 | `get_date_range` / `get_current_time` | 获取数据时间范围 / 当前时间 |
| 填充 | `replace_paragraph_text(idx, text)` | 整段占位文本替换为正文（自动去红） |
| 填充 | `fill_paragraph_run(idx, run_idx, text)` | 段落中单个 run 的精确替换 |
| 填充 | `fill_table_cell(table_idx, row, col, run, text)` | 表格单元格替换（整格原子操作） |
| 填充 | `fill_kpi_table_all` | 自动定位指标表并批量填充所有占位格 |
| 样式 | `set_block_style(idx, style_id)` | 修正段落样式（None=正文, '2'=标题） |
| 图表 | `generate_payment_chart` / `generate_duration_chart` / `generate_entry_hour_chart` / `generate_charge_hour_chart` | 生成 PNG 图表 |
| 图表 | `insert_chart_block(after_idx, path, title)` | 嵌入图表到指定位置 |
| 清理 | `delete_blocks([idx, ...])` | 删除指令块（⚠️ 会震 index） |
| 校验 | `verify_report` | 保存前自检：残留占位/样式错乱/连续标题 |
| 输出 | `save_filled_json` / `restore_to_docx` | 保存并生成 .docx |

## 工作流程（严格按此顺序）

### 1. 探索
`get_template_structure` + `load_csv_summary` → 理解模板结构和数据全貌。

### 2. 分析
根据模板内容和数据列，选择合适的分析工具计算所需指标。
（`compute_all_kpis` / `analyze_*` / `get_date_range` / `get_current_time` / `compute_statistics`）

### 3. 填充正文（⚠️ 从后往前）
**原理**：`replace_paragraph_text` 和 `fill_paragraph_run` 不改变 index。从文档末尾往前操作，每一步的 index 都准确。

操作方法：根据 Step 1 的 `get_template_structure` 结果，按 index 从大到小的顺序，逐个处理需要填充的占位块。

⚠️ **禁止自行添加元叙述**——模板中没有的句子，一句都不要加。图表标题由 `insert_chart_block` 的 `title` 参数指定。

### 4. 填表
`fill_kpi_table_all` 自动查找并填充。不受 index 影响，随时可调。

### 5. 图表（⚠️ 最后做前两步）
先生成图表，再逐个 `insert_chart_block`。每次 insert 后重新 `get_placeholder_blocks` 找下一个图表占位。

### 6. 删除指令块（⚠️ 最后做）
根据语义判断，用 `delete_blocks` 删除所有元指令块。删除后 `get_template_structure` 确认。

### 7. 校验
`verify_report` → 如果 `all_clear: false`，逐条修复后再次验证，直到通过。

### 8. 输出
`save_filled_json` → `restore_to_docx`

## 关键规则
- ⚠️ **从后往前填**：`replace_paragraph_text` / `fill_paragraph_run` 不震 index。
- ⚠️ **insert 和 delete 最后做**：它们震 index，但此时内容已填完。
- ⚠️ **fill 类工具自动去红**：占位红色 C00000 自动变黑色 000000。
- ⚠️ **保存前必须 verify_report 且 all_clear 为 true**。
- ⚠️ **每个标题后面必须有正文**。如果标题后紧跟另一个标题或空 block → 缺内容，立刻补。
- ⚠️ **语义优先**：格式标记（RED/italic/bracket）是线索，不是判决。最终由你阅读文本含义决定操作。
"""


# ═══════════════════════════════════════════════════════════
# 5. Agent 主循环
# ═══════════════════════════════════════════════════════════

def _log_llm_call(response, elapsed_ms: float, messages: list = None):
    """记录 LLM 调用日志（如果 AGENT_LOG_JOB_ID 环境变量已设置）。"""
    job_id = os.environ.get('AGENT_LOG_JOB_ID', '').strip()
    if not job_id:
        return

    msg = response.choices[0].message
    tool_details = []
    for tc in (msg.tool_calls or []):
        try:
            args = json.loads(tc.function.arguments) if tc.function.arguments else {}
        except Exception:
            args = tc.function.arguments
        tool_details.append({'name': tc.function.name, 'args': args})

    from pathlib import Path as _Path
    job_dir = os.environ.get('AGENT_JOB_DIR', '')
    if job_dir:
        jobs_root = _Path(job_dir).parent  # parent of job dir = jobs/
    else:
        jobs_root = _Path(__file__).resolve().parent / 'jobs'
    llm_log = jobs_root / job_id / 'llm_calls.jsonl'

    call_no = sum(1 for _ in open(llm_log, 'r', encoding='utf-8')) + 1 if llm_log.exists() else 1

    usage = getattr(response, 'usage', None)

    # 对话上下文快照
    messages_snapshot = []
    for m in (messages or []):
        snapshot = {'role': m.get('role', '?')}
        if m.get('content'):
            snapshot['content'] = str(m['content'])[:3000]
        if m.get('tool_calls'):
            snapshot['tool_calls'] = [{'name': tc['function']['name'],
                                        'args': tc['function']['arguments'][:300]}
                                       for tc in m['tool_calls']]
        if m.get('tool_call_id'):
            snapshot['tool_call_id'] = m['tool_call_id']
            snapshot['content'] = str(m.get('content', ''))[:1000]
        messages_snapshot.append(snapshot)

    entry = {
        'call': call_no,
        'ts': datetime.now().strftime('%Y-%m-%dT%H:%M:%SZ'),
        'model': getattr(response, 'model', '?'),
        'latency_ms': round(elapsed_ms),
        'prompt_tokens': getattr(usage, 'prompt_tokens', 0) if usage else 0,
        'completion_tokens': getattr(usage, 'completion_tokens', 0) if usage else 0,
        'finish_reason': msg.finish_reason if hasattr(msg, 'finish_reason') else None,
        'messages': messages_snapshot,
        'response': (msg.content or '')[:2000],
        'tool_calls': tool_details,
        'reasoning_content': (getattr(msg, 'reasoning_content', '') or '')[:2000],
    }

    llm_log.parent.mkdir(parents=True, exist_ok=True)
    with open(llm_log, 'a', encoding='utf-8') as f:
        f.write(json.dumps(entry, ensure_ascii=False) + '\n')


class TemplateFillerAgent:
    """停车报告模板填充 Agent"""

    def __init__(self):
        self.client = None
        self.messages = [{"role": "system", "content": SYSTEM_PROMPT}]
        self.model = os.environ.get('PRO_MODEL', 'deepseek-chat')
        self.max_steps = 50

    def setup(self):
        if not HAS_OPENAI:
            raise RuntimeError("请安装 OpenAI SDK: pip install openai")
        self.client = OpenAI(
            api_key=os.environ.get('KEY'),
            base_url=os.environ.get('BASE_URL'),
        )

    def run(self):
        """主循环：LLM 驱动工具调用直到完成"""
        self.setup()

        # 注入用户指令（如果通过环境变量传入）
        user_instr = os.environ.get('AGENT_USER_INSTRUCTIONS', '').strip()
        if user_instr:
            self.messages.append({
                "role": "user",
                "content": f"【用户特别指令】{user_instr}\n\n请严格遵守以上指令完成报告生成。现在开始执行。"
            })

        print("=" * 60)
        print("  停车明细分析报告 — 智能填充 Agent")
        print("=" * 60)
        print(f"  模板: {TEMPLATE_PATH}")
        print(f"  数据: {CSV_PATH} ({_csv_rows and len(_csv_rows) or '未加载'} 行)")
        print(f"  模型: {self.model}")
        print("=" * 60)

        step = 0
        while step < self.max_steps:
            step += 1
            print(f"\n--- Step {step} ---")

            try:
                t_call = _time.time()
                use_thinking = 'flash' in self.model.lower() or 'v3' in self.model.lower() or 'r1' in self.model.lower()
                kwargs = {
                    'model': self.model,
                    'messages': self.messages,
                    'tools': [{"type": "function", "function": f} for f in TOOLS_FUNCTIONS],
                    'tool_choice': 'auto',
                }
                if use_thinking:
                    kwargs['extra_body'] = {"thinking": {"type": "enabled"}}
                response = self.client.chat.completions.create(**kwargs)
                elapsed_ms = (_time.time() - t_call) * 1000
            except Exception as e:
                print(f"❌ API 调用失败: {e}")
                break

            msg = response.choices[0].message

            # ── LLM 调用日志 ──
            _log_llm_call(response, elapsed_ms, self.messages)

            # 检查是否有 tool_calls
            if not msg.tool_calls:
                # 没有工具调用 → 最终回复
                content = msg.content or ''
                print(f"✅ Agent 完成: {content[:300]}...")
                final_msg = {"role": "assistant", "content": content}
                if hasattr(msg, 'reasoning_content') and msg.reasoning_content:
                    final_msg['reasoning_content'] = msg.reasoning_content
                self.messages.append(final_msg)
                break

            # 处理工具调用
            for tc in msg.tool_calls:
                name = tc.function.name
                args_str = tc.function.arguments
                try:
                    args = json.loads(args_str) if args_str else {}
                except json.JSONDecodeError:
                    args = {}

                print(f"  🔧 {name}({json.dumps(args, ensure_ascii=False)[:120]})")

                result = execute_tool(name, args)
                print(f"     → {result[:200]}...")

                # 添加到消息历史（保留 reasoning_content 以兼容 DeepSeek thinking 模式）
                assistant_msg = {
                    "role": "assistant",
                    "content": None,
                    "tool_calls": [{
                        "id": tc.id,
                        "type": "function",
                        "function": {"name": name, "arguments": args_str},
                    }],
                }
                if hasattr(msg, 'reasoning_content') and msg.reasoning_content:
                    assistant_msg['reasoning_content'] = msg.reasoning_content
                self.messages.append(assistant_msg)
                self.messages.append({
                    "role": "tool",
                    "tool_call_id": tc.id,
                    "content": result,
                })

        # 完成
        print("\n" + "=" * 60)
        print(f"  ✅ 完成：{step} 步")
        if _chart_files:
            print(f"  📊 生成图表: {len(_chart_files)} 个")
            for cf in _chart_files:
                print(f"     - {cf}")
        print("=" * 60)


# ═══════════════════════════════════════════════════════════
# 6. CLI 入口
# ═══════════════════════════════════════════════════════════

if __name__ == '__main__':
    agent = TemplateFillerAgent()
    agent.run()
