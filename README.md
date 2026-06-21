# 停车明细分析报告 — Agentic Report Generation System

## 快速启动

```bash
git clone <repo> && cd backend2
cp .env.example .env          # 编辑填入 KEY / BASE_URL / PRO_MODEL
pip install -r requirements.txt
python server.py
```

浏览器 → `http://localhost:8000`

---

## 项目结构

```
backend2/
├── server.py              # FastAPI 入口（asyncio + subprocess）
├── agent_worker.py         # Job 执行进程（独立 Python 子进程）
├── core_agent.py           # Agent 引擎（16 工具 + LLM 循环）
├── docx2json.py            # DOCX → JSON 忠实转写
├── json_to_docx.py         # JSON → DOCX 忠实还原
├── job_manager.py          # Job 状态机（文件持久化）
├── logging_utils.py        # 结构化 JSONL 日志
├── data.csv                # 示例数据
├── template.docx           # 示例模板
├── templates/
│   └── index.html          # 前端单页
├── tests/
│   └── test_api.py         # API 测试 (8/8)
├── .github/
│   └── workflows/
│       └── ci.yml          # CI/CD 流水线
└── jobs/                   # 运行时创建（gitignored）
```

---

## 架构

### 并发模型：asyncio + subprocess 隔离

```
                         HTTP 请求
                            │
                            ▼
┌──────────────────────────────────────────────────┐
│  server.py (asyncio, 单进程)                      │
│                                                    │
│  POST /api/jobs  ──→  job_queue (asyncio.Queue)   │
│                           │                        │
│              ┌────────────┼────────────┐           │
│              ▼            ▼            ▼           │
│         Worker 1     Worker 2    (2 核 = 2 并发)  │
│              │            │                        │
│              ▼            ▼                        │
│    asyncio.create_subprocess_exec                  │
│              │            │                        │
└──────────────┼────────────┼────────────────────────┘
               │            │
               ▼            ▼
┌──────────────────────────────────────┐
│  agent_worker.py (独立进程)           │
│                                        │
│  1. docx2json: template.docx → JSON   │
│  2. core_agent: LLM 驱动填充          │
│  3. json_to_docx: JSON → report.docx  │
│  4. 更新 state.json                   │
│                                        │
│  每个 Job 独立进程 · 物理隔离 · 零竞态  │
└──────────────────────────────────────┘
```

**为什么用 subprocess 而不是 threading？**

旧架构每个请求启动一个 `threading.Thread`，所有线程共享 `core_agent` 模块的全局变量（`_template_data`、`_csv_rows` 等），通过 `del sys.modules` + 重新 `import` 来尝试隔离——但 `sys.modules` 操作**不是线程安全的**，并发时会出现数据串扰。改为子进程后，每个 Job 拥有独立的 Python 解释器和内存空间，彻底消除竞态。

### 三层职责分离

```
docx2json.py          Agent (core_agent.py)       json_to_docx.py
(忠实转写)            (语义决策，只改内容)          (忠实还原)
    │                      │                          │
    ▼                      ▼                          ▼
template.docx  →  template_full.json  →  filled_template.json  →  report.docx
```

### Agent 核心设计

- **LLM 语义判断**：`get_template_structure` 只提供原始事实（文本 + RED/italic/bracket 格式标签），不做预分类。LLM 通过阅读文本含义自行判断每个 block 的角色。
- **从后往前填充**：`replace_paragraph_text` / `fill_paragraph_run` 不改变 index，从文档末尾往前操作。`insert_chart_block` 和 `delete_blocks` 最后做。
- **原子替换**：`fill_table_cell` 将整格替换为单个 run，不留旧占位残渣。
- **保存前自检**：`verify_report` 扫描红色残留、样式错乱、连续标题等问题。

### Job 状态机

```
queued → running → completed
              ↘ failed → running (重试)
```

### 文件持久化（每个 job 独立目录）

```
jobs/{job_id}/
├── state.json              # status / progress / output_file
├── input/                  # 用户上传的原文件
│   ├── template.docx
│   ├── data.csv
│   └── template_full.json
├── output/
│   ├── filled_template.json
│   └── report.docx
├── charts/                 # 图表 PNG
└── llm_calls.jsonl         # 完整 LLM 调用日志
```

---

## API

| 方法 | 路径 | 说明 |
|------|------|------|
| POST | `/api/jobs` | 上传 template.docx + data.csv，创建任务 |
| GET | `/api/jobs/{id}` | 查询任务状态 |
| GET | `/api/jobs/{id}/download` | 下载生成的 .docx |

### POST /api/jobs

`multipart/form-data`:
- `template` — .docx 模板文件（必填）
- `data` — .csv 数据文件（必填）
- `instructions` — 用户指令文本（可选）

响应：
```json
{"job_id": "a1b2c3d4e5f6", "status": "queued"}
```

队列满时返回 `503 系统繁忙，请稍后重试`。

### GET /api/jobs/{job_id}

```json
{
  "job_id": "a1b2c3d4e5f6",
  "status": "running",
  "progress": 3,
  "steps": 7,
  "created_at": "2026-01-15T10:30:00Z",
  "updated_at": "2026-01-15T10:32:00Z",
  "error": null
}
```

---

## 并发配置

针对 **2 核 / 4 GB** 服务器优化，通过 `.env` 环境变量调节：

| 变量 | 默认值 | 说明 |
|------|--------|------|
| `MAX_CONCURRENT_JOBS` | 2 | 同时运行的 worker 数（匹配 CPU 核数） |
| `MAX_QUEUE_DEPTH` | 6 | 最大排队任务数，超限返回 503 |
| `JOB_TIMEOUT` | 600 | 单个 Job 超时秒数（10 分钟） |
| `SHUTDOWN_GRACE_SECONDS` | 30 | 关闭时等待子进程结束的宽限秒数 |

**内存估算**（每个 worker 子进程）：
- Python 解释器 + 基础模块：~80 MB
- matplotlib + 中文字体：~60 MB
- 模板/CSV 数据驻留：~50-200 MB（取决于数据量）
- LLM API 客户端：~20 MB
- **合计：~200-400 MB / worker**

`2 并发 × 400 MB + server 主进程 ~100 MB ≈ 900 MB`，在 4 GB 总量下安全。

---

## 部署

```bash
git clone <repo> && cd backend2
cp .env.example .env          # 编辑填入 API 密钥和配置
docker compose up --build -d
```

代码更新：
```bash
git pull
docker compose restart        # 优雅关闭 → 等待子进程 → 重启
```

---

## 日志

| 文件 | 内容 |
|------|------|
| `requests.jsonl` | 每次请求生命周期（job_created → queued → started → completed / failed / timeout） |
| `jobs/{id}/llm_calls.jsonl` | LLM 调用明细：prompt / response / reasoning_content / tool_calls / latency |

---

## 测试

```bash
python -m pytest tests/test_api.py -v   # 8 passed
```

---

## 重试设计

状态机已开放 `failed → running` 转换。输入文件保留、所有写操作覆盖写（幂等）、Agent 确定性执行。补 `POST /api/jobs/{id}/retry` 即可启用。

---

## CI/CD

```
git push main → GitHub Actions → pytest 8/8 ✅
                                    │
服务器 git pull → docker compose restart → 自动更新
```
