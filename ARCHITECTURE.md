# 停车报告 Agent 系统 — 文件架构与说明

## 项目总览

本系统是一个 **Agentic Report Generation** 后端服务：用户上传 Word 模板 + CSV 数据，LLM 驱动的 Agent 自动分析数据、填充模板、生成图表，最终输出排版完整的 `.docx` 报告。

**运行模式**：单进程同步多线程 Web 服务。FastAPI 接收请求 → `threading.Thread` 后台执行 Agent → LLM function-calling loop 驱动工具调用 → 输出报告。

---

## 文件索引

### 核心服务层

| 文件 | 行数 | 角色 |
|---|---|---|
| [`server.py`](#serverpy) | ~175 | **FastAPI 入口**。3 个 API 端点 + 前端路由 + 后台线程调度 + 并发控制 |
| [`core_agent.py`](#core_agentpy) | 2142 | **Agent 引擎**。14 个工具函数 + System Prompt + LLM function-calling 主循环 + TemplateFillerAgent 类 |
| [`job_manager.py`](#job_managerpy) | 150 | **Job 状态机**。文件级持久化：创建/状态转换/文件存取/日志追加，零外部依赖 |

### 文档转换层

| 文件 | 行数 | 角色 |
|---|---|---|
| [`docx2json.py`](#docx2jsonpy) | 498 | **DOCX → JSON 忠实转写**。解析 OOXML，输出包含 blocks/styles/numbering 的完整 JSON |
| [`json_to_docx.py`](#json_to_docxpy) | 572 | **JSON → DOCX 忠实还原**。将填充后的 JSON 还原为 .docx，保留全部原始格式 |

### 日志与工具

| 文件 | 行数 | 角色 |
|---|---|---|
| [`logging_utils.py`](#logging_utilspy) | 55 | **结构化 JSONL 日志**。`log_request()` 记录请求生命周期事件到 `requests.jsonl` |
| [`templates/index.html`](#templatesindexhtml) | 185 | **前端单页**。文件上传 + 任务列表 + 状态轮询 + 下载，纯 HTML/CSS/JS |

### 配置与部署

| 文件 | 角色 |
|---|---|
| [`.env.example`](#envexample) | 环境变量模板：API KEY / BASE_URL / FLASH_MODEL |
| [`requirements.txt`](#requirementstxt) | Python 依赖清单：FastAPI / OpenAI SDK / python-docx / matplotlib 等 8 项 |
| [`Dockerfile`](#dockerfile) | 生产镜像：Python 3.12-slim + 中文字体 + 阿里云镜像源 |
| [`docker-compose.yml`](#docker-composeyml) | 开发编排：代码热挂载 + `--reload` 热重载 + jobs 目录持久化 |
| [`.github/workflows/ci.yml`](#githubworkflowscipyml) | CI 流水线：`pytest tests/test_api.py -v`（8 个测试用例） |

### 测试与示例数据

| 文件 | 角色 |
|---|---|
| [`tests/test_api.py`](#teststest_apipy) | API 集成测试（8/8 通过）。Mock Agent 线程，验证 submit → status → download 全链路 |
| [`data.csv`](#datacsv) | 示例停车明细数据（307 KB），用于开发调试 |
| [`template.docx`](#templatedocx) | 示例报告模板（47 KB），含占位符和格式标记 |

### 运行时目录

| 路径 | 角色 |
|---|---|
| `jobs/` | Job 运行时目录（`.gitignore` 排除）。每个 job 一个子目录，含 input/output/charts 和日志 |
| `charts/` | 全局图表输出目录（Docker 内创建） |
| `requests.jsonl` | 全局请求日志（`.gitignore` 排除） |

### 其他

| 文件 | 角色 |
|---|---|
| [`README.md`](#readmemd) | 项目说明：快速启动 / API / 设计概述 / 部署 / CI/CD |
| `LICENSE` | 开源许可证（GPL-3.0） |

---

## 各文件详解

### `server.py`

**FastAPI Web 服务入口**，监听 `0.0.0.0:8000`。

**API 端点**：

| 方法 | 路径 | 功能 |
|---|---|---|
| GET | `/` | 返回前端 HTML 页面 |
| POST | `/api/jobs` | 上传 `template.docx` + `data.csv`，创建异步生成任务 |
| GET | `/api/jobs/{job_id}` | 查询任务状态（status / progress / error） |
| GET | `/api/jobs/{job_id}/download` | 下载生成的 `.docx` 报告 |

**请求处理流程**：

```
POST /api/jobs
  → 校验文件类型 (.docx + .csv)
  → JobManager 创建 job，保存输入文件
  → BoundedSemaphore 并发检查（超过 23 并发 → 503）
  → threading.Thread(daemon=True) 启动后台任务
  → 立即返回 {"job_id": "...", "status": "created"}
```

**后台任务 `_run_generation()`**：

```
Step 1: docx2json → template_full.json
Step 2-7: core_agent.run() → LLM function-calling loop
兜底: 若 Agent 未调 restore_to_docx，从 filled_template.json 还原
完成: 注册 output_file → 状态转 completed
```

**并发模型**：`BoundedSemaphore(MAX_CONCURRENT_JOBS + 20)`，默认允许最多 23 个线程同时运行。非阻塞获取，超出上限返回 503。

**代码问题**：[`server.py:163`](server.py:163) 存在残留死代码 `AgentClass.setup = _setup_with_logging`，`AgentClass` 和 `_setup_with_logging` 均未定义。[`server.py:155`](server.py:155) 和 [`server.py:163`](server.py:163) 存在重复的 `_get_agent_steps` 函数定义。

---

### `core_agent.py`

**Agent 核心引擎**（2142 行），手写 OpenAI-compatible function-calling 循环。

**架构分层**：

```
工具层（tool_* 函数）→ 引擎层（execute_tool + TOOL_MAP）→ 输出层（save + restore）
                        ↑
              System Prompt（工作流程定义）
                        ↑
              TemplateFillerAgent（主循环）
```

**14 个工具函数**：

| 类别 | 工具 | 功能 |
|---|---|---|
| 探索 | `tool_get_template_structure` | 遍历所有 blocks，返回 index + 格式标记 + 文本预览 |
| 探索 | `tool_get_placeholder_blocks` | 列出所有红色/斜体/含占位符的 run |
| 探索 | `tool_get_placeholder_rules` | 返回占位符推断规则 |
| 数据 | `tool_load_csv_summary` | CSV 列名、行数、前 5 行预览 |
| 数据 | `tool_compute_all_kpis` | 一键计算所有关键指标 |
| 数据 | `tool_compute_statistics` | 对指定列计算统计值（sum/avg/max/min/count） |
| 分析 | `tool_analyze_payment_distribution` | 支付方式分布 |
| 分析 | `tool_analyze_parking_duration` | 停车时长分布（分段统计） |
| 时间 | `tool_get_date_range` | 数据的起止日期 |
| 时间 | `tool_get_current_time` | 当前时间（用于报告生成时间） |
| 图表 | `tool_generate_chart` | 通用图表生成（柱状图/饼图/折线图） |
| 图表 | `tool_insert_chart_block` | 在指定位置插入图表 |
| 填充 | `tool_replace_paragraph_text` | 整段文本替换（自动去红） |
| 填充 | `tool_fill_paragraph_run` | 段落内单个 run 精确替换 |
| 填充 | `tool_fill_table_cell` | 表格单元格原子替换 |
| 填充 | `tool_fill_kpi_table_all` | 自动定位指标表并批量填充 |
| 样式 | `tool_set_block_style` | 修正段落样式 |
| 清理 | `tool_delete_blocks` | 批量删除指令块 |
| 校验 | `tool_verify_report` | 保存前自检：红色残留/样式/空段/连续标题 |
| 输出 | `tool_save_filled_json` | 保存填充后的 JSON |
| 输出 | `tool_restore_to_docx` | JSON → DOCX 还原 |

**Agent 主循环** (`TemplateFillerAgent.run()`)：

```
while step < 50:
    response = client.chat.completions.create(
        model=..., messages=..., tools=[21 tools], tool_choice='auto'
    )
    if no tool_calls → break（Agent 认为完成）
    for each tool_call:
        result = execute_tool(name, args)
        append assistant message + tool result to history
```

**System Prompt 核心约束**：
- 从后往前填充（避免 index 偏移）
- 图表和删除操作最后做
- 保存前必须 `verify_report` 通过
- 禁止自行添加模板中没有的内容

**LLM 日志**：通过 `_log_llm_call()` 记录每次调用的完整上下文快照（messages / response / reasoning_content / latency / token usage）到 `jobs/{id}/llm_calls.jsonl`。

---

### `job_manager.py`

**Job 状态机**，完全基于文件系统，零外部依赖。

**状态转换**：

```
created → running → completed
                  → failed → running（允许重试）
```

**目录结构**（每个 Job）：

```
jobs/{job_id}/
├── state.json          # {"status","progress","steps","created_at","updated_at","input_files","output_file","error"}
├── input/              # template.docx, data.csv, template_full.json
├── output/             # filled_template.json, report.docx
├── charts/             # 图表 PNG
├── llm_calls.jsonl     # LLM 调用日志（每行一个 JSON）
└── messages.jsonl      # Agent 对话历史（可选）
```

**核心方法**：

| 方法 | 功能 |
|---|---|
| `create(instructions)` | 创建 job 目录 + 初始化 state.json |
| `save_input(job_id, filename, bytes)` | 保存用户上传文件 |
| `transition(job_id, new_status)` | 状态转移（校验合法性） |
| `set_progress(job_id, progress, steps)` | 更新进度 |
| `set_error(job_id, error)` | 标记失败 |
| `append_llm_log(job_id, entry)` | 追加 LLM 调用日志 |
| `get_state(job_id)` | 读取状态快照 |

**设计考量**：单进程场景下文件操作无竞态。多进程需引入文件锁（`fcntl`/`msvcrt`）。

---

### `docx2json.py`

**DOCX → JSON 忠实转写**。将 `.docx` 解析为一份完整的、自描述的 JSON 结构。

**输出结构**：

```json
{
  "page": { "width": 11906, "height": 16838, "margins": {...} },
  "styles": [{ "id": "1", "name": "heading 1", ... }],
  "numbering": [...],
  "blocks": [
    { "index": 0, "type": "paragraph", "style": "1", "runs": [...] },
    { "index": 1, "type": "table", "rows": [...] }
  ]
}
```

**设计原则**：
- 每个 block 有唯一 `index`（body-children 顺序）
- 样式分三层：docDefaults → style 定义 → run 显式覆盖
- 属性名直译 OOXML，不做"友好化"转换
- 缺省值用 `null`，明确区分"未设置"和"设为默认值"

---

### `json_to_docx.py`

**JSON → DOCX 忠实还原**。将 Agent 填充后的 `filled_template.json` 还原为格式完整的 `.docx`。

**支持还原的格式**：
- 段落/表格/图片（含图表输出）
- 页面设置、页边距
- 字体、字号、颜色、粗斜体、下划线
- 单元格底色、边框、边距
- 项目编号（numbering）
- 图片嵌入（从本地路径）

---

### `logging_utils.py`

**结构化 JSONL 日志模块**。

**`log_request(event, job_id, **extra)`**：
- 记录请求生命周期事件到 `requests.jsonl`
- 同时输出到 stdout（Docker logs 可见）
- 日志格式：`{"ts":"...", "event":"job_created", "job_id":"...", ...}`

**事件类型**：`job_created` → `job_progress` → `job_started` → `job_completed` / `job_failed` → `job_downloaded`

---

### `templates/index.html`

**前端单页应用**（185 行），纯 HTML/CSS/JS，零框架依赖。

**功能**：
- 上传 `template.docx` + `data.csv` + 可选文字指令
- 实时任务列表（轮询 `GET /api/jobs/{id}` 每 2 秒）
- 任务状态卡片：等待中 / 运行中 / 已完成 / 失败
- 已完成任务一键下载
- 并发提交多个任务

---

### `.env.example`

```
KEY=your_api_key              # DeepSeek API Key
BASE_URL=https://api.deepseek.com/v1
FLASH_MODEL=deepseek-chat     # 默认模型
```

`docker-compose.yml` 通过 `env_file: - .env` 注入环境变量。Agent 内部通过 `os.environ.get('KEY')` 等读取。

---

### `requirements.txt`

```
fastapi>=0.109.0              # Web 框架
uvicorn>=0.25.0               # ASGI 服务器
python-multipart>=0.0.6       # 文件上传解析
openai>=1.6.0                 # OpenAI SDK（兼容 DeepSeek API）
python-docx>=1.0.0            # DOCX 读写
matplotlib>=3.8.0             # 图表生成
lxml>=5.0.0                   # XML 解析（OOXML）
python-dotenv>=1.0.0          # .env 加载
```

---

### `Dockerfile`

**生产镜像**：`python:3.12-slim` 基础镜像，阿里云 Debian 源 + 阿里云 PyPI 源。

**关键层**：
1. 安装中文字体 `fonts-noto-cjk`（matplotlib 图表中文渲染）
2. `pip install -r requirements.txt`
3. 复制项目代码
4. 创建 `jobs/` 和 `charts/` 目录
5. 暴露 8000 端口

---

### `docker-compose.yml`

**开发编排**：

```yaml
services:
  app:
    build: .
    ports: ["8000:8000"]
    volumes:
      - .:/app              # 代码热挂载
      - ./jobs:/app/jobs    # 任务数据持久化
    env_file: .env
    command: uvicorn server:app --host 0.0.0.0 --port 8000 --reload --reload-dir /app
    restart: unless-stopped
```

**开发 vs 生产**：
- 开发：docker-compose 覆盖 CMD，启用 `--reload` 热重载
- 生产：Dockerfile 的 CMD `python server.py`，无热重载

---

### `.github/workflows/ci.yml`

**CI 流水线**（GitHub Actions）：

```
push main/master / PR
  → checkout + setup Python 3.12
  → 安装中文字体
  → pip install -r requirements.txt pytest
  → python -m pytest tests/test_api.py -v
```

仅做代码托管 + 自动化测试。镜像构建在服务器本地完成，不依赖外部镜像仓库。

---

### `tests/test_api.py`

**API 集成测试**（8 个用例，8 通过）。

| 测试 | 验证点 |
|---|---|
| `test_frontend_serves` | GET `/` 返回 HTML |
| `test_create_job_success` | POST 上传正确文件 → 200 + job_id |
| `test_create_job_rejects_wrong_type` | 上传 .txt 冒充 .docx → 400 |
| `test_create_job_rejects_wrong_csv` | 上传 .txt 冒充 .csv → 400 |
| `test_status_nonexistent` | GET 不存在的 job → 404 |
| `test_download_nonexistent` | GET 不存在的 job/download → 404 |
| `test_full_lifecycle` | submit → poll status → download 全链路 |
| `test_download_incomplete_rejected` | 未完成 job 下载 → 404 |

**Mock 策略**：session 级别 patch `server._run_generation`，Agent 线程被替换为同步状态写入，无需真实 LLM 调用。

---

### `data.csv`

**示例停车明细数据**（307 KB），包含停车记录的完整字段：入场时间、出场时间、车牌号、支付方式、费用等。用于开发调试和 `core_agent.py` 的 CLI 独立运行。

---

### `template.docx`

**示例报告模板**（47 KB），一份包含占位符的 Word 文档。占位符以红色斜体标记（如「【起始日期 – 结束日期】」），Agent 通过格式标记 + 语义理解来定位和填充。

---

### 运行时目录

| 路径 | 说明 |
|---|---|
| `jobs/` | Job 数据根目录。每个子目录是一个独立 job，内含 input/output/charts 和日志。`.gitignore` 排除 |
| `charts/` | 全局图表输出目录。Dockerfile 中 `mkdir -p charts` 创建 |
| `requests.jsonl` | 全局请求生命周期日志。`.gitignore` 排除 |
| `filled_*.json` / `filled_*.docx` | 填充输出文件。`.gitignore` 排除 |

---

## 数据流全景

```
用户浏览器                          服务器                              DeepSeek API
   │                                 │                                     │
   │  POST /api/jobs                 │                                     │
   │  (template.docx + data.csv)     │                                     │
   ├────────────────────────────────>│                                     │
   │                                 │  docx2json.py                       │
   │                                 │  template.docx → template_full.json │
   │                                 │                                     │
   │                                 │  core_agent.py                      │
   │                                 │  ┌─ get_template_structure()        │
   │                                 │  ├─ load_csv_summary()              │
   │                                 │  ├─ compute_all_kpis()              │
   │  {"job_id":"...","status":      │  ├─ analyze_*()                     │
   │   "created"}                    │  ├─ replace_paragraph_text()        │
   │<────────────────────────────────│  ├─ fill_kpi_table_all()            │
   │                                 │  ├─ generate_*_chart()              │
   │  GET /api/jobs/{id}             │  ├─ verify_report()                 │
   │  (轮询每 2 秒)                  │  └─ restore_to_docx()               │
   │<────────────────────────────────│                                     │
   │                                 │  ┌─────────────────────────────────>│
   │                                 │  │  chat.completions.create()       │
   │                                 │  │  (function-calling loop)         │
   │                                 │  │  ← 每次返回 tool_calls 或 finish │
   │                                 │  │  → 每次发送 tool result         │
   │                                 │  └─ 最多 50 轮                      │
   │                                 │                                     │
   │  GET /api/jobs/{id}/download    │                                     │
   │  report.docx                    │                                     │
   │<────────────────────────────────│                                     │
```

---

## 关键设计决策

| 决策 | 理由 |
|---|---|
| **不用 LangChain / AutoGen** | 手写 function-calling loop 更可控，调试透明，无框架版本冲突 |
| **文件持久化，不用数据库** | 单实例部署，零运维依赖。每个 job 一个目录，天然隔离 |
| **DOCX → JSON → DOCX** | Agent 操作结构化 JSON，比直接操作 OOXML 简单可靠。`docx2json`/`json_to_docx` 保证格式无损 |
| **threading.Thread 而非 asyncio** | Agent 内部是阻塞 LLM 调用，用线程比 async 更简单。FastAPI 同步端点不阻塞事件循环 |
| **BoundedSemaphore(3+20)** | 实际允许 ~23 并发，低负载够用。高负载应改为 Semaphore(3) + Queue |
| **从后往前填充** | 避免 `replace_paragraph_text` 操作后 index 偏移影响后续填充 |
| **Agent 自行判断 block 角色** | 格式标记仅作线索，语义判断由 LLM 完成，灵活度远高于规则引擎 |
