# 停车明细分析报告 — Agentic Report Generation System

## 快速启动

```bash
git clone <repo> && cd backend2
cp .env.example .env          # 编辑填入 KEY / BASE_URL / FLASH_MODEL
pip install -r requirements.txt
uvicorn server:app --host 127.0.0.1 --port 8000 --reload
```

浏览器 → `http://localhost:8000`

---

## 项目结构

```
backend2/
├── server.py              # FastAPI 入口
├── core_agent.py           # Agent 引擎（16 工具 + LLM 循环）
├── docx2json.py            # DOCX → JSON 忠实转写
├── json_to_docx.py         # JSON → DOCX 忠实还原
├── job_manager.py          # Job 状态机（文件持久化）
├── logging_utils.py        # 结构化 JSONL 日志
├── data.csv                # 示例数据（Assignment 交付物）
├── template.docx           # 示例模板（Assignment 交付物）
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

## API

| 方法 | 路径                        | 说明                          |
| ---- | --------------------------- | ----------------------------- |
| POST | `/api/jobs`               | 上传 template.docx + data.csv |
| GET  | `/api/jobs/{id}`          | 任务状态                      |
| GET  | `/api/jobs/{id}/download` | 下载 .docx                    |

---

## 设计概述

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

### 文件持久化（每个 job 独立目录）

```
jobs/{job_id}/
├── state.json              # status / progress / output_file
├── input/                  # 用户上传的原文件
├── output/
│   ├── filled_template.json
│   └── report.docx
├── charts/                 # 图表 PNG
└── llm_calls.jsonl         # 完整 LLM 调用日志
```

### 日志

| 文件                          | 内容                                                                   |
| ----------------------------- | ---------------------------------------------------------------------- |
| `requests.jsonl`            | 每次请求生命周期（job_created → started → completed/failed）         |
| `jobs/{id}/llm_calls.jsonl` | LLM 调用：prompt / response / reasoning_content / tool_calls / latency |

---

## 部署

```bash
git clone <repo> && cd backend2
cp .env.example .env
docker-compose up --build -d
```

代码更新：`git pull` → uvicorn `--reload` 自动检测文件变化并重启。

---

## CI/CD

```
git push main → GitHub Actions → pytest 8/8 ✅
                                    │
服务器 git pull → uvicorn --reload → 自动重启
```

**设计思路**：GitHub 只做代码托管 + 测试。镜像构建在服务器本地完成（`docker-compose up --build`），不依赖外部镜像仓库。代码挂载 + `--reload` 实现热更新，git pull 后无需手动重启。

如需展示完整 CI/CD（含镜像推送），可扩展 `.github/workflows/ci.yml` 添加 `docker build && push` step。

---

## 重试设计（已预留）

状态机已开放 `failed → running` 转换。输入文件保留、所有写操作覆盖写（幂等）、Agent 确定性执行。补 `POST /api/jobs/{id}/retry` 即可启用。

---

## 测试

```bash
python -m pytest tests/test_api.py -v   # 8 passed
```

---

|  |  |  |
| - | - | - |
