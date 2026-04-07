# Design First, Code Later: Aesthetically Pleasing Template-Free Slides Generation

> ACL 2026 Findings — 官方实现（**DeepSlides**）

**DeepSlides** 是一个 AI 驱动的演示文稿自动生成系统。输入一个主题（以及可选的图片或风格描述），即可自动产出完整排版的 `.pptx` 文件。它在 [Open Deep Research](https://github.com/langchain-ai/open_deep_research) 架构基础上扩展了研究、写作、幻灯片规划和视觉设计全流程。

[English README](README.md)

---

## 工作原理

DeepSlides 采用图（Graph）驱动的规划-执行工作流，核心逻辑在 `src/open_deep_research/graph.py` 中实现：

```
输入（主题 / 图片 / 风格）
  │
  ▼
[图像分析]       ← 可选：通过视觉模型理解图片所含的研究意图
  │
  ▼
[报告规划]       ← Planner LLM 生成章节结构和搜索查询
  │
  ▼
[研究与写作]     ← 并行：每个章节独立完成网络搜索 + 内容写作
  │
  ▼
[PPT 规划]       ← 按章节分配幻灯片数量，生成幻灯片大纲
  │
  ▼
[幻灯片生成]     ← 并行：每张幻灯片 → 内容扩写 → 布局设计 → PPTX 渲染
  │
  ▼
[评分与优化]     ← LLM 评分（设计/美观/完整度）；低分幻灯片自动重试
  │
  ▼
[封面 / 章节页 / 结尾页 生成]
  │
  ▼
输出：presentation.pptx（可选通过 LibreOffice 导出 PNG 预览）
```

---

## 核心特性

- **全流程自动化** — 输入主题字符串，输出 `.pptx`，无需手动编辑。
- **多模型支持** — 规划、写作、编码、设计四个角色可独立配置不同 LLM。支持 OpenAI、Azure OpenAI、Anthropic Claude 及任何 OpenAI 兼容接口。
- **多搜索引擎** — 支持 Tavily、Perplexity、Exa、DuckDuckGo、arXiv、PubMed、Google Search、LinkUp。
- **图片输入** — 支持传入图片或截图，系统通过视觉模型分析图片内容，自动推断研究方向。
- **风格控制** — 传入风格描述、配色方案或模板路径，Designer 模型会在所有幻灯片上保持一致的视觉风格。
- **LLM 自动评分** — 每张幻灯片从设计、美观、完整度三个维度打分，低分幻灯片自动重新生成。
- **自动生成封面 / 章节页 / 结尾页** — 独立 LLM 调用生成这些框架页面，风格与正文保持统一。
- **并行执行** — 同一章节内的幻灯片通过 LangGraph 的 `Send()` API 并发生成。

---

## 关键文件

| 文件 | 用途 |
|---|---|
| `src/open_deep_research/graph.py` | 主工作流图（所有节点和边） |
| `src/open_deep_research/configuration.py` | 所有配置参数及默认值 |
| `src/open_deep_research/prompts.py` | 各阶段使用的 Prompt 模板 |
| `src/open_deep_research/state.py` | TypedDict / Pydantic 状态结构定义 |
| `src/open_deep_research/utils.py` | 搜索后端、图像描述等工具函数 |
| `src/open_deep_research/run.py` | 批量运行入口（CSV 输入） |

---

## 快速上手

### 1. 克隆并安装

```bash
git clone <your-repo-url>
cd multimodal_open_deep_research

python -m venv .venv
source .venv/bin/activate   # Windows: .venv\Scripts\activate
pip install -e .
```

如使用本地 `pptx_tools` 辅助包（推荐，用于 PPT 渲染）：

```bash
cd pptx_tools && pip install -e . && cd -
```

### 2. 配置环境变量

复制示例配置文件并填入密钥：

```bash
cp .env.example .env
```

最低要求：

```dotenv
# LLM 提供商
OPENAI_API_KEY=sk-...
OPENAI_API_BASE=https://api.openai.com/v1   # 或你的代理 / Azure 端点

# 搜索引擎（至少配置一个）
TAVILY_API_KEY=tvly-...

# 可选：为 Designer 模型配置独立密钥/端点
DESIGNER_API_BASE=https://...
```

其他支持的变量：`AZURE_OPENAI_API_VERSION`、`AZURE_OPENAI_DEPLOYMENT`、`CODER_API_BASE`、`EXA_API_KEY`、`GOOGLE_API_KEY`、`GOOGLE_CX`、`LANGCHAIN_API_KEY`。

### 3. 运行

**方式 A — LangGraph 开发服务器（交互式 UI）**

```bash
langgraph dev --no-reload
```

启动后访问 `http://localhost:8123` 打开 LangGraph Studio。可以在 UI 中提交输入、查看图状态、逐节点回放执行过程。

![LangGraph Studio](web.png)

左侧面板显示实时工作流图，右侧面板显示各节点的执行日志。在底部输入表单中填写 **Topic**、**Presentation Minutes**、**Style** 等字段，点击 **Submit** 即可启动。

**方式 B — 批量运行（CSV 输入）**

```bash
python src/open_deep_research/run.py
```

在 `run.py` 中修改 `csv_path`、`start`、`max_rows` 参数指向你的输入文件。CSV 文件至少需要 `Topic` 列，可选列：`image_path`、`style`、`presentation_minutes`。

---

## 配置参数

所有参数定义在 `configuration.py` 的 `Configuration` 类中，关键字段如下：

| 参数 | 默认值 | 说明 |
|---|---|---|
| `planner_model` | `openai/gpt-4o-mini-2024-07-18` | 报告大纲生成模型 |
| `writer_model` | `openai/gpt-4o-mini-2024-07-18` | 章节内容写作模型 |
| `coder_model` | `anthropic/claude-haiku-4.5` | PPTX 代码生成模型 |
| `designer_model` | `anthropic/claude-haiku-4.5` | 幻灯片布局设计模型 |
| `search_api` | `tavily` | 搜索引擎（`tavily`、`exa`、`duckduckgo`、`arxiv`、`pubmed`、`googlesearch`、`perplexity`、`linkup`） |
| `max_search_depth` | `5` | 每个章节的最大搜索迭代次数 |
| `number_of_queries` | `2` | 每个章节的搜索查询数 |
| `number_of_queries_for_ppt` | `1` | 幻灯片内容扩写阶段的额外搜索查询数 |

通过 `configurable` 字典可覆盖任意参数：

```python
from langchain_core.runnables import RunnableConfig

config = RunnableConfig(configurable={
    "planner_model": "openai/gpt-4o",
    "search_api": "exa",
    "max_search_depth": 3,
})
result = asyncio.run(graph.ainvoke(state, config=config))
```

---

## 依赖

核心依赖（完整列表见 `pyproject.toml`）：

- `langgraph` — 图执行引擎
- `langchain-openai`、`langchain-anthropic` — LLM 集成
- `python-pptx` — PPTX 渲染（通过 `pptx_tools`）
- `tavily-python`、`exa-py`、`duckduckgo-search`、`linkup-sdk` — 搜索后端
- `google-cloud-vision` — 可选，用于图像内容识别
- `langsmith` — 可选，用于链路追踪

可选系统依赖（PNG 导出）：

```bash
# macOS
brew install libreoffice
# Ubuntu
sudo apt install libreoffice
```

---

## 隐私与安全

- **不要提交 `.env` 文件。** 它已被 `.gitignore` 默认排除。
- 所有 API 密钥均从环境变量读取，源码中没有任何硬编码密钥。
- 可通过设置 `ABLATE_DESIGN=1` 或 `ABLATE_SCORING=1` 环境变量，在消融实验中关闭评分/优化步骤。

---

## 开源协议

MIT

---

## 引用

如果本工作对你有帮助，请引用我们的论文：

```bibtex
@inproceedings{cui2026design,
  title     = {Design First, Code Later: Aesthetically Pleasing Template-Free Slides Generation},
  author    = {Cui, Zhiyao and Wang, Chenxu and Hu, Shuyue and Zhang, Yiqun and Shao, Wenqi and Zhang, Qiaosheng and Wang, Zhen},
  booktitle = {Findings of the Association for Computational Linguistics: ACL 2026},
  year      = {2026}
}
```
