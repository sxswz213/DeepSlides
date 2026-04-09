# Design First, Code Later: Aesthetically Pleasing Template-Free Slides Generation

> ACL 2026 Findings &nbsp;|&nbsp; [📄 论文](paper.pdf)

![DeepSlides case overview](case.png)

**DeepSlides** 是一个层次化的"设计优先"演示文稿自动生成框架。它将*幻灯片级别的风格设计*与*页面级别的代码实现*显式解耦，无需依赖任何预设模板，即可生成视觉连贯、美观专业的 `.pptx` 文件。输入一个主题（以及可选的图片或风格描述），系统自动完成深度网络调研、内容写作和幻灯片生成。

[English README](README.md)

---

## 核心成果

- 🏆 **Top-1 人类偏好率 76.5%**，超过所有开源基线方法，领先第二名（9.5%）**68.0 个百分点**
- 🏆 对比商业系统（Kimi、Manus）**胜率 52.0%**，领先第二名（19.5%）**32.5 个百分点**
- 📊 VLM 评判综合得分最高（**3.78**），在布局、层次、配色、清晰度、连贯性等维度全面领先
- 🎨 人工评估中，清晰度与结构得分 **4.25–4.29**，视觉设计与美观性得分 **3.84–3.98**

---

## 工作原理

DeepSlides 采用两级架构：

### 幻灯片级设计（Slides-level Design）
系统首先为整套演示文稿确定全局视觉风格——色调、配色方案、字体颜色、装饰形状及布局多样性指导原则。页面分为**功能页**（封面、章节分隔页、结尾页）和**内容页**，分别生成定制化的风格指令。

### 页面级生成（Page-level Generation）
每张幻灯片经过三个阶段生成：

```
[内容扩写]        ← 针对每张幻灯片进行网络搜索，检索支撑文本和图片
      │
      ▼
[设计]            ← 三层设计架构：
                      背景层（纹理、装饰元素）
                      布局层（内容块排列、空间位置）
                      内容层（精确文字片段和图片）
      │
      ▼
[实现]            ← LLM Coder 将设计规格转译为可执行的 Python/PPTX 代码
      │
      ▼
[评估与优化]      ← 从完整度、合规性、美观性三维度打分
                    低分幻灯片接收针对性反馈并迭代修改
```

完整 `graph.py` 流水线：

```
输入（主题 / 图片 / 风格）
  │
  ▼
[图像分析]              ← 可选：通过视觉模型理解图片所含的研究意图
  │
  ▼
[报告规划]              ← Planner LLM 生成章节结构和搜索查询
  │
  ▼
[研究与写作]            ← 并行：每个章节独立完成网络搜索 + 内容写作
  │
  ▼
[幻灯片级风格设计]      ← 全局色调、配色、形状、字体
  │
  ▼
[封面 / 章节页 / 结尾页]← 功能页单独生成，风格与正文保持一致
  │
  ▼
[页面级生成]            ← 并行：每张幻灯片 → 扩写 → 设计 → 编码 → 评估 → 优化
  │
  ▼
输出：presentation.pptx（可选通过 LibreOffice 导出 PNG 预览）
```

---

## 核心特性

- **无模板生成** — 不依赖任何预设布局，每张幻灯片的结构根据内容从零设计，避免视觉疲劳。
- **设计与实现解耦** — 专用 *Designer* 模块在高层语义设计空间中推理；专用 *Coder* 模块将设计规格转译为稳定的 PPTX 代码，限制错误传播，保留美学意图。
- **深度调研集成** — 基于 [Open Deep Research](https://github.com/langchain-ai/open_deep_research)，自动检索并整合网页内容、图片和学术资料。
- **三维度评估** — 每张幻灯片从*完整度*、*合规性*、*美观性*三个维度打分，低分幻灯片在汇总前迭代优化。
- **多模型支持** — 规划、写作、设计、编码四个角色可分别配置不同 LLM。支持 OpenAI、Azure OpenAI、Anthropic Claude 及任何 OpenAI 兼容接口。
- **多搜索引擎** — 支持 Tavily、Perplexity、Exa、DuckDuckGo、arXiv、PubMed、Google Search、LinkUp。
- **图片输入** — 支持传入图片或截图，视觉模型自动分析内容推断研究方向。
- **并行执行** — 同一章节内的幻灯片通过 LangGraph 的 `Send()` API 并发生成。

---

## 示例输出

![Full case comparison](full_case.png)

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
git clone https://github.com/sxswz213/DeepSlides
cd DeepSlides

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

## 开源协议

MIT

---

## 引用

如果本工作对你有帮助，请引用我们的论文（[PDF](paper.pdf)）：

```bibtex
@inproceedings{cui2026design,
  title     = {Design First, Code Later: Aesthetically Pleasing Template-Free Slides Generation},
  author    = {Cui, Zhiyao and Wang, Chenxu and Hu, Shuyue and Zhang, Yiqun and Shao, Wenqi and Zhang, Qiaosheng and Wang, Zhen},
  booktitle = {Findings of the Association for Computational Linguistics: ACL 2026},
  year      = {2026}
}
```
