<div align="center">

# Design First, Code Later
### Aesthetically Pleasing Template-Free Slides Generation

[![ACL 2026 Findings](https://img.shields.io/badge/ACL-2026%20Findings-blue?style=flat-square)](paper.pdf)
[![License: MIT](https://img.shields.io/badge/License-MIT-yellow?style=flat-square)](LICENSE)
[![Python 3.10+](https://img.shields.io/badge/Python-3.10%2B-blue?style=flat-square&logo=python&logoColor=white)](https://www.python.org/)
[![Paper](https://img.shields.io/badge/📄_论文-PDF-red?style=flat-square)](paper.pdf)
[![English](https://img.shields.io/badge/文档-English-orange?style=flat-square)](README.md)

<br/>

*层次化、设计优先的演示文稿生成框架 — 将幻灯片风格设计与页面代码实现解耦，无需任何预设模板。*

<br/>

![DeepSlides case overview](case.png)

</div>

---

## 核心成果

| | |
|:---:|:---|
| 🏆 | Top-1 人类偏好率 **76.5%**，超过所有开源基线，领先第二名（9.5%）**68.0 个百分点** |
| 🏆 | 对比商业系统（Kimi、Manus）胜率 **52.0%**，领先第二名（19.5%）**32.5 个百分点** |
| 📊 | VLM 评判综合得分最高 **3.78**（布局、层次、配色、清晰度、连贯性全面领先） |
| 🎨 | 人工评估：清晰度与结构 **4.25–4.29** · 视觉设计与美观性 **3.84–3.98** |

---

## 方法对比

| 方法 | 输出格式 | 完整演示 | 无模板 |
|:---|:---:|:---:|:---:|
| EvoPresent | HTML | ✅ | ❌ |
| AutoSlides | LaTeX/PDF | ✅ | ❌ |
| SlideGen / SlideCoder | PPTX | ❌ | ❌ |
| Kimi | PPTX/PDF | ✅ | ✅ |
| Manus | PPTX/PDF | ✅ | ✅ |
| **DeepSlides（本文）** | **PPTX/Image** | ✅ | ✅ |

---

## 工作原理

DeepSlides 采用两级架构：

**幻灯片级设计（Slides-level Design）** — 首先为整套演示文稿确定全局视觉风格：色调、配色方案、字体颜色、装饰形状及布局多样性指导。页面分为*功能页*（封面、章节分隔页、结尾页）和*内容页*，分别生成定制化风格指令。

**页面级生成（Page-level Generation）** — 每张幻灯片经过四个阶段生成：

```
内容扩写  →  设计  →  实现  →  评估与优化
（网络搜索）  （三层）  （Python/PPTX）  （完整度 · 合规性 · 美观性）
```

每张幻灯片的三层设计：
- **背景层** — 纹理、装饰元素、图案
- **布局层** — 内容块排列、空间位置、结构边界
- **内容层** — 精确文字片段、图片、视觉元素

完整 `graph.py` 流水线：

```
输入（主题 / 图片 / 风格）
  │
  ├─[图像分析]          可选：通过视觉模型理解图片所含研究意图
  │
  ├─[报告规划]          生成章节结构和搜索查询
  │
  ├─[研究与写作]        并行：每章节独立完成网络搜索 + 内容写作
  │
  ├─[幻灯片级风格设计]  全局色调、配色、形状、字体
  │
  ├─[封面 / 章节页 / 结尾页]  功能页单独生成，与正文风格一致
  │
  └─[页面级生成]        并行：每张幻灯片
        扩写 → 设计（三层）→ 编码 → 评估 → 优化
                                                  │
                                                  ▼
                                       presentation.pptx
```

---

## 核心特性

- **无模板生成** — 每张幻灯片的结构根据内容从零设计，避免视觉疲劳
- **设计与实现解耦** — *Designer* 模块在语义设计空间中推理；*Coder* 模块将设计规格转译为稳定 PPTX 代码，限制错误传播，保留美学意图
- **深度调研集成** — 基于 [Open Deep Research](https://github.com/langchain-ai/open_deep_research)，自动检索并整合网页内容、图片和学术资料
- **三维度评估** — 每张幻灯片从*完整度*、*合规性*、*美观性*打分，低分幻灯片迭代优化后再汇总
- **多模型支持** — 规划、写作、设计、编码四个角色可分别配置不同 LLM（OpenAI、Azure、Anthropic Claude 或任何兼容接口）
- **多搜索引擎** — Tavily · Perplexity · Exa · DuckDuckGo · arXiv · PubMed · Google Search · LinkUp
- **图片输入** — 传入图片或截图，视觉模型自动推断研究方向
- **并行执行** — 通过 LangGraph 的 `Send()` API 并发生成同一章节内的幻灯片

---

## 示例输出

![Full case comparison](full_case.png)

---

## 关键文件

| 文件 | 用途 |
|:---|:---|
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
source .venv/bin/activate      # Windows: .venv\Scripts\activate
pip install -e .
cd pptx_tools && pip install -e . && cd -
```

### 2. 配置环境变量

```bash
cp .env.example .env
# 编辑 .env，填入你的密钥
```

最低要求：

```dotenv
OPENAI_API_KEY=sk-...
OPENAI_API_BASE=https://api.openai.com/v1   # 或你的代理 / Azure 端点
TAVILY_API_KEY=tvly-...
```

<details>
<summary>全部支持的环境变量</summary>

| 变量 | 用途 |
|:---|:---|
| `OPENAI_API_KEY` | 主 LLM 密钥 |
| `OPENAI_API_BASE` | 自定义 Base URL / 代理 |
| `TAVILY_API_KEY` | Tavily 搜索 |
| `EXA_API_KEY` | Exa 搜索 |
| `GOOGLE_API_KEY` / `GOOGLE_CX` | Google 自定义搜索 |
| `AZURE_OPENAI_API_VERSION` | Azure OpenAI 版本 |
| `AZURE_OPENAI_DEPLOYMENT` | Azure 部署名称 |
| `CODER_API_BASE` | Coder 模型独立端点 |
| `DESIGNER_API_BASE` | Designer 模型独立端点 |
| `LANGCHAIN_API_KEY` | LangSmith 追踪（可选） |

</details>

### 3. 运行

**方式 A — 交互式 UI（LangGraph Studio）**

```bash
langgraph dev --no-reload
# 访问 http://localhost:8123
```

![LangGraph Studio](web.png)

在底部输入表单中填写 **Topic**、**Presentation Minutes**、**Style** 等字段，点击 **Submit** 启动。左侧面板显示实时工作流图，右侧面板流式显示各节点执行日志。

**方式 B — 批量运行（CSV）**

```bash
python src/open_deep_research/run.py
```

在 `run.py` 中修改 `csv_path`、`start`、`max_rows` 指向你的输入文件。CSV 至少需要 `Topic` 列，可选列：`image_path`、`style`、`presentation_minutes`。

---

## 配置参数

所有参数定义在 `configuration.py` 的 `Configuration` 类中：

| 参数 | 默认值 | 说明 |
|:---|:---|:---|
| `planner_model` | `openai/gpt-4o-mini-2024-07-18` | 报告大纲生成模型 |
| `writer_model` | `openai/gpt-4o-mini-2024-07-18` | 章节内容写作模型 |
| `coder_model` | `anthropic/claude-haiku-4.5` | PPTX 代码生成模型 |
| `designer_model` | `anthropic/claude-haiku-4.5` | 幻灯片布局设计模型 |
| `search_api` | `tavily` | 搜索引擎 |
| `max_search_depth` | `5` | 每章节最大搜索迭代次数 |
| `number_of_queries` | `2` | 每章节搜索查询数 |
| `number_of_queries_for_ppt` | `1` | 幻灯片内容扩写阶段额外查询数 |

通过 `configurable` 字典覆盖任意参数：

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

```
langgraph · langchain-openai · langchain-anthropic    # LLM / 图执行引擎
python-pptx（via pptx_tools）                         # PPTX 渲染
tavily-python · exa-py · duckduckgo-search            # 搜索后端
google-cloud-vision                                   # 可选：图像识别
langsmith                                             # 可选：链路追踪
```

可选系统依赖（PNG 导出）：

```bash
brew install libreoffice      # macOS
sudo apt install libreoffice  # Ubuntu
```

---

## 开源协议

MIT © 2026 DeepSlides Authors

---

## 引用

如果本工作对你有帮助，请引用我们的论文（[PDF](paper.pdf)）：

```bibtex
@inproceedings{cui2026design,
  title     = {Design First, Code Later: Aesthetically Pleasing Template-Free Slides Generation},
  author    = {Cui, Zhiyao and Wang, Chenxu and Hu, Shuyue and Zhang, Yiqun and
               Shao, Wenqi and Zhang, Qiaosheng and Wang, Zhen},
  booktitle = {Findings of the Association for Computational Linguistics: ACL 2026},
  year      = {2026}
}
```
