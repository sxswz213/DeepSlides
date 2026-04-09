# Design First, Code Later: Aesthetically Pleasing Template-Free Slides Generation

> ACL 2026 Findings &nbsp;|&nbsp; [📄 Paper](paper.pdf)

![DeepSlides case overview](case.png)

**DeepSlides** is a hierarchical, design-first framework for automated presentation generation. It explicitly decouples *slide-level style design* from *page-level code implementation*, enabling visually coherent and aesthetically engaging `.pptx` presentations without relying on any predefined template. Given a topic (and optional image or style spec), DeepSlides conducts deep web research, synthesizes a structured report, and produces a fully formatted PowerPoint file.

[中文文档](README_zh.md)

---

## Highlights

- 🏆 **76.5% Top-1 human preference** against all open-source baselines, exceeding the second-best method (9.5%) by **68.0%**
- 🏆 **52.0% win rate** against commercial systems (Kimi, Manus), exceeding the second-best (19.5%) by **32.5%**
- 📊 Highest VLM-judge average score (**3.78**) across Layout, Hierarchy, Color, Clarity, and Coherence dimensions
- 🎨 Human evaluators rate DeepSlides **4.25–4.29** on Clarity & Structure and **3.84–3.98** on Visual Design & Aesthetics

---

## How it works

DeepSlides operates at two levels:

### Slides-level Design
Given the topic and user requirements, the system first determines the global visual identity of the entire deck — tone, color palette, font colors, decorative shapes, and layout diversity guidelines. Pages are categorized into **functional pages** (cover, section dividers, end page) and **content pages**, each receiving tailored style instructions.

### Page-level Generation
Each slide is generated through three stages:

```
[Content Expansion]   ← web search retrieves supporting text and images per slide
        │
        ▼
[Design]              ← three-layer design:
                          Background layer  (textures, decorative elements)
                          Layout layer      (block arrangement, spatial positions)
                          Content layer     (exact text snippets and images)
        │
        ▼
[Implementation]      ← LLM coder translates the design spec into executable Python/PPTX code
        │
        ▼
[Evaluation & Refinement]  ← scored on Completeness · Compliance · Aesthetics
                              low-scoring slides are revised with targeted feedback
```

The full pipeline in `graph.py`:

```
Input (topic / image / style)
  │
  ▼
[Image Analysis]            ← optional: infer research intent from a figure or screenshot
  │
  ▼
[Report Planning]           ← planner LLM outlines sections and search queries
  │
  ▼
[Research & Writing]        ← parallel web search + section writing per topic
  │
  ▼
[Slides-level Design]       ← global style, colors, shapes, tone
  │
  ▼
[Cover / Chapter / End]     ← functional pages generated with matched styling
  │
  ▼
[Page-level Generation]     ← parallel per slide: expand → design → code → evaluate → refine
  │
  ▼
Output: presentation.pptx   (+ optional PNG export via LibreOffice)
```

---

## Key features

- **Template-free** — no predefined layouts; each slide's structure is designed from scratch to match its content, preventing visual fatigue across the deck.
- **Design–implementation decoupling** — a dedicated *designer* module reasons in a high-level semantic design space; a dedicated *coder* module translates the spec into stable PPTX code. This limits error propagation and preserves aesthetic intent.
- **Deep research integration** — built on [Open Deep Research](https://github.com/langchain-ai/open_deep_research); automatically retrieves and synthesises web content, images, and academic sources per slide.
- **Three-dimension evaluation** — each slide is scored on *completeness*, *compliance*, and *aesthetics*; low-scoring slides are iteratively refined before assembly.
- **Multi-model support** — planner, writer, designer, and coder roles can each use a different LLM. Works with OpenAI, Azure OpenAI, Anthropic Claude, and any OpenAI-compatible endpoint.
- **Multi-search backend** — Tavily, Perplexity, Exa, DuckDuckGo, arXiv, PubMed, Google Search, LinkUp.
- **Image input** — provide a figure or screenshot; a vision model analyses it to infer the research direction automatically.
- **Parallel execution** — slides within a section are generated concurrently via LangGraph's `Send()` API.

---

## Example outputs

![Full case comparison](full_case.png)

---

## Key files

| File | Purpose |
|---|---|
| `src/open_deep_research/graph.py` | Main workflow graph (all nodes and edges) |
| `src/open_deep_research/configuration.py` | All configurable parameters and defaults |
| `src/open_deep_research/prompts.py` | Prompt templates for every stage |
| `src/open_deep_research/state.py` | TypedDict / Pydantic state schemas |
| `src/open_deep_research/utils.py` | Search backends, image captioning helpers |
| `src/open_deep_research/run.py` | Batch runner (CSV input) |

---

## Quick start

### 1. Clone and install

```bash
git clone https://github.com/sxswz213/DeepSlides
cd DeepSlides

python -m venv .venv
source .venv/bin/activate   # Windows: .venv\Scripts\activate
pip install -e .
```

If you use the local `pptx_tools` helper package (recommended for PPT rendering):

```bash
cd pptx_tools && pip install -e . && cd -
```

### 2. Configure environment variables

Copy the example env file and fill in your keys:

```bash
cp .env.example .env
```

Minimum required variables:

```dotenv
# LLM provider
OPENAI_API_KEY=sk-...
OPENAI_API_BASE=https://api.openai.com/v1   # or your proxy / Azure endpoint

# Search (pick at least one)
TAVILY_API_KEY=tvly-...

# Optional: separate key/endpoint for the designer model
DESIGNER_API_BASE=https://...
```

Other supported variables: `AZURE_OPENAI_API_VERSION`, `AZURE_OPENAI_DEPLOYMENT`, `CODER_API_BASE`, `EXA_API_KEY`, `GOOGLE_API_KEY`, `GOOGLE_CX`, `LANGCHAIN_API_KEY`.

### 3. Run

**Option A — LangGraph dev server (interactive UI)**

```bash
langgraph dev --no-reload
```

This starts the LangGraph Studio UI at `http://localhost:8123`. You can submit inputs, inspect graph state, and replay individual nodes interactively.

![LangGraph Studio](web.png)

The left panel shows the live workflow graph; the right panel shows the execution log per node. Fill in **Topic**, **Presentation Minutes**, **Style**, etc. in the bottom input form and click **Submit**.

**Option B — Batch runner (CSV input)**

```bash
python src/open_deep_research/run.py
```

Edit the `csv_path`, `start`, and `max_rows` arguments inside `run.py` to point at your input file. The CSV must have at minimum a `Topic` column. Optional columns: `image_path`, `style`, `presentation_minutes`.

---

## Configuration reference

All parameters live in `Configuration` (`configuration.py`). Key fields:

| Parameter | Default | Description |
|---|---|---|
| `planner_model` | `openai/gpt-4o-mini-2024-07-18` | Model for report outline generation |
| `writer_model` | `openai/gpt-4o-mini-2024-07-18` | Model for section content writing |
| `coder_model` | `anthropic/claude-haiku-4.5` | Model for PPTX code generation |
| `designer_model` | `anthropic/claude-haiku-4.5` | Model for slide layout design |
| `search_api` | `tavily` | Search backend (`tavily`, `exa`, `duckduckgo`, `arxiv`, `pubmed`, `googlesearch`, `perplexity`, `linkup`) |
| `max_search_depth` | `5` | Max reflection + search iterations per section |
| `number_of_queries` | `2` | Search queries per section |
| `number_of_queries_for_ppt` | `1` | Additional search queries during slide enrichment |

Override any field by passing it in the `configurable` dict:

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

## Dependencies

Core dependencies (see `pyproject.toml` for the full list):

- `langgraph` — graph execution engine
- `langchain-openai`, `langchain-anthropic` — LLM integrations
- `python-pptx` — PPTX rendering (via `pptx_tools`)
- `tavily-python`, `exa-py`, `duckduckgo-search`, `linkup-sdk` — search backends
- `google-cloud-vision` — optional vision API for image captioning
- `langsmith` — optional tracing

Optional system dependency for PNG export:

```bash
# macOS
brew install libreoffice
# Ubuntu
sudo apt install libreoffice
```

---

## License

MIT

---

## Citation

If you find this work useful, please cite our paper ([PDF](paper.pdf)):

```bibtex
@inproceedings{cui2026design,
  title     = {Design First, Code Later: Aesthetically Pleasing Template-Free Slides Generation},
  author    = {Cui, Zhiyao and Wang, Chenxu and Hu, Shuyue and Zhang, Yiqun and Shao, Wenqi and Zhang, Qiaosheng and Wang, Zhen},
  booktitle = {Findings of the Association for Computational Linguistics: ACL 2026},
  year      = {2026}
}
```
