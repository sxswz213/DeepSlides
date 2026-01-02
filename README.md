# DeepSlides

DeepSlides is a lightweight slide-generation workflow adapted from Open Deep Research. It follows the graph-based plan-and-execute architecture implemented in `src/open_deep_research/graph.py`:

- A planner model creates a slide/report outline and a per-section generation plan.
- For each section, writer and retrieval models produce content according to the plan.
- A designer prompt and `python-pptx` are used to render slides to a `.pptx` file; LibreOffice (`soffice`) can optionally export slides to PNG for preview.

This README focuses on the minimal steps needed to get DeepSlides running locally and on common troubleshooting tips.

## Key files

- `src/open_deep_research/graph.py` — main workflow (plan → generate → render).
- `src/open_deep_research/prompts.py` — style and layout prompt templates used by the designer model.
- `tests/` — test scripts for verifying different model and configuration setups.

## Quick start (local)

1) Clone the repository and change into it:

```bash
git clone <your-repo-url>
cd multimodal_open_deep_research
```

2) Create a virtual environment and install the package in editable mode:

```bash
python -m venv .venv
source .venv/bin/activate
pip install -e .

# Optional: if you want the LangGraph UI locally
# pip install -U "langgraph-cli[inmem]"
```

3) Configure environment variables (minimum required):

- `OPENAI_API_KEY` and optionally `OPENAI_BASE_URL` (or keys/URLs for whichever provider you use).
- `DESIGNER_API_KEY` and `DESIGNER_BASE_URL` — use these if you provide a separate key/base URL for the designer model responsible for slide formatting.
- Optional: `PLANNER_PROVIDER`, `WRITER_PROVIDER`, and other provider-specific settings can be set in `.env` or passed at runtime.

Tip: copy an example env file if present:

```bash
cp .env.example .env
# then edit .env and fill in keys and base URLs
```

4) Run a simple example (the actual CLI flags depend on `graph.py` implementation; this is a common pattern):

```bash
python -m src.open_deep_research.graph --topic "Your topic" --output ./out/presentation.pptx
```

The workflow performed by `graph.py` is: generate plan → generate per-section content → render PPTX → optionally export PNGs.


