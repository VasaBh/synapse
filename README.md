# Synapse

**A personal learning hub for AI, LLM Engineering, and technical interview preparation.**

Synapse is a self-contained static website that turns structured Word documents (`.docx`) into clean, navigable HTML reference pages. The build pipeline is pure Python — no Node.js, no bundlers, no frameworks. Write content in Word, run a script, and get a polished multi-page site.

---

## What It Is

Synapse serves two purposes:

1. **A study reference** for topics like Generative AI, RAG, LLMOps, Agentic AI, Prompt Engineering, and more — with in-depth content covering everything from first principles to production.
2. **An interview prep hub** for technical deep-dives into Python, React, OpenShift, Load Balancing, Autoscaling, and Senior Architect concepts.

The site is entirely static (no server required) and can be opened directly in a browser from the file system or served from any static host.

---

## Tech Stack

| Layer | Technology |
|---|---|
| Content authoring | Microsoft Word (`.docx`) |
| Build scripts | Python 3 (stdlib only — `zipfile`, `xml.etree.ElementTree`, `re`, `os`, `sys`, `html`) |
| Frontend | Vanilla HTML5, CSS3, JavaScript (no frameworks) |
| Styling | Custom CSS with CSS variables (`assets/style.css`, `assets/mobile.css`) |
| Interactivity | Minimal vanilla JS (`assets/scroll-top.js`) |

No `npm install`, no `pip install`, no virtual environment needed. Python's standard library handles everything.

---

## The .docx → HTML Pipeline

This is the heart of Synapse. The build scripts parse Word's internal XML format and convert it into structured HTML pages. Here's how the pipeline works:

### 1. Parsing the `.docx`

A `.docx` file is a ZIP archive. The scripts open it with `zipfile`, extract `word/document.xml`, and parse it with `xml.etree.ElementTree`. Each `<w:p>` (paragraph) element is inspected for:

- **Paragraph style** (`<w:pStyle>`) — detects `Heading1`, `Heading2`, `Heading3`, `ListParagraph`, `Normal`, etc.
- **Shading fill colour** (`<w:shd>`) — specific background hex colours (e.g. `F1F5F9`, `E2E8F0`) mark code blocks.
- **Run fonts** (`<w:rFonts>`) — `Consolas` or `Courier` font families also flag code.
- **`<w:tbl>` elements** — Word tables are extracted as row/column arrays and rendered as HTML `<table>` elements.

### 2. Converting to HTML

Paragraphs are classified and emitted as:

| Detected type | HTML output |
|---|---|
| `Heading1` / numbered section (`1. TITLE`) | `<h2>` |
| `Heading2` / numbered subsection (`1.1 Title`) | `<h3>` |
| `Heading3` / deeper (`1.1.1`) | `<h4>` |
| `ListParagraph` / Bullet style | `<ul><li>` |
| Code paragraph (shading or monospace font) | `<pre><code>` |
| ASCII diagram (high density of box-drawing chars) | `<pre class="diagram">` |
| Callout (starts with 💡 ⚠️ ✅ ❌) | `<div class="callout tip/warn/info/danger">` |
| Word table | `<div class="table-wrap"><table>` |
| Everything else | `<p>` |

### 3. Section splitting

`build_pages.py` splits each document into **sections** at every top-level heading (H1 or numbered `1. TITLE` pattern). Each section becomes its own HTML file (e.g. `topics/python/01-introduction.html`), enabling a sidebar-navigation experience with prev/next links.

### 4. Output

- **Individual section pages** land in `topics/{topic}/` — one `.html` file per section.
- **Topic index pages** (e.g. `topics/python.html`) are regenerated with a card grid linking to each section.
- The **homepage** (`index.html`) links to all topic index pages.

---

## Folder Structure

```
synapse/
│
├── index.html                  # Homepage — topic card grid
│
├── assets/
│   ├── style.css               # Main stylesheet (CSS variables, layout, components)
│   ├── mobile.css              # Responsive/mobile overrides
│   └── scroll-top.js           # Scroll-to-top button
│
├── docs/                       # Source content — Word documents
│   ├── Python Advanced Features Reference.docx
│   ├── Senior Architect Interview Prep.docx
│   ├── react-interview-guide.docx
│   ├── autoscaling-openshift-guide.docx
│   └── load-balancing-openshift-guide.docx
│
├── topics/                     # Generated HTML pages (do not edit by hand)
│   ├── generative-ai-fundamentals.html
│   ├── machine-learning.html
│   ├── rag-vector-search.html
│   ├── prompt-engineering.html
│   ├── llmops-production.html
│   ├── agentic-ai-frameworks.html
│   ├── mcp-tool-use.html
│   ├── model-fine-tuning.html
│   ├── quantum-computing.html
│   ├── ray-distributed.html
│   ├── rate-limiting.html
│   ├── commodities-complete-guide.html
│   ├── openshift.html
│   ├── python.html             # Topic index for Python
│   ├── python/                 # Section-level pages for Python
│   ├── react.html
│   ├── react/
│   ├── senior-architect.html
│   ├── senior-architect/
│   ├── autoscaling.html
│   ├── autoscaling/
│   ├── load-balancing.html
│   └── load-balancing/
│
├── build_pages.py              # PRIMARY BUILD SCRIPT — parses docs, regenerates all pages
├── generate_html.py            # Extracts docx content and injects into topic page templates
├── create_topics.py            # Scaffolds new topic pages from docx headings
├── extract_docs.py             # Utility — prints extracted paragraphs with styles (inspection)
├── extract_headings.py         # Utility — prints heading structure of all docs (inspection)
│
├── docs_extract.txt            # Cached extraction output
├── headings.txt                # Cached heading outline
│
└── .claude/
    └── settings.json           # Claude Code permissions config
```

---

## Build Scripts

### `build_pages.py` — Primary build script

The main script to run when you add or edit a `.docx`. It re-parses each document, rebuilds all section pages in the `topics/{folder}/` directories, and regenerates the topic index pages.

```bash
python build_pages.py
```

**What it does:**
- Reads each `.docx` from the `docs/` folder
- Splits the document into sections at every H1-level heading
- Writes one HTML file per section into `topics/{topic}/`
- Regenerates the topic index page (e.g. `topics/python.html`) with section cards and sidebar nav
- Prints a log of every file written

> **Note:** `DOCS_DIR` and `TOPICS_DIR` paths are hardcoded near the bottom of the script. Update them if the project moves.

---

### `generate_html.py` — Inject content into existing topic pages

Used to extract content from a single `.docx` and inject it directly into an existing topic HTML template (rather than regenerating from scratch). Useful for partial updates.

```bash
python generate_html.py
```

---

### `create_topics.py` — Scaffold new topic pages

Reads the heading structure of all docs and creates the initial HTML scaffolding for topic pages. Run this when adding a brand-new topic to the site for the first time.

```bash
python create_topics.py
```

---

### `extract_docs.py` / `extract_headings.py` — Inspection utilities

These are not part of the build pipeline — they're diagnostic tools for inspecting what the parser sees.

- `extract_docs.py` — prints the first 150 paragraphs of each doc with their detected styles.
- `extract_headings.py` — prints a structured heading outline of all docs, useful for verifying section detection before running a full build.

```bash
python extract_docs.py
python extract_headings.py
```

---

## Topics Covered

### AI & LLM Engineering
- **Generative AI Fundamentals** — Probability, tokenization, Transformer architecture, attention, context windows, temperature, hallucinations
- **Machine Learning** — Supervised & unsupervised learning, regression, decision trees, neural networks, evaluation metrics
- **RAG & Vector Search** — Retrieval-Augmented Generation, embeddings, vector databases, chunking strategies
- **Prompt Engineering** — Zero-shot, few-shot, chain-of-thought, system prompts, prompt injection, structured outputs
- **LLMOps & Production** — Model serving, monitoring, evaluation pipelines, cost management, latency optimization
- **Agentic AI Frameworks** — Agent loops, tool use, planning, memory, multi-agent systems, LangChain, CrewAI
- **MCP & Tool Use** — Model Context Protocol, tool calling patterns, function schemas, Claude SDK
- **Model Fine-Tuning** — LoRA, RLHF, instruction tuning, dataset preparation, evaluation

### Infrastructure & Systems
- **OpenShift** — Container orchestration, deployments, routes, services, operators
- **Autoscaling** — HPA, VPA, KEDA, Knative, cluster autoscaler, FastAPI metrics integration
- **Load Balancing** — Algorithms, OpenShift routes, service mesh, canary deployments, DR strategies
- **Rate Limiting** — Token bucket, sliding window, API gateway patterns
- **Ray Distributed** — Distributed computing, Ray Core, Ray Serve, parallel workloads

### Languages & Frameworks
- **Python** — Advanced features: decorators, generators, async/await, metaclasses, descriptors, GIL, memory model
- **React** — Hooks, reconciliation, advanced patterns, TypeScript integration, performance optimization

### Architecture & Specialisms
- **Senior Architect** — Architecture patterns, microservices, system design, OpenShift routing & scaling, MCP, agentic AI
- **Quantum Computing** — Qubits, superposition, entanglement, quantum algorithms
- **Commodities** — Complete guide to commodity markets

---

## Workflow: Adding New Content

1. Write or update a `.docx` file in the `docs/` folder. Use Word heading styles (`Heading 1`, `Heading 2`, `Heading 3`) to structure sections. Use a monospace font (Consolas/Courier) or a shaded paragraph background for code blocks.
2. Run the build script:
   ```bash
   python build_pages.py
   ```
3. Open `index.html` in a browser to verify the output.
4. Commit and push to GitHub:
   ```bash
   git add .
   git commit -m "Updated <topic>"
   git push
   ```

---

## Repository

```
git@github.com:VasaBh/synapse.git
```
