# doc2md-vlm

Convert PDF and PPTX files into Markdown **plus visual evidence**, so charts don't disappear the way they do in plain text extraction.

Plain text extractors flatten a dashboard into a column of orphaned numbers. This tool keeps the picture: PDF pages are rendered to PNG, PPTX native charts are exported to CSV with their series intact, and an optional vision model reads each page and writes down what the chart actually shows.

Works fully offline with a local Ollama model, or against any OpenAI-compatible endpoint.

## Features

- **PDF** — extract text and render every page to PNG at a configurable DPI.
- **PPTX** — extract slide text, embedded images, and native chart data as CSV plus a Markdown table.
- **Vision model reading** — send page images to a local Ollama model or any OpenAI-compatible API; get back chart type, axis labels, visible numbers, trends, and a one-line conclusion.
- **Batch mode** — point it at a folder.
- **Actionable errors** — a missing model, a dead endpoint, or a bad API key produces a message naming the fix, not a stack trace.
- **Stable output shape** — every conversion yields a Markdown file, an `assets/` folder, a `charts/` folder, and a `metadata.json` manifest.

## Requirements

- Python 3.9+ (tested on 3.11)
- For PPTX slide images: Windows with Microsoft PowerPoint installed (uses PowerPoint COM automation). Everything else works on macOS and Linux.

## Install

```bash
git clone https://github.com/liwenrui2333/doc2md-vlm.git
cd doc2md-vlm
pip install -r requirements.txt
```

## Quick start

```bash
# PDF → Markdown + page images
python doc2md.py input.pdf --charts

# Add a local vision model to read the charts
python doc2md.py input.pdf --charts --local-llm --vlm-model qwen2.5vl:7b

# PPTX → Markdown + chart CSVs + slide images
python doc2md.py deck.pptx --charts

# Whole folder
python doc2md.py ./docs --batch --charts

# Write somewhere else
python doc2md.py input.pdf --charts -o ./out
```

Each file produces one JSON line on stdout describing what was written. Failures print `ERROR: ...` to stderr and exit with code 1.

## Output

For `input.pdf`:

```text
input_doc2md/
  input.md
  assets/
    page-001.png
    page-002.png
  charts/          # always created; empty for PDF
  metadata.json
```

In `--batch` mode each file gets a subfolder:

```text
input_doc2md/
  input/
    input.md
    assets/
    metadata.json
```

### Sample output

````markdown
# report

Source: `report.pdf`

## Page 1

Quarterly Report
Revenue grew 18% YoY. See chart on page 2.

![Page 1](assets/page-001.png)

### Visual / chart reading

这是一张柱状图，标题为"Revenue by Quarter"，纵轴单位为万元，
四个季度的数值依次为 120、180、240、310，呈逐季上升趋势……
````

The `### Visual / chart reading` block appears whenever `--local-llm` is on. If the model call fails, the block still appears and contains `[VLM failed: ...]` instead of a reading.

### PPTX chart output

Native PowerPoint chart objects are exported to `charts/slide-002-chart-01.csv`:

```csv
plot,series,category,value
1,Spend,Google Ads,5000.0
1,Spend,Bing,1200.0
```

The first 30 rows are also inlined into the Markdown as a table. CSVs are written as UTF-8 with a BOM so Excel renders Chinese category names correctly.

### `metadata.json`

```json
{
  "source": "C:\\docs\\input.pdf",
  "markdown": "C:\\docs\\input_doc2md\\input.md",
  "assets": ["..."],
  "charts": [],
  "warnings": [],
  "options": { "charts": true, "local_llm": false, "dpi": 160, "max_pages": 120 }
}
```

Check `warnings` after a batch run — truncated page counts and unreadable charts are recorded there rather than failing the conversion.

## Options

| Flag | Default | Description |
|---|---|---|
| `-o`, `--output-dir` | `<input>_doc2md/` | Output directory |
| `--batch` | off | Treat input as a folder and convert every `.pdf` / `.pptx` inside |
| `--charts` | off | Render pages/slides to PNG; for PPTX also export chart series to CSV |
| `--local-llm` | off | Ask a vision model to read each rendered page |
| `--vlm-provider` | `ollama` | Vision backend: `ollama` (local) or `openai` (any OpenAI-compatible endpoint); env `DOC2MD_VLM_PROVIDER` |
| `--vlm-url` | provider default | Endpoint; env `DOC2MD_VLM_URL`. See below for per-provider defaults |
| `--vlm-model` | provider default | Model name; env `DOC2MD_VLM_MODEL` |
| `--vlm-api-key` | empty | Key for the `openai` provider; env `DOC2MD_VLM_API_KEY` |
| `--max-pages` | `120` | PDF page cap, `0` for no cap. Does not apply to PPTX |
| `--dpi` | `160` | PDF render resolution |
| `--timeout` | `90` | VLM request timeout in seconds |

When `--vlm-url` and `--vlm-model` are omitted they follow the provider:

| Provider | Default URL | Default model |
|---|---|---|
| `ollama` | `http://127.0.0.1:11434/api/chat` | `qwen2.5vl:7b` |
| `openai` | `https://api.openai.com/v1/chat/completions` | `gpt-4o-mini` |

`DOC2MD_VLM_MODEL` wins over the provider default. `OLLAMA_MODEL` is deliberately **not** used — it is frequently set to a non-vision model, which fails silently on images.

## Which model to use

Run `python doc2md.py --help` at any time for the short version. Before a run, the tool checks your setup and prints a warning to stderr when the model is not installed, the API key is missing, or the model is not a known vision model.

**Ollama (local, free).** Models verified to accept the `/api/chat` image payload:

| Model | Size | Pull |
|---|---|---|
| `qwen2.5vl:7b` | 6 GB | `ollama pull qwen2.5vl:7b` |
| `qwen2.5vl:3b` | 3 GB | `ollama pull qwen2.5vl:3b` |
| `minicpm-v:8b` | 5.5 GB | `ollama pull minicpm-v:8b` |
| `llama3.2-vision:11b` | 7.8 GB | `ollama pull llama3.2-vision:11b` |
| `gemma4-vision-thinking:12b` | 7.6 GB | `ollama pull gemma4-vision-thinking:12b` |
| `llava:7b` | 4.7 GB | `ollama pull llava:7b` |

```bash
ollama pull qwen2.5vl:7b
python doc2md.py input.pdf --charts --local-llm --vlm-model qwen2.5vl:7b
```

Or set it once:

```bash
export DOC2MD_VLM_MODEL=qwen2.5vl:7b     # PowerShell: $env:DOC2MD_VLM_MODEL="qwen2.5vl:7b"
```

**OpenAI-compatible (cloud or self-hosted).** Any server implementing `POST /v1/chat/completions` with the `image_url` data-URL content part works — OpenAI, DeepSeek, Qwen DashScope compatible mode, vLLM, LM Studio, OpenRouter.

```bash
export DOC2MD_VLM_API_KEY=sk-...
python doc2md.py input.pdf --charts --local-llm \
  --vlm-provider openai --vlm-model gpt-4o-mini
```

Point `--vlm-url` at a different host to use a compatible server:

```bash
python doc2md.py input.pdf --charts --local-llm \
  --vlm-provider openai \
  --vlm-url http://127.0.0.1:1234/v1/chat/completions \
  --vlm-model qwen2.5-vl-7b-instruct
```

Keyless local servers work too — leave `DOC2MD_VLM_API_KEY` unset; the `Authorization` header is simply omitted.

The prompt asks the model to answer in Chinese and to describe chart type, title, axes, visible numbers, trend, anomalies, and one business conclusion — or, on a page with no chart, just the key visible evidence. It is instructed not to invent hidden data.

**A VLM failure never fails the conversion.** The reason and a fix hint are written into the Markdown as `[VLM failed: ...]`, followed by the setup steps for the provider you chose, and the rest of the document still converts. Typical output:

```markdown
[VLM failed: model 'qwen2.5vl:7b' is not pulled. Fix: ollama pull qwen2.5vl:7b]

Local vision model setup:
  1. Install Ollama: https://ollama.com/download
  2. Pull a vision model: ollama pull qwen2.5vl:7b
  ...
```

## How it works

**PDF** (`PyMuPDF`): text is pulled page by page. When `--charts` or `--local-llm` is set, each page is also rendered to PNG at `dpi / 72` zoom and the image is referenced from the Markdown.

**PPTX** (`python-pptx`): each slide is walked shape by shape — text frames become paragraphs, picture shapes are written out as their original bytes, and chart shapes have their plots, categories, series names, and values read out into rows, saved as CSV, and inlined as a table. Slide images come from PowerPoint's own export when available.

## Limitations

- **PDF has no structural chart extraction.** `--charts` renders the page and, with `--local-llm`, adds a model reading. Only PPTX produces real chart data as CSV. The `charts` array in `metadata.json` is always empty for PDF.
- **PPTX slide images are Windows + PowerPoint only.** On macOS or Linux, `--local-llm` is skipped for PPTX with a warning in `metadata.json`; text and chart CSVs still work.
- **PDF text is flat.** Multi-column layouts and tables come out as reading-order lines, not reconstructed structure.
- **No OCR.** Pages that are pure scans yield `[No extractable text on this page.]` — the page PNG is still written.
- **`--max-pages` caps PDF only.** Long decks are converted in full.

## Troubleshooting

Every error names the problem and the fix. The common ones:

**`ERROR: Input not found: <path>`**
The path does not exist. Check quoting; on Windows with a Chinese path, quote it or use forward slashes.

**`ERROR: <path> is a directory. Add --batch ...`**
The input is a folder. Add `--batch`, or point at a single file.

**`WARNING: Model '...' is not installed in Ollama. Run: ollama pull ...`**
The model is not pulled. Run the printed command, or pick another from the table above.

**`[VLM failed: ...]` in the output**
The reason is now specific — unreachable endpoint, timeout, unpulled model, wrong content type, or 401. Each carries the fix. Note that calls to `127.0.0.1` bypass any system proxy, so a proxy setting is never the cause.

**PPTX produces no slide PNGs**
PowerPoint is not installed or not reachable via COM. A warning is recorded in `metadata.json`. On non-Windows systems this is expected; text and chart CSVs still export.

**`FutureWarning: The fitz API is deprecated` on stderr**
Cosmetic — comes from PyMuPDF 1.24+. Output is unaffected.

## License

MIT — see [LICENSE](LICENSE).
