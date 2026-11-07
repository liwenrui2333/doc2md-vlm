# doc2md-vlm

Convert PDF and PPTX files into Markdown **plus visual evidence**, so charts don't disappear the way they do in plain text extraction.

Plain text extractors flatten a dashboard into a column of orphaned numbers. This tool keeps the picture: PDF pages are rendered to PNG, PPTX native charts are exported to CSV with their series intact, and an optional local vision model reads each page and writes down what the chart actually shows.

Runs fully offline. No API keys, no cloud calls.

## Features

- **PDF** — extract text and render every page to PNG at a configurable DPI.
- **PPTX** — extract slide text, embedded images, and native chart data as CSV plus a Markdown table.
- **Local VLM reading** — optionally send page images to a local Ollama vision model, which returns chart type, axis labels, visible numbers, trends, and a one-line conclusion.
- **Batch mode** — point it at a folder.
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
| `--local-llm` | off | Ask a local vision model to read each rendered page |
| `--vlm-provider` | `ollama` | Vision backend. Only `ollama` is implemented |
| `--vlm-url` | `http://127.0.0.1:11434/api/chat` | Endpoint; env `DOC2MD_VLM_URL` |
| `--vlm-model` | `gemma4-vision-thinking:12b` | Model name; env `DOC2MD_VLM_MODEL`, then `OLLAMA_MODEL` |
| `--max-pages` | `120` | PDF page cap, `0` for no cap. Does not apply to PPTX |
| `--dpi` | `160` | PDF render resolution |
| `--timeout` | `90` | VLM request timeout in seconds |

> **Set `--vlm-model` explicitly.** The built-in default is a placeholder that is not published on the Ollama registry, so `--local-llm` will fail until you pass a real model name. See below.

## Local vision model setup

Any Ollama vision model with the `/api/chat` image API works. Verified against the Ollama registry:

| Model | Size | Pull |
|---|---|---|
| `qwen2.5vl:7b` | 6 GB | `ollama pull qwen2.5vl:7b` |
| `qwen2.5vl:3b` | 3 GB | `ollama pull qwen2.5vl:3b` |
| `minicpm-v:8b` | 5.5 GB | `ollama pull minicpm-v:8b` |
| `llama3.2-vision:11b` | 7.8 GB | `ollama pull llama3.2-vision:11b` |
| `llava:7b` | 4.7 GB | `ollama pull llava:7b` |

Then:

```bash
python doc2md.py input.pdf --charts --local-llm --vlm-model qwen2.5vl:7b
```

Or set it once:

```bash
export DOC2MD_VLM_MODEL=qwen2.5vl:7b     # PowerShell: $env:DOC2MD_VLM_MODEL="qwen2.5vl:7b"
```

The prompt asks the model to answer in Chinese and to describe chart type, title, axes, visible numbers, trend, anomalies, and one business conclusion — or, on a page with no chart, just the key visible evidence. It is instructed not to invent hidden data.

**A VLM failure never fails the conversion.** The error text is written into the Markdown as `[VLM failed: ...]` and the rest of the document still converts.

## How it works

**PDF** (`PyMuPDF`): text is pulled page by page. When `--charts` or `--local-llm` is set, each page is also rendered to PNG at `dpi / 72` zoom and the image is referenced from the Markdown.

**PPTX** (`python-pptx`): each slide is walked shape by shape — text frames become paragraphs, picture shapes are written out as their original bytes, and chart shapes have their plots, categories, series names, and values read out into rows, saved as CSV, and inlined as a table. Slide images come from PowerPoint's own export when available.

## Limitations

- **PDF has no structural chart extraction.** `--charts` renders the page and, with `--local-llm`, adds a model reading. Only PPTX produces real chart data as CSV. The `charts` array in `metadata.json` is always empty for PDF.
- **PPTX slide images are Windows + PowerPoint only.** On macOS or Linux, `--local-llm` is skipped for PPTX with a warning; text and chart CSVs still work.
- **PDF text is flat.** Multi-column layouts and tables come out as reading-order lines, not reconstructed structure.
- **No OCR.** Pages that are pure scans yield `[No extractable text on this page.]` — the page PNG is still written.
- **`--max-pages` caps PDF only.** Long decks are converted in full.

## Troubleshooting

**`ERROR: Input is a directory. Add --batch to convert supported files inside it.`**
This is also what you get when the input path does not exist. Check the path first.

**`[VLM failed: ...]` in the output**
Usually Ollama is not running (`ollama serve`) or the model name is not pulled. Verify with `ollama list`. Connections to `127.0.0.1` bypass any system proxy, so a proxy setting is not the cause.

**PPTX produces no slide PNGs**
PowerPoint is not installed or not reachable via COM. Expected on non-Windows systems.

**`FutureWarning: The fitz API is deprecated` on stderr**
Cosmetic — comes from PyMuPDF 1.24+. Output is unaffected.

## License

MIT — see [LICENSE](LICENSE).
