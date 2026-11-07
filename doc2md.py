#!/usr/bin/env python
"""
doc2md: convert PDF/PPTX documents into Markdown with visual evidence.

MVP goals:
- PDF: extract text, render page images, optionally ask a vision model.
- PPTX: extract slide text, embedded images, and native chart series as CSV.
- Output a stable folder: document.md + assets/ + charts/ + metadata.json.
"""

from __future__ import annotations

import argparse
import base64
import csv
import json
import os
import re
import shutil
import subprocess
import sys
from dataclasses import dataclass, field
from pathlib import Path
from typing import Any, Iterable

import requests


SUPPORTED_EXTS = {".pdf", ".pptx"}

PROVIDERS = ("ollama", "openai")

# Ollama models that accept images. Registry names checked at
# https://ollama.com/library; gemma4-vision-thinking verified end-to-end locally.
OLLAMA_VISION_MODELS = {
    "qwen2.5vl:7b": "6 GB, strongest small chart reader",
    "qwen2.5vl:3b": "3 GB, low VRAM",
    "minicpm-v:8b": "5.5 GB, strong OCR",
    "llama3.2-vision:11b": "7.8 GB, slower but accurate",
    "gemma4-vision-thinking:12b": "7.6 GB, reasoning-tuned",
    "llava:7b": "4.7 GB, classic baseline",
}

OPENAI_DEFAULT_BASE = "https://api.openai.com/v1"
OPENAI_DEFAULT_MODEL = "gpt-4o-mini"

VISION_PROMPT = (
    "You are reading a business document image. Reply in Chinese. "
    "If there are charts, tables, line charts, bar charts, dashboards, or numeric graphics, "
    "extract: chart type, title, axes/labels, visible numbers, trend, anomalies, and one business conclusion. "
    "If there is no chart, summarize only key visible evidence. "
    "Do not invent hidden data. Keep it concise.\n\n"
    "Context: {context}"
)


@dataclass
class ConvertOptions:
    input_path: Path
    output_dir: Path | None
    batch: bool
    charts: bool
    local_llm: bool
    vlm_provider: str
    vlm_url: str
    vlm_model: str
    vlm_api_key: str
    max_pages: int
    dpi: int
    timeout: int


@dataclass
class ConvertResult:
    source: str
    output_dir: str
    markdown: str
    assets: list[str] = field(default_factory=list)
    charts: list[str] = field(default_factory=list)
    warnings: list[str] = field(default_factory=list)


class ConversionError(RuntimeError):
    """A failure the user can act on, reported with a concrete fix."""


def safe_stem(path: Path) -> str:
    stem = re.sub(r"[^\w.-]+", "_", path.stem, flags=re.UNICODE).strip("._")
    return stem or "document"


def ensure_emptyish_dir(path: Path) -> None:
    path.mkdir(parents=True, exist_ok=True)
    for child in ("assets", "charts"):
        (path / child).mkdir(parents=True, exist_ok=True)


def rel(path: Path, base: Path) -> str:
    return path.relative_to(base).as_posix()


def write_text(path: Path, text: str) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_text(text, encoding="utf-8", newline="\n")


def import_pymupdf():
    try:
        import pymupdf
    except ImportError:
        try:
            import fitz as pymupdf
        except ImportError as exc:
            raise ConversionError(
                "PyMuPDF is not installed. Fix: pip install pymupdf"
            ) from exc
    return pymupdf


def build_vlm_help(provider: str) -> str:
    if provider == "ollama":
        listed = ", ".join(OLLAMA_VISION_MODELS)
        return (
            "Local vision model setup:\n"
            "  1. Install Ollama: https://ollama.com/download\n"
            "  2. Pull a vision model: ollama pull qwen2.5vl:7b\n"
            f"  3. Pick a name from: {listed}\n"
            "  4. Pass it explicitly: python doc2md.py input.pdf --charts --local-llm "
            "--vlm-model qwen2.5vl:7b\n"
            "  Or set DOC2MD_VLM_MODEL=qwen2.5vl:7b. Check what you have with: ollama list"
        )
    return (
        "OpenAI-compatible setup:\n"
        "  1. Export DOC2MD_VLM_API_KEY (or pass --vlm-api-key).\n"
        "  2. Point --vlm-url at any compatible endpoint, e.g. "
        "--vlm-url http://127.0.0.1:8000/v1/chat/completions\n"
        "     Works with OpenAI, DeepSeek, Qwen DashScope compatible mode, vLLM, "
        "LM Studio, OpenRouter.\n"
        "  3. Use a model that accepts images: --vlm-model gpt-4o-mini"
    )


def call_ollama_vision(image_path: Path, prompt: str, model: str, url: str, timeout: int, api_key: str = "") -> str:
    b64 = base64.b64encode(image_path.read_bytes()).decode("ascii")
    payload = {
        "model": model,
        "messages": [{"role": "user", "content": prompt, "images": [b64]}],
        "stream": False,
        "options": {"temperature": 0.1, "num_predict": 700},
    }
    response = requests.post(url, json=payload, timeout=(3, timeout), proxies={"http": "", "https": ""})
    try:
        data = response.json()
    except ValueError:
        response.raise_for_status()
        raise ConversionError(
            f"Expected JSON from {url} but got {response.headers.get('Content-Type', 'unknown')}. "
            "Is --vlm-url pointing at an Ollama server?"
        )
    if "error" in data:
        raise ConversionError(f"Ollama error: {data['error']}")
    response.raise_for_status()
    return (data.get("message") or {}).get("content", "").strip()


def call_openai_vision(
    image_path: Path, prompt: str, model: str, url: str, timeout: int, api_key: str = ""
) -> str:
    b64 = base64.b64encode(image_path.read_bytes()).decode("ascii")
    payload = {
        "model": model,
        "messages": [
            {
                "role": "user",
                "content": [
                    {"type": "text", "text": prompt},
                    {"type": "image_url", "image_url": {"url": f"data:image/png;base64,{b64}"}},
                ],
            }
        ],
        "temperature": 0.1,
        "max_tokens": 700,
        "stream": False,
    }
    headers = {"Content-Type": "application/json"}
    if api_key:
        headers["Authorization"] = f"Bearer {api_key}"
    response = requests.post(url, json=payload, headers=headers, timeout=(3, timeout))
    if response.status_code == 401:
        raise ConversionError(
            "401 Unauthorized from the vision API. Set a valid key with "
            "DOC2MD_VLM_API_KEY or --vlm-api-key. " + build_vlm_help("openai")
        )
    response.raise_for_status()
    data = response.json()
    choices = data.get("choices") or []
    if not choices:
        raise ConversionError(f"Vision API returned no choices: {str(data)[:300]}")
    return (choices[0].get("message") or {}).get("content", "").strip()


def describe_image_with_vlm(image_path: Path, opts: ConvertOptions, context: str) -> str:
    if not opts.local_llm:
        return ""
    prompt = VISION_PROMPT.format(context=context)
    caller = call_openai_vision if opts.vlm_provider == "openai" else call_ollama_vision
    try:
        return caller(
            image_path, prompt, opts.vlm_model, opts.vlm_url, opts.timeout, opts.vlm_api_key
        )
    except Exception as exc:
        reason = explain_vlm_error(exc, opts)
        return f"[VLM failed: {reason}]\n\n{build_vlm_help(opts.vlm_provider)}"



def convert_pdf(path: Path, out_dir: Path, opts: ConvertOptions) -> ConvertResult:
    fitz = import_pymupdf()

    ensure_emptyish_dir(out_dir)
    assets_dir = out_dir / "assets"
    pages_md: list[str] = [f"# {path.stem}", "", f"Source: `{path}`", ""]
    assets: list[str] = []
    warnings: list[str] = []

    with fitz.open(path) as doc:
        page_count = len(doc)
        limit = min(page_count, opts.max_pages) if opts.max_pages > 0 else page_count
        if limit < page_count:
            warnings.append(f"Only converted first {limit} of {page_count} pages.")

        for idx in range(limit):
            page = doc.load_page(idx)
            page_no = idx + 1
            text = page.get_text("text").strip()
            pages_md.extend([f"## Page {page_no}", ""])
            if text:
                pages_md.extend([text, ""])
            else:
                pages_md.extend(["[No extractable text on this page.]", ""])

            if opts.charts or opts.local_llm:
                zoom = opts.dpi / 72
                pix = page.get_pixmap(matrix=fitz.Matrix(zoom, zoom), alpha=False)
                image_path = assets_dir / f"page-{page_no:03d}.png"
                pix.save(image_path)
                assets.append(str(image_path))
                pages_md.extend([f"![Page {page_no}]({rel(image_path, out_dir)})", ""])

                vision = describe_image_with_vlm(image_path, opts, f"{path.name}, page {page_no}")
                if vision:
                    pages_md.extend(["### Visual / chart reading", "", vision, ""])

    md_path = out_dir / f"{safe_stem(path)}.md"
    write_text(md_path, "\n".join(pages_md).strip() + "\n")
    return ConvertResult(str(path), str(out_dir), str(md_path), assets, [], warnings)


def text_from_shape(shape: Any) -> str:
    if not getattr(shape, "has_text_frame", False):
        return ""
    parts: list[str] = []
    for paragraph in shape.text_frame.paragraphs:
        text = "".join(run.text for run in paragraph.runs).strip()
        if text:
            parts.append(text)
    return "\n".join(parts).strip()


def iter_chart_rows(chart: Any) -> list[dict[str, Any]]:
    rows: list[dict[str, Any]] = []
    try:
        plots = list(chart.plots)
    except Exception:
        plots = []
    for plot_idx, plot in enumerate(plots, start=1):
        categories: list[str] = []
        try:
            categories = [str(c) for c in plot.categories]
        except Exception:
            categories = []
        try:
            series_list = list(plot.series)
        except Exception:
            series_list = []
        for series_idx, series in enumerate(series_list, start=1):
            try:
                values = list(series.values)
            except Exception:
                values = []
            try:
                name = str(series.name)
            except Exception:
                name = f"series_{series_idx}"
            max_len = max(len(values), len(categories), 1)
            for i in range(max_len):
                rows.append(
                    {
                        "plot": plot_idx,
                        "series": name,
                        "category": categories[i] if i < len(categories) else i + 1,
                        "value": values[i] if i < len(values) else "",
                    }
                )
    return rows


def save_chart_csv(rows: list[dict[str, Any]], csv_path: Path) -> None:
    csv_path.parent.mkdir(parents=True, exist_ok=True)
    with csv_path.open("w", encoding="utf-8-sig", newline="") as fh:
        writer = csv.DictWriter(fh, fieldnames=["plot", "series", "category", "value"])
        writer.writeheader()
        writer.writerows(rows)


def export_pptx_slides_to_images(path: Path, out_dir: Path) -> list[Path]:
    """Best-effort slide rendering via PowerPoint COM on Windows."""
    if os.name != "nt":
        return []
    ps = shutil.which("powershell") or shutil.which("pwsh")
    if not ps:
        return []
    out_dir.mkdir(parents=True, exist_ok=True)
    script = f"""
$ErrorActionPreference = 'Stop'
$ppt = New-Object -ComObject PowerPoint.Application
$ppt.Visible = [Microsoft.Office.Core.MsoTriState]::msoTrue
$deck = $ppt.Presentations.Open('{str(path).replace("'", "''")}', $false, $false, $false)
$deck.Export('{str(out_dir).replace("'", "''")}', 'PNG')
$deck.Close()
$ppt.Quit()
"""
    try:
        subprocess.run([ps, "-NoProfile", "-Command", script], check=True, capture_output=True, text=True, timeout=120)
    except Exception:
        return []
    return sorted(out_dir.glob("*.PNG")) + sorted(out_dir.glob("*.png"))


def convert_pptx(path: Path, out_dir: Path, opts: ConvertOptions) -> ConvertResult:
    try:
        from pptx import Presentation
    except ImportError as exc:
        raise ConversionError(
            "python-pptx is not installed. Fix: pip install python-pptx"
        ) from exc

    ensure_emptyish_dir(out_dir)
    assets_dir = out_dir / "assets"
    charts_dir = out_dir / "charts"
    prs = Presentation(str(path))
    slide_images = export_pptx_slides_to_images(path, assets_dir / "slides") if opts.charts or opts.local_llm else []
    slide_image_by_index = {i + 1: image for i, image in enumerate(slide_images)}
    if (opts.charts or opts.local_llm) and not slide_images:
        reason = (
            "requires Windows with PowerPoint installed"
            if os.name != "nt"
            else "could not export slides; is PowerPoint installed?"
        )
        warnings.append(f"Slide images unavailable ({reason}). Text and chart CSV still exported.")

    md: list[str] = [f"# {path.stem}", "", f"Source: `{path}`", ""]
    assets: list[str] = []
    chart_files: list[str] = []
    warnings: list[str] = []

    for slide_idx, slide in enumerate(prs.slides, start=1):
        md.extend([f"## Slide {slide_idx}", ""])

        texts: list[str] = []
        chart_count = 0
        image_count = 0
        for shape_idx, shape in enumerate(slide.shapes, start=1):
            shape_text = text_from_shape(shape)
            if shape_text:
                texts.append(shape_text)

            if getattr(shape, "has_chart", False):
                chart_count += 1
                chart = shape.chart
                rows = iter_chart_rows(chart)
                title = ""
                try:
                    if chart.has_title:
                        title = chart.chart_title.text_frame.text.strip()
                except Exception:
                    title = ""
                if title:
                    md.extend([f"### Chart {chart_count}: {title}", ""])
                else:
                    md.extend([f"### Chart {chart_count}", ""])

                if rows:
                    csv_path = charts_dir / f"slide-{slide_idx:03d}-chart-{chart_count:02d}.csv"
                    save_chart_csv(rows, csv_path)
                    chart_files.append(str(csv_path))
                    md.extend([f"Data: [{csv_path.name}]({rel(csv_path, out_dir)})", ""])
                    md.extend(["| Series | Category | Value |", "|---|---:|---:|"])
                    for row in rows[:30]:
                        md.append(f"| {row['series']} | {row['category']} | {row['value']} |")
                    if len(rows) > 30:
                        md.append(f"| ... | ... | {len(rows) - 30} more rows in CSV |")
                    md.append("")
                else:
                    warnings.append(f"Slide {slide_idx} chart {chart_count}: no readable chart data.")

            if getattr(shape, "shape_type", None) == 13 and hasattr(shape, "image"):
                image_count += 1
                ext = shape.image.ext or "png"
                image_path = assets_dir / f"slide-{slide_idx:03d}-image-{image_count:02d}.{ext}"
                image_path.write_bytes(shape.image.blob)
                assets.append(str(image_path))
                md.extend([f"![Slide {slide_idx} image {image_count}]({rel(image_path, out_dir)})", ""])

        if texts:
            md.extend(texts)
            md.append("")
        else:
            md.extend(["[No slide text detected.]", ""])

        slide_image = slide_image_by_index.get(slide_idx)
        if slide_image:
            target = assets_dir / f"slide-{slide_idx:03d}.png"
            if slide_image.resolve() != target.resolve():
                shutil.copy2(slide_image, target)
            assets.append(str(target))
            md.extend([f"![Slide {slide_idx}]({rel(target, out_dir)})", ""])
            vision = describe_image_with_vlm(target, opts, f"{path.name}, slide {slide_idx}")
            if vision:
                md.extend(["### Visual / chart reading", "", vision, ""])
        elif opts.local_llm:
            warnings.append(f"Slide {slide_idx}: no slide image to send to the vision model.")

    md_path = out_dir / f"{safe_stem(path)}.md"
    write_text(md_path, "\n".join(md).strip() + "\n")
    return ConvertResult(str(path), str(out_dir), str(md_path), assets, chart_files, warnings)


def convert_one(path: Path, opts: ConvertOptions) -> ConvertResult:
    path = path.resolve()
    if not path.exists():
        raise ConversionError(
            f"Input not found: {path}\n"
            "Check the path, or pass a folder together with --batch."
        )
    suffix = path.suffix.lower()
    if suffix not in SUPPORTED_EXTS:
        supported = ", ".join(sorted(SUPPORTED_EXTS))
        raise ConversionError(
            f"Unsupported file type {suffix or '(no extension)'}: {path.name}. "
            f"Supported: {supported}"
        )

    base_out = opts.output_dir.resolve() if opts.output_dir else path.parent / f"{safe_stem(path)}_doc2md"
    out_dir = base_out if not opts.batch else base_out / safe_stem(path)
    ensure_emptyish_dir(out_dir)

    if suffix == ".pdf":
        result = convert_pdf(path, out_dir, opts)
    else:
        result = convert_pptx(path, out_dir, opts)

    metadata = {
        "source": result.source,
        "markdown": result.markdown,
        "assets": result.assets,
        "charts": result.charts,
        "warnings": result.warnings,
        "options": {
            "charts": opts.charts,
            "local_llm": opts.local_llm,
            "vlm_provider": opts.vlm_provider,
            "vlm_url": opts.vlm_url,
            "vlm_model": opts.vlm_model,
            "max_pages": opts.max_pages,
            "dpi": opts.dpi,
        },
    }
    write_text(out_dir / "metadata.json", json.dumps(metadata, ensure_ascii=False, indent=2))
    return result


def iter_inputs(input_path: Path, batch: bool) -> Iterable[Path]:
    if not input_path.exists():
        raise ConversionError(
            f"Input not found: {input_path.resolve()}\n"
            "Check the path or quoting. If it is a folder, add --batch."
        )
    if input_path.is_file():
        yield input_path
        return
    if not batch:
        raise ConversionError(
            f"{input_path.resolve()} is a directory. Add --batch to convert every "
            "PDF/PPTX inside it."
        )
    found = [
        child
        for child in sorted(input_path.iterdir())
        if child.is_file() and child.suffix.lower() in SUPPORTED_EXTS
    ]
    if not found:
        raise ConversionError(
            f"No PDF or PPTX files found in {input_path.resolve()}.\n"
            "Only top-level files are scanned; nested folders are skipped."
        )
    yield from found


def parse_args(argv: list[str]) -> ConvertOptions:
    parser = argparse.ArgumentParser(
        description="Convert PDF/PPTX to Markdown with chart evidence.",
        epilog=(
            "Vision model examples:\n"
            "  local Ollama:  --local-llm --vlm-model qwen2.5vl:7b\n"
            "  OpenAI compat: --local-llm --vlm-provider openai --vlm-model gpt-4o-mini\n"
            "                 (key from DOC2MD_VLM_API_KEY or --vlm-api-key)\n"
            f"  Ollama vision models verified: {', '.join(OLLAMA_VISION_MODELS)}"
        ),
        formatter_class=argparse.RawDescriptionHelpFormatter,
    )
    parser.add_argument("input", help="PDF/PPTX file or folder with --batch")
    parser.add_argument("-o", "--output-dir", help="Output directory")
    parser.add_argument("--batch", action="store_true", help="Convert all supported files in a folder")
    parser.add_argument("--charts", action="store_true", help="Extract chart data and render visual evidence")
    parser.add_argument("--local-llm", action="store_true", help="Use a vision LLM for image/chart reading")
    parser.add_argument(
        "--vlm-provider",
        default=os.environ.get("DOC2MD_VLM_PROVIDER", "ollama"),
        choices=list(PROVIDERS),
        help="Vision backend: ollama (local) or openai (any OpenAI-compatible endpoint)",
    )
    parser.add_argument(
        "--vlm-url",
        default=os.environ.get("DOC2MD_VLM_URL"),
        help="Endpoint URL. Defaults: Ollama http://127.0.0.1:11434/api/chat, "
        "OpenAI https://api.openai.com/v1/chat/completions",
    )
    parser.add_argument(
        "--vlm-model",
        default=os.environ.get("DOC2MD_VLM_MODEL"),
        help=f"Model name. Suggested Ollama vision models: {', '.join(OLLAMA_VISION_MODELS)}",
    )
    parser.add_argument(
        "--vlm-api-key",
        default=os.environ.get("DOC2MD_VLM_API_KEY", ""),
        help="API key for the openai provider (or set DOC2MD_VLM_API_KEY)",
    )
    parser.add_argument("--max-pages", type=int, default=120, help="PDF page cap; 0 means no cap")
    parser.add_argument("--dpi", type=int, default=160, help="PDF render DPI")
    parser.add_argument("--timeout", type=int, default=90, help="VLM request timeout seconds")
    args = parser.parse_args(argv)

    provider = args.vlm_provider
    vlm_url = args.vlm_url
    vlm_model = args.vlm_model
    if provider == "openai":
        vlm_url = vlm_url or f"{OPENAI_DEFAULT_BASE}/chat/completions"
        vlm_model = vlm_model or OPENAI_DEFAULT_MODEL
    else:
        vlm_url = vlm_url or "http://127.0.0.1:11434/api/chat"
        vlm_model = vlm_model or "qwen2.5vl:7b"

    return ConvertOptions(
        input_path=Path(args.input),
        output_dir=Path(args.output_dir) if args.output_dir else None,
        batch=args.batch,
        charts=args.charts,
        local_llm=args.local_llm,
        vlm_provider=provider,
        vlm_url=vlm_url,
        vlm_model=vlm_model,
        vlm_api_key=args.vlm_api_key,
        max_pages=args.max_pages,
        dpi=args.dpi,
        timeout=args.timeout,
    )


def check_vlm_config(opts: ConvertOptions) -> list[str]:
    """Warn about vision settings the user can fix before a long run starts."""
    if not opts.local_llm:
        return []
    notes: list[str] = []
    if opts.vlm_provider == "ollama":
        local = local_ollama_models()
        if local is not None and opts.vlm_model not in local:
            notes.append(
                f"Model '{opts.vlm_model}' is not installed in Ollama. "
                f"Run: ollama pull {opts.vlm_model}\n"
                f"Verified vision models: {', '.join(OLLAMA_VISION_MODELS)}"
            )
        elif local is None and opts.vlm_model not in OLLAMA_VISION_MODELS:
            notes.append(
                f"Model '{opts.vlm_model}' is not in the verified list. "
                f"Known-good Ollama vision models: {', '.join(OLLAMA_VISION_MODELS)}"
            )
    if opts.vlm_provider == "openai" and not opts.vlm_api_key:
        notes.append(
            "No API key set for the openai provider. Set DOC2MD_VLM_API_KEY or "
            "--vlm-api-key, or point --vlm-url at a keyless local server such as "
            "LM Studio or vLLM."
        )
    return notes


def local_ollama_models() -> set[str] | None:
    """Names installed in the local Ollama, or None if it cannot be reached."""
    host = os.environ.get("OLLAMA_HOST", "127.0.0.1:11434")
    if host.startswith("0.0.0.0"):
        host = host.replace("0.0.0.0", "127.0.0.1")
    if "://" not in host:
        host = f"http://{host}"
    parsed = host.split("//", 1)[1]
    if ":" not in parsed.split("/", 1)[0]:
        host = host.rstrip("/") + ":11434"
    try:
        r = requests.get(f"{host.rstrip('/')}/api/tags", timeout=5, proxies={"http": "", "https": ""})
        r.raise_for_status()
        return {m["name"] for m in r.json().get("models", [])}
    except Exception:
        return None


def explain_vlm_error(exc: Exception, opts: ConvertOptions) -> str:
    """Turn a vision call failure into a message the user can act on."""
    text = str(exc)
    if isinstance(exc, requests.exceptions.ConnectionError):
        if opts.vlm_provider == "ollama":
            return f"cannot reach Ollama at {opts.vlm_url}. Start it with: ollama serve"
        return f"cannot reach {opts.vlm_url}. Check --vlm-url and that the server is running."
    if isinstance(exc, requests.exceptions.Timeout):
        return f"timed out after {opts.timeout}s. Raise --timeout."
    lowered = text.lower()
    if "not found" in lowered and "model" in lowered:
        return f"model '{opts.vlm_model}' is not pulled. Fix: ollama pull {opts.vlm_model}"
    if "does not support images" in lowered or "vision" in lowered and "support" in lowered:
        return f"'{opts.vlm_model}' does not accept images. Use one of: {', '.join(OLLAMA_VISION_MODELS)}"
    if "401" in text or "unauthorized" in lowered:
        return (
            "401 Unauthorized. Set a valid key with DOC2MD_VLM_API_KEY or --vlm-api-key."
        )
    return text


def main(argv: list[str] | None = None) -> int:
    try:
        opts = parse_args(argv or sys.argv[1:])
        notes = check_vlm_config(opts)
        for note in notes:
            print(f"WARNING: {note}", file=sys.stderr)

        results: list[ConvertResult] = []
        for item in iter_inputs(opts.input_path, opts.batch):
            results.append(convert_one(item, opts))
    except ConversionError as exc:
        print(f"ERROR: {exc}", file=sys.stderr)
        return 1
    except KeyboardInterrupt:
        print("ERROR: interrupted", file=sys.stderr)
        return 130
    except Exception as exc:
        print(f"ERROR: {type(exc).__name__}: {exc}", file=sys.stderr)
        print(
            "If this is unexpected, re-run and report it at "
            "https://github.com/liwenrui2333/doc2md-vlm/issues",
            file=sys.stderr,
        )
        return 1

    for result in results:
        print(json.dumps(result.__dict__, ensure_ascii=False))
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
