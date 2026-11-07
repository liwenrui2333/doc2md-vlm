"""Tests for doc2md. Run with: python -m pytest -q"""

from __future__ import annotations

import csv
import json
import os
import sys
from pathlib import Path

import pytest

sys.path.insert(0, str(Path(__file__).resolve().parent.parent))

import doc2md
from tests.fixtures import make_pdf, make_pptx


@pytest.fixture
def pdf(tmp_path: Path) -> Path:
    return make_pdf(tmp_path / "report.pdf")


@pytest.fixture
def pptx(tmp_path: Path) -> Path:
    return make_pptx(tmp_path / "deck.pptx")


def run(argv: list[str]) -> int:
    return doc2md.main(argv)


# --- path and argument handling -------------------------------------------------


def test_missing_input_reports_not_found(tmp_path: Path, capsys):
    code = run([str(tmp_path / "nope.pdf")])
    err = capsys.readouterr().err
    assert code == 1
    assert "Input not found" in err
    assert "directory" not in err.lower()


def test_directory_without_batch_is_rejected(tmp_path: Path, capsys):
    code = run([str(tmp_path)])
    err = capsys.readouterr().err
    assert code == 1
    assert "is a directory" in err
    assert "--batch" in err


def test_empty_directory_batch_explains_why(tmp_path: Path, capsys):
    code = run([str(tmp_path), "--batch"])
    err = capsys.readouterr().err
    assert code == 1
    assert "No PDF or PPTX files found" in err


def test_unsupported_extension_is_named(tmp_path: Path, capsys):
    bad = tmp_path / "notes.txt"
    bad.write_text("hello", encoding="utf-8")
    code = run([str(bad)])
    err = capsys.readouterr().err
    assert code == 1
    assert ".txt" in err


# --- PDF conversion -------------------------------------------------------------


def test_pdf_text_and_images(pdf: Path, capsys):
    assert run([str(pdf), "--charts"]) == 0
    capsys.readouterr()
    out = pdf.parent / "report_doc2md"
    md = (out / "report.md").read_text(encoding="utf-8")

    # Exact glyph spacing varies between PyMuPDF versions, so assert on content.
    assert "经营复盘" in md
    assert "Page 1" in md and "Page 2" in md
    assert "assets/page-001.png" in md
    assert (out / "assets" / "page-001.png").exists()


def test_pdf_without_charts_writes_no_page_images(pdf: Path, capsys):
    assert run([str(pdf)]) == 0
    capsys.readouterr()
    out = pdf.parent / "report_doc2md"
    assert not list((out / "assets").glob("*.png"))
    assert not (out / "report.md").read_text(encoding="utf-8").count("![")


def test_max_pages_truncates_and_warns(pdf: Path, capsys):
    assert run([str(pdf), "--max-pages", "1"]) == 0
    capsys.readouterr()
    out = pdf.parent / "report_doc2md"
    md = (out / "report.md").read_text(encoding="utf-8")
    meta = json.loads((out / "metadata.json").read_text(encoding="utf-8"))

    assert md.count("## Page") == 1
    assert any("Only converted first 1" in w for w in meta["warnings"])


def test_metadata_records_source_and_options(pdf: Path, capsys):
    assert run([str(pdf), "--charts", "--dpi", "100"]) == 0
    capsys.readouterr()
    meta = json.loads((pdf.parent / "report_doc2md" / "metadata.json").read_text(encoding="utf-8"))
    assert meta["source"].endswith("report.pdf")
    assert meta["options"]["charts"] is True
    assert meta["options"]["dpi"] == 100
    assert meta["charts"] == []


# --- PPTX conversion ------------------------------------------------------------


def test_pptx_chart_becomes_csv_and_table(pptx: Path, capsys):
    assert run([str(pptx), "--charts"]) == 0
    capsys.readouterr()
    out = pptx.parent / "deck_doc2md"
    md = (out / "deck.md").read_text(encoding="utf-8")

    assert "### Chart 1: Channel efficiency" in md
    assert "| Spend | Google Ads | 5000.0 |" in md

    csvs = list((out / "charts").glob("*.csv"))
    assert len(csvs) == 1
    with csvs[0].open(encoding="utf-8-sig", newline="") as fh:
        rows = list(csv.DictReader(fh))
    assert {"plot", "series", "category", "value"} == set(rows[0])
    assert {"Spend", "Leads"} == {r["series"] for r in rows}


def test_pptx_slide_text_is_kept(pptx: Path, capsys):
    assert run([str(pptx)]) == 0
    capsys.readouterr()
    md = (pptx.parent / "deck_doc2md" / "deck.md").read_text(encoding="utf-8")
    assert "Ad spend vs leads" in md


# --- batch and output directory -------------------------------------------------


def test_batch_converts_every_supported_file(pdf: Path, pptx: Path, capsys):
    assert run([str(pdf.parent), "--batch"]) == 0
    lines = [l for l in capsys.readouterr().out.splitlines() if l.strip()]
    assert len(lines) == 2

    results = [json.loads(l) for l in lines]
    assert all("_doc2md" in r["output_dir"] for r in results)
    assert {Path(r["source"]).name for r in results} == {"report.pdf", "deck.pptx"}


def test_output_dir_override(pdf: Path, tmp_path: Path, capsys):
    target = tmp_path / "custom"
    assert run([str(pdf), "-o", str(target)]) == 0
    capsys.readouterr()
    assert (target / "report.md").exists()
    assert (target / "metadata.json").exists()


def test_unicode_filename_produces_unicode_output(pdf: Path, tmp_path: Path, capsys):
    import shutil

    renamed = tmp_path / "中文报告.pdf"
    shutil.copy(pdf, renamed)
    assert run([str(renamed)]) == 0
    capsys.readouterr()
    out = tmp_path / "中文报告_doc2md"
    assert (out / "中文报告.md").exists()


# --- vision model handling ------------------------------------------------------


def test_no_model_chosen_means_no_vlm_block(pdf: Path, capsys):
    assert run([str(pdf), "--charts"]) == 0
    capsys.readouterr()
    md = (pdf.parent / "report_doc2md" / "report.md").read_text(encoding="utf-8")
    assert "Visual / chart reading" not in md


def test_vlm_failure_is_written_into_markdown(pdf: Path, capsys):
    code = run(
        [
            str(pdf),
            "--charts",
            "--local-llm",
            "--vlm-url",
            "http://127.0.0.1:9/api/chat",
            "--vlm-model",
            "qwen2.5vl:7b",
            "--timeout",
            "2",
        ]
    )
    assert code == 0  # a vision failure must never fail the conversion
    capsys.readouterr()
    md = (pdf.parent / "report_doc2md" / "report.md").read_text(encoding="utf-8")
    assert "[VLM failed:" in md
    assert "ollama serve" in md


def test_default_model_follows_provider():
    assert doc2md.parse_args(["x.pdf"]).vlm_model == "qwen2.5vl:7b"
    assert doc2md.parse_args(["x.pdf"]).vlm_url == "http://127.0.0.1:11434/api/chat"
    openai = doc2md.parse_args(["x.pdf", "--vlm-provider", "openai"])
    assert openai.vlm_model == "gpt-4o-mini"
    assert openai.vlm_url.endswith("/v1/chat/completions")


def test_doc2md_vlm_model_env_overrides_default(monkeypatch):
    monkeypatch.setenv("DOC2MD_VLM_MODEL", "minicpm-v:8b")
    assert doc2md.parse_args(["x.pdf"]).vlm_model == "minicpm-v:8b"


def test_ollama_model_env_is_ignored(monkeypatch):
    """OLLAMA_MODEL often holds a text-only model; it must not leak in."""
    monkeypatch.setenv("OLLAMA_MODEL", "some-text-model:7b")
    monkeypatch.delenv("DOC2MD_VLM_MODEL", raising=False)
    assert doc2md.parse_args(["x.pdf"]).vlm_model == "qwen2.5vl:7b"


def test_api_key_from_env(monkeypatch):
    monkeypatch.setenv("DOC2MD_VLM_API_KEY", "sk-test")
    opts = doc2md.parse_args(["x.pdf", "--vlm-provider", "openai"])
    assert opts.vlm_api_key == "sk-test"


def test_openai_missing_key_warns(capsys):
    opts = doc2md.parse_args(["x.pdf", "--local-llm", "--vlm-provider", "openai"])
    opts.vlm_api_key = ""
    notes = doc2md.check_vlm_config(opts)
    assert any("No API key" in n for n in notes)


# --- error translation ----------------------------------------------------------


def test_unpulled_model_message_names_the_pull_command():
    opts = doc2md.parse_args(["x.pdf", "--local-llm", "--vlm-model", "ghost:1b"])
    exc = Exception("model 'ghost:1b' not found")
    assert "ollama pull ghost:1b" in doc2md.explain_vlm_error(exc, opts)


def test_connection_error_names_ollama_serve():
    import requests

    opts = doc2md.parse_args(["x.pdf", "--local-llm"])
    msg = doc2md.explain_vlm_error(requests.exceptions.ConnectionError(), opts)
    assert "ollama serve" in msg


def test_timeout_message_suggests_raising_timeout():
    import requests

    opts = doc2md.parse_args(["x.pdf", "--local-llm", "--timeout", "5"])
    assert "--timeout" in doc2md.explain_vlm_error(requests.exceptions.Timeout(), opts)
