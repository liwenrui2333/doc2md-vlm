"""Build PDF/PPTX fixtures at test time so the repo carries no binary blobs."""

from __future__ import annotations

import sys
from pathlib import Path

try:
    import pymupdf
except ImportError:  # PyMuPDF < 1.24.3 only ships the fitz alias
    import fitz as pymupdf


def make_pdf(path: Path, pages: int = 2) -> Path:
    """A PDF with Chinese text, a drawn bar chart, and one blank page."""
    path.parent.mkdir(parents=True, exist_ok=True)
    doc = pymupdf.open()
    first = doc.new_page()
    first.insert_text((72, 100), "2026年第一季度经营复盘", fontsize=20, fontname="china-s")
    first.insert_text((72, 140), "营收同比增长18%。", fontsize=12, fontname="china-s")

    if pages > 1:
        second = doc.new_page()
        second.insert_text((72, 90), "各渠道花费", fontsize=14, fontname="china-s")
        for i, (name, value) in enumerate([("Google Ads", 5000), ("Bing", 1200)]):
            height = value / 5000 * 200
            x0 = 100 + i * 130
            second.draw_rect(
                pymupdf.Rect(x0, 360 - height, x0 + 60, 360),
                color=(0.2, 0.4, 0.8),
                fill=(0.2, 0.4, 0.8),
            )
            second.insert_text((x0 + 5, 380), name, fontsize=9)

    while len(doc) < pages:
        doc.new_page()

    doc.save(path)
    doc.close()
    return path


def make_pptx(path: Path) -> Path:
    """A deck with a native chart, text, and no picture shapes."""
    from pptx import Presentation
    from pptx.chart.data import CategoryChartData
    from pptx.enum.chart import XL_CHART_TYPE
    from pptx.util import Inches

    path.parent.mkdir(parents=True, exist_ok=True)
    prs = Presentation()
    slide = prs.slides.add_slide(prs.slide_layouts[6])

    box = slide.shapes.add_textbox(Inches(0.5), Inches(0.3), Inches(8), Inches(0.6))
    box.text_frame.text = "Ad spend vs leads"

    data = CategoryChartData()
    data.categories = ["Google Ads", "Bing"]
    data.add_series("Spend", (5000, 1200))
    data.add_series("Leads", (420, 90))
    graphic = slide.shapes.add_chart(
        XL_CHART_TYPE.COLUMN_CLUSTERED, Inches(0.5), Inches(1.2), Inches(8), Inches(5), data
    )
    graphic.chart.has_title = True
    graphic.chart.chart_title.text_frame.text = "Channel efficiency"

    prs.save(path)
    return path


if __name__ == "__main__":
    target = Path(sys.argv[1] if len(sys.argv) > 1 else "fixtures")
    print(make_pdf(target / "report.pdf"))
    print(make_pptx(target / "deck.pptx"))
