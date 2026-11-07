"""Build the demo deck used in the README screenshot.

Run from the repo root:  python docs/make_demo_deck.py
Writes demo_deck.pptx next to the current directory. Needs no assets.
"""
import sys
from pptx import Presentation
from pptx.util import Inches, Pt, Emu
from pptx.dml.color import RGBColor
from pptx.chart.data import CategoryChartData
from pptx.enum.chart import XL_CHART_TYPE, XL_LEGEND_POSITION
from pptx.enum.text import PP_ALIGN

NAVY = RGBColor(0x1F, 0x3A, 0x5F)
GREY = RGBColor(0x6B, 0x72, 0x80)
ACCENT = RGBColor(0xD9, 0x77, 0x57)

prs = Presentation()
prs.slide_width = Inches(13.333)
prs.slide_height = Inches(7.5)
blank = prs.slide_layouts[6]
SW = prs.slide_width

def title_bar(slide, title, subtitle=None):
    tb = slide.shapes.add_textbox(Inches(0.9), Inches(0.55), Inches(11), Inches(0.8))
    p = tb.text_frame.paragraphs[0]
    r = p.add_run(); r.text = title
    r.font.size = Pt(30); r.font.bold = True; r.font.color.rgb = NAVY
    if subtitle:
        sb = slide.shapes.add_textbox(Inches(0.9), Inches(1.3), Inches(11), Inches(0.5))
        r2 = sb.text_frame.paragraphs[0].add_run(); r2.text = subtitle
        r2.font.size = Pt(16); r2.font.color.rgb = GREY
    line = slide.shapes.add_shape(1, Inches(0.9), Inches(1.85), Inches(2.2), Emu(28000))
    line.fill.solid(); line.fill.fore_color.rgb = ACCENT; line.line.fill.background()

# --- Slide 1: title ---
s = prs.slides.add_slide(blank)
bg = s.shapes.add_shape(1, 0, 0, SW, prs.slide_height)
bg.fill.solid(); bg.fill.fore_color.rgb = NAVY; bg.line.fill.background()
bg.shadow.inherit = False
tb = s.shapes.add_textbox(Inches(1.2), Inches(2.6), Inches(10), Inches(1.2))
r = tb.text_frame.paragraphs[0].add_run(); r.text = "2026 年度渠道投放复盘"
r.font.size = Pt(44); r.font.bold = True; r.font.color.rgb = RGBColor(255,255,255)
tb2 = s.shapes.add_textbox(Inches(1.2), Inches(3.9), Inches(10), Inches(0.6))
r2 = tb2.text_frame.paragraphs[0].add_run(); r2.text = "Q1 数据回顾  ·  增长部"
r2.font.size = Pt(20); r2.font.color.rgb = RGBColor(0xC9,0xD4,0xE3)

# --- Slide 2: KPI cards ---
s = prs.slides.add_slide(blank)
title_bar(s, "核心指标概览", "Data source: internal CRM, export date 2026-02-28")
kpis = [("¥ 4.28M", "总投放金额", "+12.4%"), ("18,420", "有效线索", "+23.1%"),
        ("¥ 232", "单条线索成本", "-8.7%"), ("6.4%", "线索转化率", "+1.2pt")]
for i, (big, label, delta) in enumerate(kpis):
    x = Inches(0.9 + i * 3.05)
    card = s.shapes.add_shape(1, x, Inches(2.3), Inches(2.75), Inches(1.9))
    card.fill.solid(); card.fill.fore_color.rgb = RGBColor(0xF3,0xF5,0xF8)
    card.line.color.rgb = RGBColor(0xDD,0xE2,0xEA)
    card.shadow.inherit = False
    v = s.shapes.add_textbox(x + Inches(0.25), Inches(2.5), Inches(2.3), Inches(0.7))
    rv = v.text_frame.paragraphs[0].add_run(); rv.text = big
    rv.font.size = Pt(28); rv.font.bold = True; rv.font.color.rgb = NAVY
    l = s.shapes.add_textbox(x + Inches(0.25), Inches(3.2), Inches(2.3), Inches(0.4))
    rl = l.text_frame.paragraphs[0].add_run(); rl.text = label
    rl.font.size = Pt(13); rl.font.color.rgb = GREY
    d = s.shapes.add_textbox(x + Inches(0.25), Inches(3.6), Inches(2.3), Inches(0.4))
    rd = d.text_frame.paragraphs[0].add_run(); rd.text = delta + " 同比"
    rd.font.size = Pt(12); rd.font.bold = True
    rd.font.color.rgb = RGBColor(0x2E,0x8B,0x57) if delta.startswith("+") else RGBColor(0x2E,0x8B,0x57)

# --- Slide 3: bar chart ---
s = prs.slides.add_slide(blank)
title_bar(s, "各渠道花费与线索对比", "Google Ads 占预算 58%，但单线索成本最高")
data = CategoryChartData()
data.categories = ["Google Ads", "Bing Ads", "Meta", "SEO", "Affiliate"]
data.add_series("Spend (¥k)", (2480, 520, 860, 240, 180))
data.add_series("Leads", (4820, 1240, 1960, 840, 560))
gf = s.shapes.add_chart(XL_CHART_TYPE.COLUMN_CLUSTERED, Inches(0.9), Inches(2.2),
                        Inches(11.5), Inches(4.6), data)
ch = gf.chart
ch.has_title = False
ch.has_legend = True
ch.legend.position = XL_LEGEND_POSITION.BOTTOM
ch.legend.include_in_layout = False
ch.font.size = Pt(12)
ch.font.color.rgb = GREY

# --- Slide 4: trend table ---
s = prs.slides.add_slide(blank)
title_bar(s, "月度趋势", "线索连续三个月增长，2 月受春节影响")
rows = [("月份", "花费 (¥k)", "线索", "CPA (¥)", "转化率"),
        ("2025-11", 612, 2840, 215, "5.8%"),
        ("2025-12", 689, 3120, 221, "6.1%"),
        ("2026-01", 704, 3410, 206, "6.5%"),
        ("2026-02", 528, 2690, 196, "6.8%")]
tbl = s.shapes.add_table(len(rows), 5, Inches(0.9), Inches(2.3), Inches(11.5), Inches(3.2)).table
for ci, w in enumerate([2.2, 2.4, 2.2, 2.3, 2.4]):
    tbl.columns[ci].width = Inches(w)
for ri, row in enumerate(rows):
    for ci, val in enumerate(row):
        cell = tbl.cell(ri, ci)
        cell.text = str(val)
        p = cell.text_frame.paragraphs[0]
        p.alignment = PP_ALIGN.LEFT if ci == 0 else PP_ALIGN.RIGHT
        for run in p.runs:
            run.font.size = Pt(14)
            run.font.bold = ri == 0
            run.font.color.rgb = RGBColor(255,255,255) if ri == 0 else NAVY

out = sys.argv[1] if len(sys.argv) > 1 else "demo_deck.pptx"
prs.save(out)
print(f"saved {out}")
