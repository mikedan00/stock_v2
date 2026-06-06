"""
exporters/ppt_exporter.py
요약본을 PPT(.pptx)로 내보내기 — pptxgenjs npm 패키지 사용
"""
from __future__ import annotations
import os, subprocess, tempfile
from datetime import date


def _safe(text: str, limit: int = 300) -> str:
    return (str(text) or "")[:limit].replace("`", "'").replace("\\", "\\\\").replace("${", "\\${").replace("\n", " ")


def _change_color(change) -> str:
    if change is None:
        return "90A4AE"
    return "EF5350" if change > 0 else "42A5F5" if change < 0 else "90A4AE"


def export_ppt(
    stocks: list[dict],
    stock_briefs: list[str],
    portfolio_brief: str,
    persona_label: str,
    persona_emoji: str,
    output_path: str,
) -> str:
    today = date.today().isoformat()

    # 슬라이드 1 — 표지
    slide1 = f"""
  // ── 슬라이드 1: 표지 ──────────────────────────────────────────────────
  let s1 = pres.addSlide();
  s1.background = {{ color: "0D1B4B" }};
  s1.addShape(pres.shapes.RECTANGLE, {{ x: 0, y: 3.8, w: 10, h: 1.0, fill: {{ color: "1565C0" }} }});
  s1.addText("📊 주식 AI 투자 브리핑", {{
    x: 0.5, y: 0.6, w: 9, h: 1.1,
    fontSize: 40, bold: true, color: "FFFFFF", fontFace: "Malgun Gothic"
  }});
  s1.addText("AI Investment Briefing Report", {{
    x: 0.5, y: 1.7, w: 9, h: 0.5,
    fontSize: 18, color: "90CAF9", fontFace: "Calibri", italic: true
  }});
  s1.addText("{_safe(persona_emoji + " " + persona_label)}", {{
    x: 0.5, y: 2.4, w: 9, h: 0.55,
    fontSize: 20, bold: true, color: "FFD54F", fontFace: "Malgun Gothic"
  }});
  s1.addText("기준일: {today}", {{
    x: 0.5, y: 3.85, w: 9, h: 0.5,
    fontSize: 16, color: "FFFFFF", fontFace: "Malgun Gothic", align: "center"
  }});
  s1.addText("본 리포트는 AI 분석 기반이며 투자 결정은 본인 책임입니다.", {{
    x: 0.5, y: 5.15, w: 9, h: 0.35,
    fontSize: 11, color: "78909C", fontFace: "Malgun Gothic", align: "center"
  }});
"""

    # 슬라이드 2 — 포트폴리오 현황 테이블
    table_rows = [
        [
            { "text": "종목명", "options": { "bold": True, "color": "FFFFFF", "fill": { "color": "1565C0" } } },
            { "text": "코드", "options": { "bold": True, "color": "FFFFFF", "fill": { "color": "1565C0" } } },
            { "text": "종가", "options": { "bold": True, "color": "FFFFFF", "fill": { "color": "1565C0" } } },
            { "text": "등락률", "options": { "bold": True, "color": "FFFFFF", "fill": { "color": "1565C0" } } },
            { "text": "시장", "options": { "bold": True, "color": "FFFFFF", "fill": { "color": "1565C0" } } },
        ]
    ]
    for s in stocks[:10]:
        close = s.get("close")
        change = s.get("change_rate")
        currency = "원" if s.get("market") == "KR" else "$"
        close_str = f"{close:,.0f}{currency}" if close else "N/A"
        change_str = f"{change:+.2f}%" if change is not None else "N/A"
        clr = _change_color(change)
        table_rows.append([
            s.get("name", s["ticker"]),
            s["ticker"],
            close_str,
            { "text": change_str, "options": { "color": clr, "bold": True } },
            s.get("market", ""),
        ])

    table_rows_js = "[\n"
    for row in table_rows:
        row_js = "    [\n"
        for cell in row:
            if isinstance(cell, dict):
                txt = _safe(cell["text"])
                opts = cell.get("options", {})
                bold_str = "true" if opts.get("bold") else "false"
                color_str = opts.get("color", "1E1E1E")
                fill_color = opts.get("fill", {}).get("color", "")
                if fill_color:
                    row_js += f'      {{ text: `{txt}`, options: {{ bold: {bold_str}, color: "{color_str}", fill: {{ color: "{fill_color}" }} }} }},\n'
                else:
                    row_js += f'      {{ text: `{txt}`, options: {{ bold: {bold_str}, color: "{color_str}" }} }},\n'
            else:
                row_js += f'      `{_safe(str(cell))}`,\n'
        row_js += "    ],\n"
        table_rows_js += row_js
    table_rows_js += "  ]"

    table_h = min(0.42 * (len(stocks) + 1), 4.0)

    slide2 = f"""
  // ── 슬라이드 2: 포트폴리오 현황 ────────────────────────────────────────
  let s2 = pres.addSlide();
  s2.background = {{ color: "F8FAFF" }};
  s2.addShape(pres.shapes.RECTANGLE, {{ x: 0, y: 0, w: 10, h: 0.9, fill: {{ color: "0D1B4B" }} }});
  s2.addText("📈 포트폴리오 종목 현황", {{
    x: 0.3, y: 0.1, w: 9.4, h: 0.7,
    fontSize: 24, bold: true, color: "FFFFFF", fontFace: "Malgun Gothic"
  }});
  s2.addTable({table_rows_js}, {{
    x: 0.3, y: 1.05, w: 9.4, h: {table_h},
    colW: [2.8, 1.4, 1.8, 1.5, 0.9],
    border: {{ pt: 1, color: "CCCCCC" }},
    fontFace: "Malgun Gothic", fontSize: 13,
    align: "center"
  }});
  s2.addText("기준일: {today}", {{
    x: 0.3, y: 5.3, w: 9.4, h: 0.3,
    fontSize: 11, color: "90A4AE", fontFace: "Malgun Gothic", align: "right"
  }});
"""

    # 슬라이드 3 — 포트폴리오 종합 전략 요약
    port_summary = portfolio_brief[:600].replace("`", "'").replace("\\", "\\\\").replace("${", "\\${")
    # 줄바꿈을 배열 항목으로 변환
    port_lines = [l.strip() for l in port_summary.split("\n") if l.strip()][:12]
    port_text_items = ""
    for line in port_lines:
        clean = line.replace("**", "").replace("##", "").strip()
        is_heading = line.startswith("**") or line.startswith("##") or line.startswith("1.") or line.startswith("2.")
        bold_str = "true" if is_heading else "false"
        port_text_items += f'  {{ text: `{_safe(clean, 150)}`, options: {{ breakLine: true, bold: {bold_str}, fontSize: {14 if is_heading else 13} }} }},\n'

    slide3 = f"""
  // ── 슬라이드 3: 종합 전략 요약 ──────────────────────────────────────────
  let s3 = pres.addSlide();
  s3.background = {{ color: "F8FAFF" }};
  s3.addShape(pres.shapes.RECTANGLE, {{ x: 0, y: 0, w: 10, h: 0.9, fill: {{ color: "1565C0" }} }});
  s3.addText("📋 포트폴리오 종합 전략", {{
    x: 0.3, y: 0.1, w: 7, h: 0.7,
    fontSize: 24, bold: true, color: "FFFFFF", fontFace: "Malgun Gothic"
  }});
  s3.addText("{_safe(persona_emoji + " " + persona_label, 60)}", {{
    x: 7.0, y: 0.15, w: 2.7, h: 0.6,
    fontSize: 13, color: "FFD54F", fontFace: "Malgun Gothic", align: "right"
  }});
  s3.addShape(pres.shapes.RECTANGLE, {{
    x: 0.3, y: 1.05, w: 9.4, h: 4.3,
    fill: {{ color: "FFFFFF" }},
    shadow: {{ type: "outer", color: "000000", blur: 8, offset: 2, angle: 135, opacity: 0.08 }}
  }});
  s3.addText([
{port_text_items}
  ], {{
    x: 0.5, y: 1.15, w: 9.0, h: 4.1,
    fontFace: "Malgun Gothic", fontSize: 13, color: "1A237E", valign: "top"
  }});
"""

    # 슬라이드 4+ — 종목별 개별 분석 (최대 10개)
    stock_slides = ""
    for idx, (stock, brief) in enumerate(zip(stocks[:10], stock_briefs[:10])):
        name = _safe(stock.get("name", stock["ticker"]), 30)
        ticker = _safe(stock["ticker"], 15)
        close = stock.get("close")
        change = stock.get("change_rate")
        market = stock.get("market", "")
        currency = "원" if market == "KR" else "$"
        close_str = _safe(f"{close:,.0f}{currency}" if close else "N/A")
        change_str = _safe(f"{change:+.2f}%" if change is not None else "N/A")
        clr = _change_color(change)
        flag = "🇰🇷" if market == "KR" else "🇺🇸"

        # 브리핑 요약 (첫 600자)
        brief_lines = [l.strip() for l in brief[:600].split("\n") if l.strip()][:10]
        brief_items = ""
        for line in brief_lines:
            clean = line.replace("**", "").replace("##", "").strip()
            is_bold = line.startswith("**") or line.startswith("##") or (len(line) > 2 and line[1] == ".")
            bold_str = "true" if is_bold else "false"
            brief_items += f'    {{ text: `{_safe(clean, 120)}`, options: {{ breakLine: true, bold: {bold_str}, fontSize: {13 if is_bold else 12} }} }},\n'

        # 주가 추이 히스토리 데이터 (주간)
        hist_week = stock.get("history", {}).get("week")
        chart_data_str = ""
        chart_code = ""
        if hist_week is not None and not hist_week.empty and "close" in hist_week.columns:
            vals = [round(float(v), 0) for v in hist_week["close"].values[-7:]]
            labels = [str(d)[:10] for d in hist_week.index[-7:]]
            vals_js = ", ".join(str(v) for v in vals)
            labels_js = ", ".join(f'"{l}"' for l in labels)
            chart_color = "EF5350" if (change or 0) >= 0 else "42A5F5"
            chart_code = f"""
  s{idx+4}.addChart(pres.charts.LINE, [{{
    name: "{name}", labels: [{labels_js}], values: [{vals_js}]
  }}], {{
    x: 5.8, y: 1.05, w: 4.0, h: 2.8,
    chartColors: ["{chart_color}"],
    chartArea: {{ fill: {{ color: "FFFFFF" }}, roundedCorners: true }},
    catAxisLabelColor: "90A4AE", valAxisLabelColor: "90A4AE",
    valGridLine: {{ color: "E8EAF6", size: 0.5 }},
    catGridLine: {{ style: "none" }},
    lineSize: 2, lineSmooth: true,
    showLegend: false,
    showTitle: true, title: "이번주 주가 추이",
    titleFontSize: 12, titleColor: "37474F"
  }});"""

        stock_slides += f"""
  // ── 슬라이드 {idx+4}: {name} ──────────────────────────────────────────
  let s{idx+4} = pres.addSlide();
  s{idx+4}.background = {{ color: "F8FAFF" }};
  s{idx+4}.addShape(pres.shapes.RECTANGLE, {{ x: 0, y: 0, w: 10, h: 0.9, fill: {{ color: "0D1B4B" }} }});
  s{idx+4}.addText("{flag} {name} ({ticker})", {{
    x: 0.3, y: 0.05, w: 7, h: 0.5,
    fontSize: 22, bold: true, color: "FFFFFF", fontFace: "Malgun Gothic"
  }});
  s{idx+4}.addText("{market} | {today}", {{
    x: 0.3, y: 0.55, w: 7, h: 0.32,
    fontSize: 12, color: "90CAF9", fontFace: "Malgun Gothic"
  }});
  s{idx+4}.addShape(pres.shapes.RECTANGLE, {{
    x: 7.3, y: 0.08, w: 2.4, h: 0.75,
    fill: {{ color: "1565C0" }}
  }});
  s{idx+4}.addText("{close_str}", {{
    x: 7.3, y: 0.08, w: 2.4, h: 0.42,
    fontSize: 18, bold: true, color: "FFFFFF", fontFace: "Malgun Gothic", align: "center"
  }});
  s{idx+4}.addText("{change_str}", {{
    x: 7.3, y: 0.5, w: 2.4, h: 0.33,
    fontSize: 15, bold: true, color: "{clr}", fontFace: "Malgun Gothic", align: "center"
  }});
  s{idx+4}.addShape(pres.shapes.RECTANGLE, {{
    x: 0.3, y: 1.05, w: 5.3, h: 4.3,
    fill: {{ color: "FFFFFF" }},
    shadow: {{ type: "outer", color: "000000", blur: 6, offset: 2, angle: 135, opacity: 0.08 }}
  }});
  s{idx+4}.addText("AI 분석 요약", {{
    x: 0.4, y: 1.1, w: 5.1, h: 0.35,
    fontSize: 13, bold: true, color: "1565C0", fontFace: "Malgun Gothic"
  }});
  s{idx+4}.addText([
{brief_items}
  ], {{
    x: 0.4, y: 1.5, w: 5.1, h: 3.75,
    fontFace: "Malgun Gothic", fontSize: 12, color: "1A237E", valign: "top"
  }});
{chart_code}
  s{idx+4}.addText("({idx+1}/{len(stocks)}) {name}", {{
    x: 0.3, y: 5.3, w: 9.4, h: 0.28,
    fontSize: 10, color: "90A4AE", fontFace: "Malgun Gothic", align: "right"
  }});
"""

    # 마지막 슬라이드 — 면책 고지
    last_idx = len(stocks) + 4
    slide_last = f"""
  // ── 마지막 슬라이드: 면책 고지 ──────────────────────────────────────────
  let s_last = pres.addSlide();
  s_last.background = {{ color: "0D1B4B" }};
  s_last.addShape(pres.shapes.RECTANGLE, {{ x: 1.5, y: 1.5, w: 7, h: 0.08, fill: {{ color: "1565C0" }} }});
  s_last.addText("감사합니다", {{
    x: 0.5, y: 1.8, w: 9, h: 1.2,
    fontSize: 44, bold: true, color: "FFFFFF", fontFace: "Malgun Gothic", align: "center"
  }});
  s_last.addText("{_safe(persona_emoji + " " + persona_label)}", {{
    x: 0.5, y: 3.1, w: 9, h: 0.6,
    fontSize: 20, color: "FFD54F", fontFace: "Malgun Gothic", align: "center"
  }});
  s_last.addText("본 리포트는 AI 분석 기반이며 투자 결정은 본인 책임입니다.", {{
    x: 0.5, y: 4.8, w: 9, h: 0.4,
    fontSize: 13, color: "90A4AE", fontFace: "Malgun Gothic", align: "center"
  }});
  s_last.addText("기준일: {today}", {{
    x: 0.5, y: 5.2, w: 9, h: 0.3,
    fontSize: 12, color: "546E7A", fontFace: "Malgun Gothic", align: "center"
  }});
"""

    js_code = f"""
const pptxgen = require("pptxgenjs");
let pres = new pptxgen();
pres.layout = 'LAYOUT_16x9';
pres.author = 'Stock AI Briefing';
pres.title = '주식 AI 투자 브리핑 {today}';
pres.subject = '{_safe(persona_label)}';
{slide1}
{slide2}
{slide3}
{stock_slides}
{slide_last}

pres.writeFile({{ fileName: `{output_path}` }})
  .then(() => console.log('OK'))
  .catch(e => {{ console.error('ERR', e.message); process.exit(1); }});
"""

    with tempfile.NamedTemporaryFile(suffix=".js", mode="w", delete=False, encoding="utf-8") as f:
        f.write(js_code)
        js_file = f.name

    try:
        result = subprocess.run(
            ["node", js_file],
            capture_output=True, text=True, timeout=90,
            env={**os.environ, "NODE_PATH": "/home/claude/.npm-global/lib/node_modules"},
        )
        if result.returncode != 0:
            raise RuntimeError(f"Node 오류: {result.stderr[:400]}")
        return output_path
    finally:
        os.unlink(js_file)
