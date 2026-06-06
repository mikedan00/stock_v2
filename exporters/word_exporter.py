"""
exporters/word_exporter.py
최종 리포트를 Word(.docx)로 내보내기 — docx npm 패키지 사용
"""
from __future__ import annotations
import os, json, subprocess, tempfile, re
from datetime import date
from pathlib import Path


def _safe_text(text: str) -> str:
    """JS 문자열에 안전하게 삽입될 수 있도록 이스케이프."""
    return (text or "")[:8000].replace("\\", "\\\\").replace("`", "'").replace("${", "\\${")


def _split_report_sections(report_text: str) -> list[dict]:
    """리포트 텍스트를 섹션으로 분할."""
    sections = []
    lines = report_text.split("\n")
    current_section = {"heading": None, "lines": []}

    for line in lines:
        stripped = line.strip()
        # 섹션 구분선
        if stripped.startswith("=") and len(stripped) > 10:
            if current_section["lines"]:
                sections.append(current_section)
            current_section = {"heading": None, "lines": []}
        elif stripped.startswith("-") and len(stripped) > 10:
            continue  # 구분선 무시
        elif stripped.startswith("【") or stripped.startswith("["):
            # 새 섹션 시작
            if current_section["lines"] or current_section["heading"]:
                sections.append(current_section)
            current_section = {"heading": stripped, "lines": []}
        else:
            current_section["lines"].append(line)

    if current_section["lines"] or current_section["heading"]:
        sections.append(current_section)

    return sections


def export_word(
    report_text: str,
    stocks: list[dict],
    persona_label: str,
    output_path: str,
) -> str:
    """
    리포트를 Word 파일로 저장.
    반환: 저장된 파일 경로
    """
    today = date.today().isoformat()
    sections = _split_report_sections(report_text)

    # 섹션별 JS 코드 생성
    section_js_parts = []
    for sec in sections:
        heading = sec.get("heading") or ""
        body = "\n".join(sec["lines"]).strip()

        if heading:
            safe_h = _safe_text(heading.strip("【】[]"))
            section_js_parts.append(
                f'new Paragraph({{ heading: HeadingLevel.HEADING_2, '
                f'children: [new TextRun({{ text: `{safe_h}`, bold: true, color: "1565C0" }})] }}),'
            )

        if body:
            for para_line in body.split("\n"):
                para_line = para_line.strip()
                if not para_line:
                    section_js_parts.append(
                        'new Paragraph({ children: [new TextRun("")] }),'
                    )
                    continue
                safe_line = _safe_text(para_line)
                bold = para_line.startswith("**") or para_line.startswith("##")
                clean_line = safe_line.replace("**", "").replace("##", "").strip()
                if bold and clean_line:
                    section_js_parts.append(
                        f'new Paragraph({{ children: [new TextRun({{ text: `{clean_line}`, bold: true }})] }}),'
                    )
                elif clean_line:
                    section_js_parts.append(
                        f'new Paragraph({{ children: [new TextRun(`{clean_line}`)] }}),'
                    )

    sections_code = "\n      ".join(section_js_parts[:300])  # 최대 300 단락

    # 주가 테이블 rows
    table_rows_code = ""
    for s in stocks[:10]:
        name = _safe_text(s.get("name", s["ticker"]))
        ticker = _safe_text(s["ticker"])
        close = s.get("close")
        change = s.get("change_rate")
        currency = "원" if s.get("market") == "KR" else "$"
        close_str = f"{close:,.0f}{currency}" if close else "N/A"
        change_str = f"{change:+.2f}%" if change is not None else "N/A"
        market = s.get("market", "")
        color = "C62828" if (change or 0) > 0 else "1565C0" if (change or 0) < 0 else "546E7A"

        table_rows_code += f"""
    new TableRow({{
      children: [
        new TableCell({{ borders, width: {{ size: 2500, type: WidthType.DXA }}, margins: cellMargins,
          children: [new Paragraph({{ children: [new TextRun({{ text: `{name}`, bold: true }})] }})] }}),
        new TableCell({{ borders, width: {{ size: 1500, type: WidthType.DXA }}, margins: cellMargins,
          children: [new Paragraph({{ children: [new TextRun(`{ticker}`)] }})] }}),
        new TableCell({{ borders, width: {{ size: 1500, type: WidthType.DXA }}, margins: cellMargins,
          children: [new Paragraph({{ alignment: AlignmentType.RIGHT, children: [new TextRun(`{close_str}`)] }})] }}),
        new TableCell({{ borders, width: {{ size: 1360, type: WidthType.DXA }}, margins: cellMargins,
          children: [new Paragraph({{ alignment: AlignmentType.RIGHT, children: [new TextRun({{ text: `{change_str}`, color: "{color}" }})] }})] }}),
        new TableCell({{ borders, width: {{ size: 700, type: WidthType.DXA }}, margins: cellMargins,
          children: [new Paragraph({{ children: [new TextRun(`{market}`)] }})] }}),
      ]
    }}),"""

    js_code = f"""
const fs = require('fs');
const {{
  Document, Packer, Paragraph, TextRun, Table, TableRow, TableCell,
  HeadingLevel, AlignmentType, WidthType, BorderStyle, ShadingType,
  PageNumber, Header, Footer
}} = require('docx');

const border = {{ style: BorderStyle.SINGLE, size: 1, color: "CCCCCC" }};
const borders = {{ top: border, bottom: border, left: border, right: border }};
const cellMargins = {{ top: 80, bottom: 80, left: 120, right: 120 }};

const doc = new Document({{
  styles: {{
    default: {{ document: {{ run: {{ font: "Malgun Gothic", size: 22 }} }} }},
    paragraphStyles: [
      {{ id: "Heading1", name: "Heading 1", basedOn: "Normal", next: "Normal", quickFormat: true,
        run: {{ size: 36, bold: true, font: "Malgun Gothic", color: "0D1B4B" }},
        paragraph: {{ spacing: {{ before: 300, after: 200 }}, outlineLevel: 0 }} }},
      {{ id: "Heading2", name: "Heading 2", basedOn: "Normal", next: "Normal", quickFormat: true,
        run: {{ size: 26, bold: true, font: "Malgun Gothic", color: "1565C0" }},
        paragraph: {{ spacing: {{ before: 200, after: 120 }}, outlineLevel: 1 }} }},
    ]
  }},
  sections: [{{
    properties: {{
      page: {{
        size: {{ width: 11906, height: 16838 }},
        margin: {{ top: 1200, right: 1200, bottom: 1200, left: 1440 }}
      }}
    }},
    headers: {{
      default: new Header({{
        children: [new Paragraph({{
          border: {{ bottom: {{ style: BorderStyle.SINGLE, size: 6, color: "1565C0", space: 1 }} }},
          children: [new TextRun({{ text: "📊 주식 AI 투자 브리핑 리포트", bold: true, size: 20, color: "1565C0" }})]
        }})]
      }})
    }},
    footers: {{
      default: new Footer({{
        children: [new Paragraph({{
          border: {{ top: {{ style: BorderStyle.SINGLE, size: 4, color: "CCCCCC", space: 1 }} }},
          alignment: AlignmentType.CENTER,
          children: [new TextRun({{ text: "본 리포트는 AI 분석 기반이며 투자 결정은 본인 책임입니다. | {today}", size: 16, color: "888888" }})]
        }})]
      }})
    }},
    children: [
      new Paragraph({{
        heading: HeadingLevel.HEADING_1,
        children: [new TextRun({{ text: "📊 주식 AI 투자 브리핑 리포트", bold: true, color: "0D1B4B" }})]
      }}),
      new Paragraph({{ children: [new TextRun({{ text: "기준일: {today}  |  분석관점: {_safe_text(persona_label)}", size: 20, color: "546E7A" }})] }}),
      new Paragraph({{ children: [new TextRun("")] }}),

      new Paragraph({{ heading: HeadingLevel.HEADING_2, children: [new TextRun({{ text: "📈 종목별 주가 현황", bold: true, color: "1565C0" }})] }}),
      new Table({{
        width: {{ size: 7560, type: WidthType.DXA }},
        columnWidths: [2500, 1500, 1500, 1360, 700],
        rows: [
          new TableRow({{
            tableHeader: true,
            children: [
              new TableCell({{ borders, width: {{ size: 2500, type: WidthType.DXA }}, margins: cellMargins,
                shading: {{ fill: "1565C0", type: ShadingType.CLEAR }},
                children: [new Paragraph({{ children: [new TextRun({{ text: "종목명", bold: true, color: "FFFFFF" }})] }})] }}),
              new TableCell({{ borders, width: {{ size: 1500, type: WidthType.DXA }}, margins: cellMargins,
                shading: {{ fill: "1565C0", type: ShadingType.CLEAR }},
                children: [new Paragraph({{ children: [new TextRun({{ text: "코드/티커", bold: true, color: "FFFFFF" }})] }})] }}),
              new TableCell({{ borders, width: {{ size: 1500, type: WidthType.DXA }}, margins: cellMargins,
                shading: {{ fill: "1565C0", type: ShadingType.CLEAR }},
                children: [new Paragraph({{ alignment: AlignmentType.RIGHT, children: [new TextRun({{ text: "종가", bold: true, color: "FFFFFF" }})] }})] }}),
              new TableCell({{ borders, width: {{ size: 1360, type: WidthType.DXA }}, margins: cellMargins,
                shading: {{ fill: "1565C0", type: ShadingType.CLEAR }},
                children: [new Paragraph({{ alignment: AlignmentType.RIGHT, children: [new TextRun({{ text: "등락률", bold: true, color: "FFFFFF" }})] }})] }}),
              new TableCell({{ borders, width: {{ size: 700, type: WidthType.DXA }}, margins: cellMargins,
                shading: {{ fill: "1565C0", type: ShadingType.CLEAR }},
                children: [new Paragraph({{ children: [new TextRun({{ text: "시장", bold: true, color: "FFFFFF" }})] }})] }}),
            ]
          }}),
          {table_rows_code}
        ]
      }}),
      new Paragraph({{ children: [new TextRun("")] }}),
      {sections_code}
    ]
  }}]
}});

Packer.toBuffer(doc).then(buf => {{
  fs.writeFileSync(`{output_path}`, buf);
  console.log('OK');
}}).catch(e => {{
  console.error('ERR', e.message);
  process.exit(1);
}});
"""

    with tempfile.NamedTemporaryFile(suffix=".js", mode="w", delete=False, encoding="utf-8") as f:
        f.write(js_code)
        js_file = f.name

    try:
        result = subprocess.run(
            ["node", js_file],
            capture_output=True, text=True, timeout=60,
            env={**os.environ, "NODE_PATH": "/home/claude/.npm-global/lib/node_modules"},
        )
        if result.returncode != 0:
            raise RuntimeError(f"Node 오류: {result.stderr[:300]}")
        return output_path
    finally:
        os.unlink(js_file)
