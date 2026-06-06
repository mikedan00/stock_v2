"""
report_generator.py — 페르소나 + RAG 컨텍스트 통합 리포트 생성
"""
from __future__ import annotations
from datetime import date
from llm_engine import call_llm
from analyst_personas import get_persona, persona_system_prompt


def _format_investor(investor: dict) -> str:
    if not investor:
        return "수급 데이터 없음"
    lines = []
    for k, label in [("외국인합계", "외국인"), ("기관합계", "기관"), ("개인", "개인")]:
        val = investor.get(k)
        if val is not None:
            lines.append(f"{label}: {val:+,}원")
    return " | ".join(lines) if lines else "수급 데이터 없음"


def _news_summary(news_list: list[dict], max_items: int = 5) -> str:
    if not news_list:
        return "뉴스 없음"
    return "\n".join(
        f"- [{i+1}] {n['title']} ({n.get('source','')})"
        for i, n in enumerate(news_list[:max_items])
    )


def _history_summary(history: dict) -> str:
    lines = []
    for period, label in [("week", "이번주"), ("month", "이번달"), ("6month", "6개월")]:
        df = history.get(period)
        if df is not None and not df.empty and "close" in df.columns:
            sp = float(df["close"].iloc[0])
            ep = float(df["close"].iloc[-1])
            ch = (ep - sp) / sp * 100 if sp else 0
            lines.append(f"{label}: {sp:,.0f}->>{ep:,.0f} ({ch:+.1f}%)")
        else:
            lines.append(f"{label}: 데이터 없음")
    return " | ".join(lines)


def generate_stock_brief(stock: dict, news: dict, persona_key: str = "A", rag_context: str = "") -> str:
    name = stock.get("name", stock["ticker"])
    close = stock.get("close")
    change = stock.get("change_rate")
    volume = stock.get("volume")
    market = stock.get("market", "")
    today = stock.get("date", date.today().isoformat())
    persona = get_persona(persona_key)
    currency = "원" if market == "KR" else "$"
    close_str = f"{close:,.0f}" if close else "N/A"
    change_str = f"{change:+.2f}%" if change is not None else "N/A"
    volume_str = f"{volume:,}" if volume else "N/A"

    rag_section = f"\n## 사용자 제공 추가 자료 (RAG)\n{rag_context[:2000]}\n" if rag_context.strip() else ""

    prompt = f"""
아래 데이터로 {name}({stock['ticker']}) 종목의 투자 브리핑을 작성하세요.
분석 관점: {persona['label']} - {persona['focus']}

## 기본 데이터 ({today})
종가: {close_str}{currency} | 등락률: {change_str} | 거래량: {volume_str}
수급: {_format_investor(stock.get('investor', {}))}

## 주가 추이
{_history_summary(stock.get('history', {}))}

## 국내 뉴스
{_news_summary(news.get('domestic', []))}

## 해외 뉴스
{_news_summary(news.get('international', []))}
{rag_section}
---
{persona['short']} 관점으로 다음 8항목을 포함한 브리핑을 작성하세요:
1. 종목 현황 요약 (핵심 판단 2~3문장)
2. 주가 추이 분석 (이번주/이번달/6개월 기술적 해석)
3. 수급 분석 (외인·기관·개인 동향)
4. 뉴스 이벤트 임팩트
5. 추가 자료 반영 분석 (RAG 자료 있을 경우)
6. 리스크 요인
7. 내일 매매 전략 (진입가·손절가·목표가 포함)
8. 내일 예상 방향 (상승/하락/횡보 + 예상 등락률%)
"""
    return call_llm(prompt, system=persona_system_prompt(persona_key), max_tokens=1800)


def generate_portfolio_brief(stocks: list[dict], all_news: list, persona_key: str = "A", rag_context: str = "") -> str:
    today = date.today().isoformat()
    persona = get_persona(persona_key)
    stock_lines = []
    for s in stocks:
        name = s.get("name", s["ticker"])
        close = s.get("close")
        change = s.get("change_rate")
        currency = "원" if s.get("market") == "KR" else "$"
        close_str = f"{close:,.0f}{currency}" if close else "N/A"
        change_str = f"{change:+.2f}%" if change is not None else "N/A"
        stock_lines.append(f"- {name}({s['ticker']}): {close_str} / {change_str}")

    rag_section = f"\n## 사용자 추가 자료\n{rag_context[:1500]}\n" if rag_context.strip() else ""

    prompt = f"""
오늘({today}) 포트폴리오 종합 전략 리포트를 작성하세요.
분석 관점: {persona['label']} - {persona['focus']}

## 포트폴리오 현황
{chr(10).join(stock_lines)}
{rag_section}
---
{persona['short']} 관점으로 다음을 포함한 종합 리포트를 작성하세요:
1. 오늘 시장 총평 (페르소나 시각)
2. 포트폴리오 종합 평가 (리스크·수익성)
3. 섹터별 테마별 분석
4. 내일 전체 투자 전략
5. Top Pick & Bottom Pick (이유 포함)
6. 리스크 관리 (손절·익절·포지션 조정)
7. 다음 주 중기 전략
"""
    return call_llm(prompt, system=persona_system_prompt(persona_key), max_tokens=2000)


def build_full_report_text(stocks, all_news, stock_briefs, portfolio_brief, persona_key="A") -> str:
    persona = get_persona(persona_key)
    today = date.today().isoformat()
    sep = "=" * 60
    sections = [
        f"주식 AI 투자 브리핑 리포트",
        f"기준일: {today}",
        f"분석 관점: {persona['label']} {persona['emoji']}",
        sep,
        "[ 포트폴리오 종합 전략 ]",
        portfolio_brief,
        sep,
    ]
    for stock, brief in zip(stocks, stock_briefs):
        name = stock.get("name", stock["ticker"])
        close = stock.get("close")
        change = stock.get("change_rate")
        currency = "원" if stock.get("market") == "KR" else "$"
        close_str = f"{close:,.0f}{currency}" if close else "N/A"
        change_str = f"{change:+.2f}%" if change is not None else "N/A"
        sections.append(f"[ {name}({stock['ticker']}) | {close_str} {change_str} ]")
        sections.append(brief)
        sections.append("-" * 40)
    sections += [sep, f"분석 AI: {persona['label']} {persona['emoji']}", "본 리포트는 AI 분석 기반이며 투자 결정은 본인 책임입니다."]
    return "\n\n".join(sections)
