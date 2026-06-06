"""
stock_data.py — pykrx + yfinance 주가 데이터 수집
종목명 입력 지원: '삼성전자' → '005930' 자동 변환
"""
from __future__ import annotations

import re
from datetime import date, timedelta
from typing import Optional

import pandas as pd

# ── 날짜 헬퍼 ────────────────────────────────────────────────────────────────

def last_trading_day(d: Optional[date] = None) -> date:
    d = d or date.today()
    while d.weekday() >= 5:
        d -= timedelta(days=1)
    return d


def date_range_for_period(period: str) -> tuple[date, date]:
    end = last_trading_day()
    if period == "week":
        start = end - timedelta(weeks=1)
    elif period == "month":
        start = end - timedelta(days=30)
    else:
        start = end - timedelta(days=183)
    return start, end


# ── 종목명 → 코드 변환 ───────────────────────────────────────────────────────

# 자주 쓰는 종목명 매핑 (pykrx 호출 실패 대비 fallback)
_KR_NAME_MAP = {
    "삼성전자": "005930", "sk하이닉스": "000660", "sk하이닉스": "000660",
    "하이닉스": "000660", "lg에너지솔루션": "373220", "삼성바이오로직스": "207940",
    "삼성바이오": "207940", "현대차": "005380", "현대자동차": "005380",
    "기아": "000270", "기아차": "000270", "셀트리온": "068270",
    "naver": "035420", "네이버": "035420", "카카오": "035720",
    "lg화학": "051910", "삼성sdi": "006400", "포스코홀딩스": "005490",
    "포스코": "005490", "kb금융": "105560", "신한지주": "055550",
    "하나금융지주": "086790", "우리금융지주": "316140", "카카오뱅크": "323410",
    "크래프톤": "259960", "넷마블": "251270", "엔씨소프트": "036570",
    "한국전력": "015760", "한전": "015760", "두산에너빌리티": "034020",
    "삼성물산": "028260", "현대모비스": "012330", "lg전자": "066570",
    "sk텔레콤": "017670", "kt": "030200", "lg유플러스": "032640",
    "카카오페이": "377300", "토스": "403300", "kakao": "035720",
}


def name_to_ticker(name: str) -> Optional[str]:
    """
    종목명 → 6자리 pykrx 코드 변환.
    1) 하드코딩 맵 → 2) pykrx 전체 종목 검색
    """
    key = name.strip().lower().replace(" ", "")
    if key in _KR_NAME_MAP:
        return _KR_NAME_MAP[key]

    try:
        from pykrx import stock as px
        # KOSPI + KOSDAQ 전체 티커 목록
        for market in ("KOSPI", "KOSDAQ"):
            tickers = px.get_market_ticker_list(market=market)
            for t in tickers:
                try:
                    n = px.get_market_ticker_name(t)
                    if n and n.strip() == name.strip():
                        return t
                except Exception:
                    continue
    except Exception:
        pass
    return None


# ── 입력 정규화 ──────────────────────────────────────────────────────────────

def normalize_ticker(raw: str) -> tuple[str, str]:
    """
    입력값을 (ticker, display_input) 로 반환.
    - '삼성전자' → ('005930', '삼성전자')
    - '005930'   → ('005930', '005930')
    - 'AAPL'     → ('AAPL',   'AAPL')
    """
    raw = raw.strip()

    # 이미 6자리 코드
    if re.fullmatch(r"\d{6}", raw):
        return raw, raw

    # 영문 티커 (영숫자+점, 1~10자)
    if re.fullmatch(r"[A-Za-z][A-Za-z0-9.\-]{0,9}", raw):
        return raw.upper(), raw.upper()

    # 한글/한자 종목명 → 코드 검색
    code = name_to_ticker(raw)
    if code:
        return code, raw  # (코드, 원래 입력명)

    # 변환 실패 → 원본 그대로 반환 (오류는 fetch 단계에서 처리)
    return raw, raw


def is_kr_code(ticker: str) -> bool:
    return bool(re.fullmatch(r"\d{6}", ticker.strip()))


# ── 국내 종목 데이터 (pykrx) ─────────────────────────────────────────────────

def fetch_kr_stock(ticker: str, display_name: str = "") -> dict:
    from pykrx import stock as px

    td = last_trading_day()
    td_str = td.strftime("%Y%m%d")

    try:
        name = px.get_market_ticker_name(ticker) or display_name or ticker
    except Exception:
        name = display_name or ticker

    # 종가
    try:
        ohlcv = px.get_market_ohlcv(td_str, td_str, ticker)
        if ohlcv.empty:
            w_start = (td - timedelta(days=7)).strftime("%Y%m%d")
            ohlcv = px.get_market_ohlcv(w_start, td_str, ticker)
        row = ohlcv.iloc[-1] if not ohlcv.empty else None
    except Exception:
        row = None

    close = float(row["종가"]) if row is not None and "종가" in row else None
    change_rate = float(row["등락률"]) if row is not None and "등락률" in row else None
    volume = int(row["거래량"]) if row is not None and "거래량" in row else None

    # 수급
    investor = {}
    try:
        inv_df = px.get_market_trading_value_by_investor(td_str, td_str, ticker)
        if not inv_df.empty:
            for label in ["외국인합계", "기관합계", "개인"]:
                if label in inv_df.index:
                    investor[label] = int(inv_df.loc[label, "순매수"])
    except Exception:
        pass

    # 기간별 OHLCV
    history = {}
    for period in ("week", "month", "6month"):
        start, end = date_range_for_period(period)
        try:
            df = px.get_market_ohlcv(
                start.strftime("%Y%m%d"), end.strftime("%Y%m%d"), ticker
            )
            history[period] = (
                df[["종가"]].rename(columns={"종가": "close"}) if not df.empty
                else pd.DataFrame()
            )
        except Exception:
            history[period] = pd.DataFrame()

    return {
        "ticker": ticker,
        "name": name,
        "market": "KR",
        "close": close,
        "change_rate": change_rate,
        "volume": volume,
        "date": td.isoformat(),
        "investor": investor,
        "history": history,
    }


# ── 해외 종목 데이터 (yfinance) ──────────────────────────────────────────────

def fetch_us_stock(ticker: str) -> dict:
    import yfinance as yf

    td = last_trading_day()
    yf_ticker = ticker.strip().upper()
    obj = yf.Ticker(yf_ticker)

    try:
        info = obj.info
        name = info.get("longName") or info.get("shortName") or yf_ticker
    except Exception:
        name = yf_ticker

    try:
        hist_1d = obj.history(period="5d")
        row = hist_1d.iloc[-1] if not hist_1d.empty else None
    except Exception:
        row = None
        hist_1d = pd.DataFrame()

    close = float(row["Close"]) if row is not None else None
    prev_close = (
        float(hist_1d.iloc[-2]["Close"])
        if row is not None and len(hist_1d) > 1 else None
    )
    change_rate = (
        (close - prev_close) / prev_close * 100
        if close and prev_close else None
    )
    volume = int(row["Volume"]) if row is not None else None

    history = {}
    for period, yf_period in [("week", "5d"), ("month", "1mo"), ("6month", "6mo")]:
        try:
            df = obj.history(period=yf_period)[["Close"]].rename(columns={"Close": "close"})
            history[period] = df
        except Exception:
            history[period] = pd.DataFrame()

    return {
        "ticker": yf_ticker,
        "name": name,
        "market": "US",
        "close": close,
        "change_rate": change_rate,
        "volume": volume,
        "date": td.isoformat(),
        "investor": {},
        "history": history,
    }


# ── 통합 수집 ────────────────────────────────────────────────────────────────

def fetch_stock(raw_input: str) -> dict:
    ticker, display = normalize_ticker(raw_input)
    if is_kr_code(ticker):
        return fetch_kr_stock(ticker, display_name=display)
    else:
        return fetch_us_stock(ticker)


def fetch_all_stocks(raw_inputs: list[str]) -> list[dict]:
    results = []
    for raw in raw_inputs[:10]:
        try:
            results.append(fetch_stock(raw))
        except Exception as e:
            results.append({
                "ticker": raw, "name": raw, "market": "?",
                "close": None, "change_rate": None, "volume": None,
                "date": last_trading_day().isoformat(),
                "investor": {}, "history": {},
                "error": str(e),
            })
    return results
