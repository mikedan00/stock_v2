"""
news_fetcher.py — 네이버/구글/yfinance 뉴스 수집 + 영→한 번역
"""
from __future__ import annotations

import re
import time
from datetime import datetime, date
from typing import Optional

import feedparser
import requests
from deep_translator import GoogleTranslator


# ── 번역 헬퍼 ────────────────────────────────────────────────────────────────

def translate_to_korean(text: str) -> str:
    """영어 텍스트를 한국어로 번역. 실패 시 원문 반환."""
    if not text:
        return text
    try:
        # 한글이 이미 포함된 경우 번역 생략
        if re.search(r"[\uAC00-\uD7A3]", text):
            return text
        return GoogleTranslator(source="auto", target="ko").translate(text[:500])
    except Exception:
        return text


# ── RSS 파싱 공통 ─────────────────────────────────────────────────────────────

def _parse_rss(url: str, max_items: int = 10, translate: bool = False) -> list[dict]:
    """RSS URL을 파싱하여 뉴스 항목 리스트 반환."""
    try:
        headers = {"User-Agent": "Mozilla/5.0 (compatible; StockBriefBot/1.0)"}
        resp = requests.get(url, headers=headers, timeout=8)
        feed = feedparser.parse(resp.text)
    except Exception:
        return []

    items = []
    for entry in feed.entries[:max_items]:
        title = re.sub(r"<[^>]+>", "", entry.get("title", "")).strip()
        link = entry.get("link", "")
        published = entry.get("published", "") or entry.get("updated", "")
        summary = re.sub(r"<[^>]+>", "", entry.get("summary", "")).strip()[:200]

        if translate:
            title = translate_to_korean(title)
            summary = translate_to_korean(summary)

        items.append({
            "title": title,
            "link": link,
            "published": published,
            "summary": summary,
        })
        time.sleep(0.1)   # 과도한 요청 방지

    return items


# ── 네이버 뉴스 ──────────────────────────────────────────────────────────────

def fetch_naver_news(query: str, max_items: int = 10) -> list[dict]:
    """네이버 뉴스 RSS (한국어)."""
    url = f"https://search.naver.com/rss?where=news&query={requests.utils.quote(query)}"
    results = _parse_rss(url, max_items)
    for r in results:
        r["source"] = "네이버뉴스"
    return results


# ── 구글 뉴스 (한국어) ───────────────────────────────────────────────────────

def fetch_google_news_kr(query: str, max_items: int = 10) -> list[dict]:
    """구글 뉴스 한국어 RSS."""
    url = (
        f"https://news.google.com/rss/search"
        f"?q={requests.utils.quote(query)}&hl=ko&gl=KR&ceid=KR:ko"
    )
    results = _parse_rss(url, max_items)
    for r in results:
        r["source"] = "구글뉴스(KR)"
    return results


# ── 구글 뉴스 (영어 → 번역) ─────────────────────────────────────────────────

def fetch_google_news_en(query: str, max_items: int = 10) -> list[dict]:
    """구글 뉴스 영어 RSS → 한국어 번역."""
    url = (
        f"https://news.google.com/rss/search"
        f"?q={requests.utils.quote(query)}&hl=en&gl=US&ceid=US:en"
    )
    results = _parse_rss(url, max_items, translate=True)
    for r in results:
        r["source"] = "구글뉴스(EN→KR)"
    return results


# ── yfinance 뉴스 ─────────────────────────────────────────────────────────────

def fetch_yfinance_news(ticker: str, max_items: int = 10) -> list[dict]:
    """yfinance Ticker.news → 한국어 번역."""
    try:
        import yfinance as yf
        obj = yf.Ticker(ticker)
        raw = obj.news or []
    except Exception:
        return []

    items = []
    for n in raw[:max_items]:
        title = n.get("title", "")
        link = n.get("link", "") or n.get("canonicalUrl", {}).get("url", "")
        published_ts = n.get("providerPublishTime", 0)
        published = (
            datetime.fromtimestamp(published_ts).strftime("%Y-%m-%d %H:%M")
            if published_ts else ""
        )
        summary = n.get("summary", "")[:200]

        title_kr = translate_to_korean(title)
        summary_kr = translate_to_korean(summary) if summary else ""

        items.append({
            "title": title_kr,
            "link": link,
            "published": published,
            "summary": summary_kr,
            "source": "yfinance",
        })

    return items


# ── 통합 뉴스 수집 ────────────────────────────────────────────────────────────

def fetch_all_news(
    stock: dict,
    kr_per_source: int = 10,
    en_per_source: int = 10,
) -> dict:
    """
    한 종목에 대한 전체 뉴스 수집.
    반환: { "domestic": [...], "international": [...] }
    """
    name = stock.get("name", stock["ticker"])
    ticker = stock["ticker"]
    market = stock.get("market", "KR")

    # 검색어 결정
    if market == "KR":
        kr_query = f"{name} 주식"
        en_query = f"{name} stock"
    else:
        kr_query = f"{name} 주식"
        en_query = f"{ticker} stock"

    # 국내 뉴스
    domestic = []
    domestic += fetch_naver_news(kr_query, kr_per_source)
    domestic += fetch_google_news_kr(kr_query, kr_per_source)
    # 중복 제거 (title 기준)
    seen = set()
    domestic_unique = []
    for item in domestic:
        if item["title"] not in seen:
            seen.add(item["title"])
            domestic_unique.append(item)
    domestic_unique = domestic_unique[:10]

    # 해외 뉴스
    international = []
    international += fetch_google_news_en(en_query, en_per_source)
    if market == "US":
        international += fetch_yfinance_news(ticker, en_per_source)
    seen_en = set()
    international_unique = []
    for item in international:
        if item["title"] not in seen_en:
            seen_en.add(item["title"])
            international_unique.append(item)
    international_unique = international_unique[:10]

    return {
        "domestic": domestic_unique,
        "international": international_unique,
    }
