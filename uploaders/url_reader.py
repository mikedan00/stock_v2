"""
uploaders/url_reader.py
사용자 입력 URL에서 텍스트 추출
"""
from __future__ import annotations
import re
import requests
from urllib.parse import urlparse


HEADERS = {
    "User-Agent": (
        "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
        "AppleWebKit/537.36 (KHTML, like Gecko) "
        "Chrome/124.0 Safari/537.36"
    ),
    "Accept-Language": "ko-KR,ko;q=0.9,en;q=0.8",
}


def fetch_url_text(url: str, max_chars: int = 5000) -> dict:
    """
    URL에서 본문 텍스트 추출.
    반환: { "url": str, "title": str, "text": str, "error": str|None }
    """
    url = url.strip()
    if not url.startswith("http"):
        url = "https://" + url

    result = {"url": url, "title": "", "text": "", "error": None}

    try:
        resp = requests.get(url, headers=HEADERS, timeout=12, allow_redirects=True)
        resp.raise_for_status()
        resp.encoding = resp.apparent_encoding or "utf-8"
        html = resp.text
    except requests.exceptions.Timeout:
        result["error"] = "요청 시간 초과 (12초)"
        return result
    except requests.exceptions.HTTPError as e:
        result["error"] = f"HTTP {e.response.status_code}"
        return result
    except Exception as e:
        result["error"] = str(e)[:100]
        return result

    try:
        from bs4 import BeautifulSoup
        soup = BeautifulSoup(html, "html.parser")

        # 제목
        title_tag = soup.find("title")
        result["title"] = title_tag.get_text(strip=True) if title_tag else urlparse(url).netloc

        # 불필요 태그 제거
        for tag in soup(["script", "style", "nav", "footer", "header",
                          "aside", "form", "iframe", "noscript", "ads"]):
            tag.decompose()

        # 본문 후보 순서로 탐색
        body = None
        for selector in ["article", "main", ".article-body", ".content",
                          ".post-content", "#article", "#content", "body"]:
            el = soup.select_one(selector)
            if el:
                body = el
                break

        raw_text = (body or soup).get_text(separator="\n")
        # 연속 빈줄 제거
        lines = [l.strip() for l in raw_text.splitlines() if l.strip()]
        text = "\n".join(lines)

        result["text"] = text[:max_chars]
        if len(text) > max_chars:
            result["text"] += f"\n...[총 {len(text)}자 중 {max_chars}자 표시]"

    except Exception as e:
        result["error"] = f"HTML 파싱 오류: {e}"

    return result


def fetch_multiple_urls(urls: list[str], max_chars_each: int = 3000) -> list[dict]:
    results = []
    for url in urls:
        if url.strip():
            results.append(fetch_url_text(url, max_chars_each))
    return results
