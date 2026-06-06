"""
llm_engine.py — HuggingFace Router 연동
모델: google/gemma-4-26B-A4B-it:deepinfra (기본)
"""
from __future__ import annotations

import os
import requests
import config


def _get_token() -> str:
    return os.environ.get("HF_TOKEN") or config.HF_TOKEN


def _get_model() -> str:
    # 런타임 사이드바 선택 → 없으면 config 기본값
    return os.environ.get("HF_MODEL_OVERRIDE") or config.HF_ROUTER_MODEL


def call_llm(prompt: str, system: str = "", max_tokens: int = 2048) -> str:
    """HF Router OpenAI-compatible endpoint 호출. 실패 시 오류 메시지 반환."""
    token = _get_token()
    model = _get_model()

    if not token:
        return "[오류] HF_TOKEN이 설정되지 않았습니다. 사이드바에서 입력해주세요."

    headers = {
        "Authorization": f"Bearer {token}",
        "Content-Type": "application/json",
    }

    messages = []
    if system:
        messages.append({"role": "system", "content": system})
    messages.append({"role": "user", "content": prompt})

    payload = {
        "model": model,
        "messages": messages,
        "max_tokens": max_tokens,
        "temperature": 0.7,
        "stream": False,
    }

    try:
        resp = requests.post(
            config.HF_API_URL, headers=headers, json=payload, timeout=120
        )
        resp.raise_for_status()
        data = resp.json()
        return data["choices"][0]["message"]["content"]
    except requests.exceptions.HTTPError as e:
        body = e.response.text[:400] if e.response else ""
        return f"[LLM HTTP 오류 {e.response.status_code}] {body}"
    except requests.exceptions.Timeout:
        return "[LLM 오류] 요청 시간 초과 (120초). 모델 서버가 바쁩니다. 잠시 후 재시도해주세요."
    except Exception as e:
        return f"[LLM 오류] {type(e).__name__}: {e}"


EXPERT_SYSTEM = """당신은 15년 경력의 국내외 주식 투자 전문가이자 애널리스트입니다.
헤지펀드 포트폴리오 매니저 경험과 기술적 분석, 기본적 분석 모두에 정통합니다.
항상 한국어로 응답하며, 데이터에 근거한 명확하고 실행 가능한 인사이트를 제공합니다.
리포트는 전문적이지만 이해하기 쉽게 작성합니다."""
