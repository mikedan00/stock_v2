"""
config.py — 환경변수 및 상수 관리
"""
import os
from dotenv import load_dotenv

load_dotenv()

# ── HuggingFace Router ───────────────────────────────────────────────────────
HF_TOKEN = os.getenv("HF_TOKEN", "")
HF_API_URL = "https://router.huggingface.co/v1/chat/completions"

# 기본 모델 (지시대로 정확히)
HF_ROUTER_MODEL = "google/gemma-4-26B-A4B-it:deepinfra"

# 후보 모델 목록 (사이드바 드롭다운용)
HF_MODEL_CANDIDATES = [
    "google/gemma-4-26B-A4B-it:deepinfra",
    "google/gemma-4-26B-A4B-it:novita",
    "google/gemma-4-31B-it:deepinfra",
    "google/gemma-4-31B-it:together",
    "Qwen/Qwen3.5-9B:together",
    "Qwen/Qwen2.5-7B-Instruct:together",
]

# ── Gmail SMTP ───────────────────────────────────────────────────────────────
GMAIL_USER = os.getenv("GMAIL_USER", "")
GMAIL_APP_PASSWORD = os.getenv("GMAIL_APP_PASSWORD", "")

# ── 주식 설정 ────────────────────────────────────────────────────────────────
MAX_STOCKS = 10
