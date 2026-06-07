"""
app.py — 주식 AI 브리핑 v2 (업그레이드)
신기능: URL 입력, 파일 업로드, RAG 분석, 5가지 페르소나, Word/PPT 내보내기
"""
import sys, os, tempfile
sys.path.insert(0, os.path.dirname(__file__))

import streamlit as st
import plotly.graph_objects as go
import pandas as pd
import numpy as np
from datetime import date

# ── 페이지 설정 ──────────────────────────────────────────────────────────────
st.set_page_config(
    page_title="주식 AI 브리핑 v2",
    page_icon="📊",
    layout="wide",
    initial_sidebar_state="expanded",
)

st.markdown("""
<style>
@import url('https://fonts.googleapis.com/css2?family=Noto+Sans+KR:wght@300;400;500;700;900&family=JetBrains+Mono:wght@400;600&display=swap');
html,body,[class*="css"]{font-family:'Noto Sans KR',sans-serif;}
.stApp{background:linear-gradient(135deg,#0a0e1a 0%,#0f1629 50%,#0a0e1a 100%);}
[data-testid="stSidebar"]{background:linear-gradient(180deg,#0d1117 0%,#161b2e 100%);border-right:1px solid #1e3a5f;}
.metric-card{background:linear-gradient(135deg,#111827,#1a2540);border:1px solid #1e3a5f;border-radius:12px;padding:16px 20px;margin:6px 0;box-shadow:0 4px 16px rgba(0,0,0,.4);}
.metric-card h3{color:#64b5f6;font-size:.78rem;font-weight:500;letter-spacing:1px;text-transform:uppercase;margin:0 0 6px 0;}
.metric-card .value{font-family:'JetBrains Mono',monospace;font-size:1.5rem;font-weight:700;color:#e2e8f0;margin:0;}
.news-item{background:#111827;border-left:3px solid #1565c0;border-radius:0 8px 8px 0;padding:10px 14px;margin:6px 0;font-size:.87rem;color:#cfd8dc;line-height:1.5;}
.news-item a{color:#90caf9;text-decoration:none;}
.news-item a:hover{text-decoration:underline;}
.news-src{font-size:.74rem;color:#546e7a;margin-top:3px;}
.report-box{background:#0d1117;border:1px solid #1e3a5f;border-radius:12px;padding:24px;font-size:.91rem;line-height:1.9;color:#cfd8dc;white-space:pre-wrap;font-family:'Noto Sans KR',sans-serif;}
.page-header{background:linear-gradient(135deg,#0d1b4b,#1a237e,#0d47a1);border-radius:16px;padding:26px 32px;margin-bottom:20px;border:1px solid #1565c0;box-shadow:0 8px 32px rgba(13,71,161,.3);}
.page-header h1{color:#fff;font-size:1.75rem;font-weight:900;margin:0;}
.page-header p{color:#90caf9;margin:5px 0 0 0;font-size:.93rem;}
.stTabs [data-baseweb="tab-list"]{background:#111827;border-radius:8px;padding:4px;gap:4px;}
.stTabs [data-baseweb="tab"]{color:#90a4ae;border-radius:6px;font-weight:500;}
.stTabs [aria-selected="true"]{background:#1565c0!important;color:white!important;}
.stButton>button{background:linear-gradient(135deg,#1565c0,#0d47a1);color:white;border:none;border-radius:8px;font-weight:600;padding:10px 20px;font-family:'Noto Sans KR',sans-serif;transition:all .2s;}
.stButton>button:hover{background:linear-gradient(135deg,#1976d2,#1565c0);box-shadow:0 4px 16px rgba(21,101,192,.4);transform:translateY(-1px);}
.section-title{color:#64b5f6;font-size:1.0rem;font-weight:700;letter-spacing:.5px;margin:20px 0 10px 0;padding-bottom:6px;border-bottom:1px solid #1e3a5f;}
.stTextInput input,.stTextArea textarea{background:#111827!important;color:#e2e8f0!important;border-color:#1e3a5f!important;border-radius:8px!important;}
.badge-up{background:rgba(239,83,80,.15);color:#ef5350;border:1px solid rgba(239,83,80,.3);border-radius:4px;padding:2px 8px;font-size:.8rem;font-weight:600;}
.badge-down{background:rgba(66,165,245,.15);color:#42a5f5;border:1px solid rgba(66,165,245,.3);border-radius:4px;padding:2px 8px;font-size:.8rem;font-weight:600;}
.badge-flat{background:rgba(144,164,174,.15);color:#90a4ae;border:1px solid rgba(144,164,174,.3);border-radius:4px;padding:2px 8px;font-size:.8rem;}
.model-badge{background:rgba(21,101,192,.2);color:#64b5f6;border:1px solid #1565c0;border-radius:6px;padding:4px 10px;font-size:.76rem;font-family:'JetBrains Mono',monospace;}
.persona-card{background:linear-gradient(135deg,#0d1b4b,#1a237e);border:2px solid #1565c0;border-radius:12px;padding:14px 18px;margin:4px 0;}
.persona-card .ptitle{color:#FFD54F;font-weight:700;font-size:1.0rem;}
.persona-card .pfocus{color:#90caf9;font-size:.82rem;margin-top:4px;line-height:1.5;}
.rag-box{background:#0a1628;border:1px dashed #1565c0;border-radius:10px;padding:14px;margin:8px 0;font-size:.83rem;color:#90a4ae;}
.url-item{background:#111827;border-radius:8px;padding:10px 14px;margin:5px 0;font-size:.85rem;}
.url-item .utitle{color:#90caf9;font-weight:600;}
.url-item .ustatus-ok{color:#66BB6A;font-size:.78rem;}
.url-item .ustatus-err{color:#EF5350;font-size:.78rem;}
.export-btn-area{background:#0d1117;border:1px solid #1e3a5f;border-radius:12px;padding:20px;margin:10px 0;}
</style>
""", unsafe_allow_html=True)

# ── 세션 상태 초기화 ──────────────────────────────────────────────────────────
for k, v in {
    "stocks_data": [], "news_data": {}, "stock_briefs": {},
    "portfolio_brief": "", "full_report": "",
    "data_loaded": False, "briefs_generated": False,
    "rag_docs": [], "url_results": [],
    "persona_key": "A",
}.items():
    if k not in st.session_state:
        st.session_state[k] = v

# ── 사이드바 ──────────────────────────────────────────────────────────────────
with st.sidebar:
    st.markdown("## ⚙️ 설정")
    st.markdown("---")

    st.markdown("### 🤖 AI 엔진")
    hf_token = st.text_input("HuggingFace Token", type="password",
        value=os.getenv("HF_TOKEN",""), placeholder="hf_xxxxxxxxxxxx",
        help="https://huggingface.co/settings/tokens")
    if hf_token:
        os.environ["HF_TOKEN"] = hf_token

    import config as cfg
    selected_model = st.selectbox("모델 선택", options=cfg.HF_MODEL_CANDIDATES, index=0)
    os.environ["HF_MODEL_OVERRIDE"] = selected_model
    st.markdown(f'<div class="model-badge">▶ {selected_model}</div>', unsafe_allow_html=True)

    st.markdown("---")
    st.markdown("### 📧 Gmail 발송")
    gmail_user = st.text_input("발신 Gmail", value=os.getenv("GMAIL_USER",""), placeholder="your@gmail.com")
    gmail_pw = st.text_input("앱 비밀번호", type="password", value=os.getenv("GMAIL_APP_PASSWORD",""))
    to_email = st.text_input("수신 이메일", placeholder="recipient@email.com")

    st.markdown("---")
    st.markdown("""
    <div style="font-size:.74rem;color:#546e7a;line-height:1.7;">
    <b style="color:#64b5f6;">종목 입력 예시</b><br>
    🇰🇷 종목명: 삼성전자, SK하이닉스<br>
    🇰🇷 코드: 005930, 000660<br>
    🇺🇸 티커: AAPL, NVDA, TSLA<br>
    ※ 최대 10종목
    </div>""", unsafe_allow_html=True)

# ── 헤더 ─────────────────────────────────────────────────────────────────────
st.markdown("""
<div class="page-header">
    <h1>📊 주식 AI 브리핑 시스템 v2</h1>
    <p>HuggingFace Router · Gemma-4 26B · URL/파일 분석 · RAG · 5가지 페르소나 · Word/PPT 내보내기</p>
</div>""", unsafe_allow_html=True)
st.markdown(f"<p style='color:#546e7a;font-size:.83rem;'>📅 {date.today().strftime('%Y년 %m월 %d일')}</p>", unsafe_allow_html=True)

# ═══════════════════════════════════════════════════════════════
# SECTION 1 — 종목 입력 + 페르소나 선택
# ═══════════════════════════════════════════════════════════════
st.markdown('<div class="section-title">🔍 STEP 1 · 분석 종목 입력 & AI 분석가 선택</div>', unsafe_allow_html=True)

col_input, col_persona = st.columns([2, 3])

with col_input:
    ticker_input = st.text_area(
        "종목 입력 (쉼표·줄바꿈 구분, 최대 10개)",
        placeholder="삼성전자, SK하이닉스, AAPL, NVDA\n또는\n005930, 000660, TSLA",
        height=130,
    )

with col_persona:
    from analyst_personas import PERSONAS, get_persona, persona_labels
    st.markdown("**AI 분석가 선택**")
    persona_choice = st.radio(
        "분석가",
        options=list(PERSONAS.keys()),
        format_func=lambda k: PERSONAS[k]["emoji"] + " " + PERSONAS[k]["label"],
        index=list(PERSONAS.keys()).index(st.session_state.persona_key),
        label_visibility="collapsed",
    )
    st.session_state.persona_key = persona_choice
    p = get_persona(persona_choice)
    st.markdown(f"""
    <div class="persona-card">
        <div class="ptitle">{p['emoji']} {p['label']}</div>
        <div class="pfocus">분석 관점: {p['focus']}</div>
    </div>""", unsafe_allow_html=True)

# ═══════════════════════════════════════════════════════════════
# SECTION 2 — URL 입력
# ═══════════════════════════════════════════════════════════════
st.markdown('<div class="section-title">🔗 STEP 2 · 뉴스/분석 사이트 URL 입력 (선택)</div>', unsafe_allow_html=True)

with st.expander("📌 URL 입력 — 특정 뉴스기사, 분석 사이트를 추가 분석에 반영", expanded=False):
    url_input = st.text_area(
        "URL 입력 (줄바꿈으로 구분, 최대 10개)",
        placeholder="https://finance.naver.com/news/...\nhttps://www.hankyung.com/...\nhttps://www.bloomberg.com/...",
        height=110,
    )
    col_url1, col_url2 = st.columns([1, 3])
    with col_url1:
        btn_fetch_urls = st.button("🌐 URL 콘텐츠 수집", use_container_width=True)
    with col_url2:
        if st.session_state.url_results:
            ok = sum(1 for r in st.session_state.url_results if not r.get("error"))
            fail = len(st.session_state.url_results) - ok
            st.markdown(f"<span style='color:#66BB6A;'>✅ {ok}개 수집 성공</span>"
                        + (f" <span style='color:#EF5350;'>/ {fail}개 실패</span>" if fail else ""),
                        unsafe_allow_html=True)

    if btn_fetch_urls:
        urls = [u.strip() for u in url_input.split("\n") if u.strip()][:10]
        if urls:
            from uploaders.url_reader import fetch_multiple_urls
            from rag.rag_engine import get_rag
            with st.spinner(f"{len(urls)}개 URL 수집 중..."):
                results = fetch_multiple_urls(urls, max_chars_each=4000)
                st.session_state.url_results = results
                rag = get_rag()
                for r in results:
                    if not r.get("error") and r.get("text"):
                        rag.add_document(r["text"], source=r["url"][:60], doc_type="url")
            ok = sum(1 for r in results if not r.get("error"))
            st.success(f"✅ {ok}/{len(urls)}개 URL 수집 완료 → RAG 인덱스 추가됨")
            st.rerun()

    if st.session_state.url_results:
        for r in st.session_state.url_results:
            status_cls = "ustatus-err" if r.get("error") else "ustatus-ok"
            status_txt = f"❌ {r['error']}" if r.get("error") else f"✅ {len(r.get('text',''))}자 수집"
            title = r.get("title") or r["url"][:60]
            st.markdown(f"""
            <div class="url-item">
                <div class="utitle">{title}</div>
                <div class="{status_cls}">{status_txt} · {r['url'][:70]}</div>
            </div>""", unsafe_allow_html=True)

# ═══════════════════════════════════════════════════════════════
# SECTION 3 — 파일 업로드
# ═══════════════════════════════════════════════════════════════
st.markdown('<div class="section-title">📎 STEP 3 · 분석 자료 파일 업로드 (선택)</div>', unsafe_allow_html=True)

with st.expander("📁 파일 업로드 — Excel, PDF, Word, PPT, TXT, HWP, 이미지", expanded=False):
    uploaded_files = st.file_uploader(
        "파일 선택 (복수 선택 가능)",
        type=["xlsx","xls","pdf","docx","doc","pptx","ppt","txt","md","csv","hwp","png","jpg","jpeg"],
        accept_multiple_files=True,
    )
    col_f1, col_f2 = st.columns([1, 3])
    with col_f1:
        btn_process_files = st.button("📂 파일 분석 & RAG 추가", use_container_width=True,
                                      disabled=not uploaded_files)
    with col_f2:
        try:
            from rag.rag_engine import get_rag
            st.markdown(f"<span style='color:#64b5f6;font-size:.83rem;'>📚 RAG 인덱스: {get_rag().doc_count}개 청크</span>",
                        unsafe_allow_html=True)
        except Exception:
            pass

    if btn_process_files and uploaded_files:
        from uploaders.file_reader import extract_text, summarize_file_content
        from rag.rag_engine import get_rag
        rag = get_rag()
        results = []
        prog = st.progress(0)
        for i, f in enumerate(uploaded_files):
            prog.progress((i+1)/len(uploaded_files), text=f"📂 {f.name} 처리 중...")
            data = f.read()
            text = extract_text(data, f.name)
            summary = summarize_file_content(text, f.name, max_chars=4000)
            if not text.startswith("[파일 읽기 오류") and not text.startswith("[지원하지"):
                rag.add_document(summary, source=f.name, doc_type="file")
                results.append({"name": f.name, "chars": len(text), "ok": True})
            else:
                results.append({"name": f.name, "error": text[:100], "ok": False})
        st.session_state.rag_docs = results
        prog.progress(1.0, text="완료!")
        ok = sum(1 for r in results if r["ok"])
        st.success(f"✅ {ok}/{len(results)}개 파일 RAG 인덱스 추가 완료! (총 {rag.doc_count}개 청크)")
        st.rerun()

    if st.session_state.rag_docs:
        for r in st.session_state.rag_docs:
            if r.get("ok"):
                st.markdown(f"<div class='rag-box'>✅ <b>{r['name']}</b> — {r['chars']:,}자 추출 → RAG 추가됨</div>",
                            unsafe_allow_html=True)
            else:
                st.markdown(f"<div class='rag-box' style='border-color:#EF5350;'>❌ <b>{r['name']}</b> — {r.get('error','오류')}</div>",
                            unsafe_allow_html=True)

    # RAG 초기화 버튼
    if st.session_state.rag_docs or st.session_state.url_results:
        if st.button("🗑️ RAG 인덱스 초기화", use_container_width=True):
            from rag.rag_engine import reset_rag
            reset_rag()
            st.session_state.rag_docs = []
            st.session_state.url_results = []
            st.success("RAG 초기화 완료")
            st.rerun()

# ═══════════════════════════════════════════════════════════════
# SECTION 4 — 데이터 수집 & AI 리포트 버튼
# ═══════════════════════════════════════════════════════════════
st.markdown('<div class="section-title">🚀 STEP 4 · 데이터 수집 & AI 분석 실행</div>', unsafe_allow_html=True)

def parse_inputs(raw: str) -> list[str]:
    import re
    tokens = re.split(r"[,\n]+", raw.strip())
    seen, result = set(), []
    for t in [t.strip() for t in tokens]:
        if t and t not in seen:
            seen.add(t); result.append(t)
    return result[:10]

col_b1, col_b2, col_b3 = st.columns([1.5, 1.5, 3])
with col_b1:
    btn_load = st.button("📥 주가·뉴스 수집", type="primary", use_container_width=True)
with col_b2:
    btn_report = st.button("🤖 AI 리포트 생성", use_container_width=True,
                           disabled=not st.session_state.data_loaded,
                           help="데이터 수집 후 활성화됩니다")
with col_b3:
    if st.session_state.data_loaded:
        persona = get_persona(st.session_state.persona_key)
        st.markdown(f"<span style='color:#FFD54F;font-size:.85rem;'>선택된 분석가: {persona['emoji']} {persona['label']}</span>",
                    unsafe_allow_html=True)

# ── 데이터 수집 실행 ──────────────────────────────────────────────────────────
if btn_load:
    inputs = parse_inputs(ticker_input)
    if not inputs:
        st.error("종목을 입력해주세요.")
    else:
        from stock_data import fetch_all_stocks
        from news_fetcher import fetch_all_news
        with st.spinner(f"📡 {len(inputs)}개 종목 주가 수집 중..."):
            stocks = fetch_all_stocks(inputs)
            st.session_state.stocks_data = stocks

        prog = st.progress(0, text="📰 뉴스 수집 중...")
        news_data = {}
        for i, stock in enumerate(stocks):
            prog.progress((i+1)/len(stocks), text=f"📰 {stock['name']} 뉴스 수집 중...")
            news_data[stock["ticker"]] = fetch_all_news(stock)
        st.session_state.news_data = news_data
        st.session_state.data_loaded = True
        st.session_state.briefs_generated = False
        st.session_state.stock_briefs = {}
        st.session_state.portfolio_brief = ""
        st.session_state.full_report = ""

        ok = [s for s in stocks if s.get("close") is not None]
        fail = [s for s in stocks if s.get("close") is None]
        st.success(f"✅ {len(ok)}개 종목 수집 완료!" + (f" ({len(fail)}개 실패)" if fail else ""))
        if fail:
            st.warning("실패: " + ", ".join(f"{s['name']}({s.get('error','?')[:40]})" for s in fail))
        st.rerun()

# ── AI 리포트 생성 ────────────────────────────────────────────────────────────
if btn_report:
    token_check = os.environ.get("HF_TOKEN") or cfg.HF_TOKEN
    if not token_check:
        st.error("❌ HuggingFace Token이 없습니다. 사이드바에서 입력하세요.")
    elif not st.session_state.data_loaded or not st.session_state.stocks_data:
        st.error("먼저 데이터를 수집하세요.")
    else:
        from report_generator import generate_stock_brief, generate_portfolio_brief, build_full_report_text
        from rag.rag_engine import get_rag
        rag = get_rag()
        persona_key = st.session_state.persona_key
        stocks = st.session_state.stocks_data
        news_data = st.session_state.news_data
        stock_briefs = {}
        total = len(stocks) + 1
        prog = st.progress(0, text="🤖 AI 분석 시작...")

        for i, stock in enumerate(stocks):
            ticker = stock["ticker"]
            name = stock.get("name", ticker)
            prog.progress(i/total, text=f"🤖 {name} 분석 중... ({i+1}/{len(stocks)})")
            # RAG 컨텍스트
            rag_ctx = rag.get_context_for_stock(name, ticker, top_k=5) if rag.doc_count > 0 else ""
            with st.spinner(f"{name} AI 브리핑 생성 중..."):
                brief = generate_stock_brief(stock, news_data.get(ticker,{}), persona_key, rag_ctx)
            stock_briefs[ticker] = brief

        prog.progress(len(stocks)/total, text="📋 포트폴리오 종합 분석 중...")
        # 포트폴리오용 RAG: 전체 컨텍스트
        port_rag_ctx = ""
        if rag.doc_count > 0:
            all_names = " ".join(s.get("name","") for s in stocks)
            port_rag_ctx = rag.get_context_for_stock(all_names, "", top_k=6)
        with st.spinner("포트폴리오 종합 리포트 생성 중..."):
            portfolio_brief = generate_portfolio_brief(stocks, [], persona_key, port_rag_ctx)

        prog.progress(1.0, text="✅ 완료!")
        persona = get_persona(persona_key)
        full_report = build_full_report_text(stocks, [], list(stock_briefs.values()), portfolio_brief, persona_key)

        st.session_state.stock_briefs = stock_briefs
        st.session_state.portfolio_brief = portfolio_brief
        st.session_state.full_report = full_report
        st.session_state.briefs_generated = True
        st.success(f"✅ [{persona['emoji']} {persona['short']}] AI 리포트 생성 완료!")
        st.rerun()

# ═══════════════════════════════════════════════════════════════
# 결과 탭
# ═══════════════════════════════════════════════════════════════
if st.session_state.data_loaded and st.session_state.stocks_data:
    stocks = st.session_state.stocks_data
    news_data = st.session_state.news_data

    tab1, tab2, tab3, tab4, tab5, tab6 = st.tabs([
        "📈 주가 현황", "📰 뉴스", "📊 차트", "🤖 AI 브리핑", "📋 최종 리포트", "💾 내보내기"
    ])

    # ── 탭1: 주가 현황 ───────────────────────────────────────────────────────
    with tab1:
        st.markdown('<div class="section-title">📈 오늘 종가 현황</div>', unsafe_allow_html=True)
        for i in range(0, len(stocks), 3):
            cols = st.columns(3)
            for col, stock in zip(cols, stocks[i:i+3]):
                with col:
                    name = stock.get("name", stock["ticker"])
                    close = stock.get("close")
                    change = stock.get("change_rate")
                    market = stock.get("market","")
                    currency = "원" if market=="KR" else "$"
                    flag = "🇰🇷" if market=="KR" else "🇺🇸"
                    close_str = f"{close:,.0f}" if close else "N/A"
                    volume = stock.get("volume")
                    vol_str = f"{volume:,}" if volume else "-"
                    if change is not None:
                        badge = (f'<span class="badge-up">▲ {change:+.2f}%</span>' if change>0
                                 else f'<span class="badge-down">▼ {change:.2f}%</span>' if change<0
                                 else '<span class="badge-flat">─ 0.00%</span>')
                        val_color = "#ef5350" if change>0 else "#42a5f5" if change<0 else "#90a4ae"
                    else:
                        badge='<span class="badge-flat">N/A</span>'; val_color="#90a4ae"
                    st.markdown(f"""
                    <div class="metric-card">
                        <h3>{flag} {stock['ticker']} · {market}</h3>
                        <div style="font-weight:700;font-size:1.05rem;color:#e2e8f0;margin-bottom:4px;">{name}</div>
                        <div class="value" style="color:{val_color};">{close_str} {currency}</div>
                        <div style="margin-top:8px;">{badge}</div>
                        <div style="margin-top:6px;font-size:.76rem;color:#546e7a;">거래량: {vol_str} | {stock.get('date','')}</div>
                    </div>""", unsafe_allow_html=True)
                    if stock.get("error"):
                        st.caption(f"⚠️ {stock['error'][:60]}")

        kr_inv = [s for s in stocks if s.get("market")=="KR" and s.get("investor")]
        if kr_inv:
            st.markdown('<div class="section-title">💹 외인·기관·개인 수급 (당일)</div>', unsafe_allow_html=True)
            rows = []
            for s in kr_inv:
                inv = s.get("investor",{})
                def fmt(v): return f"{v:+,}" if v is not None else "-"
                rows.append({"종목":s.get("name",s["ticker"]),"코드":s["ticker"],
                             "외국인(원)":fmt(inv.get("외국인합계")),
                             "기관(원)":fmt(inv.get("기관합계")),
                             "개인(원)":fmt(inv.get("개인"))})
            st.dataframe(pd.DataFrame(rows), use_container_width=True, hide_index=True)

    # ── 탭2: 뉴스 ───────────────────────────────────────────────────────────
    with tab2:
        if not news_data:
            st.info("데이터 수집 후 뉴스가 표시됩니다.")
        else:
            col_n1, col_n2 = st.columns([2,1])
            with col_n1:
                sel = st.selectbox("종목 선택", [s["ticker"] for s in stocks],
                    format_func=lambda t: next((f"{s['name']} ({t})" for s in stocks if s["ticker"]==t),t))
            with col_n2:
                if st.session_state.url_results:
                    st.markdown(f"<div style='padding-top:28px;font-size:.83rem;color:#64b5f6;'>🔗 URL 자료 {len(st.session_state.url_results)}개 RAG 반영됨</div>",
                                unsafe_allow_html=True)
            if sel in news_data:
                nd = news_data[sel]
                c1, c2 = st.columns(2)
                with c1:
                    st.markdown('<div class="section-title">🇰🇷 국내 뉴스</div>', unsafe_allow_html=True)
                    for n in nd.get("domestic",[]):
                        st.markdown(f"""<div class="news-item">
                            <a href="{n.get('link','#')}" target="_blank">{n.get('title','')}</a>
                            <div class="news-src">{n.get('source','')} · {str(n.get('published',''))[:16]}</div>
                        </div>""", unsafe_allow_html=True)
                    if not nd.get("domestic"):
                        st.info("국내 뉴스 없음")
                with c2:
                    st.markdown('<div class="section-title">🌐 해외 뉴스 (번역)</div>', unsafe_allow_html=True)
                    for n in nd.get("international",[]):
                        st.markdown(f"""<div class="news-item">
                            <a href="{n.get('link','#')}" target="_blank">{n.get('title','')}</a>
                            <div class="news-src">{n.get('source','')} · {str(n.get('published',''))[:16]}</div>
                        </div>""", unsafe_allow_html=True)
                    if not nd.get("international"):
                        st.info("해외 뉴스 없음")

            # URL 수집 결과 표시
            if st.session_state.url_results:
                st.markdown('<div class="section-title">🔗 수집된 URL 자료 (RAG 반영)</div>', unsafe_allow_html=True)
                for r in st.session_state.url_results:
                    if not r.get("error") and r.get("text"):
                        title = r.get("title") or r["url"][:50]
                        preview = r["text"][:200].replace("\n"," ")
                        st.markdown(f"""<div class="url-item">
                            <div class="utitle">🔗 {title}</div>
                            <div style="font-size:.8rem;color:#78909C;">{r['url'][:70]}</div>
                            <div style="font-size:.82rem;color:#90a4ae;margin-top:4px;">{preview}...</div>
                        </div>""", unsafe_allow_html=True)

    # ── 탭3: 차트 ───────────────────────────────────────────────────────────
    with tab3:
        sel2 = st.selectbox("차트 종목", [s["ticker"] for s in stocks],
            format_func=lambda t: next((f"{s['name']} ({t})" for s in stocks if s["ticker"]==t),t),
            key="chart_sel")
        period_sel = st.radio("기간", ["week","month","6month"],
            format_func={"week":"이번주","month":"이번달","6month":"6개월"}.get,
            horizontal=True)
        s_info = next((s for s in stocks if s["ticker"]==sel2), None)
        if s_info:
            hist = s_info.get("history",{}).get(period_sel, pd.DataFrame())
            name = s_info.get("name", sel2)
            if hist is not None and not hist.empty and "close" in hist.columns:
                fig = go.Figure()
                fig.add_trace(go.Scatter(x=hist.index, y=hist["close"], mode="lines", name=name,
                    line=dict(color="#42a5f5",width=2), fill="tozeroy", fillcolor="rgba(66,165,245,0.07)"))
                xn = np.arange(len(hist))
                if len(xn) > 1:
                    z = np.polyfit(xn, hist["close"].values, 1)
                    p_fn = np.poly1d(z)
                    tc = "#ef5350" if z[0]>0 else "#1976d2"
                    fig.add_trace(go.Scatter(x=hist.index, y=p_fn(xn), mode="lines", name="추세선",
                        line=dict(color=tc, width=1.5, dash="dash")))
                label = {"week":"이번주","month":"이번달","6month":"6개월"}[period_sel]
                currency = "원" if s_info.get("market")=="KR" else "$"
                fig.update_layout(
                    title=dict(text=f"{name} 주가 추이 ({label})", font=dict(color="#e2e8f0",size=16)),
                    plot_bgcolor="#0d1117", paper_bgcolor="#0d1117",
                    font=dict(color="#90a4ae"),
                    xaxis=dict(gridcolor="#1e3a5f", linecolor="#1e3a5f"),
                    yaxis=dict(gridcolor="#1e3a5f", linecolor="#1e3a5f", tickformat=","),
                    legend=dict(bgcolor="#111827", bordercolor="#1e3a5f"),
                    height=420, margin=dict(l=10,r=10,t=40,b=10),
                )
                st.plotly_chart(fig, use_container_width=True)
                sp = float(hist["close"].iloc[0]); ep = float(hist["close"].iloc[-1])
                ch = (ep-sp)/sp*100 if sp else 0
                c1,c2,c3 = st.columns(3)
                c1.metric("시작가", f"{sp:,.0f}{currency}")
                c2.metric("현재가", f"{ep:,.0f}{currency}")
                c3.metric("기간 수익률", f"{ch:+.2f}%")
            else:
                st.info("차트 데이터가 없습니다.")

    # ── 탭4: AI 브리핑 ───────────────────────────────────────────────────────
    with tab4:
        if not st.session_state.briefs_generated:
            persona = get_persona(st.session_state.persona_key)
            st.info(f"⬆️ STEP 4의 **🤖 AI 리포트 생성** 버튼을 눌러주세요.")
            st.markdown(f"""
            <div class="persona-card" style="max-width:500px;">
                <div class="ptitle">{persona['emoji']} 선택된 분석가: {persona['label']}</div>
                <div class="pfocus">{persona['focus']}</div>
            </div>
            <div style="margin-top:12px;background:#111827;border-radius:10px;padding:14px;font-size:.86rem;color:#90a4ae;">
            <b style="color:#64b5f6;">AI 브리핑 포함 내용</b><br>
            ① 종목 현황 요약 &nbsp;② 주가 추이 분석 &nbsp;③ 수급 분석<br>
            ④ 뉴스 이벤트 임팩트 &nbsp;⑤ RAG 자료 반영 분석<br>
            ⑥ 리스크 요인 &nbsp;⑦ 내일 매매 전략 (진입가·손절가·목표가)<br>
            ⑧ 내일 예상 방향 & 등락률
            </div>""", unsafe_allow_html=True)
        else:
            persona = get_persona(st.session_state.persona_key)
            st.markdown(f"""
            <div style="background:#0d1b4b;border:1px solid #1565c0;border-radius:8px;padding:8px 16px;margin-bottom:12px;display:inline-block;">
                <span style="color:#FFD54F;font-weight:700;">{persona['emoji']} {persona['label']}</span>
                <span style="color:#90caf9;font-size:.83rem;margin-left:12px;">관점으로 분석</span>
            </div>""", unsafe_allow_html=True)

            sel3 = st.selectbox("브리핑 종목", [s["ticker"] for s in stocks],
                format_func=lambda t: next((f"{s['name']} ({t})" for s in stocks if s["ticker"]==t),t),
                key="brief_sel")
            brief = st.session_state.stock_briefs.get(sel3,"")
            if brief:
                si = next((s for s in stocks if s["ticker"]==sel3),{})
                name = si.get("name", sel3)
                close = si.get("close"); change = si.get("change_rate")
                currency = "원" if si.get("market")=="KR" else "$"
                chg_color = "#ef5350" if (change or 0)>0 else "#42a5f5"
                c1,c2 = st.columns([4,1])
                with c1:
                    st.markdown(f"### 📌 {name} ({sel3}) 투자 분석")
                with c2:
                    if close:
                        st.markdown(f"""<div style="text-align:right;">
                            <div style="font-size:1.2rem;font-weight:700;color:{chg_color};">{close:,.0f}{currency}</div>
                            <div style="font-size:.9rem;color:{chg_color};">{f'{change:+.2f}%' if change else 'N/A'}</div>
                        </div>""", unsafe_allow_html=True)

                if brief.startswith("[LLM") or brief.startswith("[오류"):
                    st.error(brief)
                    st.info("💡 HF_TOKEN 확인 또는 모델 변경 후 재시도하세요.")
                else:
                    st.markdown(f'<div class="report-box">{brief}</div>', unsafe_allow_html=True)

            if st.session_state.portfolio_brief:
                st.markdown("---")
                st.markdown(f"### 📋 포트폴리오 종합 전략 [{persona['emoji']} {persona['short']}]")
                pb = st.session_state.portfolio_brief
                if pb.startswith("[LLM") or pb.startswith("[오류"):
                    st.error(pb)
                else:
                    st.markdown(f'<div class="report-box">{pb}</div>', unsafe_allow_html=True)

    # ── 탭5: 최종 리포트 + 이메일 ────────────────────────────────────────────
    with tab5:
        if not st.session_state.full_report:
            st.info("AI 리포트를 먼저 생성해주세요.")
        else:
            persona = get_persona(st.session_state.persona_key)
            st.markdown(f'<div class="section-title">📋 최종 브리핑 리포트 [{persona["emoji"]} {persona["label"]}]</div>',
                        unsafe_allow_html=True)
            st.markdown(f'<div class="report-box">{st.session_state.full_report}</div>',
                        unsafe_allow_html=True)
            st.download_button("📥 TXT 다운로드",
                data=st.session_state.full_report.encode("utf-8"),
                file_name=f"stock_briefing_{date.today()}.txt",
                mime="text/plain", use_container_width=True)

            st.markdown("---")
            st.markdown('<div class="section-title">📧 이메일 발송</div>', unsafe_allow_html=True)
            c1,c2 = st.columns([3,1])
            with c1:
                recipient = st.text_input("수신 이메일", value=to_email, placeholder="recipient@email.com", key="send_to")
            with c2:
                st.markdown("<br>",unsafe_allow_html=True)
                btn_send = st.button("📤 발송", type="primary", use_container_width=True)
            if btn_send:
                from email_sender import send_report_email
                with st.spinner("발송 중..."):
                    ok, msg = send_report_email(recipient, st.session_state.full_report, gmail_user, gmail_pw)
                st.success(msg) if ok else st.error(msg)

    # ── 탭6: 내보내기 (Word / PPT) ────────────────────────────────────────────
    with tab6:
        st.markdown('<div class="section-title">💾 리포트 내보내기</div>', unsafe_allow_html=True)

        if not st.session_state.briefs_generated:
            st.info("AI 리포트를 먼저 생성해야 내보내기가 가능합니다.")
        else:
            persona = get_persona(st.session_state.persona_key)
            st.markdown(f"""
            <div style="background:#0d1b4b;border:1px solid #FFD54F;border-radius:10px;padding:12px 18px;margin-bottom:16px;">
                <span style="color:#FFD54F;font-weight:700;">{persona['emoji']} {persona['label']}</span>
                <span style="color:#90caf9;font-size:.85rem;"> 분석 리포트를 Word 또는 PPT로 내보내기</span>
            </div>""", unsafe_allow_html=True)

            col_w, col_p = st.columns(2)

            # ── Word 내보내기
            with col_w:
                st.markdown("""
                <div class="export-btn-area">
                    <div style="color:#64b5f6;font-weight:700;font-size:1.05rem;margin-bottom:8px;">📄 Word (.docx)</div>
                    <div style="color:#90a4ae;font-size:.83rem;line-height:1.6;">
                    • 전체 리포트 완전 포함<br>
                    • 종목 현황 표 포함<br>
                    • 헤더/푸터 자동 생성<br>
                    • 섹션별 계층 구조<br>
                    • 프로 서식 (Malgun Gothic)
                    </div>
                </div>""", unsafe_allow_html=True)
                btn_word = st.button("📄 Word 파일 생성", type="primary", use_container_width=True, key="btn_word")

                if btn_word:
                    with st.spinner("📄 Word 파일 생성 중..."):
                        try:
                            from exporters.word_exporter import export_word_bytes
                            word_bytes = export_word_bytes(
                                report_text=st.session_state.full_report,
                                stocks=st.session_state.stocks_data,
                                persona_label=f"{persona['emoji']} {persona['label']}",
                            )
                            st.download_button(
                                "⬇️ Word 다운로드",
                                data=word_bytes,
                                file_name=f"stock_briefing_{date.today()}.docx",
                                mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                                use_container_width=True,
                                key="dl_word",
                            )
                            st.success("✅ Word 파일 생성 완료! (python-docx 기반)")
                        except Exception as e:
                            st.error(f"Word 생성 오류: {e}")
                            st.info("pip install python-docx 를 실행하세요.")

            # ── PPT 내보내기
            with col_p:
                st.markdown("""
                <div class="export-btn-area">
                    <div style="color:#64b5f6;font-weight:700;font-size:1.05rem;margin-bottom:8px;">📊 PowerPoint (.pptx)</div>
                    <div style="color:#90a4ae;font-size:.83rem;line-height:1.6;">
                    • 표지 + 현황표 슬라이드<br>
                    • 종합 전략 요약 슬라이드<br>
                    • 종목별 개별 분석 슬라이드<br>
                    • 주가 추이 차트 (인라인)<br>
                    • 다크 프리미엄 디자인
                    </div>
                </div>""", unsafe_allow_html=True)
                btn_ppt = st.button("📊 PPT 파일 생성", type="primary", use_container_width=True, key="btn_ppt")

                if btn_ppt:
                    with st.spinner("📊 PPT 파일 생성 중..."):
                        try:
                            from exporters.ppt_exporter import export_ppt_bytes
                            ppt_bytes = export_ppt_bytes(
                                stocks=st.session_state.stocks_data,
                                stock_briefs=list(st.session_state.stock_briefs.values()),
                                portfolio_brief=st.session_state.portfolio_brief,
                                persona_label=persona['label'],
                                persona_emoji=persona['emoji'],
                            )
                            st.download_button(
                                "⬇️ PPT 다운로드",
                                data=ppt_bytes,
                                file_name=f"stock_briefing_{date.today()}.pptx",
                                mime="application/vnd.openxmlformats-officedocument.presentationml.presentation",
                                use_container_width=True,
                                key="dl_ppt",
                            )
                            st.success("✅ PPT 파일 생성 완료! (python-pptx 기반)")
                        except Exception as e:
                            st.error(f"PPT 생성 오류: {e}")
                            st.info("pip install python-pptx 를 실행하세요.")

            # TXT 백업 다운로드 (항상 가능)
            st.markdown("---")
            st.markdown("**📥 TXT 텍스트 다운로드 (항상 가능)**")
            st.download_button(
                "📥 전체 리포트 TXT",
                data=st.session_state.full_report.encode("utf-8"),
                file_name=f"stock_briefing_{date.today()}.txt",
                mime="text/plain",
                use_container_width=True,
                key="dl_txt2",
            )
