import streamlit as st
import os, json, re, math, io
import plotly.graph_objects as go
from openai import OpenAI
from tavily import TavilyClient

st.set_page_config(
    page_title="ISO 4.1 外部環境分析平台",
    page_icon=None,
    layout="centered",
    initial_sidebar_state="collapsed",
)

# ── Session defaults ──────────────────────────────────────────────────────────
for k, v in {"dark": False, "results": None, "lang": "zh"}.items():
    if k not in st.session_state:
        st.session_state[k] = v

# ── i18n ──────────────────────────────────────────────────────────────────────
_LANG = {
    "zh": {
        "badge": "ISO 9001 · CLAUSE 4.1 · CONTEXT OF THE ORGANIZATION",
        "title": "外部環境分析平台",
        "subtitle": "輸入企業基本資訊，AI 自動爬取最新情報，生成 PESTLE 分析、機會威脅評估及稽核員建議報告。",
        "settings": "設定",
        "dark_mode": "深色模式",
        "language": "語言",
        "company_label": "公司名稱",
        "company_ph": "例：鼎智國際技術服務集團",
        "country_label": "所在國家／地區",
        "industry_label": "行業別（中華民國行業標準分類）",
        "website_label": "官方網站",
        "warn_website": "請填寫官方網站後，再開始分析。",
        "website_ph": "例：https://www.tksg.global/",
        "model_label": "AI 分析模型",
        "run_btn": "執行 ISO 4.1 外部環境分析",
        "warn_company": "請填寫公司名稱後，再開始分析。",
        "tab_info": "企業資訊 & 總部", "tab_dist": "情報分佈", "tab_graph": "情報知識圖",
        "tab_pestle": "PESTLE 分析", "tab_ot": "機會與威脅", "tab_audit": "稽核建議",
        "metric_total": "情報總量", "metric_types": "議題類型",
        "metric_company": "分析企業", "metric_model": "使用模型",
        "globe_cap": "可拖曳旋轉地球儀",
        "dist_cap": "依 {n} 篇新聞自動分類 — 雷達圖顯示各議題覆蓋廣度，長條圖顯示篇數分佈",
        "graph_cap": "中心為分析企業，第一環為議題類型，外環為各篇新聞。點擊節點可放大，懸停查看標題",
        "report_hdr": "ISO 4.1 外部環境分析報告",
        "pdf_btn": "下載 PDF 報告",
        "summary_lbl": "智慧摘要",
        "no_data": "此區塊暫無資料。",
        "sources_title": "情報來源",
        "sources_cap": "共 {n} 筆，含類型標籤與相關度評分。",
        "open": "開啟", "relevance": "相關度",
        "risk_title": "外部風險指數",
        "risk_high": "高風險", "risk_mid": "中風險", "risk_low": "低風險",
    },
    "en": {
        "badge": "ISO 9001 · CLAUSE 4.1 · CONTEXT OF THE ORGANIZATION",
        "title": "External Environment Analysis",
        "subtitle": "Enter company details. AI crawls the latest intelligence and generates PESTLE, O&T, and auditor recommendations.",
        "settings": "Settings",
        "dark_mode": "Dark Mode",
        "language": "Language",
        "company_label": "Company Name",
        "company_ph": "e.g. TKSG International",
        "country_label": "Country / Region",
        "industry_label": "Industry",
        "website_label": "Official Website",
        "warn_website": "Please enter the official website before running the analysis.",
        "website_ph": "e.g. https://www.tksg.global/",
        "model_label": "AI Model",
        "run_btn": "Run ISO 4.1 External Analysis",
        "warn_company": "Please enter a company name before running the analysis.",
        "tab_info": "Company & HQ", "tab_dist": "Intel Distribution", "tab_graph": "Knowledge Graph",
        "tab_pestle": "PESTLE Analysis", "tab_ot": "Opportunities & Threats", "tab_audit": "Auditor Recommendations",
        "metric_total": "Total Intel", "metric_types": "Topic Types",
        "metric_company": "Company", "metric_model": "Model",
        "globe_cap": "Drag to rotate globe",
        "dist_cap": "Auto-classified from {n} articles — radar shows breadth, bar chart shows count",
        "graph_cap": "Center: company · Ring 1: categories · Ring 2: articles. Click to zoom, hover for details",
        "report_hdr": "ISO 4.1 External Environment Report",
        "pdf_btn": "Download PDF Report",
        "summary_lbl": "EXECUTIVE SUMMARY",
        "no_data": "No data available for this section.",
        "sources_title": "Intelligence Sources",
        "sources_cap": "{n} sources with topic tags and relevance scores.",
        "open": "Open", "relevance": "Relevance",
        "risk_title": "External Risk Index",
        "risk_high": "High Risk", "risk_mid": "Medium Risk", "risk_low": "Low Risk",
    },
    "de": {
        "badge": "ISO 9001 · KLAUSEL 4.1 · KONTEXT DER ORGANISATION",
        "title": "Externe Umweltanalyse",
        "subtitle": "Unternehmensdaten eingeben. KI crawlt Informationen und generiert PESTLE-Analyse, Chancen/Risiken und Prüferempfehlungen.",
        "settings": "Einstellungen",
        "dark_mode": "Dunkelmodus",
        "language": "Sprache",
        "company_label": "Unternehmensname",
        "company_ph": "z.B. Volkswagen AG",
        "country_label": "Land / Region",
        "industry_label": "Branche",
        "website_label": "Website",
        "warn_website": "Bitte Website eingeben.",
        "website_ph": "z.B. https://www.tksg.global/",
        "model_label": "KI-Modell",
        "run_btn": "ISO 4.1 Analyse starten",
        "warn_company": "Bitte Unternehmensnamen eingeben.",
        "tab_info": "Unternehmen & HQ", "tab_dist": "Informationsverteilung", "tab_graph": "Wissensgraph",
        "tab_pestle": "PESTLE-Analyse", "tab_ot": "Chancen & Risiken", "tab_audit": "Prüferempfehlungen",
        "metric_total": "Gesamtquellen", "metric_types": "Thementypen",
        "metric_company": "Unternehmen", "metric_model": "Modell",
        "globe_cap": "Globus ziehen zum Drehen",
        "dist_cap": "{n} Artikel klassifiziert — Radar zeigt Themenbreite, Balken zeigt Anzahl",
        "graph_cap": "Zentrum: Unternehmen · Ring 1: Kategorien · Ring 2: Artikel. Klicken zum Zoomen",
        "report_hdr": "ISO 4.1 Externe Umweltanalyse",
        "pdf_btn": "PDF-Bericht herunterladen",
        "summary_lbl": "EXECUTIVE SUMMARY",
        "no_data": "Keine Daten verfügbar.",
        "sources_title": "Informationsquellen",
        "sources_cap": "{n} Quellen mit Themen-Tags und Relevanzbewertungen.",
        "open": "Öffnen", "relevance": "Relevanz",
        "risk_title": "Externer Risikoindex",
        "risk_high": "Hohes Risiko", "risk_mid": "Mittleres Risiko", "risk_low": "Niedriges Risiko",
    },
}

def T(key: str, **kw) -> str:
    row = _LANG.get(st.session_state.lang, _LANG["zh"])
    txt = row.get(key, _LANG["zh"].get(key, key))
    return txt.format(**kw) if kw else txt

# ── Theme ─────────────────────────────────────────────────────────────────────
def palette(dark):
    if dark:
        return dict(
            bg="#0d0a1a", surface="#160f2e", surface2="#0f0a22",
            border="#2d1f50", border2="#1a1035",
            text="#ede9fe", text2="#a78bfa", text3="#6d5fa8", text4="#3b2a6e",
            accent="#a78bfa", accent_dim="rgba(167,139,250,0.18)",
            accent_border="rgba(167,139,250,0.30)",
            chart_grid="rgba(45,31,80,0.6)",
            plot_bg="rgba(0,0,0,0)", paper_bg="rgba(0,0,0,0)",
            input_bg="#160f2e", placeholder="#3b2a6e",
            globe_land="#1e1048", globe_ocean="#0d0a1a",
            globe_coast="#6d28d9", globe_country="#2d1f50",
        )
    else:
        return dict(
            bg="#FAFBFC", surface="#ffffff", surface2="#f5f3ff",
            border="#e2e8f0", border2="#f1f5f9",
            text="#0f172a", text2="#475569", text3="#94a3b8", text4="#cbd5e1",
            accent="#7C3AED", accent_dim="rgba(124,58,237,0.10)",
            accent_border="rgba(124,58,237,0.25)",
            chart_grid="rgba(226,232,240,0.8)",
            plot_bg="rgba(0,0,0,0)", paper_bg="rgba(0,0,0,0)",
            input_bg="#ffffff", placeholder="#cbd5e1",
            globe_land="#ede9fe", globe_ocean="#ddd6fe",
            globe_coast="#7C3AED", globe_country="#c4b5fd",
        )

P = palette(st.session_state.dark)

# ── CSS ───────────────────────────────────────────────────────────────────────
st.markdown(f"""
<style>
@import url('https://fonts.googleapis.com/css2?family=Inter:ital,wght@0,300;0,400;0,500;0,600;0,700;0,800&display=swap');

#MainMenu, footer, header {{ visibility: hidden; }}
.stDeployButton {{ display: none !important; }}

html, body, .stApp {{
    font-family: 'Inter', -apple-system, sans-serif !important;
    background-color: {P['bg']} !important;
    color: {P['text']} !important;
}}
.main .block-container {{ max-width: 960px; padding: 0 2.4rem 6rem; }}

label[data-testid="stWidgetLabel"] p,
.stTextInput label, .stSelectbox label, .stToggle label {{
    color: {P['text3']} !important;
    font-size: 0.71rem !important;
    font-weight: 600 !important;
    letter-spacing: 0.08em !important;
    text-transform: uppercase !important;
}}
.stTextInput input {{
    background-color: {P['input_bg']} !important;
    border: 1px solid {P['border']} !important;
    border-radius: 8px !important;
    color: {P['text']} !important;
    font-family: 'Inter', sans-serif !important;
    font-size: 0.88rem !important;
}}
.stTextInput input:focus {{
    border-color: {P['accent']} !important;
    box-shadow: 0 0 0 3px {P['accent_dim']} !important;
}}
.stTextInput input::placeholder {{ color: {P['placeholder']} !important; }}
[data-baseweb="select"] > div {{
    background-color: {P['input_bg']} !important;
    border-color: {P['border']} !important;
    border-radius: 8px !important;
    color: {P['text']} !important;
}}
[data-baseweb="select"] svg {{ fill: {P['text3']} !important; }}
.stButton > button[kind="primary"] {{
    background: #7C3AED !important;
    color: #ffffff !important;
    border: none !important;
    border-radius: 8px !important;
    font-weight: 600 !important;
    font-family: 'Inter', sans-serif !important;
    font-size: 0.9rem !important;
    letter-spacing: 0.01em !important;
    box-shadow: 0 4px 14px rgba(124,58,237,0.28) !important;
    transition: all 0.18s ease !important;
}}
.stButton > button[kind="primary"]:hover {{
    background: #6d28d9 !important;
    box-shadow: 0 6px 22px rgba(124,58,237,0.42) !important;
    transform: translateY(-1px) !important;
}}
[data-testid="stToggle"] label {{
    color: {P['text2']} !important;
    font-size: 0.78rem !important;
    font-weight: 500 !important;
    text-transform: none !important;
    letter-spacing: 0 !important;
}}
[data-testid="stAlert"] {{ border-radius: 8px !important; border: none !important; font-size: 0.88rem !important; }}
[data-testid="stMarkdown"] p {{ color: {P['text2']}; line-height: 1.82; font-size: 0.9rem; }}
[data-testid="stMarkdown"] h3 {{
    color: {P['text']} !important; font-size: 0.95rem !important; font-weight: 600 !important;
    border-bottom: 1px solid {P['border']}; padding-bottom: 0.4rem; margin-top: 1.6rem !important;
}}
[data-testid="stMarkdown"] li {{ color: {P['text2']}; line-height: 1.78; font-size: 0.9rem; }}
[data-testid="stMarkdown"] strong {{ color: {P['text']} !important; }}
h2, [data-testid="stHeadingWithActionElements"] h2 {{
    color: {P['text']} !important; font-size: 1rem !important;
    font-weight: 600 !important; font-family: 'Inter', sans-serif !important;
}}
[data-testid="stMetric"] {{
    background: {P['surface']} !important; border: 1px solid {P['border']} !important;
    border-radius: 10px !important; padding: 0.75rem 1rem !important;
}}
[data-testid="stMetricLabel"] p {{
    color: {P['text3']} !important; font-size: 0.7rem !important;
    text-transform: uppercase !important; letter-spacing: 0.07em !important; font-weight: 600 !important;
}}
[data-testid="stMetricValue"] {{
    color: {P['text']} !important; font-size: 1.35rem !important; font-weight: 700 !important;
}}
hr {{ border-color: {P['border2']} !important; margin: 1.8rem 0 !important; }}
[data-testid="stCaptionContainer"] p {{ color: {P['text3']} !important; font-size: 0.78rem !important; }}
::-webkit-scrollbar {{ width: 4px; height: 4px; }}
::-webkit-scrollbar-track {{ background: transparent; }}
::-webkit-scrollbar-thumb {{ background: {P['border']}; border-radius: 4px; }}

/* Tabs */
[data-testid="stTabs"] [role="tablist"] {{
    gap: 0 !important;
    border-bottom: 1px solid {P['border']} !important;
    margin-bottom: 1rem !important;
}}
[data-testid="stTabs"] [role="tab"] {{
    font-family: 'Inter', sans-serif !important;
    font-size: 0.8rem !important;
    font-weight: 500 !important;
    color: {P['text3']} !important;
    padding: 0.45rem 1.1rem !important;
    border: none !important;
    border-bottom: 2px solid transparent !important;
    background: transparent !important;
    border-radius: 0 !important;
    letter-spacing: 0.01em !important;
}}
[data-testid="stTabs"] [role="tab"][aria-selected="true"] {{
    color: {P['accent']} !important;
    border-bottom: 2px solid {P['accent']} !important;
    font-weight: 600 !important;
}}
[data-testid="stTabs"] [role="tab"]:hover {{
    color: {P['text2']} !important;
    background: transparent !important;
}}

/* Download button */
.stDownloadButton > button {{
    background: transparent !important;
    border: 1.5px solid {P['accent']} !important;
    color: {P['accent']} !important;
    border-radius: 8px !important;
    font-size: 0.82rem !important;
    font-weight: 600 !important;
    font-family: 'Inter', sans-serif !important;
    padding: 0.42rem 0.9rem !important;
    transition: all 0.16s ease !important;
    letter-spacing: 0.01em !important;
}}
.stDownloadButton > button:hover {{
    background: {P['accent_dim']} !important;
}}

/* Premium report markdown */
[data-testid="stMarkdown"] h3 {{
    border-left: 3px solid {P['accent']} !important;
    padding-left: 0.75rem !important;
    border-bottom: none !important;
    margin-top: 1.8rem !important;
    font-size: 1rem !important;
    color: {P['text']} !important;
    font-weight: 600 !important;
}}
[data-testid="stMarkdown"] strong {{
    color: {P['text']} !important;
    font-weight: 650 !important;
}}
[data-testid="stMarkdown"] ul li, [data-testid="stMarkdown"] ol li {{
    padding-bottom: 0.25rem !important;
}}
</style>
""", unsafe_allow_html=True)

# ── Data ──────────────────────────────────────────────────────────────────────
COUNTRIES = [
    "中華民國（台灣）","中國大陸","美國","日本","韓國","新加坡","香港",
    "德國","英國","法國","荷蘭","瑞士","瑞典","義大利","西班牙",
    "印度","越南","泰國","馬來西亞","印尼","菲律賓",
    "澳洲","紐西蘭","加拿大","巴西","墨西哥","阿聯酋","沙烏地阿拉伯","其他",
]
COUNTRY_COORDS = {
    "中華民國（台灣）": (23.97, 120.97),  "中國大陸": (35.86, 104.19),
    "美國": (37.09, -95.71),              "日本": (36.20, 138.25),
    "韓國": (35.91, 127.77),              "新加坡": (1.35, 103.82),
    "香港": (22.32, 114.17),              "德國": (51.17, 10.45),
    "英國": (55.38, -3.44),               "法國": (46.23, 2.21),
    "荷蘭": (52.13, 5.29),                "瑞士": (46.82, 8.23),
    "瑞典": (60.13, 18.64),               "義大利": (41.87, 12.57),
    "西班牙": (40.46, -3.75),             "印度": (20.59, 78.96),
    "越南": (14.06, 108.28),              "泰國": (15.87, 100.99),
    "馬來西亞": (4.21, 101.98),           "印尼": (-0.79, 113.92),
    "菲律賓": (12.88, 121.77),            "澳洲": (-25.27, 133.78),
    "紐西蘭": (-40.90, 174.89),           "加拿大": (56.13, -106.35),
    "巴西": (-14.24, -51.93),             "墨西哥": (23.63, -102.55),
    "阿聯酋": (23.42, 53.85),             "沙烏地阿拉伯": (23.89, 45.08),
    "其他": (20.0, 0.0),
}
INDUSTRIES = [
    "A - 農、林、漁、牧業","B - 礦業及土石採取業","C - 製造業",
    "D - 電力及燃氣供應業","E - 用水供應及污染整治業","F - 營建工程業",
    "G - 批發及零售業","H - 運輸及倉儲業","I - 住宿及餐飲業",
    "J - 出版、影音製作、傳播及資通訊服務業","K - 金融及保險業","L - 不動產業",
    "M - 專業、科學及技術服務業","N - 支援服務業",
    "O - 公共行政及國防；強制性社會安全","P - 教育業",
    "Q - 醫療保健及社會工作服務業","R - 藝術、娛樂及休閒服務業","S - 其他服務業",
]
MODELS = {
    "GPT-5.4 Mini（預設）": "gpt-5.4-mini",
    "GPT-4.1 Mini": "gpt-4.1-mini",
    "GPT-4.1": "gpt-4.1",
    "GPT-4o": "gpt-4o",
    "GPT-4o Mini": "gpt-4o-mini",
    "o4-Mini": "o4-mini",
}
NEWS_CATEGORIES = {
    "政策法規": ["法規","政策","法律","監管","裁罰","條例","標準","規定","合規","regulation","policy","law","sanction"],
    "市場趨勢": ["市場","需求","成長","趨勢","規模","份額","競爭","market","growth","trend","demand"],
    "技術創新": ["技術","創新","研發","AI","數位","轉型","自動化","智慧","technology","innovation","digital","automation"],
    "財務數據": ["營收","獲利","財報","股價","投資","資金","虧損","revenue","profit","financial","investment"],
    "風險事件": ["風險","危機","衝突","地緣","制裁","供應鏈","shortage","risk","crisis","war","conflict","supply"],
    "國際動態": ["國際","全球","美中","貿易","出口","進口","關稅","global","trade","tariff","international"],
}
CATEGORY_COLORS = {
    "政策法規":"#8B5CF6","市場趨勢":"#3B82F6","技術創新":"#10B981",
    "財務數據":"#F59E0B","風險事件":"#EF4444","國際動態":"#EC4899","其他":"#6B7280",
}

# ── Helpers ───────────────────────────────────────────────────────────────────
def categorize_article(title, content):
    text = (title + " " + content).lower()
    scores = {cat: sum(1 for kw in kws if kw.lower() in text) for cat, kws in NEWS_CATEGORIES.items()}
    best = max(scores, key=scores.get)
    return best if scores[best] > 0 else "其他"

def render_steps(steps, p):
    rows = ""
    for i, step in enumerate(steps):
        s = step["status"]
        last = i == len(steps) - 1
        if s == "done":
            dot = f'<span style="width:9px;height:9px;border-radius:50%;background:#10b981;flex-shrink:0;display:inline-block;box-shadow:0 0 5px #10b98166;"></span>'
            label_color = p["text2"]
            weight = "400"
        elif s == "running":
            dot = f'<span style="width:9px;height:9px;border-radius:50%;background:{p["accent"]};flex-shrink:0;display:inline-block;box-shadow:0 0 8px {p["accent"]};"></span>'
            label_color = p["text"]
            weight = "600"
        else:
            dot = f'<span style="width:9px;height:9px;border-radius:50%;border:1.5px solid {p["border"]};flex-shrink:0;display:inline-block;background:transparent;"></span>'
            label_color = p["text3"]
            weight = "400"

        detail = step.get("detail", "")
        detail_html = f'<span style="font-size:0.71rem;color:{p["text3"]};margin-left:auto;white-space:nowrap;padding-left:8px;">{detail}</span>' if detail else ""
        border_bottom = "" if last else f'border-bottom:1px solid {p["border2"]};'

        rows += (
            f'<div style="display:flex;align-items:center;gap:10px;padding:0.48rem 0;{border_bottom}">'
            f'{dot}'
            f'<span style="font-size:0.86rem;color:{label_color};font-weight:{weight};'
            f'font-family:Inter,sans-serif;">{step["label"]}</span>'
            f'{detail_html}'
            f'</div>'
        )
    header_color = p["text3"]
    return (
        f'<div style="background:{p["surface"]};border:1px solid {p["border"]};'
        f'border-radius:12px;padding:1rem 1.25rem;">'
        f'<div style="font-size:0.67rem;font-weight:700;color:{header_color};'
        f'letter-spacing:0.12em;text-transform:uppercase;margin-bottom:0.9rem;">'
        f'分析流程</div>'
        f'{rows}</div>'
    )

def build_globe(country, company_name, dark, p):
    lat, lon = COUNTRY_COORDS.get(country, (20.0, 0.0))
    fig = go.Figure()
    # Glow ring (larger, transparent)
    fig.add_trace(go.Scattergeo(
        lat=[lat], lon=[lon], mode="markers",
        marker=dict(size=22, color="rgba(124,58,237,0.15)",
                    line=dict(color="rgba(167,139,250,0.4)", width=1)),
        hoverinfo="skip", showlegend=False,
    ))
    # Main marker
    fig.add_trace(go.Scattergeo(
        lat=[lat], lon=[lon],
        mode="markers+text",
        marker=dict(size=11, color="#7C3AED",
                    line=dict(color="#c4b5fd", width=2),
                    symbol="circle"),
        text=[f"  {company_name}"],
        textposition="middle right",
        textfont=dict(color="#a78bfa" if dark else "#6d28d9", size=11, family="Inter"),
        hovertemplate=f"<b>{company_name}</b><br>{country}<extra></extra>",
        showlegend=False,
    ))
    fig.update_geos(
        projection_type="orthographic",
        projection_rotation=dict(lon=lon, lat=lat, roll=0),
        showland=True, landcolor=p["globe_land"],
        showocean=True, oceancolor=p["globe_ocean"],
        showcoastlines=True, coastlinecolor=p["globe_coast"],
        coastlinewidth=0.6,
        showcountries=True, countrycolor=p["globe_country"],
        countrywidth=0.4,
        showlakes=False,
        bgcolor="rgba(0,0,0,0)",
    )
    fig.update_layout(
        height=380, margin=dict(l=0, r=0, t=0, b=0),
        paper_bgcolor="rgba(0,0,0,0)",
        showlegend=False,
        geo=dict(bgcolor="rgba(0,0,0,0)"),
    )
    return fig

def build_pin_globe(country, company_name, dark, p):
    """Reactive globe for the landing form — pins the selected country, labels with company name."""
    lat, lon = COUNTRY_COORDS.get(country, (20.0, 0.0))
    label = company_name.strip() if company_name and company_name.strip() else ""
    fig = go.Figure()
    # Outer glow ring
    fig.add_trace(go.Scattergeo(
        lat=[lat], lon=[lon], mode="markers",
        marker=dict(size=36, color="rgba(124,58,237,0.08)",
                    line=dict(color="rgba(167,139,250,0.25)", width=1)),
        hoverinfo="skip", showlegend=False,
    ))
    # Inner glow
    fig.add_trace(go.Scattergeo(
        lat=[lat], lon=[lon], mode="markers",
        marker=dict(size=20, color="rgba(124,58,237,0.18)",
                    line=dict(color="rgba(167,139,250,0.0)", width=0)),
        hoverinfo="skip", showlegend=False,
    ))
    # Pin dot
    fig.add_trace(go.Scattergeo(
        lat=[lat], lon=[lon], mode="markers",
        marker=dict(size=11, color="#7C3AED",
                    line=dict(color="#c4b5fd", width=2), symbol="circle"),
        hovertemplate=f"<b>{label or country}</b><extra></extra>",
        showlegend=False,
    ))
    # Company name label (only if provided)
    if label:
        fig.add_trace(go.Scattergeo(
            lat=[lat], lon=[lon], mode="text",
            text=[f"  {label}"],
            textposition="middle right",
            textfont=dict(color="#a78bfa" if dark else "#5b21b6", size=11, family="Inter"),
            hoverinfo="skip", showlegend=False,
        ))
    fig.update_geos(
        projection_type="orthographic",
        projection_rotation=dict(lon=lon, lat=lat, roll=0),
        showland=True, landcolor=p["globe_land"],
        showocean=True, oceancolor=p["globe_ocean"],
        showcoastlines=True, coastlinecolor=p["globe_coast"], coastlinewidth=0.6,
        showcountries=True, countrycolor=p["globe_country"], countrywidth=0.4,
        showlakes=False, bgcolor="rgba(0,0,0,0)",
    )
    fig.update_layout(
        height=380, margin=dict(l=0, r=0, t=0, b=0),
        paper_bgcolor="rgba(0,0,0,0)", showlegend=False,
        geo=dict(bgcolor="rgba(0,0,0,0)"),
    )
    return fig


def hex_rgba(hex_c: str, alpha: float) -> str:
    h = hex_c.lstrip("#")
    r, g, b = int(h[0:2], 16), int(h[2:4], 16), int(h[4:6], 16)
    return f"rgba({r},{g},{b},{alpha})"


def calc_risk_score(cat_counts, total):
    weights = {"風險事件": 3.0, "國際動態": 1.8, "政策法規": 1.2,
               "財務數據": 0.8, "市場趨勢": 0.4, "技術創新": 0.2, "其他": 0.3}
    raw = sum(cat_counts.get(c, 0) * w for c, w in weights.items())
    cap = total * 3.0
    score = min(100, int(raw / cap * 100)) if cap > 0 else 0
    level = "高" if score >= 60 else "中" if score >= 30 else "低"
    color = "#EF4444" if score >= 60 else "#F59E0B" if score >= 30 else "#10B981"
    return score, level, color


def build_risk_gauge(score, level, risk_color, dark, p):
    fig = go.Figure(go.Indicator(
        mode="gauge+number",
        value=score,
        number={"font": {"size": 38, "color": p["text"], "family": "Inter"},
                "suffix": ""},
        title={"text": f"外部風險指數<br><span style='font-size:13px;color:{risk_color};font-weight:600'>{level}風險</span>",
               "font": {"size": 12, "color": p["text2"], "family": "Inter"}},
        gauge={
            "axis": {"range": [0, 100], "tickwidth": 1,
                     "tickfont": {"size": 9, "color": p["text3"]}, "tickcolor": p["border"]},
            "bar": {"color": risk_color, "thickness": 0.28},
            "bgcolor": "rgba(0,0,0,0)", "borderwidth": 0,
            "steps": [
                {"range": [0, 30],  "color": "rgba(16,185,129,0.18)"},
                {"range": [30, 60], "color": "rgba(245,158,11,0.18)"},
                {"range": [60, 100],"color": "rgba(239,68,68,0.18)"},
            ],
            "threshold": {"line": {"color": risk_color, "width": 3},
                          "thickness": 0.8, "value": score},
        },
    ))
    fig.update_layout(height=210, margin=dict(l=20, r=20, t=30, b=0),
                      paper_bgcolor="rgba(0,0,0,0)", font=dict(family="Inter"))
    return fig


def build_radar(cat_counts, dark, p):
    cats = ["政策法規", "市場趨勢", "技術創新", "財務數據", "風險事件", "國際動態"]
    vals = [cat_counts.get(c, 0) for c in cats]
    vals_closed = vals + [vals[0]]
    cats_closed  = cats + [cats[0]]
    fig = go.Figure()
    fig.add_trace(go.Scatterpolar(
        r=vals_closed, theta=cats_closed, fill="toself",
        fillcolor=hex_rgba(p["accent"], 0.16),
        line=dict(color=p["accent"], width=2.5),
        marker=dict(size=8, color=p["accent"], line=dict(color="#ffffff", width=1.5)),
        name="情報分佈",
    ))
    fig.update_layout(
        polar=dict(
            bgcolor="rgba(0,0,0,0)",
            angularaxis=dict(tickfont=dict(size=10, color=p["text2"], family="Inter"),
                             linecolor=p["border"], gridcolor=p["chart_grid"]),
            radialaxis=dict(visible=True, tickformat="d",
                            tickfont=dict(size=8, color=p["text3"]),
                            gridcolor=p["chart_grid"], linecolor=p["border"]),
        ),
        height=300, margin=dict(l=60, r=60, t=20, b=20),
        paper_bgcolor="rgba(0,0,0,0)", showlegend=False, font=dict(family="Inter"),
    )
    return fig


def build_sunburst(company_name, raw_results, dark, p):
    if not raw_results:
        return None
    ids, labels, parents, colors, hover = [], [], [], [], []

    short_co = (company_name[:9] + "…") if len(company_name) > 9 else company_name
    ids.append("root"); labels.append(short_co); parents.append("")
    colors.append(p["accent"]); hover.append(f"<b>{company_name}</b><br>分析主體")

    cat_map: dict[str, list[int]] = {}
    for i, r in enumerate(raw_results):
        cat = categorize_article(r.get("title", ""), r.get("content", ""))
        cat_map.setdefault(cat, []).append(i)

    for cat, idxs in cat_map.items():
        cid = f"cat::{cat}"
        ids.append(cid); labels.append(cat); parents.append("root")
        colors.append(CATEGORY_COLORS.get(cat, "#6B7280"))
        hover.append(f"<b>{cat}</b><br>{len(idxs)} 篇新聞")

    for cat, idxs in cat_map.items():
        cid = f"cat::{cat}"
        base = CATEGORY_COLORS.get(cat, "#6B7280")
        for idx in idxs:
            r = raw_results[idx]
            title = r.get("title", f"新聞 {idx+1}")
            short = (title[:22] + "…") if len(title) > 22 else title
            score = r.get("score")
            tip = f"<b>{title}</b>"
            if score:
                tip += f"<br>相關度：{score:.2f}"
            ids.append(f"art::{idx}"); labels.append(short)
            parents.append(cid); colors.append(hex_rgba(base, 0.75)); hover.append(tip)

    fig = go.Figure(go.Sunburst(
        ids=ids, labels=labels, parents=parents,
        marker=dict(colors=colors, line=dict(width=1.5, color="rgba(255,255,255,0.3)")),
        hovertext=hover, hoverinfo="text",
        textfont=dict(size=10, color="#ffffff", family="Inter"),
        insidetextorientation="radial",
        maxdepth=3,
        branchvalues="remainder",
    ))
    fig.update_layout(
        height=480, margin=dict(l=0, r=0, t=10, b=0),
        paper_bgcolor="rgba(0,0,0,0)",
        hoverlabel=dict(bgcolor=p["surface"], bordercolor=p["border"],
                        font=dict(family="Inter", size=12, color=p["text"])),
    )
    return fig


def extract_summary(report):
    for line in report.split("\n"):
        ls = line.strip()
        if len(ls) > 40 and not ls.startswith("#") and not ls.startswith("-"):
            ls = re.sub(r'\*\*(.*?)\*\*', r'\1', ls)
            return ls[:220] + ("…" if len(ls) > 220 else "")
    return ""


def split_report_sections(report):
    parts = re.split(r'(?m)(?=^#{1,3}\s)', report)
    parts = [s.strip() for s in parts if s.strip()]
    sec1 = sec2 = sec3 = ""
    for s in parts:
        if not sec1 and re.search(r'PESTLE|外部環境|政治|Economic|Social|Technology', s, re.I):
            sec1 = s
        elif not sec2 and re.search(r'機會|威脅|Opportunity|Threat', s):
            sec2 = s
        elif not sec3 and re.search(r'稽核|建議|總結|結論|Recommendation', s, re.I):
            sec3 = s
    if not sec1 and not sec2 and not sec3:
        n = len(report); sec1 = report
    elif not sec1:
        sec1 = report
    return sec1, sec2, sec3


def generate_pdf_bytes(res, report_text):
    from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, HRFlowable
    from reportlab.lib.pagesizes import A4
    from reportlab.lib.styles import ParagraphStyle
    from reportlab.lib.units import mm
    from reportlab.pdfbase import pdfmetrics
    from reportlab.pdfbase.cidfonts import UnicodeCIDFont
    from reportlab.lib.colors import HexColor, white

    try:
        pdfmetrics.registerFont(UnicodeCIDFont("STSong-Light"))
        FONT = "STSong-Light"
    except Exception:
        FONT = "Helvetica"

    VIOLET   = HexColor("#7C3AED")
    VIOLET_L = HexColor("#a78bfa")
    VIOLET_B = HexColor("#f5f3ff")
    DARK     = HexColor("#0f172a")
    MID      = HexColor("#475569")
    LIGHT    = HexColor("#94a3b8")
    BORDER   = HexColor("#e2d9f3")
    W, H     = A4
    buf      = io.BytesIO()

    def draw_page(canvas, doc):
        canvas.saveState()
        # Dot-grid tech background
        canvas.setFillColor(HexColor("#fdfcff"))
        canvas.rect(0, 0, W, H, fill=1, stroke=0)
        canvas.setFillColor(HexColor("#ede9fe"))
        for gx in range(0, int(W) + 20, 20):
            for gy in range(0, int(H) + 20, 20):
                canvas.circle(gx, gy, 0.55, fill=1, stroke=0)
        # Violet header band
        canvas.setFillColor(VIOLET)
        canvas.rect(0, H - 28*mm, W, 28*mm, fill=1, stroke=0)
        canvas.setFillColor(VIOLET_L)
        canvas.rect(0, H - 28*mm, W, 1.5, fill=1, stroke=0)
        # ISO badge
        canvas.setFillColor(HexColor("#6d28d9"))
        canvas.roundRect(18*mm, H - 14*mm, 22*mm, 6*mm, 2, fill=1, stroke=0)
        canvas.setFillColor(white)
        canvas.setFont(FONT, 7)
        canvas.drawString(19.5*mm, H - 11.5*mm, "ISO 9001  ·  4.1")
        # Report title
        canvas.setFillColor(white)
        canvas.setFont(FONT, 15)
        canvas.drawString(44*mm, H - 13.5*mm, "外部環境分析報告")
        # Meta line
        canvas.setFont(FONT, 9)
        canvas.setFillColor(HexColor("#c4b5fd"))
        canvas.drawString(18*mm, H - 23*mm,
            f"{res['company_name']}  ·  {res['country']}  ·  {res['industry_short']}  ·  {res['model_label'].split('（')[0]}")
        # Footer
        canvas.setFillColor(VIOLET_B)
        canvas.rect(0, 0, W, 16*mm, fill=1, stroke=0)
        canvas.setFillColor(VIOLET)
        canvas.rect(0, 16*mm, W, 0.7, fill=1, stroke=0)
        canvas.setFont(FONT, 8)
        canvas.setFillColor(MID)
        canvas.drawString(18*mm, 7*mm, f"ISO 4.1 外部環境分析平台  ·  AI Powered  ·  情報來源：{len(res['raw'])} 筆")
        canvas.setFillColor(LIGHT)
        canvas.drawRightString(W - 18*mm, 7*mm, f"第 {doc.page} 頁")
        canvas.restoreState()

    doc = SimpleDocTemplate(
        buf, pagesize=A4,
        topMargin=33*mm, bottomMargin=22*mm,
        leftMargin=18*mm, rightMargin=18*mm,
        title=f"ISO 4.1 外部環境分析 — {res['company_name']}",
        author="ISO 4.1 外部環境分析平台",
    )

    def S(**kw):
        return ParagraphStyle("s", fontName=FONT, **kw)

    sh2   = S(fontSize=13, textColor=VIOLET,  spaceBefore=14, spaceAfter=5,  leading=20)
    sbody = S(fontSize=10, textColor=MID,     spaceAfter=3,  leading=17)
    sbull = S(fontSize=10, textColor=MID,     spaceAfter=2,  leading=17, leftIndent=10)
    smeta = S(fontSize=9,  textColor=LIGHT,   spaceAfter=8,  leading=14)

    story = []
    story.append(Paragraph(res["company_name"], S(fontSize=18, textColor=DARK, spaceAfter=4, leading=24)))
    story.append(Paragraph(
        f"{res['country']}  ·  {res['industry_short']}  ·  情報來源 {len(res['raw'])} 筆", smeta))
    story.append(HRFlowable(width="100%", thickness=0.7, color=BORDER, spaceAfter=14))

    for line in report_text.split("\n"):
        ls = line.strip()
        if not ls:
            story.append(Spacer(1, 5)); continue
        ls = re.sub(r'\*\*(.*?)\*\*', r'<b>\1</b>', ls)
        if re.match(r'^#{1,3}\s', ls):
            story.append(Paragraph(re.sub(r'^#+\s*', '', ls), sh2))
        elif ls.startswith("- ") or ls.startswith("• "):
            story.append(Paragraph(f"&bull; {ls[2:].strip()}", sbull))
        else:
            story.append(Paragraph(ls, sbody))

    doc.build(story, onFirstPage=draw_page, onLaterPages=draw_page)
    return buf.getvalue()

# ── Hero + Form ───────────────────────────────────────────────────────────────
_ap = P["accent"]
_ad = P["accent_dim"]
_ab = P["accent_border"]
_t  = P["text"]
_t2 = P["text2"]
_t3 = P["text3"]
_sf = P["surface"]
_br = P["border"]

# Settings button row (flush top-right)
_, _scol = st.columns([9, 1])
with _scol:
    with st.popover("⚙", use_container_width=True):
        dark_on = st.toggle(T("dark_mode"), value=st.session_state.dark, key="_theme")
        if dark_on != st.session_state.dark:
            st.session_state.dark = dark_on
            st.rerun()
        st.divider()
        lang_map = {"中文": "zh", "English": "en", "Deutsch": "de"}
        lang_labels = list(lang_map.keys())
        cur_label = next((k for k, v in lang_map.items() if v == st.session_state.lang), "中文")
        chosen = st.radio(T("language"), lang_labels,
                          index=lang_labels.index(cur_label), key="_lang_radio")
        if lang_map[chosen] != st.session_state.lang:
            st.session_state.lang = lang_map[chosen]
            st.rerun()

# Hero: left text+form | right globe
hero_l, hero_r = st.columns([6, 4], gap="large")

with hero_l:
    # Badge
    st.markdown(
        f'<p style="font-size:0.58rem;color:{_ap};letter-spacing:0.09em;text-transform:uppercase;'
        f'font-weight:700;margin:0.8rem 0 1rem;font-family:Inter,sans-serif;white-space:nowrap;">'
        f'{T("badge")}</p>',
        unsafe_allow_html=True,
    )
    # Big title
    st.markdown(
        f'<h1 style="font-size:2.7rem;font-weight:800;letter-spacing:-0.03em;line-height:1.05;'
        f'color:{_t};margin:0 0 0.8rem;font-family:Inter,sans-serif;">'
        f'{T("title")}</h1>',
        unsafe_allow_html=True,
    )
    # Subtitle
    st.markdown(
        f'<p style="font-size:0.88rem;color:{_t2};line-height:1.75;'
        f'margin:0 0 1.6rem;font-family:Inter,sans-serif;">{T("subtitle")}</p>',
        unsafe_allow_html=True,
    )
    # ── Form ───────────────────────────────────────────────────────────────────
    st.markdown(f'<div style="height:0.1rem;border-top:1px solid {_br};margin:0.6rem 0 1.2rem;"></div>',
                unsafe_allow_html=True)
    company_name = st.text_input(T("company_label"), placeholder=T("company_ph"))
    country      = st.selectbox(T("country_label"), COUNTRIES)
    industry     = st.selectbox(T("industry_label"), INDUSTRIES, index=12)
    website      = st.text_input(T("website_label"), placeholder=T("website_ph"))

    model_label = st.selectbox(T("model_label"), list(MODELS.keys()), index=0)
    selected_model = MODELS[model_label]

    st.markdown("<div style='height:0.3rem'></div>", unsafe_allow_html=True)
    run = st.button(T("run_btn"), use_container_width=True, type="primary")

with hero_r:
    st.markdown("<div style='height:2.2rem'></div>", unsafe_allow_html=True)
    pin_fig = build_pin_globe(country, company_name, st.session_state.dark, P)
    pin_fig.update_layout(height=190, margin=dict(l=0, r=0, t=0, b=0))
    st.plotly_chart(pin_fig, use_container_width=True, key="pin_globe")

# ── Run ───────────────────────────────────────────────────────────────────────
if run:
    if not company_name:
        st.warning(T("warn_company"))
    elif not website:
        st.warning(T("warn_website"))
    elif not os.environ.get("OPENAI_API_KEY"):
        st.error("找不到 OPENAI_API_KEY，請在環境變數中設定您的 OpenAI API 金鑰。")
    elif not os.environ.get("TAVILY_API_KEY"):
        st.error("找不到 TAVILY_API_KEY，請在環境變數中設定您的 Tavily API 金鑰。")
    else:
        industry_short = industry.split(" - ")[-1] if " - " in industry else industry

        steps = [
            {"id":"init",     "label":"初始化搜尋引擎",                     "status":"pending","detail":""},
            {"id":"s1",       "label":"搜尋趨勢與風險相關新聞（15 筆）",    "status":"pending","detail":""},
            {"id":"s2",       "label":"搜尋法規與政策相關新聞（10 筆）",    "status":"pending","detail":""},
            {"id":"dedup",    "label":"合併去重，篩選最相關情報",            "status":"pending","detail":""},
            {"id":"website",  "label":"爬取官方網站資料",                   "status":"pending","detail":""},
            {"id":"ai",       "label":f"AI 模型分析（{model_label.split('（')[0]}）","status":"pending","detail":""},
            {"id":"parse",    "label":"解析報告與知識圖譜標籤",              "status":"pending","detail":""},
        ]
        step_box = st.empty()

        def mark(sid, status, detail=""):
            for s in steps:
                if s["id"] == sid:
                    s["status"] = status
                    if detail:
                        s["detail"] = detail
                    break
            step_box.markdown(render_steps(steps, P), unsafe_allow_html=True)

        mark("init", "running")
        step_box.markdown(render_steps(steps, P), unsafe_allow_html=True)
        tavily_client = TavilyClient(api_key=os.environ["TAVILY_API_KEY"])
        mark("init", "done", "就緒")

        mark("s1", "running")
        q1 = f"{country} {industry_short} {company_name} 趨勢 風險 最新新聞 2026"
        r1 = tavily_client.search(query=q1, max_results=15).get("results", [])
        mark("s1", "done", f"取得 {len(r1)} 筆")

        mark("s2", "running")
        q2 = f"{country} {industry_short} {company_name} 法規 政策 監管 產業標準 2026"
        r2 = tavily_client.search(query=q2, max_results=10).get("results", [])
        mark("s2", "done", f"取得 {len(r2)} 筆")

        mark("dedup", "running")
        seen, raw = set(), []
        for r in r1 + r2:
            u = r.get("url", "")
            if u not in seen:
                seen.add(u)
                raw.append(r)
        raw = raw[:25]
        mark("dedup", "done", f"共 {len(raw)} 筆")

        co_ctx = ""
        if website:
            mark("website", "running")
            try:
                ext = tavily_client.extract(urls=[website]).get("results", [])
                if ext:
                    co_ctx = f"\n\n【企業官網摘要（{website}）】\n{ext[0].get('raw_content','')[:2000]}"
                mark("website", "done", "擷取完成")
            except Exception:
                mark("website", "done", "擷取失敗，略過")
        else:
            mark("website", "done", "略過（未填寫）")

        mark("ai", "running", "生成中...")
        ctx = "\n\n---\n\n".join(
            f"[{i}] 【{r.get('title','')}】\n來源：{r.get('url','')}\n{r.get('content','')}"
            for i, r in enumerate(raw)
        )
        sys_p = f"""你是一位擁有 20 年經驗的 ISO 9001 首席稽核員及企業戰略顧問。

請根據最新新聞情報（以及企業官網資料，若有提供），針對以下企業進行 ISO 條文 4.1 外部議題分析：
- 企業名稱：{company_name}
- 所在國家：{country}
- 所屬產業：{industry_short}
{f"- 官方網站：{website}（已爬取官網內容作為企業背景）" if website else ""}

請以繁體中文輸出：

### 1. PESTLE 外部環境分析
結合 {country} 政經環境與產業特性，選 3-4 個最相關面向深入說明。

### 2. 外部機會與威脅分析
列出 2 項機會、2 項威脅，需與 {country} 產業政策及市場現況高度相關。

### 3. 稽核員總結建議
具體且針對性強的建議，點名最需要立即處理的外部風險。

---
報告最後輸出 JSON（```json 包住），為每篇新聞提供簡短繁體中文標題（≤15字）與 3~5 個關鍵主題標籤：
```json
{{"article_tags":[{{"index":0,"short_title":"簡短標題","tags":["主題A","主題B"]}},...] }}
```
相關文章請共用標籤，以利關聯網絡呈現。"""

        usr_p = f"以下是最新搜尋情報（共 {len(raw)} 筆）：\n\n{ctx}{co_ctx}\n\n請針對【{company_name}】撰寫完整 ISO 4.1 分析報告，並附上 JSON 標籤資料。"
        client = OpenAI(api_key=os.environ["OPENAI_API_KEY"])
        resp = client.chat.completions.create(
            model=selected_model,
            messages=[{"role":"system","content":sys_p},{"role":"user","content":usr_p}],
            temperature=0.7,
        )
        full = resp.choices[0].message.content
        mark("ai", "done", "完成")

        mark("parse", "running")
        jm = re.search(r"```json\s*([\s\S]*?)\s*```", full)
        tags = []
        if jm:
            try:
                tags = json.loads(jm.group(1)).get("article_tags", [])
            except Exception:
                pass

        # ── Clean report: strip JSON block + any label lines before it ──────
        report = full
        # Remove code blocks
        report = re.sub(r"```json[\s\S]*?```", "", report)
        # Remove common separator + JSON label lines (e.g. "---\n以下是...標籤資料")
        report = re.sub(r"\n?---+\n[^\n]*(?:JSON|json|標籤|article)[^\n]*", "", report)
        # Remove stray trailing separators
        report = re.sub(r"\n?---+\s*$", "", report).strip()

        # ── Parse into 3 sections at store time (most reliable) ─────────────
        def _parse_sections(text):
            parts = re.split(r'\n(?=### )', text)
            sec1 = sec2 = sec3 = ""
            for sp in parts:
                sp = sp.strip()
                if not sp:
                    continue
                fl = sp.split('\n')[0]
                if not sec1 and re.search(r'PESTLE|外部環境', sp, re.I):
                    sec1 = sp
                elif not sec2 and re.search(r'機會|威脅', sp):
                    sec2 = sp
                elif not sec3 and re.search(r'稽核|建議|總結', sp, re.I):
                    sec3 = sp
            # Fallback: use position (1st, 2nd, 3rd heading)
            if not (sec1 or sec2 or sec3):
                numbered = [p.strip() for p in parts if p.strip()]
                if len(numbered) >= 1: sec1 = numbered[0]
                if len(numbered) >= 2: sec2 = numbered[1]
                if len(numbered) >= 3: sec3 = numbered[2]
            return sec1 or text, sec2, sec3

        sec1, sec2, sec3 = _parse_sections(report)
        mark("parse", "done", f"標籤 {len(tags)} 篇")

        cat_counts, cat_list = {}, []
        for r in raw:
            c = categorize_article(r.get("title",""), r.get("content",""))
            cat_list.append(c)
            cat_counts[c] = cat_counts.get(c, 0) + 1

        st.session_state.results = dict(
            raw=raw, report=report, tags=tags,
            sec1=sec1, sec2=sec2, sec3=sec3,
            cat_counts=cat_counts, cat_list=cat_list,
            company_name=company_name, country=country,
            industry_short=industry_short, model_label=model_label,
            website=website,
        )
        step_box.empty()
        st.rerun()

# ── Results ───────────────────────────────────────────────────────────────────
res = st.session_state.results
if res:
    dark = st.session_state.dark
    P2 = palette(dark)

    # ── Tab group: Data panels ────────────────────────────────────────────────
    st.divider()
    td1, td2, td3 = st.tabs([T("tab_info"), T("tab_dist"), T("tab_graph")])

    with td1:
        mc = st.columns(4)
        mc[0].metric(T("metric_total"), f"{len(res['raw'])} 筆")
        mc[1].metric(T("metric_types"), f"{len(res['cat_counts'])} 類")
        mc[2].metric(T("metric_company"), res["company_name"])
        mc[3].metric(T("metric_model"), res["model_label"].split("（")[0])
        st.markdown("<div style='height:0.4rem'></div>", unsafe_allow_html=True)
        g_col, r_col = st.columns([3, 2])
        with g_col:
            st.caption(f"{res['company_name']} · {res['country']} — {T('globe_cap')}")
            globe_fig = build_globe(res["country"], res["company_name"], dark, P2)
            globe_fig.update_layout(height=230, margin=dict(l=0, r=0, t=0, b=0))
            st.plotly_chart(globe_fig, use_container_width=True)
        with r_col:
            r_score, r_level, r_color = calc_risk_score(res["cat_counts"], len(res["raw"]))
            gauge_fig = build_risk_gauge(r_score, r_level, r_color, dark, P2)
            st.plotly_chart(gauge_fig, use_container_width=True)

    with td2:
        st.caption(T("dist_cap", n=len(res['raw'])))
        rc, bc = st.columns([1, 1])
        with rc:
            radar_fig = build_radar(res["cat_counts"], dark, P2)
            st.plotly_chart(radar_fig, use_container_width=True)
        with bc:
            bar_fig = go.Figure(go.Bar(
                x=list(res["cat_counts"].keys()), y=list(res["cat_counts"].values()),
                marker=dict(color=[CATEGORY_COLORS.get(c, "#6B7280") for c in res["cat_counts"]],
                            line=dict(color="rgba(0,0,0,0)", width=0)),
                text=list(res["cat_counts"].values()), textposition="outside",
                textfont=dict(color=P2["text3"], size=12, family="Inter"),
            ))
            bar_fig.update_layout(
                height=300,
                yaxis=dict(tickformat="d", dtick=1, showgrid=True, gridcolor=P2["chart_grid"],
                           tickfont=dict(color=P2["text3"], size=11, family="Inter")),
                xaxis=dict(tickfont=dict(color=P2["text2"], size=11, family="Inter"), showgrid=False),
                plot_bgcolor=P2["plot_bg"], paper_bgcolor=P2["paper_bg"],
                margin=dict(l=10, r=10, t=20, b=10), font=dict(family="Inter"),
            )
            st.plotly_chart(bar_fig, use_container_width=True)

    with td3:
        st.caption(T("graph_cap"))
        sb_fig = build_sunburst(res["company_name"], res["raw"], dark, P2)
        if sb_fig:
            st.plotly_chart(sb_fig, use_container_width=True)
        else:
            st.info(T("no_data"))

    # ── Report ────────────────────────────────────────────────────────────────
    st.divider()
    rh_col, pdf_col = st.columns([5, 2])
    with rh_col:
        st.subheader(f"{res['company_name']} — {T('report_hdr')}")
        st.caption(f"{res['country']}  ·  {res['industry_short']}  ·  {res['model_label'].split('（')[0]}")
    with pdf_col:
        st.markdown("<div style='height:1.8rem'></div>", unsafe_allow_html=True)
        try:
            pdf_bytes = generate_pdf_bytes(res, res["report"])
            st.download_button(
                T("pdf_btn"),
                data=pdf_bytes,
                file_name=f"ISO41_{res['company_name']}_分析報告.pdf",
                mime="application/pdf",
                use_container_width=True,
            )
        except Exception:
            st.caption("PDF 暫時無法生成")

    # Executive summary banner
    _summary = extract_summary(res["report"])
    if _summary:
        st.markdown(
            f'<div style="border-left:3px solid {P2["accent"]};padding:0.65rem 1rem;'
            f'background:{P2["surface"]};border-radius:0 8px 8px 0;'
            f'margin:0.6rem 0 1.2rem;font-size:0.88rem;color:{P2["text2"]};'
            f'font-family:Inter,sans-serif;line-height:1.65;">'
            f'<span style="font-size:0.68rem;font-weight:700;letter-spacing:0.1em;'
            f'color:{P2["accent"]};text-transform:uppercase;">{T("summary_lbl")}</span>'
            f'<br>{_summary}</div>',
            unsafe_allow_html=True,
        )

    # ── Full-page continuous report (no tabs) ─────────────────────────────────
    sec1 = res.get("sec1") or res["report"]
    sec2 = res.get("sec2", "")
    sec3 = res.get("sec3", "")

    def _clean_md(text: str) -> str:
        """Strip markdown decorations; keep numbered lists and blank lines, remove the rest."""
        out = []
        for line in text.split('\n'):
            # Strip heading markers, keep the text
            line = re.sub(r'^#{1,6}\s+', '', line)
            # Strip bold / italic / bold-italic
            line = re.sub(r'\*{1,3}([^*\n]+?)\*{1,3}', r'\1', line)
            line = re.sub(r'_{1,2}([^_\n]+?)_{1,2}', r'\1', line)
            # Convert bullet lines to plain paragraphs
            line = re.sub(r'^[-*•]\s+', '', line)
            out.append(line)
        return '\n'.join(out)

    def _sec_strip_heading(text):
        lines = text.strip().split('\n')
        if lines and re.match(r'^#{1,4}\s', lines[0]):
            return '\n'.join(lines[1:]).strip()
        return text

    def _sec_header(num, label):
        border_c = hex_rgba(P2["accent"], 0.20)
        return (
            f'<div style="display:flex;align-items:center;gap:12px;'
            f'margin:2.2rem 0 1rem;padding-bottom:0.6rem;'
            f'border-bottom:1.5px solid {border_c};">'
            f'<span style="background:{P2["accent"]};color:#fff;font-size:0.6rem;'
            f'font-weight:700;letter-spacing:0.12em;padding:3px 10px;border-radius:3px;">'
            f'0{num}</span>'
            f'<span style="font-size:1.1rem;font-weight:700;color:{P2["text"]};'
            f'font-family:Inter,sans-serif;">{label}</span>'
            f'</div>'
        )

    if sec1:
        st.markdown(_sec_header(1, T("tab_pestle")), unsafe_allow_html=True)
        st.markdown(_clean_md(_sec_strip_heading(sec1)))
    if sec2:
        st.markdown(_sec_header(2, T("tab_ot")), unsafe_allow_html=True)
        st.markdown(_clean_md(_sec_strip_heading(sec2)))
    if sec3:
        st.markdown(_sec_header(3, T("tab_audit")), unsafe_allow_html=True)
        st.markdown(_clean_md(_sec_strip_heading(sec3)))
    if not (sec1 or sec2 or sec3):
        st.info(T("no_data"))

    # ── Sources ───────────────────────────────────────────────────────────────
    st.divider()
    st.subheader(T("sources_title"))
    st.caption(T("sources_cap", n=len(res['raw'])))
    rows = ""
    for i, r in enumerate(res["raw"]):
        title = r.get("title", f"來源 {i+1}")
        url   = r.get("url", "")
        score = r.get("score", None)
        cat   = res["cat_list"][i] if i < len(res["cat_list"]) else "其他"
        cc    = CATEGORY_COLORS.get(cat, "#6B7280")
        sc_html = (f'<span style="color:{P2["text3"]};font-size:0.7rem;white-space:nowrap;flex-shrink:0;">'
                   f'{T("relevance")} {score:.2f}</span>') if score else ""
        lk_html = (f'<a href="{url}" target="_blank" style="flex-shrink:0;font-size:0.72rem;'
                   f'color:{P2["accent"]};text-decoration:none;white-space:nowrap;'
                   f'border:1px solid {P2["accent_border"]};border-radius:4px;padding:2px 8px;">{T("open")}</a>') if url else ""
        ts = (title[:68] + "…") if len(title) > 68 else title
        rows += (
            f'<div style="display:flex;align-items:center;gap:10px;padding:0.6rem 0.85rem;'
            f'border-radius:8px;background:{P2["surface"]};border:1px solid {P2["border2"]};margin-bottom:5px;">'
            f'<span style="font-size:0.62rem;font-weight:700;letter-spacing:0.04em;color:{cc};'
            f'border:1px solid {cc}50;border-radius:4px;padding:2px 6px;white-space:nowrap;flex-shrink:0;">{cat}</span>'
            f'<span style="font-size:0.84rem;color:{P2["text2"]};flex:1;overflow:hidden;'
            f'text-overflow:ellipsis;white-space:nowrap;" title="{title}">{ts}</span>'
            f'{sc_html}{lk_html}</div>'
        )
    st.markdown(rows, unsafe_allow_html=True)
    st.balloons()
