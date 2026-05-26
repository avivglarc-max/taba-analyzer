"""
app.py – מנתח תב"ע AI
עיצוב זהה לדמו HTML המקורי
"""

import os, json
import streamlit as st
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
from dotenv import load_dotenv

load_dotenv()

from pdf_processor import extract_text_from_pdf
from ai_analyzer import (
    create_client, analyze_single_chunk,
    analyze_long_document, validate_calculations
)
from excel_exporter import export_to_excel

# ─── הגדרות עמוד ────────────────────────────────────────
st.set_page_config(
    page_title='מנתח תב"ע AI',
    page_icon="🏗️",
    layout="wide",
    initial_sidebar_state="expanded"
)

# ─── CSS – עיצוב זהה לדמו ───────────────────────────────
st.markdown("""
<style>
@import url('https://fonts.googleapis.com/css2?family=Heebo:wght@300;400;500;600;700;800;900&family=IBM+Plex+Mono:wght@400;600&display=swap');

:root {
  --ink:    #0f1923;
  --paper:  #f4f1eb;
  --blue:   #1b3f72;
  --teal:   #1d7874;
  --amber:  #e8a020;
  --red:    #c0392b;
  --green:  #1a7a4a;
  --muted:  #8a8a8a;
  --border: #d6d0c4;
  --card:   #ffffff;
}

html, body, [class*="css"] {
  font-family: 'Heebo', sans-serif !important;
  direction: rtl;
}

/* רקע עמוד */
.stApp { background: var(--paper) !important; }

/* סרגל צד */
section[data-testid="stSidebar"] {
  background: #0f1923 !important;
  border-left: none !important;
}
section[data-testid="stSidebar"] * { color: #e8e4dc !important; }
section[data-testid="stSidebar"] .stTextInput input,
section[data-testid="stSidebar"] .stSelectbox select,
section[data-testid="stSidebar"] .stNumberInput input {
  background: rgba(255,255,255,0.07) !important;
  border: 1px solid rgba(255,255,255,0.12) !important;
  color: #fff !important;
  border-radius: 8px !important;
  font-family: 'Heebo', sans-serif !important;
}
section[data-testid="stSidebar"] label {
  font-size: 0.68rem !important;
  font-weight: 700 !important;
  letter-spacing: 1.5px !important;
  text-transform: uppercase !important;
  color: rgba(255,255,255,0.35) !important;
}

/* כפתור ניתוח */
.stButton > button {
  background: linear-gradient(135deg, #1d7874 0%, #155c59 100%) !important;
  color: #fff !important;
  border: none !important;
  border-radius: 10px !important;
  padding: 13px !important;
  font-family: 'Heebo', sans-serif !important;
  font-size: 0.95rem !important;
  font-weight: 700 !important;
  width: 100% !important;
  transition: all 0.2s !important;
}
.stButton > button:hover {
  background: linear-gradient(135deg, #155c59 0%, #0f4440 100%) !important;
  transform: translateY(-1px) !important;
}

/* KPI כרטיסים */
.kpi-card {
  background: var(--card);
  border: 1px solid var(--border);
  border-radius: 12px;
  padding: 18px 20px;
  position: relative;
  overflow: hidden;
  margin-bottom: 14px;
}
.kpi-card::before {
  content: '';
  position: absolute;
  top: 0; right: 0;
  width: 4px; height: 100%;
  background: var(--accent, var(--teal));
}
.kpi-label {
  font-size: 0.7rem; font-weight: 700;
  letter-spacing: 1px; text-transform: uppercase;
  color: var(--muted); margin-bottom: 8px;
}
.kpi-value {
  font-size: 1.9rem; font-weight: 900;
  line-height: 1; color: var(--ink);
  font-family: 'IBM Plex Mono', monospace;
}
.kpi-sub { font-size: 0.72rem; color: var(--muted); margin-top: 5px; }
.progress-bar {
  height: 6px; background: #e9ecef;
  border-radius: 3px; overflow: hidden; margin-top: 6px;
}
.progress-fill {
  height: 100%; border-radius: 3px;
  background: linear-gradient(90deg, #1d7874, #2ab5af);
}

/* טבלאות */
.tbl-wrap {
  background: var(--card);
  border: 1px solid var(--border);
  border-radius: 12px;
  overflow: hidden;
  margin-bottom: 20px;
}
.tbl-header {
  padding: 14px 20px;
  border-bottom: 1px solid var(--border);
  font-size: 0.8rem; font-weight: 700;
  letter-spacing: 0.5px; color: var(--muted);
  text-transform: uppercase;
}
table { width: 100%; border-collapse: collapse; font-size: 0.875rem; }
thead th {
  background: #f8f7f4; padding: 10px 16px;
  text-align: right; font-weight: 700;
  font-size: 0.75rem; color: var(--muted);
  text-transform: uppercase;
  border-bottom: 1px solid var(--border);
}
tbody tr { border-bottom: 1px solid rgba(0,0,0,0.04); }
tbody tr:hover { background: rgba(29,120,116,0.03); }
tbody td { padding: 12px 16px; }
.total-row td { background: #0f1923; color: #fff; font-weight: 700; }

.badge { display: inline-block; padding: 2px 8px; border-radius: 4px; font-size: 0.7rem; font-weight: 700; }
.badge-1 { background:#dbeafe; color:#1e3a8a; }
.badge-2 { background:#dcfce7; color:#14532d; }
.badge-3 { background:#fef3c7; color:#713f12; }
.badge-4 { background:#fce7f3; color:#831843; }
.badge-5 { background:#ede9fe; color:#4c1d95; }

/* כרטיסי סיכון */
.risk-card {
  background: var(--card); border: 1px solid var(--border);
  border-radius: 10px; padding: 14px 16px 14px 20px;
  border-right: 4px solid; margin-bottom: 10px;
  display: grid; grid-template-columns: auto 1fr auto;
  align-items: start; gap: 12px;
}
.risk-card.high   { border-right-color: #c0392b; background: #fdf2f2; }
.risk-card.medium { border-right-color: #e8a020; background: #fffbf0; }
.risk-card.low    { border-right-color: #1a7a4a; background: #f0faf5; }
.risk-icon { font-size: 1.2rem; }
.risk-body h4 { font-size: 0.85rem; font-weight: 700; margin-bottom: 4px; color: var(--ink); }
.risk-body p  { font-size: 0.78rem; color: var(--muted); line-height: 1.5; }
.risk-tag { font-size: 0.65rem; font-weight: 700; padding: 2px 8px; border-radius: 4px; white-space: nowrap; }
.risk-card.high   .risk-tag { background:#fde8e8; color:#c0392b; }
.risk-card.medium .risk-tag { background:#fef3cd; color:#92400e; }
.risk-card.low    .risk-tag { background:#d1fae5; color:#1a7a4a; }

/* זכויות */
.rights-card {
  background: var(--card); border: 1px solid var(--border);
  border-radius: 12px; overflow: hidden; margin-bottom: 16px;
}
.rights-card-head {
  background: #1b3f72; color: #fff;
  padding: 12px 16px; font-size: 0.8rem; font-weight: 700;
}
.rights-card-head.teal { background: #1d7874; }
.rights-row {
  display: flex; justify-content: space-between;
  padding: 10px 16px; border-bottom: 1px solid rgba(0,0,0,0.04);
  font-size: 0.83rem;
}
.rights-row:last-child { border-bottom: none; }
.rights-row .key { color: var(--muted); font-weight: 500; }
.rights-row .val { font-weight: 700; font-family: 'IBM Plex Mono', monospace; }

/* סיכום */
.summary-block {
  background: var(--card); border: 1px solid var(--border);
  border-radius: 12px; padding: 22px 24px; margin-bottom: 16px;
}
.summary-block h3 {
  font-size: 0.75rem; font-weight: 700;
  text-transform: uppercase; letter-spacing: 1px;
  color: var(--muted); margin-bottom: 12px;
}
.summary-block p { font-size: 0.9rem; line-height: 1.8; color: #333; }

/* topbar */
.topbar {
  background: var(--card); border-bottom: 1px solid var(--border);
  padding: 12px 24px; margin-bottom: 24px;
  display: flex; justify-content: space-between; align-items: center;
  border-radius: 12px;
}
.topbar-crumb { font-size: 0.8rem; color: var(--muted); font-weight: 500; }
.topbar-crumb span { color: var(--ink); font-weight: 700; }
.status-pill {
  display: inline-flex; align-items: center; gap: 6px;
  font-size: 0.75rem; font-weight: 600;
  padding: 4px 12px; border-radius: 99px;
  background: #e8f5f0; color: #1a7a4a;
}

/* העלאה */
.upload-zone {
  border: 2px dashed var(--border); border-radius: 14px;
  padding: 32px; text-align: center; margin-bottom: 24px;
  background: rgba(255,255,255,0.6);
}

/* download bar */
.dl-bar {
  display: flex; align-items: center; justify-content: space-between;
  background: #0f1923; border-radius: 12px; padding: 16px 22px; margin-top: 8px;
}
.dl-bar p { color: rgba(255,255,255,0.6); font-size: 0.82rem; margin:0; }
.dl-bar strong { color: #fff; display: block; font-size: 0.95rem; }

/* הסתר Streamlit elements */
#MainMenu, footer, header { visibility: hidden; }
.block-container { padding-top: 1rem !important; }

div[data-testid="stDataFrame"] { direction: rtl; }
</style>
""", unsafe_allow_html=True)


# ─── Sidebar ─────────────────────────────────────────────
with st.sidebar:
    st.markdown("""
    <div style="padding:8px 0 20px; border-bottom:1px solid rgba(255,255,255,0.08); margin-bottom:16px;">
      <div style="font-size:2rem;">🏗️</div>
      <div style="font-size:1.15rem; font-weight:800; color:#fff; line-height:1.2;">מנתח תב"ע AI</div>
      <div style="font-size:0.72rem; color:rgba(255,255,255,0.4); margin-top:4px;">ניתוח זכויות בנייה · תמהיל דירות</div>
    </div>
    """, unsafe_allow_html=True)

    api_key = st.text_input("🔑 API Key", value=os.getenv("ANTHROPIC_API_KEY",""), type="password")

    st.markdown("<div style='height:8px'></div>", unsafe_allow_html=True)
    project_name = st.text_input("שם הפרויקט", placeholder="מגדל הים – חיפה")
    city = st.text_input("עיר / ישוב", placeholder="חיפה")
    known_plot_area = st.number_input('שטח מגרש ידוע מ"ר', min_value=0.0, value=0.0, step=100.0)

    st.markdown("<div style='height:8px'></div>", unsafe_allow_html=True)
    mix_preference = st.radio("🎯 העדפת תמהיל", [
        "שמרני – דירות גדולות",
        "מאוזן – תמהיל מגוון",
        "אגרסיבי – מקסימום יח\"ד"
    ], index=1)

    st.markdown("<div style='height:16px'></div>", unsafe_allow_html=True)
    analyze_btn = st.button("🔍  נתח מסמך", type="primary", use_container_width=True)

    st.markdown("""
    <div style="margin-top:auto; padding-top:24px; font-size:0.68rem; color:rgba(255,255,255,0.2); line-height:1.7;">
    מופעל על ידי Claude Sonnet 4<br>
    PyMuPDF · Streamlit · openpyxl
    </div>
    """, unsafe_allow_html=True)


# ─── Topbar ──────────────────────────────────────────────
proj_display = project_name or "ללא שם"
city_display = city or ""
st.markdown(f"""
<div class="topbar">
  <div class="topbar-crumb">ניתוח תב"ע / <span>{proj_display}{' – ' + city_display if city_display else ''}</span></div>
  <div class="status-pill"><span style="width:6px;height:6px;border-radius:50%;background:#1a7a4a;display:inline-block;"></span> Claude מחובר</div>
</div>
""", unsafe_allow_html=True)


# ─── העלאת קובץ ──────────────────────────────────────────
st.markdown('<div class="upload-zone">', unsafe_allow_html=True)
uploaded_file = st.file_uploader("📄 העלאת מסמך תב\"ע – גרור PDF לכאן או לחץ להעלאה", type=["pdf"])
if uploaded_file:
    st.markdown(f"""
    <div style="display:inline-flex;align-items:center;gap:8px;background:#1d7874;color:#fff;
    border-radius:8px;padding:6px 14px;font-size:0.82rem;font-weight:600;margin-top:8px;">
    📄 {uploaded_file.name} · {uploaded_file.size/1024/1024:.1f} MB
    </div>""", unsafe_allow_html=True)
st.markdown('</div>', unsafe_allow_html=True)


# ─── ניתוח ───────────────────────────────────────────────
if analyze_btn:
    if not uploaded_file:
        st.error("❌ אנא העלה קובץ PDF")
        st.stop()
    if not api_key:
        st.error("❌ הזן API Key בסרגל הצד")
        st.stop()
    if not project_name:
        project_name = uploaded_file.name.replace(".pdf","")
    if not city:
        city = "לא צוינה"

    try:
        with st.spinner("📖 מחלץ טקסט מה-PDF..."):
            pdf_data = extract_text_from_pdf(uploaded_file.read())

        client = create_client(api_key)
        plot_area = known_plot_area if known_plot_area > 0 else None

        if pdf_data["is_long"] and len(pdf_data["chunks"]) > 1:
            progress = st.progress(0)
            def upd(cur, tot): progress.progress(cur/tot)
            with st.spinner("🤖 Claude מנתח..."):
                analysis = analyze_long_document(client, pdf_data["chunks"], project_name, city, mix_preference, plot_area, progress_callback=upd)
            progress.progress(1.0)
        else:
            with st.spinner("🤖 Claude מנתח את המסמך..."):
                analysis = analyze_single_chunk(client, pdf_data["full_text"], project_name, city, mix_preference, plot_area)

        analysis = validate_calculations(analysis)
        st.session_state["analysis"] = analysis
        st.session_state["project_name"] = project_name
        st.session_state["city"] = city
        st.success("✅ הניתוח הושלם!")

    except Exception as e:
        st.error(f"❌ שגיאה: {str(e)}")
        st.stop()


# ─── תצוגת תוצאות ────────────────────────────────────────
if "analysis" in st.session_state:
    analysis = st.session_state["analysis"]
    project_name = st.session_state["project_name"]
    city = st.session_state["city"]

    if "error" in analysis:
        st.error(f"שגיאת AI: {analysis['error']}")
        st.stop()

    financial = analysis.get("financial_estimate", {})
    rights    = analysis.get("zoning_rights", {})
    units     = financial.get("סה_כ_יחידות_מוצעות", "—")
    main_used = analysis.get("_total_apt_area_calculated", 0)
    main_allowed = rights.get("שטח_עיקרי_מותר_מ2", 0)
    pct = round(main_used / main_allowed * 100) if main_allowed else financial.get("אחוז_ניצול_זכויות","—")
    income = financial.get("הכנסה_משוערת_מיליון_ש_ח","—")
    building_pct = rights.get("אחוזי_בנייה_מותרים","—")

    # ── KPI ROW ──
    c1,c2,c3,c4 = st.columns(4)
    with c1:
        st.markdown(f"""<div class="kpi-card" style="--accent:#1d7874">
          <div class="kpi-label">יחידות דיור</div>
          <div class="kpi-value">{units}</div>
          <div class="kpi-sub">מתמהיל {mix_preference.split('–')[0].strip()}</div>
        </div>""", unsafe_allow_html=True)
    with c2:
        st.markdown(f"""<div class="kpi-card" style="--accent:#1b3f72">
          <div class="kpi-label">שטח עיקרי מ"ר</div>
          <div class="kpi-value">{int(main_used):,}</div>
          <div class="kpi-sub">מתוך {int(main_allowed):,} מותר</div>
          <div class="progress-bar"><div class="progress-fill" style="width:{pct if isinstance(pct,int) else 0}%"></div></div>
        </div>""", unsafe_allow_html=True)
    with c3:
        st.markdown(f"""<div class="kpi-card" style="--accent:#e8a020">
          <div class="kpi-label">ניצול זכויות</div>
          <div class="kpi-value">{pct}%</div>
          <div class="kpi-sub">אחוזי בנייה {building_pct}%</div>
        </div>""", unsafe_allow_html=True)
    with c4:
        st.markdown(f"""<div class="kpi-card" style="--accent:#1a7a4a">
          <div class="kpi-label">הכנסה משוערת</div>
          <div class="kpi-value">₪{income}M</div>
          <div class="kpi-sub">הנחת מחיר 2.2M ליח"ד</div>
        </div>""", unsafe_allow_html=True)

    st.markdown("<div style='height:8px'></div>", unsafe_allow_html=True)

    # ── TABS ──
    tab1, tab2, tab3, tab4 = st.tabs(["🏠 תמהיל דירות", "📐 זכויות בנייה", "⚠️ סיכוני תכנון", "📝 סיכום מנהלים"])

    # ── טאב 1: תמהיל ──
    with tab1:
        mix = analysis.get("apartment_mix", [])
        if mix:
            df = pd.DataFrame(mix)

            # גרפים
            g1, g2 = st.columns(2)
            with g1:
                if "כמות_יחידות" in df.columns and "סוג_דירה" in df.columns:
                    fig_pie = px.pie(df, names="סוג_דירה", values="כמות_יחידות",
                        title="התפלגות יחידות לפי סוג",
                        color_discrete_sequence=['#3b82f6','#10b981','#f59e0b','#8b5cf6','#ef4444'],
                        hole=0.62)
                    fig_pie.update_layout(font_family="Heebo", title_font_size=12,
                        paper_bgcolor='white', plot_bgcolor='white',
                        legend=dict(orientation="h", yanchor="bottom", y=-0.3))
                    st.plotly_chart(fig_pie, use_container_width=True)
            with g2:
                area_col = next((c for c in df.columns if 'שטח_כולל' in c or 'שטח כולל' in c), None)
                type_col  = next((c for c in df.columns if 'סוג' in c), None)
                if area_col and type_col:
                    fig_bar = px.bar(df, x=type_col, y=area_col,
                        title='שטח מנוצל לפי סוג דירה (מ"ר)',
                        color=type_col,
                        color_discrete_sequence=['#3b82f6','#10b981','#f59e0b','#8b5cf6','#ef4444'])
                    fig_bar.update_layout(font_family="Heebo", title_font_size=12,
                        showlegend=False, paper_bgcolor='white', plot_bgcolor='white',
                        xaxis=dict(showgrid=False), yaxis=dict(gridcolor='rgba(0,0,0,0.05)'))
                    st.plotly_chart(fig_bar, use_container_width=True)

            # טבלה
            BADGES = ["badge-1","badge-2","badge-3","badge-4","badge-5"]
            cols_map = {c: c.replace("_"," ") for c in df.columns}
            rows_html = ""
            total_units = 0; total_area = 0
            for i, row in df.iterrows():
                sug = row.get("סוג_דירה") or row.get("סוג דירה","")
                shat = row.get("שטח_ממוצע") or row.get("שטח ממוצע","—")
                kamut = row.get("כמות_יחידות") or row.get("כמות יחידות",0)
                achuz = row.get("אחוז_מהתמהיל") or row.get("אחוז מהתמהיל","—")
                shat_k = row.get("שטח_כולל_מ2") or row.get("שטח כולל",0)
                hera = row.get("הערה","")
                total_units += int(kamut) if str(kamut).isdigit() else 0
                total_area  += int(shat_k) if str(shat_k).replace(",","").isdigit() else 0
                rows_html += f"""<tr>
                  <td><span class="badge {BADGES[i%5]}">{sug}</span></td>
                  <td>{shat}</td><td>{kamut}</td><td>{achuz}%</td>
                  <td>{int(shat_k):,}</td><td>{hera}</td></tr>"""
            rows_html += f"""<tr class="total-row">
              <td>סה"כ</td><td>—</td><td>{total_units}</td><td>100%</td><td>{total_area:,}</td><td>—</td></tr>"""

            st.markdown(f"""
            <div class="tbl-wrap">
              <div class="tbl-header">🏠 תמהיל דירות מוצע</div>
              <table><thead><tr>
                <th>סוג דירה</th><th>שטח ממוצע מ"ר</th><th>כמות יחידות</th>
                <th>% מהתמהיל</th><th>שטח כולל מ"ר</th><th>הערה</th>
              </tr></thead><tbody>{rows_html}</tbody></table>
            </div>""", unsafe_allow_html=True)

            # אזהרות
            for w in analysis.get("validation_warnings", []):
                sev = w.get("חומרה","נמוך")
                bg = "#fff8e1" if sev != "גבוה" else "#fdf2f2"
                border = "#e8a020" if sev != "גבוה" else "#c0392b"
                st.markdown(f"""<div style="background:{bg};border:1px solid {border};border-right:4px solid {border};
                border-radius:10px;padding:12px 16px;font-size:0.82rem;margin-bottom:12px;">
                ⚠️ <strong>בדיקת עקביות:</strong> {w.get('פירוט','')}</div>""", unsafe_allow_html=True)

    # ── טאב 2: זכויות ──
    with tab2:
        r = rights
        z_rows = [
            ("שטח מגרש", f"{r.get('שטח_מגרש_מ2','—')} מ\"ר"),
            ("אחוזי בנייה מותרים", f"{r.get('אחוזי_בנייה_מותרים','—')}%"),
            ("שטח כולל מותר", f"{r.get('שטח_כולל_מותר_מ2','—')} מ\"ר"),
            ("שטח עיקרי מותר", f"{r.get('שטח_עיקרי_מותר_מ2','—')} מ\"ר"),
            ("שטח שירות מותר", f"{r.get('שטח_שירות_מותר_מ2','—')} מ\"ר"),
            ("מספר קומות מקסימלי", f"{r.get('מספר_קומות_מקסימלי','—')} קומות"),
            ("גובה מקסימלי", f"{r.get('גובה_מקסימלי_מ','—')} מ'"),
            ("אחוז כיסוי מגרש", f"{r.get('אחוז_כיסוי_מגרש','—')}%"),
        ]
        z_html = "".join(f'<div class="rights-row"><span class="key">{k}</span><span class="val">{v}</span></div>' for k,v in z_rows)

        parking = analysis.get("parking",{})
        p_rows = [
            ("חניות למגורים", f"{parking.get('חניות_חובה_למגורים','—')} ליח\"ד"),
            ("סה\"כ חניות נדרשות", f"{parking.get('סה_כ_חניות_נדרשות','—')}"),
            ("קו בניין קדמי", f"{r.get('קו_בניין_קדמי_מ','—')} מ'"),
            ("קו בניין צדדי", f"{r.get('קו_בניין_צדדי_מ','—')} מ'"),
        ]
        p_html = "".join(f'<div class="rights-row"><span class="key">{k}</span><span class="val">{v}</span></div>' for k,v in p_rows)

        rc1, rc2 = st.columns(2)
        with rc1:
            st.markdown(f"""<div class="rights-card">
              <div class="rights-card-head">📐 זכויות בנייה</div>{z_html}</div>""", unsafe_allow_html=True)
        with rc2:
            st.markdown(f"""<div class="rights-card">
              <div class="rights-card-head teal">🅿️ חנייה וקווי בניין</div>{p_html}</div>""", unsafe_allow_html=True)

        # בונוסים
        bonuses = analysis.get("bonuses_and_exceptions",[])
        if bonuses:
            b_rows = ""
            BADGES = ["badge-2","badge-1","badge-3","badge-4"]
            for i,b in enumerate(bonuses):
                sug = b.get("סוג_הקלה") or b.get("סוג","")
                teur = b.get("תיאור","")
                tos = b.get("תוספת_שטח_מ2","") or b.get("תוספת","")
                tna = b.get("תנאים","")
                b_rows += f"""<tr><td><span class="badge {BADGES[i%4]}">{sug}</span></td>
                  <td>{teur}</td><td>{tos}</td><td>{tna}</td></tr>"""
            st.markdown(f"""<div class="tbl-wrap">
              <div class="tbl-header">🎁 הקלות ובונוסים מזוהים</div>
              <table><thead><tr><th>סוג</th><th>תיאור</th><th>תוספת זכויות מ"ר</th><th>תנאים</th></tr></thead>
              <tbody>{b_rows}</tbody></table></div>""", unsafe_allow_html=True)

        # שימושים
        permitted = analysis.get("permitted_uses",{})
        if permitted:
            u_rows = ""
            for use, val in permitted.items():
                icon = "✅ מותר" if val is True else "❌ אסור" if val is False else f"⚠️ {val}"
                u_rows += f"<tr><td>{use}</td><td>{icon}</td><td></td></tr>"
            st.markdown(f"""<div class="tbl-wrap">
              <div class="tbl-header">🏗️ שימושים מותרים</div>
              <table><thead><tr><th>שימוש</th><th>סטטוס</th><th>הערה</th></tr></thead>
              <tbody>{u_rows}</tbody></table></div>""", unsafe_allow_html=True)

    # ── טאב 3: סיכונים ──
    with tab3:
        risks = analysis.get("risk_flags",[])
        levels = {"גבוה":"high","בינוני":"medium","נמוך":"low"}
        icons  = {"גבוה":"🔴","בינוני":"🟡","נמוך":"🟢"}
        risks_html = ""
        for risk in risks:
            level = risk.get("סוג_סיכון","נמוך")
            css = levels.get(level,"low")
            icon = icons.get(level,"🟢")
            title = risk.get("תיאור","")
            detail = risk.get("פירוט") or risk.get("תיאור_מפורט","")
            section = risk.get("סעיף_רלוונטי","—")
            risks_html += f"""<div class="risk-card {css}">
              <span class="risk-icon">{icon}</span>
              <div class="risk-body"><h4>{title}</h4>
              <p>{detail} <br><small>סעיף: {section}</small></p></div>
              <span class="risk-tag">{level}</span>
            </div>"""
        st.markdown(risks_html, unsafe_allow_html=True)

        quality = analysis.get("data_quality",{})
        if quality:
            st.markdown(f"""<div class="tbl-wrap" style="padding:16px 20px;">
              <div class="tbl-header">🎯 רמת ביטחון בנתונים</div>
              <div style="padding:14px 16px;font-size:0.83rem;">
              רמה כללית: <strong>{quality.get('רמת_ביטחון','—')}</strong><br>
              <span style="color:var(--muted)">{quality.get('הערות_לאנליסט','')}</span>
              </div></div>""", unsafe_allow_html=True)

    # ── טאב 4: סיכום ──
    with tab4:
        summary = analysis.get("summary",{})
        st.markdown(f"""<div class="summary-block">
          <h3>📋 סיכום מנהלים</h3>
          <p>{summary.get('סיכום_מנהלים','')}</p>
        </div>""", unsafe_allow_html=True)

        st.markdown(f"""<div class="summary-block" style="background:linear-gradient(135deg,#f0faf5,#e8f8f3);border-color:#a8dfc8;">
          <h3>✅ המלצה עיקרית</h3>
          <p style="font-weight:600;color:#1a7a4a;">{summary.get('המלצה_עיקרית','')}</p>
        </div>""", unsafe_allow_html=True)

        col_s, col_a = st.columns(2)
        strengths = "".join(f"<li>{s}</li>" for s in summary.get("נקודות_חוזק",[]))
        challenges = "".join(f"<li>{c}</li>" for c in summary.get("אתגרים_עיקריים",[]))
        with col_s:
            st.markdown(f"""<div class="summary-block">
              <h3>💪 נקודות חוזק</h3>
              <ul style="list-style:none;padding:0">{strengths}</ul>
            </div>""", unsafe_allow_html=True)
        with col_a:
            st.markdown(f"""<div class="summary-block">
              <h3>🚧 אתגרים</h3>
              <ul style="list-style:none;padding:0">{challenges}</ul>
            </div>""", unsafe_allow_html=True)

        steps = summary.get("שלבים_מומלצים",[])
        colors = ["#1b3f72","#1b3f72","#1d7874","#1d7874","#e8a020"]
        steps_html = "".join(f"""<div style="display:flex;gap:12px;align-items:center;margin-bottom:10px;">
          <span style="background:{colors[min(i,4)]};color:#fff;border-radius:50%;width:24px;height:24px;
          display:flex;align-items:center;justify-content:center;font-size:0.75rem;font-weight:700;flex-shrink:0;">{i+1}</span>
          <span style="font-size:0.88rem;">{step}</span></div>""" for i,step in enumerate(steps))
        st.markdown(f"""<div class="summary-block">
          <h3>🗺️ שלבים מומלצים</h3>{steps_html}</div>""", unsafe_allow_html=True)

    # ── הורדת Excel ──
    st.markdown("""<div class="dl-bar">
      <div><strong>⬇️ הורד דוח Excel מלא</strong>
      <p>4 גיליונות: סיכום מנהלים · תמהיל דירות · זכויות בנייה · סיכוני תכנון</p></div>
    </div>""", unsafe_allow_html=True)

    with st.spinner("מכין Excel..."):
        excel_bytes = export_to_excel(analysis, project_name, city)
    st.download_button(
        label="⬇️ הורד דוח Excel",
        data=excel_bytes,
        file_name=f"דוח_תבע_{project_name}.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        use_container_width=True
    )
