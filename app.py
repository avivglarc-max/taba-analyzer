"""
app.py
------
אפליקציית Streamlit לניתוח תב"ע והפקת דוח תמהיל דירות.
מחבר בין: ממשק משתמש → עיבוד PDF → Claude AI → Excel

להרצה:
    streamlit run app.py

הגדרת API Key (בחר אחת):
    1. משתנה סביבה: export ANTHROPIC_API_KEY=sk-ant-...
    2. קובץ .env: ANTHROPIC_API_KEY=sk-ant-...
    3. ישירות בקוד: API_KEY = "sk-ant-..." (לא מומלץ!)
"""

import os
import streamlit as st
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
from dotenv import load_dotenv

# !! טעינת API Key !!
# יצור קובץ .env בתיקיית הפרויקט עם השורה: ANTHROPIC_API_KEY=sk-ant-...
load_dotenv()

from pdf_processor import extract_text_from_pdf
from ai_analyzer import (
    create_client,
    analyze_single_chunk,
    analyze_long_document,
    validate_calculations
)
from excel_exporter import export_to_excel


# ─── הגדרות עמוד ────────────────────────────────────────
st.set_page_config(
    page_title="מנתח תב\"ע AI",
    page_icon="🏗️",
    layout="wide",
    initial_sidebar_state="expanded"
)

# ─── CSS מותאם ──────────────────────────────────────────
st.markdown("""
<style>
    @import url('https://fonts.googleapis.com/css2?family=Heebo:wght@300;400;600;700;900&display=swap');

    html, body, [class*="css"] { font-family: 'Heebo', sans-serif; direction: rtl; }

    .main-header {
        background: linear-gradient(135deg, #1A3A5C 0%, #2B7A78 100%);
        padding: 2rem; border-radius: 16px;
        color: white; text-align: center;
        margin-bottom: 2rem;
        box-shadow: 0 8px 32px rgba(26,58,92,0.3);
    }
    .main-header h1 { font-size: 2.4rem; font-weight: 900; margin: 0; }
    .main-header p  { font-size: 1.05rem; opacity: 0.85; margin-top: 0.5rem; }

    .metric-card {
        background: white; border-radius: 12px; padding: 1.2rem;
        border: 1px solid #e0e7ef;
        box-shadow: 0 2px 12px rgba(0,0,0,0.06);
        text-align: center;
    }
    .metric-card .value { font-size: 2rem; font-weight: 700; color: #1A3A5C; }
    .metric-card .label { font-size: 0.85rem; color: #6c757d; margin-top: 0.2rem; }

    .risk-high   { background: #f8d7da; border-right: 4px solid #dc3545; }
    .risk-medium { background: #fff3cd; border-right: 4px solid #ffc107; }
    .risk-low    { background: #d1e7dd; border-right: 4px solid #198754; }

    .stButton > button {
        background: linear-gradient(135deg, #1A3A5C, #2B7A78);
        color: white; border: none; border-radius: 8px;
        padding: 0.6rem 2rem; font-weight: 600;
        font-family: 'Heebo', sans-serif;
        transition: all 0.2s;
    }
    .stButton > button:hover { transform: translateY(-2px); box-shadow: 0 4px 12px rgba(26,58,92,0.4); }

    .info-box {
        background: #EBF4FF; border-right: 4px solid #1A3A5C;
        padding: 1rem; border-radius: 8px; margin: 0.5rem 0;
    }

    div[data-testid="stDataFrame"] { direction: rtl; }
    .stTabs [data-baseweb="tab"] { font-family: 'Heebo', sans-serif; font-weight: 600; }
</style>
""", unsafe_allow_html=True)


# ─── Header ─────────────────────────────────────────────
st.markdown("""
<div class="main-header">
    <h1>🏗️ מנתח תב"ע AI</h1>
    <p>ניתוח אוטומטי של מסמכי תכנון ובנייה · הפקת דוח תמהיל דירות · ניתוח זכויות בנייה</p>
</div>
""", unsafe_allow_html=True)


# ─── Sidebar – הגדרות ────────────────────────────────────
with st.sidebar:
    st.markdown("### ⚙️ הגדרות")

    # !! כאן מזינים את ה-API Key !!
    api_key = st.text_input(
        "🔑 Anthropic API Key",
        value=os.getenv("ANTHROPIC_API_KEY", ""),
        type="password",
        help="הזן את ה-API Key שלך מ-console.anthropic.com"
    )

    st.markdown("---")
    st.markdown("### 📋 פרטי הפרויקט")

    project_name = st.text_input("שם הפרויקט", placeholder="פרויקט המגדל – תל אביב")
    city = st.text_input("עיר / ישוב", placeholder="תל אביב")

    mix_preference = st.selectbox(
        "העדפת תמהיל",
        ["שמרני (דירות גדולות, פחות יחידות)",
         "מאוזן (תמהיל מגוון)",
         "אגרסיבי (מקסימום יחידות, דירות קטנות)"]
    )

    known_plot_area = st.number_input(
        "שטח מגרש ידוע מ\"ר (אופציונלי)",
        min_value=0.0, value=0.0, step=100.0,
        help="אם ידוע לך שטח המגרש, הזן אותו לבדיקת עקביות"
    )

    st.markdown("---")
    st.markdown("### 🔧 הגדרות מתקדמות")

    force_chunking = st.checkbox(
        "כפה חלוקה לפרקים",
        help="גם אם המסמך קצר, חלק אותו לפרקים לניתוח מדויק יותר"
    )

    st.markdown("---")
    st.markdown("""
    <div style="font-size:0.75rem; color:#6c757d; text-align:center;">
    מופעל על ידי Claude Sonnet 4<br>
    נבנה עם Streamlit + PyMuPDF
    </div>
    """, unsafe_allow_html=True)


# ─── אזור העלאת קובץ ────────────────────────────────────
col1, col2 = st.columns([2, 1])

with col1:
    st.markdown("#### 📄 העלאת מסמך תב\"ע")
    uploaded_file = st.file_uploader(
        "גרור PDF לכאן או לחץ להעלאה",
        type=["pdf"],
        help="תומך ב-PDF עם טקסט (לא תמונות סרוקות)"
    )

with col2:
    st.markdown("#### 📊 מה תקבל?")
    st.markdown("""
    <div class="info-box">
    ✅ תמהיל דירות אופטימלי<br>
    ✅ ניתוח זכויות בנייה<br>
    ✅ זיהוי הקלות ובונוסים<br>
    ✅ ניתוח סיכוני תכנון<br>
    ✅ בדיקת עקביות מתמטית<br>
    ✅ דוח Excel להורדה
    </div>
    """, unsafe_allow_html=True)


# ─── כפתור ניתוח ────────────────────────────────────────
if uploaded_file:
    st.markdown(f"**📂 קובץ שהועלה:** {uploaded_file.name} ({uploaded_file.size / 1024:.1f} KB)")

analyze_btn = st.button("🔍 נתח מסמך", type="primary", use_container_width=True)


# ─── לוגיקת ניתוח ───────────────────────────────────────
if analyze_btn:
    # ולידציה
    if not uploaded_file:
        st.error("❌ אנא העלה קובץ PDF.")
        st.stop()
    if not api_key:
        st.error("❌ הזן Anthropic API Key בסרגל הצד.")
        st.stop()
    if not project_name:
        st.warning("⚠️ לא הוזן שם פרויקט – ייעשה שימוש בשם הקובץ.")
        project_name = uploaded_file.name.replace(".pdf", "")
    if not city:
        city = "לא צוינה"

    try:
        # שלב 1: חילוץ PDF
        with st.spinner("📖 מחלץ טקסט מה-PDF..."):
            pdf_data = extract_text_from_pdf(uploaded_file.read())

        col_a, col_b, col_c = st.columns(3)
        col_a.metric("📄 עמודים", pdf_data["page_count"])
        col_b.metric("🔤 תווים", f"{pdf_data['char_count']:,}")
        col_c.metric("📦 מצב", "ארוך – Chunking" if pdf_data["is_long"] else "רגיל")

        if pdf_data["char_count"] < 500:
            st.warning("⚠️ הטקסט שחולץ קצר מאוד. ייתכן שה-PDF מבוסס תמונה (סרוק) ולא טקסט.")

        # שלב 2: ניתוח AI
        client = create_client(api_key)
        use_chunking = pdf_data["is_long"] or force_chunking
        plot_area = known_plot_area if known_plot_area > 0 else None

        if use_chunking and len(pdf_data["chunks"]) > 1:
            st.info(f"📦 המסמך חולק ל-{len(pdf_data['chunks'])} פרקים לניתוח.")
            progress_bar = st.progress(0)
            status_text = st.empty()

            def update_progress(current, total):
                progress_bar.progress(current / total)
                status_text.text(f"🔄 מנתח פרק {current}/{total}...")

            with st.spinner("🤖 Claude מנתח את המסמך..."):
                analysis = analyze_long_document(
                    client,
                    pdf_data["chunks"],
                    project_name, city, mix_preference,
                    plot_area,
                    progress_callback=update_progress
                )
            progress_bar.progress(1.0)
            status_text.text("✅ ניתוח הושלם!")
        else:
            with st.spinner("🤖 Claude מנתח את המסמך..."):
                analysis = analyze_single_chunk(
                    client,
                    pdf_data["full_text"],
                    project_name, city, mix_preference, plot_area
                )

        # שלב 3: בדיקת עקביות
        analysis = validate_calculations(analysis)

        # שמירה ב-session state
        st.session_state["analysis"] = analysis
        st.session_state["project_name"] = project_name
        st.session_state["city"] = city
        st.success("✅ הניתוח הושלם בהצלחה!")

    except Exception as e:
        st.error(f"❌ שגיאה בניתוח: {str(e)}")
        if "api_key" in str(e).lower() or "authentication" in str(e).lower():
            st.error("נראה שה-API Key שגוי. בדוק שוב.")
        st.stop()


# ─── תצוגת תוצאות ────────────────────────────────────────
if "analysis" in st.session_state:
    analysis = st.session_state["analysis"]
    project_name = st.session_state["project_name"]
    city = st.session_state["city"]

    if "error" in analysis:
        st.error(f"שגיאת AI: {analysis['error']}")
        with st.expander("תגובה גולמית"):
            st.text(analysis.get("raw_response", ""))
        st.stop()

    st.markdown("---")
    st.markdown("## 📊 תוצאות הניתוח")

    # KPIs ראשיים
    financial = analysis.get("financial_estimate", {})
    rights = analysis.get("zoning_rights", {})

    kpi1, kpi2, kpi3, kpi4 = st.columns(4)
    with kpi1:
        st.markdown(f"""
        <div class="metric-card">
            <div class="value">{financial.get('סה_כ_יחידות_מוצעות', '—')}</div>
            <div class="label">יחידות דיור</div>
        </div>""", unsafe_allow_html=True)
    with kpi2:
        st.markdown(f"""
        <div class="metric-card">
            <div class="value">{rights.get('אחוזי_בנייה_מותרים', '—')}%</div>
            <div class="label">אחוזי בנייה</div>
        </div>""", unsafe_allow_html=True)
    with kpi3:
        st.markdown(f"""
        <div class="metric-card">
            <div class="value">{financial.get('אחוז_ניצול_זכויות', '—')}%</div>
            <div class="label">ניצול זכויות</div>
        </div>""", unsafe_allow_html=True)
    with kpi4:
        st.markdown(f"""
        <div class="metric-card">
            <div class="value">₪{financial.get('הכנסה_משוערת_מיליון_ש_ח', '—')}M</div>
            <div class="label">הכנסה משוערת</div>
        </div>""", unsafe_allow_html=True)

    st.markdown("<br>", unsafe_allow_html=True)

    # טאבים
    tab1, tab2, tab3, tab4, tab5 = st.tabs([
        "🏠 תמהיל דירות",
        "📐 זכויות בנייה",
        "⚠️ סיכוני תכנון",
        "📝 סיכום מנהלים",
        "🔍 JSON גולמי"
    ])

    # ── טאב 1: תמהיל דירות ──
    with tab1:
        mix = analysis.get("apartment_mix", [])
        if mix:
            df_mix = pd.DataFrame(mix)
            df_mix.columns = [c.replace("_", " ") for c in df_mix.columns]

            col_t, col_c = st.columns([3, 2])
            with col_t:
                st.dataframe(df_mix, use_container_width=True, hide_index=True)

            with col_c:
                if "כמות יחידות" in df_mix.columns and "סוג דירה" in df_mix.columns:
                    fig = px.pie(
                        df_mix,
                        names="סוג דירה",
                        values="כמות יחידות",
                        title="התפלגות יחידות לפי סוג",
                        color_discrete_sequence=px.colors.qualitative.Bold
                    )
                    fig.update_layout(font_family="Arial", title_font_size=14)
                    st.plotly_chart(fig, use_container_width=True)

            # בדיקת אזהרות
            warnings = analysis.get("validation_warnings", [])
            for w in warnings:
                severity = w.get("חומרה", "נמוך")
                css = "risk-high" if severity == "גבוה" else "risk-medium"
                st.markdown(f'<div class="{css}" style="padding:0.8rem; border-radius:8px; margin:0.4rem 0;">⚠️ {w.get("פירוט","")}</div>', unsafe_allow_html=True)
        else:
            st.info("לא זוהה תמהיל דירות מפורש במסמך.")

    # ── טאב 2: זכויות בנייה ──
    with tab2:
        col_r, col_p = st.columns(2)

        with col_r:
            st.markdown("#### 📐 זכויות בנייה")
            rights_data = {
                "שטח מגרש מ\"ר": rights.get("שטח_מגרש_מ2"),
                "אחוזי בנייה": f"{rights.get('אחוזי_בנייה_מותרים', '—')}%",
                "שטח כולל מותר מ\"ר": rights.get("שטח_כולל_מותר_מ2"),
                "שטח עיקרי מותר מ\"ר": rights.get("שטח_עיקרי_מותר_מ2"),
                "שטח שירות מותר מ\"ר": rights.get("שטח_שירות_מותר_מ2"),
                "מספר קומות מקסימלי": rights.get("מספר_קומות_מקסימלי"),
                "גובה מקסימלי מ\"": rights.get("גובה_מקסימלי_מ"),
                "קו בניין קדמי מ\"": rights.get("קו_בניין_קדמי_מ"),
                "קו בניין צדדי מ\"": rights.get("קו_בניין_צדדי_מ"),
                "אחוז כיסוי מגרש": f"{rights.get('אחוז_כיסוי_מגרש', '—')}%",
            }
            df_rights = pd.DataFrame(rights_data.items(), columns=["פרמטר", "ערך"])
            st.dataframe(df_rights, use_container_width=True, hide_index=True)

        with col_p:
            st.markdown("#### 🅿️ חנייה")
            parking = analysis.get("parking", {})
            parking_data = {
                "חניות למגורים (ליח\"ד)": parking.get("חניות_חובה_למגורים"),
                "חניות למסחר": parking.get("חניות_חובה_למסחר"),
                "סה\"כ חניות נדרשות": parking.get("סה_כ_חניות_נדרשות"),
                "הערות": parking.get("הערות_חנייה"),
            }
            df_parking = pd.DataFrame(parking_data.items(), columns=["פרמטר", "ערך"])
            st.dataframe(df_parking, use_container_width=True, hide_index=True)

            st.markdown("#### 🏗️ שימושים מותרים")
            permitted = analysis.get("permitted_uses", {})
            for use, val in permitted.items():
                icon = "✅" if val is True else "❌" if val is False else "ℹ️"
                st.markdown(f"{icon} **{use}**: {val if isinstance(val, str) else ''}")

        # הקלות ובונוסים
        bonuses = analysis.get("bonuses_and_exceptions", [])
        if bonuses:
            st.markdown("#### 🎁 הקלות ובונוסים")
            df_bonus = pd.DataFrame(bonuses)
            if not df_bonus.empty:
                df_bonus.columns = [c.replace("_", " ") for c in df_bonus.columns]
                st.dataframe(df_bonus, use_container_width=True, hide_index=True)

        # גרף ניצול זכויות
        main_area = rights.get("שטח_עיקרי_מותר_מ2")
        total_calc = analysis.get("_total_apt_area_calculated", 0)
        if main_area and total_calc:
            fig_bar = go.Figure()
            fig_bar.add_trace(go.Bar(
                x=["שטח מותר", "שטח מתוכנן"],
                y=[main_area, total_calc],
                marker_color=["#1A3A5C", "#2B7A78"],
                text=[f"{main_area:,.0f} מ\"ר", f"{total_calc:,.0f} מ\"ר"],
                textposition="auto"
            ))
            fig_bar.update_layout(
                title="ניצול שטח עיקרי", font_family="Arial",
                yaxis_title="מ\"ר", showlegend=False
            )
            st.plotly_chart(fig_bar, use_container_width=True)

    # ── טאב 3: סיכוני תכנון ──
    with tab3:
        risks = analysis.get("risk_flags", [])
        if risks:
            risk_colors = {"גבוה": "risk-high", "בינוני": "risk-medium", "נמוך": "risk-low"}
            for risk in risks:
                level = risk.get("סוג_סיכון", "נמוך")
                css = risk_colors.get(level, "risk-low")
                st.markdown(f"""
                <div class="{css}" style="padding:1rem; border-radius:8px; margin:0.5rem 0;">
                    <strong>{'🔴' if level=='גבוה' else '🟡' if level=='בינוני' else '🟢'} {level}</strong>
                    – {risk.get('תיאור', '')}<br>
                    <small style="opacity:0.7">סעיף: {risk.get('סעיף_רלוונטי', '—')}</small>
                </div>""", unsafe_allow_html=True)

        quality = analysis.get("data_quality", {})
        if quality:
            st.markdown("---")
            st.markdown(f"**רמת ביטחון בנתונים:** {quality.get('רמת_ביטחון', '—')}")
            missing = quality.get("שדות_חסרים", [])
            if missing:
                st.warning(f"**שדות חסרים:** {', '.join(str(m) for m in missing)}")
            if quality.get("הערות_לאנליסט"):
                st.info(f"📌 {quality['הערות_לאנליסט']}")

    # ── טאב 4: סיכום ──
    with tab4:
        summary = analysis.get("summary", {})
        if summary.get("סיכום_מנהלים"):
            st.markdown("### 📋 סיכום מנהלים")
            st.markdown(f'<div class="info-box">{summary["סיכום_מנהלים"]}</div>',
                        unsafe_allow_html=True)

        if summary.get("המלצה_עיקרית"):
            st.markdown("### ✅ המלצה עיקרית")
            st.success(summary["המלצה_עיקרית"])

        col_s, col_a = st.columns(2)
        with col_s:
            st.markdown("#### 💪 נקודות חוזק")
            for s in summary.get("נקודות_חוזק", []):
                st.markdown(f"✔ {s}")
        with col_a:
            st.markdown("#### 🚧 אתגרים")
            for c in summary.get("אתגרים_עיקריים", []):
                st.markdown(f"◆ {c}")

        steps = summary.get("שלבים_מומלצים", [])
        if steps:
            st.markdown("#### 🗺️ שלבים מומלצים")
            for i, step in enumerate(steps, 1):
                st.markdown(f"**{i}.** {step}")

    # ── טאב 5: JSON גולמי ──
    with tab5:
        import json
        st.code(json.dumps(analysis, ensure_ascii=False, indent=2), language="json")

    # ─── כפתור הורדת Excel ──────────────────────────────────
    st.markdown("---")
    st.markdown("### 📥 הורדת הדוח")

    col_dl1, col_dl2 = st.columns([1, 3])
    with col_dl1:
        with st.spinner("מכין קובץ Excel..."):
            excel_bytes = export_to_excel(analysis, project_name, city)

        st.download_button(
            label="⬇️ הורד דוח Excel",
            data=excel_bytes,
            file_name=f"דוח_תבע_{project_name}_{city}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True
        )
    with col_dl2:
        st.markdown("""
        <div class="info-box">
        הדוח כולל 4 גיליונות: סיכום מנהלים · תמהיל דירות · זכויות בנייה · סיכוני תכנון
        </div>""", unsafe_allow_html=True)
