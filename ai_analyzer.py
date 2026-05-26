"""
ai_analyzer.py
--------------
לוגיקת ניתוח ה-AI באמצעות Claude.
אחראי על בניית הפרומפטים, שליחה ל-API,
קבלת תגובה בפורמט JSON, וניתוח מסמכים ארוכים.
"""

import json
import re
from typing import Dict, List, Optional
import anthropic


# ─── הגדרת ה-API ──────────────────────────────────────────
# !! כאן מזינים את ה-API Key של Anthropic !!
# אפשר גם להשתמש במשתנה סביבה: ANTHROPIC_API_KEY
# (מומלץ: להגדיר קובץ .env עם ANTHROPIC_API_KEY=sk-ant-...)
MODEL = "claude-sonnet-4-20250514"


# ─── פרומפטים ─────────────────────────────────────────────

SYSTEM_PROMPT = """אתה מומחה בתכנון עירוני ישראלי, בקי בתקנות תב"ע, תמ"א 38, ותוכניות מיתאר.
תפקידך לנתח מסמכי תכנון ולהפיק דוח זכויות בנייה מפורט ומדויק.
תמיד תחזיר תשובה בפורמט JSON בלבד, ללא טקסט נוסף לפני או אחרי ה-JSON.
אם מידע מסוים לא קיים במסמך - ציין null ולא תמציא נתונים."""

ANALYSIS_PROMPT_TEMPLATE = """נתח את מסמך התב"ע הבא והפק דוח JSON מלא.

=== פרמטרי הפרויקט ===
שם פרויקט: {project_name}
עיר: {city}
העדפת תמהיל: {mix_preference}
שטח מגרש ידוע (אם הוזן): {known_plot_area} מ"ר

=== תוכן המסמך ===
{document_text}

=== הנחיות ===
נתח את המסמך והחזר JSON עם המבנה הבא בדיוק:

{{
  "metadata": {{
    "תוכנית_מספר": "...",
    "שם_תוכנית": "...",
    "עיר_ישוב": "...",
    "ייעוד_קרקע": "...",
    "תאריך_אישור": "...",
    "סטטוס_תוכנית": "..."
  }},
  "zoning_rights": {{
    "שטח_מגרש_מ2": null,
    "אחוזי_בנייה_מותרים": null,
    "שטח_כולל_מותר_מ2": null,
    "שטח_עיקרי_מותר_מ2": null,
    "שטח_שירות_מותר_מ2": null,
    "מספר_קומות_מקסימלי": null,
    "גובה_מקסימלי_מ": null,
    "קו_בניין_קדמי_מ": null,
    "קו_בניין_צדדי_מ": null,
    "קו_בניין_אחורי_מ": null,
    "אחוז_כיסוי_מגרש": null
  }},
  "apartment_mix": [
    {{
      "סוג_דירה": "1 חדרים",
      "שטח_ממוצע_מ2": null,
      "כמות_יחידות": null,
      "אחוז_מהתמהיל": null,
      "שטח_כולל_מ2": null,
      "הערה": "..."
    }}
  ],
  "parking": {{
    "חניות_חובה_למגורים": null,
    "חניות_חובה_למסחר": null,
    "סה_כ_חניות_נדרשות": null,
    "הערות_חנייה": "..."
  }},
  "permitted_uses": {{
    "מגורים": true,
    "מסחר": null,
    "משרדים": null,
    "מלאכה": null,
    "ציבורי": null,
    "אחר": "..."
  }},
  "bonuses_and_exceptions": [
    {{
      "סוג": "תמ\"א 38 / הקלה / בונוס",
      "תיאור": "...",
      "תוספת_זכויות_מ2": null,
      "תנאים": "..."
    }}
  ],
  "risk_flags": [
    {{
      "סוג_סיכון": "גבוה/בינוני/נמוך",
      "תיאור": "...",
      "סעיף_רלוונטי": "..."
    }}
  ],
  "financial_estimate": {{
    "תמהיל_אופטימלי": "{mix_preference}",
    "סה_כ_יחידות_מוצעות": null,
    "שטח_עיקרי_מנוצל_מ2": null,
    "אחוז_ניצול_זכויות": null,
    "הכנסה_משוערת_מיליון_ש_ח": null,
    "הנחות_חישוב": "..."
  }},
  "summary": {{
    "סיכום_מנהלים": "...",
    "המלצה_עיקרית": "...",
    "נקודות_חוזק": ["...", "..."],
    "אתגרים_עיקריים": ["...", "..."],
    "שלבים_מומלצים": ["...", "..."]
  }},
  "data_quality": {{
    "רמת_ביטחון": "גבוה/בינוני/נמוך",
    "שדות_חסרים": ["..."],
    "הערות_לאנליסט": "..."
  }}
}}
"""

CHUNK_SUMMARY_PROMPT = """זהו חלק {chunk_num} מתוך {total_chunks} של מסמך תב"ע.
חלץ מחלק זה את כל המידע הרלוונטי לזכויות בנייה, תמהיל דירות, שימושים, וחנייה.
החזר JSON עם המפתחות הרלוונטיים בלבד שמצאת בחלק הזה.

=== תוכן ===
{chunk_text}
"""


# ─── פונקציות ────────────────────────────────────────────

def create_client(api_key: str) -> anthropic.Anthropic:
    """יצירת לקוח Anthropic עם ה-API Key."""
    return anthropic.Anthropic(api_key=api_key)


def analyze_single_chunk(
    client: anthropic.Anthropic,
    text: str,
    project_name: str,
    city: str,
    mix_preference: str,
    known_plot_area: Optional[float] = None
) -> Dict:
    """
    ניתוח מסמך קצר (פרק בודד).
    מחזיר dict מה-JSON שהחזיר Claude.
    """
    prompt = ANALYSIS_PROMPT_TEMPLATE.format(
        project_name=project_name,
        city=city,
        mix_preference=mix_preference,
        known_plot_area=known_plot_area if known_plot_area else "לא הוזן",
        document_text=text[:60000]  # חסימת בטיחות
    )

    response = client.messages.create(
        model=MODEL,
        max_tokens=4096,
        system=SYSTEM_PROMPT,
        messages=[{"role": "user", "content": prompt}]
    )

    raw = response.content[0].text
    return _safe_parse_json(raw)


def analyze_long_document(
    client: anthropic.Anthropic,
    chunks: List[str],
    project_name: str,
    city: str,
    mix_preference: str,
    known_plot_area: Optional[float] = None,
    progress_callback=None
) -> Dict:
    """
    ניתוח מסמך ארוך באמצעות:
    1. ניתוח כל צ'אנק בנפרד
    2. מיזוג התוצאות
    3. ניתוח סיכום סופי
    """
    chunk_summaries = []

    for i, chunk in enumerate(chunks):
        if progress_callback:
            progress_callback(i + 1, len(chunks))

        prompt = CHUNK_SUMMARY_PROMPT.format(
            chunk_num=i + 1,
            total_chunks=len(chunks),
            chunk_text=chunk
        )

        response = client.messages.create(
            model=MODEL,
            max_tokens=2048,
            system=SYSTEM_PROMPT,
            messages=[{"role": "user", "content": prompt}]
        )

        summary = _safe_parse_json(response.content[0].text)
        chunk_summaries.append(summary)

    # מיזוג כל הסיכומים לניתוח סופי
    merged_text = json.dumps(chunk_summaries, ensure_ascii=False, indent=2)

    final_prompt = ANALYSIS_PROMPT_TEMPLATE.format(
        project_name=project_name,
        city=city,
        mix_preference=mix_preference,
        known_plot_area=known_plot_area if known_plot_area else "לא הוזן",
        document_text=f"להלן סיכומי הפרקים שנותחו:\n{merged_text}"
    )

    final_response = client.messages.create(
        model=MODEL,
        max_tokens=4096,
        system=SYSTEM_PROMPT,
        messages=[{"role": "user", "content": final_prompt}]
    )

    return _safe_parse_json(final_response.content[0].text)


def validate_calculations(analysis: Dict) -> Dict:
    """
    בדיקה מתמטית של עקביות הנתונים:
    - סה"כ שטח יח"ד vs. שטח עיקרי מותר
    - מספר חניות vs. דרישה
    - אחוז ניצול זכויות
    """
    warnings = []
    rights = analysis.get("zoning_rights", {})
    mix = analysis.get("apartment_mix", [])

    # חישוב שטח כולל של כל הדירות
    total_apt_area = sum(
        (apt.get("שטח_ממוצע_מ2") or 0) * (apt.get("כמות_יחידות") or 0)
        for apt in mix
    )

    allowed_main = rights.get("שטח_עיקרי_מותר_מ2")
    if allowed_main and total_apt_area > 0:
        if total_apt_area > allowed_main * 1.05:
            warnings.append({
                "סוג": "חריגת שטח",
                "פירוט": f"שטח כולל של התמהיל ({total_apt_area:,.0f} מ\"ר) חורג משטח עיקרי מותר ({allowed_main:,.0f} מ\"ר)",
                "חומרה": "גבוה"
            })
        utilization = (total_apt_area / allowed_main) * 100
        if "financial_estimate" in analysis:
            analysis["financial_estimate"]["אחוז_ניצול_זכויות"] = round(utilization, 1)

    # בדיקת חנייה
    parking = analysis.get("parking", {})
    total_units = sum(apt.get("כמות_יחידות") or 0 for apt in mix)
    req_parking = parking.get("חניות_חובה_למגורים")
    if req_parking and total_units:
        expected = total_units * req_parking
        actual = parking.get("סה_כ_חניות_נדרשות")
        if actual and abs(actual - expected) > 2:
            warnings.append({
                "סוג": "אי-התאמה חנייה",
                "פירוט": f"חניות מחושבות: {expected}, מצוין במסמך: {actual}",
                "חומרה": "בינוני"
            })

    analysis["validation_warnings"] = warnings
    analysis["_total_apt_area_calculated"] = total_apt_area
    return analysis


def _safe_parse_json(raw: str) -> Dict:
    """
    ניתוח JSON בטוח עם fallback.
    מנקה קודי markdown אם קיימים.
    """
    # מסיר ```json ... ``` אם קיים
    cleaned = re.sub(r'^```(?:json)?\s*', '', raw.strip(), flags=re.MULTILINE)
    cleaned = re.sub(r'\s*```$', '', cleaned.strip(), flags=re.MULTILINE)

    try:
        return json.loads(cleaned)
    except json.JSONDecodeError:
        # ניסיון לחלץ JSON מתוך הטקסט
        match = re.search(r'\{.*\}', cleaned, re.DOTALL)
        if match:
            try:
                return json.loads(match.group())
            except Exception:
                pass
        return {"error": "לא הצלחתי לנתח את תגובת ה-AI", "raw_response": raw[:2000]}
