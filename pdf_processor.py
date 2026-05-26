"""
pdf_processor.py
----------------
אחראי על חילוץ טקסט מ-PDF תב"ע בצורה נקייה,
כולל תמיכה בקבצים ארוכים באמצעות חלוקה לפרקים (Chunking).
"""

import fitz  # PyMuPDF
import re
from typing import List, Dict


# ─── קבועים ──────────────────────────────────────────────
MAX_CHARS_PER_CHUNK = 60_000   # ~15,000 טוקן בממוצע לצ'אנק


# ─── פונקציות עיבוד ──────────────────────────────────────

def clean_text(text: str) -> str:
    """
    מנקה את הטקסט שחולץ מה-PDF:
    - מסיר שורות ריקות מרובות
    - מסיר תווים מיותרים
    - מנרמל רווחים
    """
    # מסיר תווים לא-ASCII בעייתיים אך שומר על עברית
    text = re.sub(r'[\x00-\x08\x0b\x0c\x0e-\x1f\x7f]', '', text)
    # מנרמל שורות ריקות מרובות
    text = re.sub(r'\n{3,}', '\n\n', text)
    # מנרמל רווחים מרובים
    text = re.sub(r'[ \t]{2,}', ' ', text)
    return text.strip()


def extract_text_from_pdf(pdf_bytes: bytes) -> Dict:
    """
    חילוץ מלא של טקסט מ-PDF.
    מחזיר dict עם:
    - full_text: הטקסט המלא
    - pages: רשימת טקסטים לפי עמוד
    - page_count: מספר עמודים
    - is_long: האם המסמך ארוך מהמכסה
    - chunks: רשימת צ'אנקים אם המסמך ארוך
    """
    doc = fitz.open(stream=pdf_bytes, filetype="pdf")
    pages = []

    for page_num in range(len(doc)):
        page = doc[page_num]
        # חילוץ טקסט עם שמירת מבנה כמה שיותר
        text = page.get_text("text", flags=fitz.TEXT_PRESERVE_WHITESPACE)
        pages.append({
            "page_num": page_num + 1,
            "text": clean_text(text)
        })

    doc.close()

    full_text = "\n\n--- עמוד {page} ---\n\n".join(
        p["text"] for p in pages
    )
    full_text = clean_text(full_text)

    is_long = len(full_text) > MAX_CHARS_PER_CHUNK

    return {
        "full_text": full_text,
        "pages": pages,
        "page_count": len(pages),
        "char_count": len(full_text),
        "is_long": is_long,
        "chunks": split_to_chunks(pages) if is_long else [full_text]
    }


def split_to_chunks(pages: List[Dict]) -> List[str]:
    """
    מחלק את עמודי המסמך לצ'אנקים לפי גודל תווים.
    כל צ'אנק לא יחרוג מ-MAX_CHARS_PER_CHUNK.
    שומר על שלמות עמוד (לא חוצה עמוד באמצע).
    """
    chunks = []
    current_chunk = ""
    current_page_start = 1

    for page in pages:
        page_text = f"\n--- עמוד {page['page_num']} ---\n{page['text']}"

        if len(current_chunk) + len(page_text) > MAX_CHARS_PER_CHUNK:
            if current_chunk:
                chunks.append(current_chunk.strip())
            current_chunk = page_text
            current_page_start = page['page_num']
        else:
            current_chunk += page_text

    if current_chunk:
        chunks.append(current_chunk.strip())

    return chunks


def extract_section_hints(text: str) -> List[str]:
    """
    זיהוי מקטעים/פרקים רלוונטיים בתב"ע לפי מילות מפתח.
    עוזר להתמקד בחלקים הרלוונטיים לזכויות הבנייה.
    """
    keywords = [
        "תמהיל", "דירות", "יחידות דיור", "שטח עיקרי",
        "שטח שירות", "אחוזי בנייה", "מספר קומות", "גובה",
        "קווי בניין", "שטח מגרש", "סה\"כ זכויות", "הקלה",
        "בונוס", "תמ\"א", "חנייה", "שימושים מותרים",
        "מסחר", "משרדים", "ייעוד", "זכויות בנייה"
    ]

    hints = []
    lines = text.split('\n')

    for i, line in enumerate(lines):
        if any(kw in line for kw in keywords):
            # לוקח את השורה ועוד 2 שורות הקשר
            context_start = max(0, i - 1)
            context_end = min(len(lines), i + 3)
            hints.append(" | ".join(lines[context_start:context_end]).strip())

    return list(set(hints))[:30]  # מגביל ל-30 רמזים
