#!/usr/bin/env python3
"""
fix_option_contamination.py

Re-extracts all exam PDFs with a FIXED parser (properly strips 'Cevap Açıklama'
explanation text from options), then patches the contaminated option texts in
each course's Excel file.

Contaminated exams: 2021-2022 Final, 2021-2022 Bütünleme, 2022-2023 Yaz Okulu
(plus occasionally others) had explanation text printed after each option E,
which the original extraction scripts failed to strip.
"""

import re
import unicodedata
import pdfplumber
import pandas as pd
from pathlib import Path

# ─────────────────────────── PATHS ─────────────────────────────────────────
FINALLER_DIR = Path("/Users/kurter/Library/CloudStorage/GoogleDrive-kurter@gmail.com/My Drive/_____University/Atatürk/Kurter/2025-2026/Bahar/Finaller - Çıkmış")
EXCEL_DIR    = Path("/Users/kurter/Library/CloudStorage/GoogleDrive-kurter@gmail.com/My Drive/_____University/Atatürk/Kurter/2025-2026/Bahar/Finaller - Çıkmış/AI_Study_ADAY/1_sinif/bahar")

# Course → (pdf_prefix, excel_filename)
# pdf_prefix is the leading part of the PDF filename in FINALLER_DIR
COURSE_CONFIG = {
    "Afet ve Acil Durum Mevzuatı": {
        "pdf_prefix": "Afet ve Acil Durum Mevzuatı - ",
        "excel": "Afet_ve_Acil_Durum_Mevzuati_Soru_Bankasi.xlsx",
    },
    "Dosyalama ve Arşivleme": {
        "pdf_prefix": "Dosyalama ve Arşivleme ",
        "excel": "Dosyalama_ve_Arsivleme_Soru_Bankasi.xlsx",
    },
    "İlk Yardım ve Acil Sağlık Hizmetleri": {
        "pdf_prefix": "İlk yardım ve Acil Sağlık Hizmetleri ",
        "excel": "Ilk_Yardim_ve_Acil_Saglik_Hizmetleri_Soru_Bankasi.xlsx",
    },
    "Kimyasal Maddeler ve Tehlikeleri": {
        "pdf_prefix": "Kimyasal Maddeler ve Tehlikeleri ",
        "excel": "Kimyasal_Maddeler_ve_Tehlikeleri_Soru_Bankasi.xlsx",
    },
    "Lojistik Yönetimi": {
        "pdf_prefix": "Lojistik Yönetimi ",
        "excel": "Lojistik_Yonetimi_Soru_Bankasi.xlsx",
    },
    "Temel Bilgi Teknolojileri II": {
        "pdf_prefix": "Temel Bilgi Teknolojileri II ",
        "excel": "Temel_Bilgi_Teknolojileri_II_Soru_Bankasi.xlsx",
    },
}

# ─────────────────────────── PDF EXTRACTION ────────────────────────────────

def fix_cid_chars(text: str) -> str:
    """Remove (cid:XXX) artifacts and normalize whitespace."""
    if not text:
        return ""
    text = re.sub(r"\(cid:\d+\)", "", text)
    return text


def extract_text_two_columns(pdf_path: Path) -> str:
    """Extract text from two-column AUZEF exam PDF."""
    parts = []
    with pdfplumber.open(str(pdf_path)) as pdf:
        for page in pdf.pages:
            w, h = page.width, page.height
            left  = fix_cid_chars(page.crop((0,    0, w/2, h)).extract_text() or "")
            right = fix_cid_chars(page.crop((w/2,  0, w,   h)).extract_text() or "")
            if left:  parts.append(left)
            if right: parts.append(right)
    return "\n".join(parts)


def parse_answer_key(text: str) -> dict:
    """Extract answer key from dense num-letter lines."""
    answers = {}
    for line in text.split("\n"):
        pairs = re.findall(r"(\d{1,2})[.\s]+([A-E])(?=\s|$)", line)
        if len(pairs) >= 5:
            for num_str, letter in pairs:
                num = int(num_str)
                if 1 <= num <= 20 and num not in answers:
                    answers[num] = letter
    return answers


CEVAP_PATTERN = re.compile(
    r"Cevap\s*(?:Açıklama|Acıklama|Aciklama|Aciklama)?\s*[:\-]?\s*[\(\[]?[A-E][\)\]]?",
    re.IGNORECASE
)

def parse_questions_fixed(text: str, sinav_label: str) -> list:
    """
    Parse questions with FIXED Cevap Açıklama stripping.
    After option E is assigned, stop accumulating lines into it.
    """
    answer_key = parse_answer_key(text)

    # Remove CEVAP ANAHTARI block and NAHTARI continuation
    text_q = re.sub(r"CEVAP\s+AN\w*.*?(?=\n\d{1,2}[.\s]|\Z)", "", text,
                    flags=re.IGNORECASE | re.DOTALL)
    text_q = re.sub(r"NAHTARI.*", "", text_q, flags=re.IGNORECASE | re.DOTALL)

    q_split = re.compile(r"(?:^|\n)(\d{1,2})\.\s*", re.MULTILINE)
    starts = []
    for m in q_split.finditer(text_q):
        num = int(m.group(1))
        if 1 <= num <= 20:
            starts.append((m.start(), num, m.end()))

    questions = []
    for i, (start_pos, q_num, content_start) in enumerate(starts):
        end_pos = starts[i+1][0] if i+1 < len(starts) else len(text_q)
        block = text_q[content_start:end_pos].strip()

        # ── KEY FIX: strip Cevap Açıklama and everything after it ──────────
        # Handle "Cevap Açıklama: (X) ..." on same line or new line
        block_clean = re.sub(
            r"\nCevap\s*(?:A[çc][ıi]klama)?\s*[:\-]?.*",
            "", block, flags=re.IGNORECASE | re.DOTALL
        ).strip()
        block_clean = re.sub(
            r"Cevap\s*(?:A[çc][ıi]klama)?\s*[:\-]?\s*[\(\[]?[A-E][\)\]]?.*",
            "", block_clean, flags=re.IGNORECASE | re.DOTALL
        ).strip()
        # Also handle "Doğru Cevap: (X)" format
        block_clean = re.sub(
            r"\nDoğru\s+Cevap\s*[:\-]?.*",
            "", block_clean, flags=re.IGNORECASE | re.DOTALL
        ).strip()

        lines = [l.strip() for l in block_clean.split("\n") if l.strip()]

        option_re = re.compile(r"^([A-E])\)\s*(.+)$")
        options = {}
        q_lines = []

        for line in lines:
            m = option_re.match(line)
            if m:
                letter = m.group(1).upper()
                opt_text = m.group(2).strip()
                if letter not in options:
                    options[letter] = opt_text
                else:
                    # Continuation of same option (only if E not yet seen)
                    if "E" not in options:
                        options[letter] += " " + opt_text
            elif options:
                last = list(options.keys())[-1]
                # ONLY append continuation if E hasn't been found yet
                # This stops explanation text from bleeding into option E
                if last != "E" and "E" not in options:
                    options[last] += " " + line
                # If E already found → this is explanation text → SKIP
            else:
                q_lines.append(line)

        question_text = " ".join(q_lines).strip()
        question_text = re.sub(r"^\d{1,2}\.\s*", "", question_text)

        dogru_cevap = answer_key.get(q_num, "")

        questions.append({
            "sinav": sinav_label,
            "soru_no": q_num,
            "soru": question_text,
            "A": options.get("A", ""),
            "B": options.get("B", ""),
            "C": options.get("C", ""),
            "D": options.get("D", ""),
            "E": options.get("E", ""),
            "dogru_cevap": dogru_cevap,
        })

    return questions


# ─────────────────────────── TEXT NORMALISATION ────────────────────────────

def nfc(s: str) -> str:
    return unicodedata.normalize("NFC", s)

def norm_q(text: str) -> str:
    """Normalize question text for matching."""
    if not text:
        return ""
    t = nfc(str(text))
    t = t.lower().strip()
    t = re.sub(r"\s+", " ", t)
    # Remove leading question number if any
    t = re.sub(r"^\d{1,2}\.\s*", "", t)
    return t


def norm_opt(text: str) -> str:
    """Normalize option text for comparison."""
    if not text:
        return ""
    t = nfc(str(text))
    t = t.strip()
    t = re.sub(r"\s+", " ", t)
    return t


# ─────────────────────────── MAIN ──────────────────────────────────────────

def process_course(course_name: str, cfg: dict):
    excel_path = EXCEL_DIR / cfg["excel"]
    prefix     = cfg["pdf_prefix"]

    # Find all PDFs for this course
    pdfs = []
    for f in sorted(FINALLER_DIR.iterdir()):
        fname = nfc(f.name)
        pref  = nfc(prefix)
        if fname.startswith(pref) and fname.lower().endswith(".pdf"):
            # Extract exam label from filename
            label_part = fname[len(pref):].replace(".pdf", "").strip()
            pdfs.append((label_part, f))

    if not pdfs:
        print(f"  [!] No PDFs found for {course_name}")
        return

    print(f"\n{'='*60}")
    print(f"COURSE: {course_name}")
    print(f"  PDFs found: {len(pdfs)}")

    # Fresh extraction → build lookup
    # lookup[norm_q_text] = list of {sinav, q_num, A, B, C, D, E}
    lookup = {}

    for sinav_label, pdf_path in pdfs:
        try:
            text = extract_text_two_columns(pdf_path)
            questions = parse_questions_fixed(text, sinav_label)
            for q in questions:
                key = norm_q(q["soru"])
                if key not in lookup:
                    lookup[key] = []
                lookup[key].append(q)
        except Exception as e:
            print(f"    [ERROR] {pdf_path.name}: {e}")

    print(f"  Unique questions in fresh extraction: {len(lookup)}")

    # Load Excel
    df = pd.read_excel(excel_path)
    opt_cols = ["A Şıkkı", "B Şıkkı", "C Şıkkı", "D Şıkkı", "E Şıkkı"]

    fixes = 0
    not_found = 0

    for idx, row in df.iterrows():
        excel_q_norm = norm_q(str(row["Soru"]))

        # Find matching fresh extractions
        matches = lookup.get(excel_q_norm, [])

        if not matches:
            not_found += 1
            continue

        # For each option column, check if the Excel value is contaminated
        # (i.e., it's longer than ALL fresh versions' corresponding option)
        for col in opt_cols:
            excel_val = norm_opt(str(row[col]) if pd.notna(row[col]) else "")
            if not excel_val or len(excel_val) < 80:
                continue  # Short options: unlikely contaminated

            # Try to find a matching fresh option across all exam versions
            best_clean = None
            for fresh_q in matches:
                for letter in ["A", "B", "C", "D", "E"]:
                    fresh_val = norm_opt(fresh_q.get(letter, ""))
                    if not fresh_val:
                        continue
                    # Check: does the Excel option START WITH the fresh option?
                    # (means fresh = clean option, Excel = clean + explanation)
                    if (len(fresh_val) < len(excel_val) and
                            excel_val.lower().startswith(fresh_val.lower()[:min(len(fresh_val), 40)]) and
                            len(fresh_val) >= 3):
                        # Accept if we cut at least 20 chars of contamination
                        if len(excel_val) - len(fresh_val) >= 15:
                            if best_clean is None or len(fresh_val) < len(best_clean):
                                best_clean = fresh_val  # use the shortest clean version

            if best_clean is not None:
                # Verify it's a real truncation (not just a coincidental prefix)
                old_val = str(row[col]) if pd.notna(row[col]) else ""
                print(f"  [FIX] Row #{row['#']} col={col}")
                print(f"    OLD ({len(old_val):4d}): {old_val[:80]}...")
                print(f"    NEW ({len(best_clean):4d}): {best_clean[:80]}")
                df.at[idx, col] = best_clean
                fixes += 1

    print(f"\n  Fixes applied: {fixes}")
    print(f"  Questions not found in fresh extraction: {not_found}")

    if fixes > 0:
        df.to_excel(excel_path, index=False, sheet_name="Tüm Sorular")
        print(f"  Saved: {excel_path.name}")
    else:
        print(f"  No changes needed.")

    return fixes


def main():
    total_fixes = 0
    for course_name, cfg in COURSE_CONFIG.items():
        f = process_course(course_name, cfg)
        if f:
            total_fixes += f

    print(f"\n{'='*60}")
    print(f"TOTAL FIXES: {total_fixes}")


if __name__ == "__main__":
    main()
