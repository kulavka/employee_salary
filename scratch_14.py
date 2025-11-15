import re
from pathlib import Path
import pandas as pd

# --- Имя/фамилия ---

NAME_WORD = r"[A-Za-zĀ-ž'’\-]+"
NAME_PART = rf"{NAME_WORD}(?: {NAME_WORD}){{0,3}}"

TAIL_STOPWORDS = (
    "Social security number",
    "Personal ID",
    "Employee number",
    "Tax number",
    "Address",
    "IBAN",
)

PAGE_MARKER_RE = re.compile(r"(?m)^\s*account\s+number\s*:?\b", flags=re.IGNORECASE)

PATTERN_WITH_COMMA = re.compile(
    rf"\b(?P<surname>{NAME_PART})\s*,\s*(?P<name>{NAME_PART})\b"
    rf"(?![^\r\n]*?(?:{'|'.join(map(re.escape, TAIL_STOPWORDS))}))"
)
PATTERN_NO_COMMA = re.compile(
    rf"\b(?P<surname>{NAME_PART})\s+(?P<name>{NAME_PART})\b"
    rf"(?![^\r\n]*?(?:{'|'.join(map(re.escape, TAIL_STOPWORDS))}))"
)

# --- Табличные строки/метки ---

# Ставка "Normaali työ" / "Hourly salary"
RATE_LABEL_RE = re.compile(r"(?i)normaali\s+ty[oö]|hourly\s+salary")

# Overtime 50% (уже было)
OT50_LABEL_RE = re.compile(r"(?i)ylity[oö]\s*50\s*%|overtime(?:[, ]+\w+)?\s*50\s*%")

# Новые: 100 / 150 / 200 / 300
OT100_LABEL_RE = re.compile(r"(?i)ylity[oö]\s*100\s*%|overtime(?:[, ]+\w+)?\s*100\s*%")
OT150_LABEL_RE = re.compile(r"(?i)ylity[oö]\s*,?\s*vrk\s*150\s*%|overtime(?:[, ]+\w+)?\s*150\s*%")
OT200_LABEL_RE = re.compile(r"(?i)ylity[oö]\s*,?\s*vrk\s*200\s*%|overtime(?:[, ]+\w+)?\s*200\s*%")
# допускаем странный вариант "ylityöl+300%"
OT300_LABEL_RE = re.compile(r"(?i)sunnuntai\s+ylity[oö](?:l\+)?\s*300\s*%")

# Evening shift hours
EVENING_LABEL_RE = re.compile(r"(?i)evening\s+work\s+allowance")

# Числа (с поддержкой тысячных разделителей и знака)
NUMBER_RE = re.compile(r"(?<!\d)(-?\d{1,3}(?:[ .]\d{3})*(?:[.,]\d+)|-?\d+[.,]\d+|-?\d+)(?!\d)")

# --- Утилиты ---

def strip_tails(s: str) -> str:
    for tail in TAIL_STOPWORDS:
        idx = s.find(tail)
        if idx != -1:
            s = s[:idx]
    return s.strip()

def normalize_page_text(text: str) -> str:
    text = (text or "").replace("\u00A0", " ")
    text = re.sub(r",\s*[\r\n]+\s*", ", ", text)
    for tail in TAIL_STOPWORDS:
        text = text.replace(" " + tail, "\n" + tail)
    text = re.sub(r"[ \t]{2,}", " ", text)
    return text

def handle_page_text(text: str) -> tuple[str, str] | None:
    text = normalize_page_text(text)

    for ln in (ln.strip() for ln in text.splitlines() if ln.strip()):
        m = PATTERN_WITH_COMMA.search(ln)
        if m:
            return strip_tails(m["name"]), strip_tails(m["surname"])
    for ln in (ln.strip() for ln in text.splitlines() if ln.strip()):
        m = PATTERN_NO_COMMA.search(ln)
        if m:
            return strip_tails(m["name"]), strip_tails(m["surname"])
    m = PATTERN_WITH_COMMA.search(text) or PATTERN_NO_COMMA.search(text)
    if m:
        return strip_tails(m["name"]), strip_tails(m["surname"])
    return None

def _to_float(num_str: str) -> float | None:
    s = num_str.strip().replace("\u00A0", " ")
    if "," in s and "." in s:
        s = s.replace(" ", "").replace(".", "").replace(",", ".")
    else:
        s = s.replace(" ", "").replace(",", ".")
    try:
        return float(s)
    except ValueError:
        return None

def _numbers_after(label_pat: re.Pattern, text: str, need: int = 3) -> list[str]:
    m = label_pat.search(text)
    if not m:
        return []
    tail = text[m.end():]
    return [g[0] if isinstance(g, tuple) else g for g in NUMBER_RE.findall(tail)][:need]

def extract_rate(text: str) -> float | None:
    """Unit price: второе число после метки ставки."""
    text = normalize_page_text(text)
    nums = _numbers_after(RATE_LABEL_RE, text, need=3)
    return _to_float(nums[1]) if len(nums) >= 2 else None

def extract_qty(text: str, label_pat: re.Pattern) -> float | None:
    """Quantity: первое число после заданной метки."""
    text = normalize_page_text(text)
    nums = _numbers_after(label_pat, text, need=1)
    return _to_float(nums[0]) if nums else None

def page_has_marker(text: str) -> bool:
    return bool(PAGE_MARKER_RE.search(text or ""))

# --- Основная функция ---

def extract_names_from_pdf(pdf_path: str) -> pd.DataFrame:
    rows: list[dict] = []
    used_pypdf2 = False

    def process_text(text: str):
        res = handle_page_text(text)
        if not res:
            return
        name, surname = res
        row = {
            "Surname": surname,
            "Name": name,
            "Rate per hour": extract_rate(text),
            "Overtime 50%":  extract_qty(text, OT50_LABEL_RE),
            "Overtime 100%": extract_qty(text, OT100_LABEL_RE),
            "Overtime 150%": extract_qty(text, OT150_LABEL_RE),
            "Overtime 200%": extract_qty(text, OT200_LABEL_RE),
            "Overtime 300%": extract_qty(text, OT300_LABEL_RE),
            "Evening shift/ hours": extract_qty(text, EVENING_LABEL_RE),
        }
        rows.append(row)

    # PyPDF2
    try:
        import PyPDF2
        with open(pdf_path, "rb") as f:
            reader = PyPDF2.PdfReader(f)
            for page in reader.pages:
                text = page.extract_text() or ""
                if page_has_marker(text):
                    process_text(text)
        used_pypdf2 = True
    except Exception:
        pass

    # pdfminer fallback
    if not used_pypdf2 and not rows:
        try:
            from pdfminer.high_level import extract_text_to_fp
            from pdfminer.layout import LAParams
            from pdfminer.pdfpage import PDFPage
            import io
            with open(pdf_path, "rb") as fh:
                for i, _ in enumerate(PDFPage.get_pages(fh)):
                    buf = io.StringIO()
                    with open(pdf_path, "rb") as f2:
                        extract_text_to_fp(f2, buf, laparams=LAParams(), page_numbers=[i])
                    text = buf.getvalue()
                    if page_has_marker(text):
                        process_text(text)
        except Exception:
            pass

    cols = ["Surname", "Name", "Rate per hour",
            "Overtime 50%", "Overtime 100%", "Overtime 150%",
            "Overtime 200%", "Overtime 300%",
            "Evening shift/ hours"]
    return pd.DataFrame(rows, columns=cols)

# --- Пример запуска ---
if __name__ == "__main__":
    pdf_file = r"C:\Users\nikit\AppData\Roaming\JetBrains\PyCharm2023.3\scratches\w21-w22 payslips copy.pdf"
    df = extract_names_from_pdf(pdf_file)
    print(df.to_string(index=False))
    df.to_excel(r"names.xlsx", index=False)
