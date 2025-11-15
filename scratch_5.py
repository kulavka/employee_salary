# -*- coding: utf-8 -*-
import re
import unicodedata
from pathlib import Path
import pandas as pd
import pdfplumber

PDF_PATH = r"C:\Users\nikit\Downloads\w21 inv 28_04_2025_25_05_2025_subcontractor_followup_2025001863 copy (1).pdf"
OUT_XLSX = "tyontekijat_table_all_strings.xlsx"

COLS = ["Name", "Aika", "Norm", "50%", "100%", "Iltalisä", "Yövuoro", "Kaikki yhteensä"]

def strip_accents(s: str) -> str:
    return "".join(ch for ch in unicodedata.normalize("NFD", s) if unicodedata.category(ch) != "Mn")

def low_noacc(s: str) -> str:
    return strip_accents(s).lower()

def rm_parens(s: str) -> str:
    # убрать полностью участки в скобках вместе со скобками
    return re.sub(r"\([^)]*\)", "", s).strip()

def join_words(words):
    # words: список слов с полями {'text','x0','x1','top','bottom'}
    return " ".join(w["text"] for w in words)

def cluster_lines(words, y_tol=3):
    """Группируем слова в строки по Y с допуском."""
    lines = []
    for w in sorted(words, key=lambda x: x["top"]):
        if not lines:
            lines.append([w]); continue
        last_line = lines[-1]
        if abs(w["top"] - last_line[-1]["top"]) <= y_tol:
            last_line.append(w)
        else:
            lines.append([w])
    # внутри строки отсортируем по X
    for ln in lines:
        ln.sort(key=lambda x: x["x0"])
    return lines

def find_header_line(lines):
    """Ищем индекс строки-шапки: содержит все ключи шапки."""
    required = ["tyontekijat", "aika", "norm", "yovuoro", "kaikki", "yhteensa"]
    for idx, ln in enumerate(lines):
        t = low_noacc(join_words(ln))
        if all(k in t for k in required):
            return idx
    return None

def header_column_xs(header_words):
    """
    Возвращаем x-координаты начала колонок по словам шапки.
    Для 'Kaikki yhteensä' берём x по слову 'Kaikki'.
    """
    labels_map = {
        "tyontekijat": "Name",
        "aika": "Aika",
        "norm": "Norm",
        "50%": "50%",
        "100%": "100%",
        "iltalisa": "Iltalisä",
        "yovuoro": "Yövuoro",
        "kaikki": "Kaikki yhteensä",  # берём x по 'Kaikki'
    }

    xs = {}
    for w in header_words:
        txt_raw = w["text"]
        txt = low_noacc(txt_raw)
        txt = txt.replace("ö", "o").replace("ä", "a")  # на всякий
        if txt in labels_map:
            xs[labels_map[txt]] = w["x0"]
        # случаи типа "Iltalisä" -> "iltalisa", "Yövuoro" -> "yovuoro"
        elif "iltalisa".startswith(txt):
            xs["Iltalisä"] = xs.get("Iltalisä", w["x0"])
        elif "yovuoro".startswith(txt):
            xs["Yövuoro"] = xs.get("Yövuoro", w["x0"])
        elif txt.endswith("%") and txt in ("50%", "100%"):
            xs[txt] = w["x0"]

    # ожидаемый порядок колонок
    order = ["Name", "Aika", "Norm", "50%", "100%", "Iltalisä", "Yövuoro", "Kaikki yhteensä"]
    # если что-то не нашли, попробуем дооценить позициями ближайших слов
    missing = [c for c in order if c not in xs]
    if missing:
        # возьмём ровно столько первых слов, сколько колонок, по возрастанию x0
        sorted_by_x = sorted(header_words, key=lambda x: x["x0"])
        approx = [w["x0"] for w in sorted_by_x[:len(order)]]
        xs = {col: approx[i] for i, col in enumerate(order)}
    return [xs[c] for c in order]

def build_bins(col_xs):
    """По x-координатам колонок строим границы ячеек (полубисы)."""
    xs_sorted = sorted(col_xs)
    # середины между соседями
    mids = [(xs_sorted[i] + xs_sorted[i+1]) / 2 for i in range(len(xs_sorted)-1)]
    # биновые границы: (-inf, mid1], (mid1, mid2], ..., (last_mid, +inf)
    return [-float("inf")] + mids + [float("inf")]

def put_words_into_cells(line_words, col_labels, col_bins):
    """
    Разбрасываем слова строки по колонкам согласно x-координате центра слова.
    Склеиваем тексты внутри ячейки через пробел.
    """
    cells = {c: [] for c in col_labels}
    for w in line_words:
        cx = (w["x0"] + w["x1"]) / 2
        # найдём индекс бина
        bi = None
        for i in range(len(col_bins)-1):
            if col_bins[i] <= cx <= col_bins[i+1]:
                bi = i; break
        if bi is None:
            continue
        col = col_labels[bi]
        cells[col].append(w["text"])
    # склеим
    return {c: " ".join(v).strip() for c, v in cells.items()}

def is_total_line(line_words):
    t = low_noacc(join_words(line_words))
    return "kaikki yhteensa" in t  # строка-итог — не включаем в таблицу

def parse_pdf_to_df(pdf_path: str) -> pd.DataFrame:
    rows = []
    with pdfplumber.open(pdf_path) as pdf:
        for page_idx, page in enumerate(pdf.pages, start=1):
            words = page.extract_words(
                x_tolerance=2,
                y_tolerance=2,
                keep_blank_chars=False,
                use_text_flow=True
            )
            if not words:
                continue
            lines = cluster_lines(words, y_tol=3)

            hdr_i = find_header_line(lines)
            if hdr_i is None:
                continue  # на странице нет нужной таблицы

            header_words = lines[hdr_i]
            col_labels = ["Name", "Aika", "Norm", "50%", "100%", "Iltalisä", "Yövuoro", "Kaikki yhteensä"]
            col_xs = header_column_xs(header_words)
            col_bins = build_bins(col_xs)

            # строки под шапкой до итога или до конца страницы/следующей шапки
            for ln in lines[hdr_i+1:]:
                if is_total_line(ln):
                    break
                # если встретили новую шапку (редко), прерываем
                if find_header_line([ln]) is not None:
                    break

                cells = put_words_into_cells(ln, col_labels, col_bins)
                # если строка явно пустая/шум — пропустим
                if not cells["Name"] and not cells["Aika"]:
                    continue

                # убираем скобки целиком в каждой ячейке
                cleaned = {k: rm_parens(v) for k, v in cells.items()}
                rows.append(cleaned)

    return pd.DataFrame(rows, columns=COLS)

if __name__ == "__main__":
    df = parse_pdf_to_df(PDF_PATH)

    mask = df.apply(lambda r: not any("tekijä:" in str(v).lower() for v in r.values), axis=1)
    df = df[mask].reset_index(drop=True)

    print("Rows:", len(df))
    df.to_excel(OUT_XLSX, index=False)
    print(f"Saved: {Path(OUT_XLSX).resolve()}")
