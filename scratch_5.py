# -*- coding: utf-8 -*-
import re
import unicodedata
from pathlib import Path
import pandas as pd
import pdfplumber

# Пути к обоим PDF
PDF_PATH_W21 = r"C:\Users\nikit\Downloads\w21 inv 28_04_2025_25_05_2025_subcontractor_followup_2025001863 copy (1).pdf"
PDF_PATH_W22 = r"C:\Users\nikit\Downloads\w22 inv 06_05_2025_01_06_2025_subcontractor_followup_2025001952 copy.pdf"

OUT_XLSX = "tyontekijat_weeks.xlsx"

REQ_FIRST = ["Työntekijät", "Aika", "Norm"]          # обязательные в начале
REQ_LAST  = ["Kaikki yhteensä"]                      # обязательная последняя

def strip_accents(s: str) -> str:
    return "".join(ch for ch in unicodedata.normalize("NFD", s) if unicodedata.category(ch) != "Mn")

def low_noacc(s: str) -> str:
    return strip_accents(s).lower()

def rm_parens(s: str) -> str:
    return re.sub(r"\([^)]*\)", "", s).strip()

def join_words(words):
    return " ".join(w["text"] for w in words)

def cluster_lines(words, y_tol=3):
    lines = []
    for w in sorted(words, key=lambda x: x["top"]):
        if not lines:
            lines.append([w]); continue
        last = lines[-1]
        if abs(w["top"] - last[-1]["top"]) <= y_tol:
            last.append(w)
        else:
            lines.append([w])
    for ln in lines:
        ln.sort(key=lambda x: x["x0"])
    return lines

def is_header_line(line_words):
    t = low_noacc(join_words(line_words))
    # минимум: Työntekijät, Aika, Norm, Kaikki, Yhteensä
    need = ["tyontekijat", "aika", "norm", "kaikki", "yhteensa"]
    return all(k in t for k in need)

def normalize_header(line_words):
    """
    Возвращаем список (label, x0) в порядке X для всех столбцов из шапки,
    где 'Kaikki yhteensä' склеиваем как одно имя.
    Остальные названия берём как есть (сырые), чтобы поддерживать любые новые колонки.
    """
    items = []
    i = 0
    while i < len(line_words):
        w = line_words[i]
        txt = w["text"].strip()
        t = low_noacc(txt)

        # склейка "Kaikki yhteensä"
        if t.startswith("kaikki"):
            label = "Kaikki yhteensä"
            x0 = w["x0"]
            # если следующее слово похоже на yhteensä — пропустим его
            if i + 1 < len(line_words):
                t2 = low_noacc(line_words[i+1]["text"].strip())
                if "yhteensa" in t2 or "yhteensä" in strip_accents(t2):
                    i += 1
            items.append((label, x0))
            i += 1
            continue

        # иначе — берём текст как есть (это и есть динамическая колонка)
        items.append((txt, w["x0"]))
        i += 1

    # нормализация имён + дубликаты
    seen = {}
    normed = []
    for label, x in items:
        label = label.strip()
        if not label:
            continue
        # стандартизируем основные
        if low_noacc(label) in ("tyontekajat", "tyontekijat"):
            label = "Työntekijät"
        elif low_noacc(label) == "aika":
            label = "Aika"
        elif low_noacc(label).startswith("norm"):
            label = "Norm"
        # дубликаты -> Label (2)
        key = (label.lower(),)
        if key in seen:
            seen[key] += 1
            label = f"{label} ({seen[key]})"
        else:
            seen[key] = 1
        normed.append((label, x))

    # сортировка по X
    normed.sort(key=lambda kv: kv[1])

    # проверка обязательных
    labels_only = [l for l,_ in normed]
    if not all(req in labels_only for req in REQ_FIRST + REQ_LAST):
        return []  # шапка не подходит

    # "Kaikki yhteensä" должна быть последней
    if normed and normed[-1][0] != "Kaikki yhteensä":
        for j,(lab,xx) in enumerate(normed):
            if lab == "Kaikki yhteensä":
                kept = (lab, xx)
                del normed[j]
                normed.append(kept)
                break

    return normed

def build_bins(x_list):
    xs_sorted = list(x_list)
    mids = [(xs_sorted[i] + xs_sorted[i+1]) / 2 for i in range(len(xs_sorted)-1)]
    return [-float("inf")] + mids + [float("inf")]

def assign_cells(line_words, columns, bins):
    cells = {c: [] for c in columns}
    for w in line_words:
        cx = (w["x0"] + w["x1"]) / 2
        bi = None
        for i in range(len(bins)-1):
            if bins[i] <= cx <= bins[i+1]:
                bi = i
                break
        if bi is None:
            continue
        cells[columns[bi]].append(w["text"])
    # склеим и уберём скобки
    return {c: rm_parens(" ".join(v)).strip() for c, v in cells.items()}

def is_total_line(line_words):
    t = low_noacc(join_words(line_words))
    return ("kaikki yhteensa" in t) and ("tyontekijat" not in t)

def parse_pdf_any_columns(pdf_path: str) -> pd.DataFrame:
    all_rows = []
    dynamic_order_global = []  # порядок появления нестандартных колонок

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
            headers = [i for i, ln in enumerate(lines) if is_header_line(ln)]
            if not headers:
                continue

            for hi in headers:
                header_items = normalize_header(lines[hi])
                if not header_items:
                    continue

                cols = [lab for lab,_ in header_items]
                xs   = [x   for _,x in header_items]

                bins = build_bins(xs)

                # глобальный порядок динамических колонок
                for lab in cols:
                    if lab in REQ_FIRST or lab in REQ_LAST:
                        continue
                    if lab not in dynamic_order_global:
                        dynamic_order_global.append(lab)

                # строки под шапкой
                stop_at = len(lines)
                for hj in headers:
                    if hj > hi:
                        stop_at = min(stop_at, hj)

                for li in range(hi+1, stop_at):
                    ln = lines[li]
                    if is_total_line(ln):
                        break
                    cells = assign_cells(ln, cols, bins)
                    # фильтруем Tekijä:
                    if any("tekijä:" in low_noacc(v) for v in cells.values() if v):
                        continue
                    # пустые строки выкидываем
                    if not cells.get("Työntekijät") and not cells.get("Aika"):
                        continue
                    all_rows.append(cells)

    # Собираем полный список колонок:
    all_cols = []
    for c in REQ_FIRST:
        if c not in all_cols:
            all_cols.append(c)
    for c in dynamic_order_global:
        if c not in all_cols and c not in REQ_LAST and c not in REQ_FIRST:
            all_cols.append(c)
    for c in REQ_LAST:
        if c not in all_cols:
            all_cols.append(c)

    df = pd.DataFrame(all_rows)
    for c in all_cols:
        if c not in df.columns:
            df[c] = ""
    df = df[all_cols]

    # финальная подстраховка против Tekijä:
    mask = df.apply(lambda r: not any("tekijä:" in str(v).lower() for v in r.values), axis=1)
    df = df[mask].reset_index(drop=True)
    return df

if __name__ == "__main__":
    # парсим оба файла
    df_w21 = parse_pdf_any_columns(PDF_PATH_W21)
    df_w22 = parse_pdf_any_columns(PDF_PATH_W22)

    # переименуем колонки в обоих датафреймах
    rename_map = {
        "Työntekijät": "Name",
        "Aika": "Dates",
        "Iltalisä": "Evening shift bonus",
        "Yövuoro": "Night shift bonus",
        "Kaikki yhteensä": "Salary",
    }

    df_w21 = df_w21.rename(columns={old: new for old, new in rename_map.items() if old in df_w21.columns})
    df_w22 = df_w22.rename(columns={old: new for old, new in rename_map.items() if old in df_w22.columns})

    print("1st week rows:", len(df_w21))
    print("2nd week rows:", len(df_w22))

    # записываем в один Excel с двумя листами
    with pd.ExcelWriter(OUT_XLSX, engine="openpyxl") as writer:
        df_w21.to_excel(writer, sheet_name="1st week", index=False)
        df_w22.to_excel(writer, sheet_name="2nd week", index=False)

    print(f"Saved: {Path(OUT_XLSX).resolve()}")
