"""
Дилерский прайс: разбор эксель-файла поставщика.

Формат отличается от прейскуранта КазНИИСА принципиально: кода АГСК в нём
нет вообще, позиция опознаётся по артикулу, а часть строк приходит и без
него — только с наименованием. Поэтому у файла своя читалка, а не ветка
внутри разбора прейскуранта.

Колонки ищутся по заголовкам, а не по номерам: у разных поставщиков они
стоят в разном порядке, зато называются похоже. Заголовок ищется в первых
двадцати строках — над таблицей обычно висит шапка с названием прайса.

Из трёх ценовых колонок осмысленных данных ровно в одной. В разобранных
файлах «Цена КП» = «Цена Себес» × 1.2, а «Цена ГЭ» = «Цена КП» × 1.2 —
последние две вычисляются из первой. Читаем всё равно все три: считать
наценку своей формулой значит однажды разойтись с поставщиком.
"""
from __future__ import annotations

import os
import re
from typing import Dict, List, Optional, Tuple

VAT = 1.16

# Сколько строк сверху просматриваем в поисках шапки таблицы
_HEADER_SCAN_ROWS = 20
# Сколько колонок читаем: прайсы шире семи колонок не встречались,
# но запас на служебные поля не мешает
_MAX_COLS = 24

# Заголовок → поле. Сравнение по нормализованной строке: регистр, пробелы
# и точки у поставщиков пляшут, а слова одни и те же.
_HEADER_MAP = {
    "article": ("артикул", "код товара", "article", "sku"),
    "name":    ("наименование", "название", "товар", "name"),
    "unit":    ("едизм", "единица", "единицаизмерения", "ед", "unit"),
    "qty":     ("колво", "количество", "qty"),
    "seb":     ("ценасебес", "себес", "себестоимость", "ценазакупа",
                "закупочнаяцена", "закуп", "дилерскаяцена", "партнерская"),
    "kp":      ("ценакп", "кп", "ценапродажи", "розница"),
    "ge":      ("ценагэ", "гэ"),
}
# Поля, без которых таблица не таблица
_REQUIRED = ("name",)
_PRICE_FIELDS = ("seb", "kp", "ge")

_GROUP_CHARS = "     ⁠'’`"


def _norm_header(value) -> str:
    """Заголовок к виду, по которому его можно сравнивать."""
    s = str(value or "").strip().lower()
    s = s.replace("ё", "е")
    return re.sub(r"[^a-zа-я0-9]+", "", s)


def parse_number(val) -> Optional[float]:
    """Число из ячейки прайса. None, если разобрать нельзя.

    Повторяет логику импорта базы: точка с тремя цифрами после неё — это
    разделитель разрядов, а не дробная часть. «293.710» — это 293 710,
    а не 293 рубля 71 копейка.
    """
    if val is None:
        return None
    if isinstance(val, bool):
        return None
    if isinstance(val, (int, float)):
        f = float(val)
        return f if f > 0 else None

    s = str(val).strip()
    if not s:
        return None
    for ch in _GROUP_CHARS:
        s = s.replace(ch, "")
    s = s.replace("−", "-")
    if not s:
        return None

    dot, comma = s.rfind("."), s.rfind(",")
    if dot >= 0 and comma >= 0:
        dec = "." if dot > comma else ","
        s = s.replace("." if dec == "," else ",", "").replace(dec, ".")
    else:
        sep = "." if dot >= 0 else ("," if comma >= 0 else "")
        if sep:
            head, tail = s.rsplit(sep, 1)
            grouping = (s.count(sep) > 1
                        or (len(tail) == 3 and tail.isdigit()
                            and head.lstrip("-+") not in ("", "0")))
            s = s.replace(sep, "") if grouping else s.replace(sep, ".")
    try:
        f = float(s)
    except (ValueError, TypeError):
        return None
    return f if f > 0 else None


def _norm_text(value) -> str:
    return " ".join(str(value or "").split())


def row_key(article: str, name: str) -> str:
    """Ключ позиции: артикул, а если его нет — наименование.

    Тот же ключ, что и при импорте базы, включая префикс «~» для имён.
    Без него позиции без артикула схлопнулись бы в одну.
    """
    a = (article or "").strip()
    if a:
        return a.upper()
    n = " ".join((name or "").split()).lower()
    return f"~{n}" if n else ""


# ─── Чтение файла ────────────────────────────────────────────────────────────

def _read_xls(path: str) -> Tuple[str, List[list]]:
    """Старый формат BIFF. openpyxl его не открывает, нужен xlrd."""
    try:
        import xlrd
    except ImportError as e:
        raise RuntimeError(
            "Для файлов .xls нужен модуль xlrd. Установите его "
            "(pip install xlrd) или пересохраните прайс в .xlsx."
        ) from e

    wb = xlrd.open_workbook(path)
    sh = wb.sheet_by_index(0)
    rows = []
    for r in range(sh.nrows):
        rows.append([sh.cell_value(r, c) for c in range(min(sh.ncols, _MAX_COLS))])
    return sh.name, rows


def _read_xlsx(path: str) -> Tuple[str, List[list]]:
    import openpyxl
    wb = openpyxl.load_workbook(path, data_only=True, read_only=True)
    try:
        ws = wb[wb.sheetnames[0]]
        rows = []
        for row in ws.iter_rows(values_only=True):
            rows.append(list(row[:_MAX_COLS]))
        return ws.title, rows
    finally:
        wb.close()


def _read_any(path: str) -> Tuple[str, List[list]]:
    ext = os.path.splitext(path)[1].lower()
    if ext == ".xls":
        return _read_xls(path)
    if ext in (".xlsx", ".xlsm"):
        return _read_xlsx(path)
    raise RuntimeError(f"Неподдерживаемый формат файла: {ext or '—'}")


# ─── Поиск шапки ─────────────────────────────────────────────────────────────

def _find_header(rows: List[list]) -> Tuple[int, Dict[str, int]]:
    """Строка заголовков и раскладка колонок. (-1, {}) — не нашлось."""
    best_i, best_map = -1, {}
    for i, row in enumerate(rows[:_HEADER_SCAN_ROWS]):
        found: Dict[str, int] = {}
        for c, cell in enumerate(row):
            h = _norm_header(cell)
            if not h:
                continue
            for field, variants in _HEADER_MAP.items():
                if field in found:
                    continue
                if h in variants or any(h.startswith(v) for v in variants):
                    found[field] = c
                    break
        # Годится шапка, где есть наименование и хотя бы одна цена
        if all(f in found for f in _REQUIRED) and any(f in found for f in _PRICE_FIELDS):
            if len(found) > len(best_map):
                best_i, best_map = i, found
    return best_i, best_map


def looks_like_dealer_price(path: str) -> bool:
    """Похож ли файл на дилерский прайс. Ошибки чтения — это «нет»."""
    if os.path.splitext(path)[1].lower() not in (".xls", ".xlsx", ".xlsm"):
        return False
    try:
        _sheet, rows = _read_any(path)
    except Exception:
        return False
    i, _m = _find_header(rows)
    return i >= 0


# ─── Разбор ──────────────────────────────────────────────────────────────────

def parse_dealer_price(path: str, progress_cb=None) -> dict:
    """Разбирает дилерский прайс.

    Возвращает {"rows": [...], "sheet": str, "columns": {...},
                "skipped_no_price": int, "skipped_empty": int}.

    Строка: row (номер в файле, 1-based), article, name, unit, qty,
    seb, kp, ge, key.
    """
    sheet, raw = _read_any(path)
    hdr_i, cols = _find_header(raw)
    if hdr_i < 0:
        raise RuntimeError(
            "В файле не найдена таблица с колонками «Наименование» и ценой. "
            "Ожидается шапка вида: Артикул | Наименование | Ед. изм. | "
            "Кол-во | Цена Себес | Цена КП | Цена ГЭ."
        )

    def cell(row, field):
        c = cols.get(field)
        if c is None or c >= len(row):
            return None
        return row[c]

    out: List[dict] = []
    seen: set = set()
    skipped_no_price = skipped_empty = dupes = 0
    total = max(1, len(raw) - hdr_i - 1)

    for i in range(hdr_i + 1, len(raw)):
        row = raw[i]
        if not row or all(v in (None, "") for v in row):
            continue
        article = _norm_text(cell(row, "article"))
        name    = _norm_text(cell(row, "name"))
        if not (article or name):
            skipped_empty += 1
            continue

        seb = parse_number(cell(row, "seb"))
        kp  = parse_number(cell(row, "kp"))
        ge  = parse_number(cell(row, "ge"))
        if not (seb or kp):
            skipped_no_price += 1
            continue

        key = row_key(article, name)
        if not key:
            skipped_empty += 1
            continue
        if key in seen:
            # Дубль внутри прайса: побеждает первая строка. Брать последнюю
            # значило бы молча предпочесть ту, что ниже в файле, — а порядок
            # строк у поставщика ничего не значит.
            dupes += 1
            continue
        seen.add(key)

        out.append({
            "row":     i + 1,
            "article": article,
            "name":    name,
            "unit":    _norm_text(cell(row, "unit")) or "шт.",
            "qty":     parse_number(cell(row, "qty")),
            "seb":     seb,
            "kp":      kp,
            "ge":      ge,
            "key":     key,
        })
        if progress_cb and (i - hdr_i) % 500 == 0:
            progress_cb(min(99, int((i - hdr_i) / total * 100)),
                        f"Разбор прайса: {len(out):,} поз.")

    return {
        "rows":    out,
        "sheet":   sheet,
        "columns": cols,
        "header_row": hdr_i + 1,
        "skipped_no_price": skipped_no_price,
        "skipped_empty":    skipped_empty,
        "duplicates":       dupes,
    }


def prices_of(row: dict, vat_included: bool) -> Tuple[Optional[float], Optional[float]]:
    """Цены позиции для базы: (Партнёр, КазНИИСА).

    «Цена Себес» — закупка, ложится в «Партнёр», из неё считается
    себестоимость. «Цена КП» — цена продажи, ложится в «КазНИИСА»,
    из неё считается Цена КП в предпросмотре.

    Если в прайсе цены без НДС, обе домножаются на 1.16: база хранит
    цены с налогом, иначе себестоимость окажется занижена на 16 %.
    """
    k = 1.0 if vat_included else VAT
    partner = round(row["seb"] * k, 2) if row.get("seb") else None
    kaznisa = round(row["kp"] * k, 2) if row.get("kp") else None
    return partner, kaznisa


# ─── Сопоставление с эксель-базой сегмента ───────────────────────────────────

# Колонки базы, 1-based. Совпадают с раскладкой импорта.
COL_ARTICLE, COL_NAME, COL_UNIT = 2, 3, 4
COL_KAZNISA = 5          # цена продажи
COL_PARTNER = 9          # закупка


def _base_key(article: str, name: str) -> str:
    return row_key(article, name)


def match_base_to_dealer(base_rows: List[dict], dealer_rows: List[dict],
                         vat_included: bool, progress_cb=None) -> dict:
    """Сверяет строки эксель-базы с позициями прайса.

    Возвращает {"changed": [...], "unmatched_base": N,
                "unmatched_dealer": [...], "stats": {...}}.

    В changed попадают только строки, где цена реально меняется: писать
    в файл то же самое число значит без нужды трогать сотни ячеек и
    портить сравнение копий.
    """
    by_article: Dict[str, dict] = {}
    by_name:    Dict[str, dict] = {}
    for d in dealer_rows:
        if d["article"]:
            by_article.setdefault(d["article"].upper(), d)
        nm = d["name"].lower()
        if nm:
            by_name.setdefault(nm, d)

    changed: List[dict] = []
    used: set = set()
    unmatched_base = 0
    by_art_n = by_name_n = same_n = 0
    total = max(1, len(base_rows))

    for i, b in enumerate(base_rows):
        d = None
        how = ""
        art = (b.get("article") or "").strip().upper()
        if art:
            d = by_article.get(art)
            if d is not None:
                how = "артикул"
        if d is None:
            nm = (b.get("name") or "").strip().lower()
            if nm:
                d = by_name.get(nm)
                if d is not None:
                    how = "наименование"
        if d is None:
            unmatched_base += 1
            continue

        used.add(d["key"])
        if how == "артикул":
            by_art_n += 1
        else:
            by_name_n += 1

        new_partner, new_kaznisa = prices_of(d, vat_included)
        old_partner = b.get("partner")
        old_kaznisa = b.get("kaznisa")

        def _same(a, x):
            return a is not None and x is not None and abs(float(a) - float(x)) < 0.01

        if (new_partner is None or _same(old_partner, new_partner)) and \
           (new_kaznisa is None or _same(old_kaznisa, new_kaznisa)):
            same_n += 1
            continue

        _diff = None
        _pct = None
        if new_kaznisa is not None and old_kaznisa:
            _diff = round(new_kaznisa - float(old_kaznisa), 2)
            _pct = round((new_kaznisa / float(old_kaznisa) - 1) * 100, 1)

        changed.append({
            "row":          b["row"],
            "article":      b.get("article", ""),
            "name":         b.get("name", ""),
            "name_pl":      d["name"],
            "brand":        b.get("brand", ""),
            "unit":         b.get("unit", ""),
            "how":          how,
            "price_old":    old_kaznisa,
            "price_new":    new_kaznisa,
            "partner_old":  old_partner,
            "partner_new":  new_partner,
            "diff_abs":     _diff,
            "diff_pct":     _pct,
        })
        if progress_cb and i % 5000 == 0:
            progress_cb(min(99, int(i / total * 100)),
                        f"Сверка с базой: {i:,} из {len(base_rows):,}")

    unmatched_dealer = [d for d in dealer_rows if d["key"] not in used]
    return {
        "changed":          changed,
        "unmatched_dealer": unmatched_dealer,
        "stats": {
            "base_rows":        len(base_rows),
            "dealer_rows":      len(dealer_rows),
            "matched":          by_art_n + by_name_n,
            "by_article":       by_art_n,
            "by_name":          by_name_n,
            "unchanged":        same_n,
            "changed":          len(changed),
            "unmatched_base":   unmatched_base,
            "unmatched_dealer": len(unmatched_dealer),
        },
    }


def write_dealer_prices_to_base(path: str, changed: List[dict],
                                sheet_name: str = "") -> dict:
    """Пишет «Партнёр» и «КазНИИСА» в исходный файл базы.

    Файл собирается во временный и подменяется одним движением: обрыв
    записи не должен оставить менеджера с половиной базы.

    Про formula_cells: openpyxl сохраняет формулу, но не её значение.
    Если цены в базе заданы формулами, файл после записи нужно один раз
    открыть в Excel и сохранить, иначе следующий импорт увидит пустоту.
    """
    import openpyxl
    from services.pricelist_sync import _find_db_sheet

    if not changed:
        return {"written": 0, "formula_cells": 0}

    wb = openpyxl.load_workbook(path, data_only=False,
                                keep_vba=path.lower().endswith(".xlsm"))
    try:
        sheet = sheet_name if sheet_name in wb.sheetnames else _find_db_sheet(wb)
        ws = wb[sheet]

        written = 0
        for m in changed:
            if m.get("price_new") is not None:
                c = ws.cell(row=m["row"], column=COL_KAZNISA)
                c.value = m["price_new"]
                c.number_format = "#,##0.00"
            if m.get("partner_new") is not None:
                c = ws.cell(row=m["row"], column=COL_PARTNER)
                c.value = m["partner_new"]
                c.number_format = "#,##0.00"
            written += 1

        formula_cells = 0
        for row in ws.iter_rows(min_row=2, max_col=12):
            for c in row:
                if isinstance(c.value, str) and c.value.startswith("="):
                    formula_cells += 1

        tmp = path + ".tmp"
        wb.save(tmp)
    finally:
        wb.close()

    os.replace(tmp, path)
    return {"written": written, "formula_cells": formula_cells}


def read_base_rows_ext(path: str):
    """Строки эксель-базы вместе с ценой «Партнёр».

    Штатная читалка сверки с прейскурантом колонку «Партнёр» не берёт —
    там правится только КазНИИСА. Здесь она нужна: дилерский прайс
    обновляет обе цены, и без старого значения не видно, что меняется.
    """
    import openpyxl
    from services.pricelist_sync import _find_db_sheet

    wb = openpyxl.load_workbook(path, data_only=True, read_only=True)
    try:
        sheet = _find_db_sheet(wb)
        ws = wb[sheet]
        rows = []
        for idx, row in enumerate(ws.iter_rows(min_row=2, values_only=True), start=2):
            if not row or all(v is None for v in row):
                continue

            def cell(i):
                return row[i - 1] if len(row) >= i else None

            article = _norm_text(cell(COL_ARTICLE))
            name    = _norm_text(cell(COL_NAME))
            if not (article or name):
                continue
            rows.append({
                "row":     idx,
                "article": article,
                "name":    name,
                "unit":    _norm_text(cell(COL_UNIT)) or "шт.",
                "brand":   _norm_text(cell(10)),
                "kaznisa": parse_number(cell(COL_KAZNISA)),
                "partner": parse_number(cell(COL_PARTNER)),
            })
        return rows, sheet
    finally:
        wb.close()
