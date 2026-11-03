"""
Поля позиции базы — как колонки листа «БД» шаблона WV.

Общее для вкладок «Продукция» и «Ввод позиций»: порядок и названия колонок,
разбор значения ячейки и разбор вставки из буфера обмена. Порядок колонок
повторяет лист «БД» (A…L), поэтому блок, скопированный из шаблона, ложится
на сетку столбец в столбец.
"""
from __future__ import annotations

import csv
import io
import re
from typing import Dict, List, Optional, Tuple

from services.dealer_price import parse_number

# (ключ, заголовок как в листе БД, ширина, тип)
#   тип: text | price | int
FIELDS: List[Tuple[str, str, int, str]] = [
    ("num",          "№",            50,  "int"),
    ("article",      "Артикул",      150, "text"),
    ("name",         "Наименование", 320, "text"),
    ("unit",         "Ед. изм.",     60,  "text"),
    ("kaznisa",      "КазНИИСА",     95,  "price"),
    ("rrts",         "РРЦ",          85,  "price"),
    ("mrc",          "МРЦ",          85,  "price"),
    ("opt",          "Опт",          85,  "price"),
    ("partner",      "Партнёр",      90,  "price"),
    ("brand",        "Бренд",        110, "text"),
    ("multiplicity", "Кратность",    70,  "int"),
    ("kaznisa_code", "Код КазНИИСА", 130, "text"),
]
KEYS = [f[0] for f in FIELDS]
KIND = {f[0]: f[3] for f in FIELDS}
TITLE = {f[0]: f[1] for f in FIELDS}
PRICE_KEYS = [k for k in KEYS if KIND[k] == "price"]

SEGMENTS = [("ss", "Слаботочные"), ("os", "Освещение"),
            ("sil", "Силовые"), ("gen", "Общая база")]
SEGMENT_TITLE = dict(SEGMENTS)


def parse_cell(key: str, text) -> object:
    """Значение ячейки по типу поля. Пустое — None.

    Неверное значение — ValueError с понятным текстом: его показывают
    менеджеру как есть.
    """
    s = "" if text is None else str(text).strip()
    kind = KIND.get(key, "text")
    if kind == "text":
        s = " ".join(s.split())
        return s or None
    if not s:
        return None
    if kind == "int":
        v = parse_number(s)
        if v is None or abs(v - round(v)) > 1e-9:
            raise ValueError(f"«{TITLE[key]}»: нужно целое число, а не «{s}»")
        return int(round(v))
    # Цена: «1 240,50», «1240.5», «1 240» — всё понимается
    v = parse_number(s)
    if v is None:
        if re.fullmatch(r"[-−]?\s*[\d\s.,]+", s):
            raise ValueError(f"«{TITLE[key]}»: цена должна быть больше нуля")
        raise ValueError(f"«{TITLE[key]}»: «{s}» — не число")
    return round(v, 2)


def fmt_cell(key: str, value) -> str:
    """Значение для показа в таблице."""
    if value is None or value == "":
        return ""
    kind = KIND.get(key, "text")
    if kind == "price":
        try:
            # 1 240,50 — как принято в КП
            return f"{float(value):,.2f}".replace(",", " ").replace(".", ",")
        except (TypeError, ValueError):
            return str(value)
    if kind == "int":
        try:
            return str(int(value))
        except (TypeError, ValueError):
            return str(value)
    return str(value)


# ─── Вставка из буфера ───────────────────────────────────────────────────────

def _header_key(title: str) -> Optional[str]:
    """Колонка листа БД по заголовку — для вставки блока с шапкой."""
    h = re.sub(r"[^a-zа-я0-9№]+", "", str(title or "").lower().replace("ё", "е"))
    if not h:
        return None
    if h.startswith("№") or h in ("номер", "ngq", "№gq") or h.startswith("ngq"):
        return "num"
    if "артикул" in h:
        return "article"
    if "наимен" in h or "номенкл" in h:
        return "name"
    if ("казнии" in h or "казнис" in h or "агск" in h) and "код" in h:
        return "kaznisa_code"
    if "казнии" in h or "казнис" in h:
        return "kaznisa"
    if "ррц" in h:
        return "rrts"
    if "мрц" in h:
        return "mrc"
    if h.startswith("опт"):
        return "opt"
    if "партн" in h or "проект" in h or "дистр" in h:
        return "partner"
    if "бренд" in h or "производ" in h:
        return "brand"
    if "кратн" in h:
        return "multiplicity"
    if h.startswith("ед"):
        return "unit"
    return None


def parse_clipboard(text: str, start_col: int = 0) -> Tuple[List[Dict[str, str]], str]:
    """Блок из буфера обмена → строки {ключ: текст}.

    Excel кладёт в буфер строки через перевод строки и ячейки через
    табуляцию; ячейку с переводом строки внутри берёт в кавычки — это
    понимает модуль csv.

    Если первая строка блока — шапка (узнаются хотя бы две колонки листа
    БД), колонки раскладываются по заголовкам, в любом порядке. Иначе —
    по месту, начиная с колонки start_col, как вставка в Excel.

    Возвращает (строки, как разложено — для подсказки менеджеру).
    """
    text = (text or "").replace("\r\n", "\n").replace("\r", "\n")
    if not text.strip():
        return [], ""
    rows = list(csv.reader(io.StringIO(text), delimiter="\t"))
    rows = [r for r in rows if any((c or "").strip() for c in r)]
    if not rows:
        return [], ""

    head = [_header_key(c) for c in rows[0]]
    known = [k for k in head if k]
    if len(set(known)) >= 2:
        mapping = {i: k for i, k in enumerate(head) if k}
        body = rows[1:]
        how = "по заголовкам: " + ", ".join(TITLE[k] for k in dict.fromkeys(mapping.values()))
    else:
        mapping = {i: KEYS[start_col + i] for i in range(len(KEYS) - start_col)}
        body = rows
        cols = [KEYS[start_col + i] for i in range(min(max(len(r) for r in rows),
                                                       len(KEYS) - start_col))]
        how = "по месту: " + ", ".join(TITLE[k] for k in cols)

    out = []
    for r in body:
        d = {}
        for i, cell in enumerate(r):
            k = mapping.get(i)
            if k and k not in d:
                d[k] = (cell or "").strip()
        if any(v for v in d.values()):
            out.append(d)
    return out, how


def rows_to_tsv(rows: List[Dict], keys: List[str] = KEYS, header: bool = True) -> str:
    """Строки → текст для буфера: вставляется в Excel как таблица."""
    lines = []
    if header:
        lines.append("\t".join(TITLE[k] for k in keys))
    for r in rows:
        cells = []
        for k in keys:
            v = r.get(k)
            if v is None:
                cells.append("")
            elif KIND.get(k) == "price":
                cells.append(f"{float(v):.2f}".replace(".", ","))
            else:
                cells.append(str(v).replace("\t", " ").replace("\n", " "))
        lines.append("\t".join(cells))
    return "\n".join(lines)
