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

# Сколько строк сверху просматриваем в поисках шапки: над таблицей бывает
# шапка с реквизитами поставщика, условиями и датой
_HEADER_SCAN_ROWS = 40

# Сколько колонок читаем: шире прейскуранты не бывают, а лишние колонки
# справа только замедляют чтение больших файлов
_MAX_COLS = 32

# Листы, где лежат снятые с продажи и архивные позиции, — не прейскурант
_SKIP_SHEET_RE = re.compile(r"удал|выводят|архив|снят|исключ", re.I)

# Латинские буквы, неотличимые на глаз от кириллических: «Aртикул», «Цена KП»
_HDR_HOMO = str.maketrans("AaBCcEeHKkMOoPpTXxYy", "АаВСсЕеНКкМОоРрТХхУу")
_WORD_RE = re.compile(r"[a-zа-яё0-9]+", re.I)

_GROUP_CHARS = "     ⁠'’`"


def _hdr_words(value) -> Tuple[List[str], List[str]]:
    """Слова заголовка: как есть и с латинскими двойниками в кириллице.

    Обе версии нужны: «Aртикул» опознаётся только во второй, а английское
    «article» — только в первой.
    """
    raw = str(value or "").lower().replace("ё", "е")
    cyr = str(value or "").translate(_HDR_HOMO).lower().replace("ё", "е")
    return _WORD_RE.findall(raw), _WORD_RE.findall(cyr)


def _classify_header(value) -> Tuple[Optional[str], Optional[bool]]:
    """Что за колонка: (поле, НДС). Поле None — колонка не нужна.

    Поля: article, name, unit, qty, seb (закупка), kp (продажа), ge,
    price (просто «Цена»), total (итоговая сумма — признак КП или сметы),
    skip (код АГСК, штрихкод и прочее, что артикулом не является).
    НДС: True — «с НДС», False — «без НДС», None — не сказано.
    """
    raw, cyr = _hdr_words(value)
    if not raw:
        return None, None
    # Смысл ищем в обеих версиях и берём более конкретный: в «Цена KП» с
    # латинской K версия как есть видит только «цену», а кириллическая — КП
    field = None
    for words in (raw, cyr):
        f = _field_of(words)
        if f and (field is None or _FIELD_RANK[f] < _FIELD_RANK[field]):
            field = f
    if field is None:
        return None, None

    vat = None
    if field in ("seb", "kp", "ge", "price"):
        w = raw + cyr
        if "ндс" in w:
            vat = not ("без" in w or "облагается" in w and "не" in w)
    return field, vat


# Чем меньше номер, тем конкретнее смысл колонки
_FIELD_RANK = {f: i for i, f in enumerate(
    ("total", "skip", "ge", "seb", "kp", "article", "name", "price", "unit", "qty"))}


def _field_of(words: List[str]) -> Optional[str]:
    """Смысл колонки по словам заголовка. None — колонка не нужна."""
    field = None
    joined = " ".join(words)
    first = words[0]
    if (any(w.startswith("сумм") for w in words)
            or "общая стоимость" in joined
            or first in ("всего", "итого")):
        field = "total"
    elif any(w in ("агск", "казниса", "казнииса", "ntin", "ean", "gtin",
                   "штрихкод", "штрих", "продукции") for w in words):
        field = "skip"
    elif "гэ" in words:
        field = "ge"
    elif any(w.startswith(("себес", "закуп", "входн", "дилер", "партнер"))
             for w in words):
        field = "seb"
    elif "кп" in words or any(w.startswith(("продаж", "реализац"))
                              for w in words):
        field = "kp"
    elif (any(w.startswith(("артикул", "референс")) for w in words)
          or first in ("код", "sku", "article", "reference")
          or "каталожный номер" in joined):
        field = "article"
    elif (any(w.startswith(("наименован", "номенклатур", "названи",
                            "описани")) for w in words)
          or words == ["товар"] or first == "name"):
        field = "name"
    elif first.startswith(("цена", "прайс", "тариф")) or first == "price":
        field = "price"
    elif first in ("ед", "unit") or first.startswith("единиц"):
        field = "unit"
    elif first.startswith(("кол", "qty")):
        field = "qty"
    return field


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

_XLRD_MISSING = (
    "Старый формат .xls читается модулем xlrd, а он не установлен.\n\n"
    "Самое быстрое решение: откройте прайс в Excel и сохраните как "
    "«Книга Excel (*.xlsx)» — такой файл приложение читает само.\n\n"
    "Чтобы принимать .xls напрямую, нужно доустановить модуль и "
    "пересобрать приложение — из папки client:\n"
    "    pip install xlrd\n"
    "    pyinstaller aps_parser.spec --clean --noconfirm"
)


def _read_sheets(path: str, max_rows: Optional[int] = None,
                 only: Optional[str] = None) -> List[Tuple[str, List[list]]]:
    """Листы книги: [(имя, строки)]. max_rows — читать только верх листа."""
    ext = os.path.splitext(path)[1].lower()
    out: List[Tuple[str, List[list]]] = []
    if ext == ".xls":
        try:
            import xlrd
        except ImportError as e:
            raise RuntimeError(_XLRD_MISSING) from e
        wb = xlrd.open_workbook(path)
        for sh in wb.sheets():
            if only is not None and sh.name != only:
                continue
            n = sh.nrows if max_rows is None else min(sh.nrows, max_rows)
            out.append((sh.name, [[sh.cell_value(r, c)
                                   for c in range(min(sh.ncols, _MAX_COLS))]
                                  for r in range(n)]))
        return out
    if ext in (".xlsx", ".xlsm"):
        import openpyxl
        wb = openpyxl.load_workbook(path, data_only=True, read_only=True)
        try:
            for name in wb.sheetnames:
                if only is not None and name != only:
                    continue
                ws = wb[name]
                rows = [list(r[:_MAX_COLS])
                        for r in ws.iter_rows(values_only=True, max_row=max_rows)]
                out.append((name, rows))
        finally:
            wb.close()
        return out
    raise RuntimeError(f"Неподдерживаемый формат файла: {ext or '—'}")


# ─── Поиск шапки ─────────────────────────────────────────────────────────────

def _header_candidate(cells: List[str]) -> dict:
    """Разбор одной строки-кандидата в шапку."""
    cols: Dict[str, int] = {}
    vat: Dict[str, Optional[bool]] = {}
    titles: Dict[str, str] = {}
    totals: List[str] = []
    for c, text in enumerate(cells):
        field, v = _classify_header(text)
        if field is None or field == "skip":
            continue
        if field == "total":
            totals.append(_norm_text(text))
            continue
        if field in cols:          # первая колонка с этим смыслом побеждает
            continue
        cols[field] = c
        vat[field] = v
        titles[field] = _norm_text(text)
    # Одна колонка «Цена» без себес и КП — это цена поставщика, то есть
    # для нас закупка
    if "price" in cols and "seb" not in cols and "kp" not in cols:
        cols["seb"], vat["seb"], titles["seb"] = cols["price"], vat["price"], titles["price"]
    cols.pop("price", None)
    has_price = "seb" in cols or "kp" in cols
    return {"cols": cols, "vat": vat, "titles": titles, "totals": totals,
            "valid": "name" in cols and has_price,
            "score": len(cols)}


def _find_header(rows: List[list]) -> Optional[dict]:
    """Лучшая шапка в верхних строках листа: в одну строку или в две."""
    best = None
    top = rows[:_HEADER_SCAN_ROWS]
    for i, row in enumerate(top):
        one = [_norm_text(v) for v in row]
        cands = [(one, 1)]
        if i + 1 < len(rows):
            nxt = [_norm_text(v) for v in rows[i + 1]]
            width = max(len(one), len(nxt))
            one += [""] * (width - len(one))
            nxt += [""] * (width - len(nxt))
            # Объединённая ячейка верхней строки («Цена» над «Себес» и «КП»)
            # хранит текст только в первой ячейке — протягиваем его вправо,
            # но лишь туда, где снизу есть подзаголовок
            two, carry = [], ""
            for top_t, bot_t in zip(one, nxt):
                carry = top_t or carry
                t = top_t or (carry if bot_t else "")
                two.append(f"{t} {bot_t}".strip())
            cands.append((two, 2))
        for cells, depth in cands:
            h = _header_candidate(cells)
            if not (h["valid"] or h["totals"]):
                continue
            h["row"], h["depth"] = i, depth
            # Шапка в две строки выигрывает, только если узнаёт больше полей
            key = (h["valid"] and not h["totals"], h["score"], -depth)
            if best is None or key > best["_key"]:
                h["_key"] = key
                best = h
    return best


def _locate(path: str) -> Tuple[Optional[str], Optional[dict], str]:
    """Лист и шапка прейскуранта. (лист, шапка, причина отказа)."""
    usable, rejected = [], []
    for name, rows in _read_sheets(path, max_rows=_HEADER_SCAN_ROWS + 2):
        if _SKIP_SHEET_RE.search(name or ""):
            continue
        h = _find_header(rows)
        if h is None:
            continue
        (usable if h["valid"] and not h["totals"] else rejected).append((name, h))
    if usable:
        # Больше узнанных полей — лучше; при равенстве — лист раньше в книге
        name, h = max(usable, key=lambda t: t[1]["score"])
        return name, h, ""
    if rejected:
        name, h = rejected[0]
        cols = ", ".join(f"«{t}»" for t in h["totals"][:2])
        return None, None, (
            f"Это похоже на КП, смету или спецификацию, а не на прейскурант: "
            f"на листе «{name}» есть колонки итоговых сумм ({cols}).\n\n"
            f"В прейскуранте должны быть цены за единицу — без количеств "
            f"и сумм.")
    return None, None, ""


def inspect_dealer_price(path: str) -> Tuple[bool, str]:
    """Похож ли файл на прейскурант. Возвращает (да/нет, причина).

    Причина возвращается, а не проглатывается: «не распознан», «не смог
    прочитать» и «это КП, а не прейскурант» лечатся по-разному.
    Пустая причина при отказе — шапка просто не найдена.
    """
    ext = os.path.splitext(path)[1].lower()
    if ext not in (".xls", ".xlsx", ".xlsm"):
        return False, f"Неподдерживаемое расширение файла: {ext or '—'}."
    try:
        sheet, h, why = _locate(path)
    except Exception as e:
        return False, str(e) or e.__class__.__name__
    if sheet is None:
        return False, why
    return True, ""


def looks_like_dealer_price(path: str) -> bool:
    """Короткая форма проверки для мест, где причина не нужна."""
    ok, _why = inspect_dealer_price(path)
    return ok


# ─── Разбор ──────────────────────────────────────────────────────────────────

def _cell_text(v) -> str:
    """Текст ячейки. Число-артикул без хвоста «.0»: 36008.0 → «36008»."""
    if isinstance(v, float) and v.is_integer():
        return str(int(v))
    return _norm_text(v)


def _suspicious_name(name: str, series: str) -> bool:
    """Наименование, от которого остался обрывок.

    После марки или серии — одна цифра без букв: «ВВГнгLSLTx ,5» вместо
    «ВВГнгLSLTx 3х1,5». Такую позицию в базу класть можно, но найти её по
    спецификации нельзя, поэтому о ней надо сказать.
    """
    n = name.strip()
    if len(n) < 4:
        return True
    if series and n.lower().startswith(series.lower()):
        rest = n[len(series):]
        if (not re.search(r"[a-zа-яё]", rest, re.I)
                and len(rest.strip(" ,.;:-/")) <= 4):
            return True
    return False


# Строка итогов — не товар, даже если у неё есть цена
_TOTAL_ROW_RE = re.compile(r"^(итого|всего|итог|total)\b", re.I)


def parse_dealer_price(path: str, progress_cb=None) -> dict:
    """Разбирает прейскурант.

    Возвращает {"rows": [...], "sheet", "header_row", "columns", "titles",
                "vat_from_header", "skipped_no_price", "skipped_empty",
                "skipped_total", "duplicates", "ambiguous_articles",
                "ambiguous_rows"}.

    Строка: row (номер в файле, 1-based), article, name, unit, qty,
    seb, kp, ge, seb_vat, kp_vat, key, article_src.
    """
    sheet, hdr, why = _locate(path)
    if sheet is None:
        raise RuntimeError(why or (
            "В файле не найдена таблица с колонками «Наименование» и ценой. "
            "Ожидается шапка вида: Артикул | Наименование | Ед. изм. | "
            "Цена Себес | Цена КП."))
    raw = _read_sheets(path, only=sheet)[0][1]
    cols = hdr["cols"]
    start = hdr["row"] + hdr["depth"]          # первая строка данных
    seb_vat = hdr["vat"].get("seb")
    kp_vat  = hdr["vat"].get("kp")

    def cell(row, field):
        c = cols.get(field)
        if c is None or c >= len(row):
            return None
        return row[c]

    def priced(row):
        return parse_number(cell(row, "seb")) or parse_number(cell(row, "kp"))

    # Какие значения «Артикула» не определяют товар: одно и то же значение
    # у позиций с разными наименованиями. В прайсе на кабель там марка
    # («ВВГнгLS» у всех сечений), в других — артикул упаковки и штуки.
    names_by_art: Dict[str, set] = {}
    for i in range(start, len(raw)):
        row = raw[i]
        if not row or not priced(row):
            continue
        a = _cell_text(cell(row, "article")).upper()
        n = _norm_text(cell(row, "name")).lower()
        if a and n:
            names_by_art.setdefault(a, set()).add(n)
    ambiguous = {a for a, names in names_by_art.items() if len(names) > 1}

    out: List[dict] = []
    seen: set = set()
    suspicious: List[dict] = []
    skipped_no_price = skipped_empty = skipped_total = dupes = 0
    ambiguous_rows = 0
    total = max(1, len(raw) - start)

    for i in range(start, len(raw)):
        row = raw[i]
        if not row or all(v in (None, "") for v in row):
            continue
        article = _cell_text(cell(row, "article"))
        name    = _norm_text(cell(row, "name"))
        if not (article or name):
            skipped_empty += 1
            continue
        if not article and _TOTAL_ROW_RE.match(name):
            skipped_total += 1
            continue

        seb = parse_number(cell(row, "seb"))
        kp  = parse_number(cell(row, "kp"))
        ge  = parse_number(cell(row, "ge"))
        if not (seb or kp):
            skipped_no_price += 1
            continue

        # Неуникальный «артикул» товар не определяет — позицию опознаём по
        # наименованию, а исходное значение сохраняем для отчёта
        article_src = article
        if article and article.upper() in ambiguous:
            article = ""
            ambiguous_rows += 1

        key = row_key(article, name)
        if not key:
            skipped_empty += 1
            continue
        if key in seen:
            # Настоящий дубль — и артикул, и наименование те же: первая строка
            dupes += 1
            continue
        seen.add(key)
        if _suspicious_name(name, article_src):
            suspicious.append({"row": i + 1, "name": name})

        out.append({
            "row":     i + 1,
            "article": article,
            "name":    name,
            "unit":    _norm_text(cell(row, "unit")) or "шт.",
            "qty":     parse_number(cell(row, "qty")),
            "seb":     seb,
            "kp":      kp,
            "ge":      ge,
            "seb_vat": seb_vat,
            "kp_vat":  kp_vat,
            "key":     key,
            "article_src": article_src,
        })
        if progress_cb and (i - start) % 500 == 0:
            progress_cb(min(99, int((i - start) / total * 100)),
                        f"Разбор прейскуранта: {len(out):,} поз.")

    used = [v for f, v in (("seb", seb_vat), ("kp", kp_vat)) if f in cols]
    if used and all(v is True for v in used):
        vat_from_header = True
    elif used and all(v is False for v in used):
        vat_from_header = False
    elif any(v is not None for v in used):
        vat_from_header = "mixed"
    else:
        vat_from_header = None

    return {
        "rows":    out,
        "sheet":   sheet,
        "header_row": hdr["row"] + hdr["depth"],   # последняя строка шапки, 1-based
        "header_depth": hdr["depth"],
        "columns": cols,
        "titles":  hdr["titles"],
        "vat_from_header": vat_from_header,
        "skipped_no_price": skipped_no_price,
        "skipped_empty":    skipped_empty,
        "skipped_total":    skipped_total,
        "duplicates":       dupes,
        # Значений «Артикула», которые оказались не уникальными, и позиций,
        # опознанных из-за этого по наименованию
        "ambiguous_articles": len(ambiguous),
        "ambiguous_rows":     ambiguous_rows,
        # Позиции с оборванным наименованием — их стоит поправить в файле
        "suspicious":         suspicious,
    }


def _vat_k(column_vat: Optional[bool], vat_included: bool) -> float:
    """Множитель НДС для колонки.

    Если в заголовке сказано «с НДС» или «без НДС» — верим заголовку.
    Переключатель «Цены уже с НДС» решает только за колонки, где не сказано.
    """
    if column_vat is True:
        return 1.0
    if column_vat is False:
        return VAT
    return 1.0 if vat_included else VAT


def prices_of(row: dict, vat_included: bool) -> Tuple[Optional[float], Optional[float]]:
    """Цены позиции для базы: (Партнёр, КазНИИСА), с НДС.

    «Цена Себес» — закупка, ложится в «Партнёр», из неё считается
    себестоимость. «Цена КП» — цена продажи, ложится в «КазНИИСА»,
    из неё считается Цена КП в предпросмотре. База хранит цены с налогом:
    без НДС себестоимость окажется занижена на 16 %.
    """
    ks = _vat_k(row.get("seb_vat"), vat_included)
    kk = _vat_k(row.get("kp_vat"), vat_included)
    partner = round(row["seb"] * ks, 2) if row.get("seb") else None
    kaznisa = round(row["kp"] * kk, 2) if row.get("kp") else None
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
