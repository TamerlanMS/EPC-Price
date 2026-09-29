"""
Отчёт по сверке дилерского прайса.

Листов столько, сколько исходов у сверки, — смешивать их нельзя, у них
разный набор колонок и разные действия по итогу:

    «Итоги»          — что за файл, в каком режиме и с каким результатом;
    «Обновление цен» — строки приложенной базы, где цена меняется;
    «Есть в базах»   — позиции прайса, найденные в сегментах (какой именно);
    «В общую базу»   — позиции, которых нет нигде.

Позиции базы, которых нет в прайсе, отдельным листом не выводятся: на
полной базе сегмента туда попадает 119 тысяч строк из 120. Их количество
видно в итогах.
"""
import logging
import os
from datetime import datetime
from typing import Dict, List

import openpyxl
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side

logger = logging.getLogger(__name__)

FONT_NAME = "Arial"

HDR_FILL   = PatternFill("solid", fgColor="1F3864")
HDR_FONT   = Font(name=FONT_NAME, size=10, bold=True, color="FFFFFF")
BASE_FONT  = Font(name=FONT_NAME, size=10)
BOLD_FONT  = Font(name=FONT_NAME, size=10, bold=True)
TITLE_FONT = Font(name=FONT_NAME, size=13, bold=True, color="1F3864")

UP_FILL   = PatternFill("solid", fgColor="F8CBAD")   # цена выросла
DOWN_FILL = PatternFill("solid", fgColor="D5E8D4")   # цена упала
NEW_FILL  = PatternFill("solid", fgColor="DEEBF7")   # цены не было

_THIN  = Side(style="thin", color="BFBFBF")
BORDER = Border(left=_THIN, right=_THIN, top=_THIN, bottom=_THIN)

MONEY = "#,##0.00"
PCT   = '+0.0%;-0.0%;"—"'

SEGMENT_NAMES = {
    "ss":  "Слаботочка",
    "os":  "Освещение",
    "sil": "Силовое",
    "gen": "Общая база",
}


def _header(ws, headers, widths):
    for c, (h, w) in enumerate(zip(headers, widths), 1):
        cell = ws.cell(row=1, column=c, value=h)
        cell.fill, cell.font = HDR_FILL, HDR_FONT
        cell.alignment = Alignment(horizontal="center", vertical="center",
                                   wrap_text=True)
        cell.border = BORDER
        ws.column_dimensions[openpyxl.utils.get_column_letter(c)].width = w
    ws.freeze_panes = "A2"
    ws.auto_filter.ref = f"A1:{openpyxl.utils.get_column_letter(len(headers))}1"


def _put(ws, r, c, value, fmt=None, fill=None, font=None):
    cell = ws.cell(row=r, column=c, value=value)
    cell.font = font or BASE_FONT
    cell.border = BORDER
    if fmt:
        cell.number_format = fmt
    if fill:
        cell.fill = fill
    return cell


# ─── Листы ───────────────────────────────────────────────────────────────────

def _sheet_changes(wb, rows: List[Dict], applied: bool):
    ws = wb.active
    ws.title = "Обновление цен"
    _header(ws,
            ["Стр.", "Артикул", "Наименование (база)", "Наименование (прайс)",
             "Бренд", "Ед.", "Совпало по",
             "Партнёр было", "Партнёр станет",
             "КазНИИСА было", "КазНИИСА станет", "Разница, ₸", "Разница, %"],
            [8, 20, 42, 42, 14, 8, 14, 16, 16, 16, 16, 14, 12])

    for i, m in enumerate(rows, start=2):
        old = m.get("price_old")
        new = m.get("price_new")
        if old is None:
            fill = NEW_FILL
        elif new is not None and new > old:
            fill = UP_FILL
        elif new is not None and new < old:
            fill = DOWN_FILL
        else:
            fill = None

        _put(ws, i, 1,  m["row"])
        _put(ws, i, 2,  m.get("article", ""))
        _put(ws, i, 3,  m.get("name", ""))
        _put(ws, i, 4,  m.get("name_pl", ""))
        _put(ws, i, 5,  m.get("brand", ""))
        _put(ws, i, 6,  m.get("unit", ""))
        _put(ws, i, 7,  m.get("how", ""))
        _put(ws, i, 8,  m.get("partner_old"), MONEY)
        _put(ws, i, 9,  m.get("partner_new"), MONEY, fill)
        _put(ws, i, 10, old, MONEY)
        _put(ws, i, 11, new, MONEY, fill, BOLD_FONT)
        _put(ws, i, 12, m.get("diff_abs"), MONEY, fill)
        d = m.get("diff_pct")
        _put(ws, i, 13, (d / 100.0) if d is not None else None, PCT, fill)

    r = len(rows) + 3
    ws.cell(row=r, column=1,
            value=("Цены записаны в исходный файл базы."
                   if applied else
                   "Режим сравнения: файл базы НЕ изменялся.")).font = BOLD_FONT
    ws.cell(row=r + 1, column=1,
            value="Красный — цена выросла, зелёный — упала, "
                  "голубой — цены в базе не было.").font = BASE_FONT


def _sheet_found(wb, rows: List[Dict]):
    ws = wb.create_sheet("Есть в базах")
    _header(ws,
            ["Артикул", "Наименование (прайс)", "Наименование (база)",
             "Сегмент", "Совпало по",
             "Партнёр в базе", "Партнёр в прайсе",
             "КазНИИСА в базе", "КазНИИСА в прайсе"],
            [20, 46, 46, 16, 14, 16, 16, 16, 16])
    for i, m in enumerate(rows, start=2):
        seg = m.get("segment", "")
        _put(ws, i, 1, m.get("article", ""))
        _put(ws, i, 2, m.get("name", ""))
        _put(ws, i, 3, m.get("db_name", ""))
        _put(ws, i, 4, SEGMENT_NAMES.get(seg, seg))
        _put(ws, i, 5, "артикул" if m.get("how") == "key" else "артикул (нечёткий)")
        _put(ws, i, 6, m.get("old_partner"), MONEY)
        _put(ws, i, 7, m.get("new_partner"), MONEY)
        _put(ws, i, 8, m.get("old_kaznisa"), MONEY)
        _put(ws, i, 9, m.get("new_kaznisa"), MONEY, font=BOLD_FONT)

    r = len(rows) + 3
    ws.cell(row=r, column=1,
            value="Эти позиции уже есть в базах — в общую базу они не "
                  "загружались.").font = BOLD_FONT


def _sheet_to_general(wb, rows: List[Dict], loaded: bool):
    ws = wb.create_sheet("В общую базу")
    _header(ws,
            ["Артикул", "Наименование", "Ед.", "Партнёр", "КазНИИСА"],
            [22, 62, 12, 18, 18])
    for i, e in enumerate(rows, start=2):
        _put(ws, i, 1, e.get("article", ""))
        _put(ws, i, 2, e.get("name", ""))
        _put(ws, i, 3, e.get("unit", ""))
        _put(ws, i, 4, e.get("partner"), MONEY)
        _put(ws, i, 5, e.get("kaznisa"), MONEY, font=BOLD_FONT)

    r = len(rows) + 3
    ws.cell(row=r, column=1,
            value=("Позиции загружены в общую базу (сегмент «Общая база»)."
                   if loaded else
                   "Режим сравнения: в базу ничего не загружалось.")).font = BOLD_FONT
    ws.cell(row=r + 1, column=1,
            value="Позиции без артикула опознаются по наименованию.").font = BASE_FONT


def _sheet_summary(wb, stats: Dict, price_name: str, base_path: str,
                   applied: bool, backup: str, vat_included: bool,
                   update_existing: bool):
    ws = wb.create_sheet("Итоги", 0)
    ws.column_dimensions["A"].width = 46
    ws.column_dimensions["B"].width = 38

    ws["A1"] = "Сверка дилерского прайса"
    ws["A1"].font = TITLE_FONT
    r = 3

    def line(label, value, bold=False):
        nonlocal r
        ws.cell(row=r, column=1, value=label).font = BOLD_FONT if bold else BASE_FONT
        c = ws.cell(row=r, column=2, value=value)
        c.font = BOLD_FONT if bold else BASE_FONT
        r += 1

    line("Дата сверки", datetime.now().strftime("%d.%m.%Y %H:%M"))
    line("Прайс", price_name)
    line("Цены в прайсе", "уже с НДС" if vat_included else "без НДС, ×1.16")
    line("Обновлять имеющиеся", "да" if update_existing else "нет", True)
    if update_existing and base_path:
        line("Файл базы", os.path.basename(base_path))
        line("Режим",
             "Цены применены к файлу" if applied
             else "ТОЛЬКО СРАВНЕНИЕ — файл базы не изменялся", True)
        if backup:
            line("Копия до изменений", os.path.basename(backup))
    r += 1

    line("Позиций в прайсе", stats.get("dealer_rows", 0), True)
    if stats.get("base_rows"):
        line("Строк в базе сегмента", stats.get("base_rows", 0))
        line("Сопоставлено с базой", stats.get("matched", 0), True)
        line("    по артикулу", stats.get("by_article", 0))
        line("    по наименованию", stats.get("by_name", 0))
        line("Цена изменится", stats.get("changed", 0), True)
        line("Цена совпадает", stats.get("unchanged", 0))
        line("Нет в прайсе (строк базы)", stats.get("unmatched_base", 0))
        r += 1
    if stats.get("db_matched") is not None:
        line("Найдено в базах (по БД)", stats.get("db_matched", 0), True)
        for seg, n in sorted((stats.get("by_segment") or {}).items(),
                             key=lambda kv: -kv[1]):
            line(f"    {SEGMENT_NAMES.get(seg, seg)}", n)
    line("Новых позиций", stats.get("new", 0), True)
    line("    добавлено в общую базу", stats.get("added", 0))
    line("    обновлено в общей базе", stats.get("updated", 0))
    r += 1
    line("Пропущено без цены", stats.get("skipped_no_price", 0))
    line("Дублей в прайсе", stats.get("duplicates", 0))
    r += 1
    ws.cell(row=r, column=1,
            value="«Цена Себес» из прайса → колонка «Партнёр» (закупка), "
                  "«Цена КП» → колонка «КазНИИСА» (цена продажи).").font = BASE_FONT
    ws.cell(row=r + 1, column=1,
            value="Сопоставление по артикулу и точному наименованию. "
                  "Нечёткий поиск не применяется.").font = BASE_FONT


# ─── Точка входа ─────────────────────────────────────────────────────────────

def default_report_path(src_path: str) -> str:
    """Отчёт ложится рядом с исходным файлом, с датой в имени."""
    folder = os.path.dirname(os.path.abspath(src_path))
    stem   = os.path.splitext(os.path.basename(src_path))[0]
    stamp  = datetime.now().strftime("%Y-%m-%d_%H-%M")
    return os.path.join(folder, f"Сверка прайса — {stem} — {stamp}.xlsx")


def build_dealer_report(out_path: str, stats: Dict, price_name: str,
                        changed: List[Dict] = None,
                        found: List[Dict] = None,
                        to_general: List[Dict] = None,
                        base_path: str = "", applied: bool = False,
                        backup: str = "", loaded: bool = False,
                        vat_included: bool = False,
                        update_existing: bool = True) -> str:
    wb = openpyxl.Workbook()
    _sheet_changes(wb, changed or [], applied)
    if found:
        _sheet_found(wb, found)
    _sheet_to_general(wb, to_general or [], loaded)
    _sheet_summary(wb, stats, price_name, base_path, applied, backup,
                   vat_included, update_existing)
    wb.active = 0
    wb.save(out_path)
    logger.info("dealer report saved: %s", out_path)
    return out_path
