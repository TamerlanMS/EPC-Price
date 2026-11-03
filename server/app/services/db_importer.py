import openpyxl
import io
import json
from sqlalchemy.ext.asyncio import AsyncSession
from sqlalchemy import select, update, delete
from app.models.models import Product, BrandConstant, CurrencyRate, ImportLog, Manager, PriceHistory, AuditLog
from typing import Tuple
import logging

_ANOMALY_THRESHOLD = 35.0  # % изменения цены для записи аномалии

logger = logging.getLogger(__name__)


def _open_workbook(file_bytes: bytes):
    """Открывает xlsx или xlsm. keep_vba=True обязателен для xlsm."""
    buf = io.BytesIO(file_bytes)
    try:
        wb = openpyxl.load_workbook(buf, keep_vba=True, data_only=True)
    except Exception:
        buf.seek(0)
        wb = openpyxl.load_workbook(buf, data_only=True)
    return wb


# Всё, что в выгрузках встречается как разделитель разрядов:
# обычный, неразрывный, узкий неразрывный и тонкий пробелы, апостроф.
_GROUP_CHARS = " \u00a0\u202f\u2009\u2007\u2060'\u2019`"


def parse_number(val) -> float | None:
    """Число из ячейки эксель-базы. None, если разобрать нельзя.

    Роль разделителя определяется по длине группы после него:
      «1 240,04»  → 1240.04   (две цифры — дробная часть)
      «293.710»   → 293710.0  (три цифры — разряды)
      «2.937.100» → 2937100.0 (несколько разделителей — точно разряды)
      «293,710.00»→ 293710.0  (два вида — дробным считается последний)

    Раньше любая точка считалась дробной, и «293.710» превращалось
    в 293.71 — цена падала на порядок.
    """
    if val is None:
        return None
    if isinstance(val, (int, float)) and not isinstance(val, bool):
        f = float(val)
        return f if f > 0 else None

    s = str(val).strip()
    if not s:
        return None
    for ch in _GROUP_CHARS:
        s = s.replace(ch, "")
    s = s.replace("\u2212", "-")          # типографский минус
    if not s:
        return None

    dot, comma = s.rfind("."), s.rfind(",")
    if dot >= 0 and comma >= 0:
        # Есть оба знака: дробный — тот, что правее
        dec = "." if dot > comma else ","
        s = s.replace("." if dec == "," else ",", "")
        s = s.replace(dec, ".")
    else:
        sep = "." if dot >= 0 else ("," if comma >= 0 else "")
        if sep:
            tail = s.rsplit(sep, 1)[1]
            head = s.rsplit(sep, 1)[0]
            many = s.count(sep) > 1
            # Ровно три цифры после единственного разделителя — разряды,
            # но не когда целая часть нулевая: «0.125» — это доля, не 125.
            grouping = many or (len(tail) == 3 and tail.isdigit()
                                and head.lstrip("-+") not in ("", "0"))
            s = s.replace(sep, "") if grouping else s.replace(sep, ".")

    try:
        f = float(s)
    except (ValueError, TypeError):
        return None
    return f if f > 0 else None


def safe_float(val) -> float | None:
    return parse_number(val)


def safe_int(val) -> int | None:
    try:
        return int(float(str(val).strip()))
    except (ValueError, TypeError):
        return None


def safe_str(val) -> str | None:
    if val is None:
        return None
    s = str(val).strip()
    return s if s not in ("None", "-", "—", "") else None


async def import_products_from_excel(
    file_bytes: bytes,
    db: AsyncSession,
    filename: str = "import.xlsm",
    segment: str = "ss",
    changed_by: str = "",
) -> Tuple[int, int]:
    wb = _open_workbook(file_bytes)

    sheet_name = None
    for candidate in ["БД", "бд", "BD", "bd", "База", "Sheet1"]:
        if candidate in wb.sheetnames:
            sheet_name = candidate
            break

    if not sheet_name:
        raise ValueError(
            f"Лист 'БД' не найден. Доступные листы: {', '.join(wb.sheetnames)}"
        )

    ws = wb[sheet_name]
    added = 0
    updated = 0

    # ── Считаем кол-во ДО операции ───────────────────────────────────────────
    from sqlalchemy import func as _func
    _cnt_before_q = await db.execute(
        select(_func.count()).select_from(Product)
        .where(Product.segment == segment, Product.is_active == True)
    )
    count_before = _cnt_before_q.scalar() or 0

    # Сегмент больше не гасится перед импортом: файл добавляет новые позиции
    # и обновляет цены существующих, но ничего не удаляет. Удаление — только
    # во вкладке «Продукция».

    all_rows = list(ws.iter_rows(min_row=1, values_only=True))
    if all_rows:
        header_row = all_rows[0]
        logger.info("БД sheet header row (%d cols): %s", len(header_row),
                    [str(v)[:20] if v else "" for v in header_row[:15]])
    if len(all_rows) > 1:
        first_data = all_rows[1]
        logger.info("БД sheet first data row: %s",
                    [str(v)[:20] if v else "" for v in first_data[:15]])

    parsed: list = []
    no_article = 0
    for row in all_rows[1:]:
        if not row or len(row) < 4:
            continue
        if all(v is None for v in row):
            continue
        article = safe_str(row[1])
        name    = safe_str(row[2])
        if not ((article or "").strip() or (name or "").strip()):
            continue          # нет ни артикула, ни наименования — пустая строка
        # Строка заголовков колонок внутри листа (у шаблона — вторая строка)
        if (article or "").strip().lower() == "артикул":
            continue
        if not (article or "").strip():
            no_article += 1
        parsed.append(dict(
            num=safe_int(row[0]), article=article, name=name,
            unit=safe_str(row[3]) or "шт.",
            kaznisa=safe_float(row[4]) if len(row) > 4 else None,
            rrts=safe_float(row[5])    if len(row) > 5 else None,
            mrc=safe_float(row[6])     if len(row) > 6 else None,
            opt=safe_float(row[7])     if len(row) > 7 else None,
            partner=safe_float(row[8]) if len(row) > 8 else None,
            brand=safe_str(row[9])     if len(row) > 9 else None,
            multiplicity=safe_int(row[10]) if len(row) > 10 else None,
            kaznisa_code=safe_str(row[11]) if len(row) > 11 else None,
        ))

    from app.services.product_store import upsert_products
    res = await upsert_products(db, segment, parsed, changed_by=changed_by)
    st = res["stats"]
    added   = st["new"]
    updated = st["updated"] + st["restored"]

    # ── Считаем кол-во ПОСЛЕ операции ────────────────────────────────────────
    _cnt_after_q = await db.execute(
        select(_func.count()).select_from(Product)
        .where(Product.segment == segment, Product.is_active == True)
    )
    count_after = _cnt_after_q.scalar() or 0

    db.add(ImportLog(
        filename=filename,
        segment=segment,
        rows_added=added,
        rows_updated=updated,
        status="success",
        action="import",
        count_before=count_before,
        count_after=count_after,
        changed_by=changed_by or None,
        message=(
            f"[{segment.upper()}] {count_before} → {count_after}  "
            f"(+{added} новых, ~{updated} обновлено"
            + (f", из корзины не возвращено {st['deleted']}" if st["deleted"] else "")
            + ")"
            + (f"  | {changed_by}" if changed_by else "")
        ),
    ))

    await db.commit()
    wb.close()

    # Diagnostic: how many imported products have prices?
    from sqlalchemy import func
    count_q = await db.execute(
        select(func.count()).select_from(Product).where(
            (Product.rrts.isnot(None)) | (Product.kaznisa.isnot(None)) | (Product.mrc.isnot(None))
        )
    )
    with_prices = count_q.scalar() or 0
    total_q = await db.execute(select(func.count()).select_from(Product))
    total_p = total_q.scalar() or 0
    logger.info(
        "Import products done: +%d / ~%d (без артикула: %d). "
        "In DB now: %d total, %d with prices, %d WITHOUT prices.",
        added, updated, no_article, total_p, with_prices, total_p - with_prices
    )
    return {
        "added":      added,
        "updated":    st["updated"],
        "restored":   st["restored"],
        "unchanged":  st["same"],
        "deleted":    st["deleted"],       # в корзине — не возвращены
        "duplicates": st["duplicate"],
        "no_article": no_article,
        "count_before": count_before,
        "count_after":  count_after,
    }


async def clear_segment_products(
    db: AsyncSession,
    segment: str,
    hard_delete: bool = False,
    changed_by: str = "",
) -> int:
    """
    Очищает базу продуктов для указанного сегмента.
    hard_delete=False  → деактивирует (is_active=False), данные сохраняются
    hard_delete=True   → физически удаляет строки
    Возвращает количество затронутых позиций.
    """
    from sqlalchemy import func as _func
    _cnt_q = await db.execute(
        select(_func.count()).select_from(Product)
        .where(Product.segment == segment, Product.is_active == True)
    )
    count_before = _cnt_q.scalar() or 0

    if hard_delete:
        result = await db.execute(
            delete(Product).where(Product.segment == segment)
        )
        count = result.rowcount
    else:
        result = await db.execute(
            update(Product)
            .where(Product.segment == segment, Product.is_active == True)
            .values(is_active=False)
        )
        count = result.rowcount

    count_after = 0 if hard_delete else (count_before - count)
    action = "hard_delete" if hard_delete else "clear"

    db.add(ImportLog(
        filename=f"<очистка сегмента {segment.upper()}>",
        segment=segment,
        rows_added=0,
        rows_updated=count,
        status="success",
        action=action,
        count_before=count_before,
        count_after=count_after,
        changed_by=changed_by or None,
        message=(
            f"[{segment.upper()}] Очищено: {count} позиций "
            f"({'удалено' if hard_delete else 'деактивировано'})"
            + (f"  | {changed_by}" if changed_by else "")
        ),
    ))
    await db.commit()
    logger.info("clear_segment[%s]: %d products %s by %s",
                segment, count, "deleted" if hard_delete else "deactivated", changed_by or "?")
    return count


async def get_segment_stats(db: AsyncSession) -> dict:
    """Возвращает количество активных товаров по каждому сегменту."""
    from sqlalchemy import func
    from app.models.models import ALL_SEGMENTS
    stats = {}
    for seg in ALL_SEGMENTS:
        q = await db.execute(
            select(func.count())
            .select_from(Product)
            .where(Product.segment == seg, Product.is_active == True)
        )
        stats[seg] = q.scalar() or 0
    total = await db.execute(
        select(func.count()).select_from(Product).where(Product.is_active == True)
    )
    stats["total"] = total.scalar() or 0
    return stats


async def import_constants_from_excel(
    file_bytes: bytes,
    db: AsyncSession,
    filename: str = "const.xlsm",
) -> int:
    wb = _open_workbook(file_bytes)

    sheet_name = None
    for candidate in ["Const", "const", "Константы", "Constants"]:
        if candidate in wb.sheetnames:
            sheet_name = candidate
            break

    if not sheet_name:
        raise ValueError(
            f"Лист 'Const' не найден. Доступные листы: {', '.join(wb.sheetnames)}"
        )

    ws = wb[sheet_name]
    count = 0
    currency_map = {}
    managers_seen = set()

    for row in ws.iter_rows(min_row=2, values_only=True):
        if not row or all(v is None for v in row):
            continue

        # ── Менеджеры: колонки B-E (индексы 1-4) ─────────────────────────
        if len(row) > 4:
            full_name = safe_str(row[1])
            position  = safe_str(row[2])
            email     = safe_str(row[3])
            phone     = safe_str(row[4])
            # Пропускаем заголовок и плейсхолдеры
            if full_name and full_name not in ("Фамилия, имя", "ФИО") and full_name not in managers_seen:
                managers_seen.add(full_name)
                ex = await db.execute(select(Manager).where(Manager.full_name == full_name))
                m = ex.scalar_one_or_none()
                if m:
                    if position: m.position = position
                    if email:    m.email    = email
                    if phone:    m.phone    = phone
                    m.is_active = True
                else:
                    db.add(Manager(
                        full_name=full_name,
                        position=position,
                        email=email,
                        phone=phone,
                        is_active=True,
                    ))

        if len(row) > 13:
            brand = safe_str(row[7])
            if brand and brand not in ["Бренд", "Brand"]:
                margin    = safe_float(row[8])
                logistics = safe_float(row[9])
                rate_val  = safe_float(row[10])
                curr      = safe_float(row[11])
                nds       = safe_float(row[12])
                gp        = safe_float(row[13])

                existing = await db.execute(
                    select(BrandConstant).where(BrandConstant.brand == brand)
                )
                bc = existing.scalar_one_or_none()

                if bc:
                    if margin:    bc.margin       = margin
                    if logistics: bc.logistics     = logistics
                    if rate_val:  bc.rate          = rate_val
                    if curr:      bc.currency_rate = curr
                    if nds:       bc.nds           = nds
                    if gp:        bc.gp            = gp
                else:
                    db.add(BrandConstant(
                        brand=brand,
                        margin=margin or 1.2,
                        logistics=logistics or 1.03,
                        rate=rate_val or 4.0,
                        currency_rate=curr or 1.0,
                        nds=nds or 1.16,
                        gp=gp or 0.8,
                    ))
                count += 1

        if len(row) > 22:
            cname = safe_str(row[21])
            crate = safe_float(row[22])
            if cname and crate:
                currency_map[cname] = crate

    for cname, crate in currency_map.items():
        ex = await db.execute(select(CurrencyRate).where(CurrencyRate.name == cname))
        cr = ex.scalar_one_or_none()
        if cr:
            cr.rate = crate
        else:
            db.add(CurrencyRate(name=cname, rate=crate))

    db.add(ImportLog(
        filename=filename,
        rows_added=count,
        rows_updated=len(currency_map),
        status="success",
        message=f"Брендов: {count}, курсов валют: {len(currency_map)}"
    ))

    await db.commit()
    wb.close()
    logger.info(f"Import constants: brands={count}, currencies={len(currency_map)}")
    return count
