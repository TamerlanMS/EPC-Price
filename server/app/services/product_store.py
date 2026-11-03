"""
Добавление и обновление позиций базы — одно правило для всех путей.

Им пользуются и импорт из Excel, и вкладка «Ввод позиций». Правило:

  • новая позиция (такого ключа в сегменте нет) — создаётся со всеми полями;
  • существующая — меняются только цены: КазНИИСА, РРЦ, МРЦ, Опт, Партнёр.
    Пустая цена во входных данных старую не стирает. Наименование, бренд,
    код и прочее правятся только во вкладке «Продукция»;
  • удалённая в интерфейсе (лежит в корзине) — не трогается и не
    возвращается: удаление — осознанное решение, старый файл его не отменяет;
  • погашенная прежним импортом (неактивна, но в корзину её никто не клал) —
    возвращается, если снова пришла. Прежний импорт гасил весь сегмент и
    включал только то, что было в файле, — такие позиции не удалялись
    сознательно.

Ничего не удаляется: удалить позицию можно только во вкладке «Продукция».

Ключ позиции — как у импорта: артикул, а без него — «~наименование».
Позиция опознаётся внутри сегмента: один артикул в двух сегментах — две
разные позиции.
"""
from __future__ import annotations

import json
import logging
from datetime import datetime, timezone
from typing import Dict, List, Optional

from sqlalchemy import select, update
from sqlalchemy.ext.asyncio import AsyncSession

from app.models.models import AuditLog, PriceHistory, Product

logger = logging.getLogger(__name__)

PRICE_FIELDS = ("kaznisa", "rrts", "mrc", "opt", "partner")
TEXT_FIELDS = ("article", "name", "unit", "brand", "kaznisa_code")
# Изменение цены больше чем на столько процентов записывается как аномалия
ANOMALY_PCT = 35.0


def row_key(article: Optional[str], name: Optional[str]) -> str:
    """Ключ позиции: артикул, а если его нет — наименование."""
    a = (article or "").strip()
    if a:
        return a
    n = " ".join((name or "").split()).lower()
    return f"~{n}" if n else ""


def _clean(row: Dict) -> Dict:
    """Входная строка к виду полей Product."""
    out: Dict = {}
    for f in TEXT_FIELDS:
        v = row.get(f)
        v = " ".join(str(v).split()) if v is not None else ""
        out[f] = v or None
    for f in PRICE_FIELDS:
        v = row.get(f)
        try:
            v = float(v) if v not in (None, "") else None
        except (TypeError, ValueError):
            v = None
        out[f] = v if (v is not None and v > 0) else None
    try:
        m = row.get("multiplicity")
        out["multiplicity"] = int(float(m)) if m not in (None, "") else None
    except (TypeError, ValueError):
        out["multiplicity"] = None
    try:
        n = row.get("num")
        out["num"] = int(float(n)) if n not in (None, "") else None
    except (TypeError, ValueError):
        out["num"] = None
    out["unit"] = out["unit"] or "шт."
    return out


async def upsert_products(
    db: AsyncSession,
    segment: str,
    rows: List[Dict],
    changed_by: str = "",
    dry_run: bool = False,
) -> Dict:
    """Добавляет новые позиции и обновляет цены существующих.

    Возвращает {"stats": {...}, "results": [{"index", "status", "id",
    "changes": {поле: [было, стало]}, "message"}]}.

    Статусы строк:
      new         — будет/была создана;
      updated     — цены изменятся/изменены;
      restored    — была погашена прежним импортом, возвращается;
      same        — позиция есть, цены те же;
      deleted     — лежит в корзине, пропущена;
      duplicate   — повтор ключа внутри этих же данных, пропущена;
      error       — нет ни артикула, ни наименования.

    dry_run=True — ничего не пишет, только говорит, что будет.
    """
    now = datetime.now(timezone.utc)

    # Все позиции сегмента, включая неактивные: ключ → строка
    res = await db.execute(
        select(Product.id, Product.article, Product.name, Product.brand,
               Product.is_active, Product.deleted_at,
               Product.kaznisa, Product.rrts, Product.mrc, Product.opt,
               Product.partner)
        .where(Product.segment == segment)
    )
    existing: Dict[str, Dict] = {}
    for r in res.mappings().all():
        k = row_key(r["article"], r["name"])
        if not k:
            continue
        prev = existing.get(k)
        # Дубли ключа в базе уже бывают: предпочитаем живую позицию
        if prev is None or (not prev["is_active"] and r["is_active"]):
            existing[k] = dict(r)

    results: List[Dict] = []
    seen: set = set()
    stats = {k: 0 for k in ("new", "updated", "restored", "same",
                            "deleted", "duplicate", "error")}
    to_insert: List[Dict] = []

    for idx, raw in enumerate(rows):
        row = _clean(raw)
        key = row_key(row["article"], row["name"])
        if not key:
            stats["error"] += 1
            results.append({"index": idx, "status": "error",
                            "message": "нет ни артикула, ни наименования"})
            continue
        if key in seen:
            stats["duplicate"] += 1
            results.append({"index": idx, "status": "duplicate",
                            "message": "повтор в этих же данных — взята первая строка"})
            continue
        seen.add(key)

        cur = existing.get(key)
        if cur is None:
            stats["new"] += 1
            results.append({"index": idx, "status": "new"})
            to_insert.append(row)
            continue

        if cur["deleted_at"] is not None:
            stats["deleted"] += 1
            results.append({"index": idx, "status": "deleted", "id": cur["id"],
                            "message": "позиция в корзине — не изменена"})
            continue

        changes = {}
        for f in PRICE_FIELDS:
            new = row[f]
            old = cur[f]
            if new is None:
                continue                      # пустое не стирает
            if old is None or abs(float(old) - new) > 0.005:
                changes[f] = [old, new]

        reactivate = not cur["is_active"]
        status = ("restored" if reactivate
                  else "updated" if changes else "same")
        stats[status] += 1
        results.append({"index": idx, "status": status, "id": cur["id"],
                        "changes": changes})

        if dry_run or not (changes or reactivate):
            continue

        values = {f: v[1] for f, v in changes.items()}
        values["updated_by"] = changed_by or None
        values["updated_at"] = now
        if reactivate:
            values["is_active"] = True
        await db.execute(update(Product).where(Product.id == cur["id"]).values(**values))

        if changes:
            _record_price_change(db, segment, cur, changes, changed_by)

    if not dry_run and to_insert:
        created = [Product(segment=segment, is_active=True,
                           updated_by=changed_by or None, **row)
                   for row in to_insert]
        db.add_all(created)
        await db.flush()                 # после flush у объектов есть id
        it = iter(created)
        for r in results:
            if r["status"] == "new":
                r["id"] = next(it).id

    return {"stats": stats, "results": results}


def _record_price_change(db: AsyncSession, segment: str, cur: Dict,
                         changes: Dict, changed_by: str) -> None:
    """История цен (старые значения) и аномалии — как у импорта."""
    article = cur.get("article") or (cur.get("name") or "")[:60]
    db.add(PriceHistory(
        article=article, segment=segment, brand=cur.get("brand"),
        name=cur.get("name"), kaznisa=cur.get("kaznisa"), rrts=cur.get("rrts"),
        mrc=cur.get("mrc"), opt=cur.get("opt"),
    ))
    for f, (old, new) in changes.items():
        if not old or not new:
            continue
        pct = (new - float(old)) / float(old) * 100
        if abs(pct) >= ANOMALY_PCT:
            db.add(AuditLog(
                action="price_anomaly",
                resource=f"{segment}/{article}",
                username=changed_by or "system",
                details=json.dumps({
                    "brand": cur.get("brand") or "", "name": cur.get("name") or "",
                    "field": f, "old_price": round(float(old), 4),
                    "new_price": round(new, 4), "pct_change": round(pct, 1),
                }, ensure_ascii=False),
                status="warning",
            ))
