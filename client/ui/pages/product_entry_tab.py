"""
Вкладка «Ввод позиций» раздела «База данных».

Сетка в колонках листа «БД» шаблона: позиции можно набрать руками или
вставить блоком из Excel (Ctrl+V). «Проверить» показывает по каждой
строке, что произойдёт; «Сохранить в базу» записывает.

Правило записи — то же, что у импорта (сервер, product_store):
  • новая позиция создаётся со всеми полями;
  • у существующей меняются только цены, пустая ячейка цену не стирает;
  • позиции из корзины не меняются и не возвращаются.
"""
from __future__ import annotations

import threading
import tkinter as tk
from tkinter import messagebox, ttk

import customtkinter as ctk

from assets.theme import *  # noqa: F401,F403
from services.product_fields import (FIELDS, KEYS, KIND, SEGMENTS, SEGMENT_TITLE,
                                     TITLE, fmt_cell, parse_cell, parse_clipboard,
                                     rows_to_tsv)
from ui.pages.grid_cell_editor import CellEditor

# Результат проверки — сразу за номером строки, чтобы был виден без прокрутки;
# дальше колонки строго как в листе «БД».
COLS = ["row", "status"] + KEYS
START_ROWS = 20
KEY_C, KEY_V, KEY_A = 67, 86, 65

TAGS = {
    "new":       ROW_EXACT,
    "updated":   ROW_MULTIPLE,
    "restored":  ROW_EDITED,
    "deleted":   "#E5E7E9",
    "duplicate": "#FDEBD0",
    "error":     ROW_NOTFOUND,
}
PRICE_SHORT = {"kaznisa": "КазНИИСА", "rrts": "РРЦ", "mrc": "МРЦ",
               "opt": "Опт", "partner": "Партнёр"}


def _changes_text(changes: dict) -> str:
    parts = []
    for k, (old, new) in (changes or {}).items():
        o = fmt_cell(k, old) if old is not None else "—"
        parts.append(f"{PRICE_SHORT.get(k, k)} {o} → {fmt_cell(k, new)}")
    return "; ".join(parts)


def _status_text(r: dict, saved: bool) -> str:
    s = r.get("status")
    if s == "new":
        return "✓ Создана" if saved else "＋ Новая позиция"
    if s == "updated":
        return ("✓ Цены обновлены: " if saved else "Цены: ") + _changes_text(r.get("changes"))
    if s == "restored":
        return "✓ Возвращена в базу" if saved else "↺ Вернётся в базу (была отключена импортом)"
    if s == "same":
        return "Уже есть в базе, цены те же"
    if s == "deleted":
        return "В корзине — не изменится (восстановите во вкладке «Продукция»)"
    if s == "duplicate":
        return "Повтор строки выше — пропущена"
    return r.get("message") or "Ошибка"


class ProductEntryTab(ctk.CTkFrame):
    def __init__(self, parent, api, on_error, on_changed=None,
                 default_segment: str = "ss"):
        super().__init__(parent, fg_color="transparent")
        self.api = api
        self.on_error = on_error
        self.on_changed = on_changed or (lambda: None)

        self._data: dict = {}        # iid → {ключ: текст}
        self._status: dict = {}      # iid → (статус, текст)
        self._col = 1                # текущая колонка (индекс в KEYS): «Артикул»
        self._busy = False
        self._seq = 0

        self.grid_columnconfigure(0, weight=1)
        self.grid_rowconfigure(2, weight=1)
        seg = default_segment if default_segment in SEGMENT_TITLE else "ss"
        self._build_toolbar(SEGMENT_TITLE[seg])
        self._build_tree()
        self._build_bottom()
        self._add_rows(START_ROWS)
        self._set_headings()

    # ── Построение ───────────────────────────────────────────────────────────

    def _btn(self, parent, text, cmd, width=120, color=NAVY_LIGHT, hover=NAVY):
        return ctk.CTkButton(parent, text=text, command=cmd, width=width, height=32,
                             font=FONT_SMALL, fg_color=color, hover_color=hover,
                             corner_radius=RADIUS_SM)

    def _build_toolbar(self, seg_title: str):
        bar = ctk.CTkFrame(self, fg_color="transparent")
        bar.grid(row=0, column=0, sticky="ew", padx=12, pady=(10, 4))
        bar.grid_columnconfigure(6, weight=1)

        ctk.CTkLabel(bar, text="Сегмент базы:", font=FONT_SMALL,
                     text_color=TEXT_PRIMARY).grid(row=0, column=0, padx=(0, 6))
        self._seg_var = ctk.StringVar(value=seg_title)
        ctk.CTkOptionMenu(
            bar, values=[t for _c, t in SEGMENTS], variable=self._seg_var, width=160,
            font=FONT_SMALL, fg_color=NAVY_LIGHT, button_color=NAVY,
            command=lambda _v: self._invalidate_all(),
        ).grid(row=0, column=1, padx=(0, 12))

        self._btn(bar, "📋 Вставить из Excel", self._paste, 150).grid(row=0, column=2, padx=(0, 6))
        self._btn(bar, "＋ 10 строк", lambda: self._add_rows(10), 95).grid(row=0, column=3, padx=(0, 6))
        self._btn(bar, "✕ Удалить строки", self._delete_rows, 130).grid(row=0, column=4, padx=(0, 6))
        self._btn(bar, "Очистить", self._clear, 90, "#7F8C8D", "#626567").grid(row=0, column=5)

        self._check_btn = self._btn(bar, "🔍 Проверить", self._check, 120)
        self._check_btn.grid(row=0, column=7, padx=(0, 6))
        self._save_btn = self._btn(bar, "💾 Сохранить в базу", self._save, 160,
                                   "#27AE60", "#1E8449")
        self._save_btn.grid(row=0, column=8)

        ctk.CTkLabel(
            self, font=FONT_SMALL, text_color=TEXT_SECONDARY, anchor="w", justify="left",
            text=("Колонки — как в листе «БД» шаблона: блок, скопированный из Excel, "
                  "вставляется в выбранную ячейку (с шапкой — по заголовкам). "
                  "Двойной щелчок или ввод с клавиатуры — правка, Enter — вниз, Tab — вправо, "
                  "Del — очистить ячейку.\n"
                  "Новые позиции создаются целиком. У позиций, которые уже есть в сегменте, "
                  "меняются только цены — пустая цена старую не стирает; наименование, бренд "
                  "и код правятся во вкладке «Продукция». Позиции из корзины не меняются."),
        ).grid(row=1, column=0, sticky="ew", padx=14, pady=(0, 4))

    def _build_tree(self):
        wrap = tk.Frame(self, bg=BG_CARD)
        wrap.grid(row=2, column=0, sticky="nsew", padx=12)
        wrap.grid_rowconfigure(0, weight=1)
        wrap.grid_columnconfigure(0, weight=1)

        st = ttk.Style()
        st.configure("Entry.Treeview", rowheight=24, font=("Calibri", 11))
        st.configure("Entry.Treeview.Heading", background=NAVY, foreground="white",
                     font=("Calibri", 11, "bold"))
        st.map("Entry.Treeview", background=[("selected", BLUE_PALE)],
               foreground=[("selected", TEXT_PRIMARY)])

        self.tree = ttk.Treeview(wrap, columns=COLS, show="headings",
                                 style="Entry.Treeview", selectmode="extended")
        widths = {"row": 40, "status": 280}
        widths.update({k: w for k, _t, w, _k in FIELDS})
        for c in COLS:
            self.tree.column(c, width=widths[c], minwidth=34,
                             stretch=(c == "name"),
                             anchor=("e" if KIND.get(c) == "price"
                                     else "center" if c == "row" else "w"))
        for tag, color in TAGS.items():
            self.tree.tag_configure(tag, background=color)

        vsb = ttk.Scrollbar(wrap, orient="vertical", command=self.tree.yview)
        hsb = ttk.Scrollbar(wrap, orient="horizontal", command=self.tree.xview)
        self.tree.configure(yscrollcommand=lambda *a: (vsb.set(*a), self._editor.commit()),
                            xscrollcommand=lambda *a: (hsb.set(*a), self._editor.commit()))
        self.tree.grid(row=0, column=0, sticky="nsew")
        vsb.grid(row=0, column=1, sticky="ns")
        hsb.grid(row=1, column=0, sticky="ew")

        self._editor = CellEditor(self.tree, font=("Calibri", 11))
        tr = self.tree
        tr.bind("<Button-1>", self._on_click, add="+")
        tr.bind("<Double-1>", self._on_double)
        tr.bind("<Return>", lambda _e: self._edit_current())
        tr.bind("<F2>", lambda _e: self._edit_current())
        tr.bind("<Delete>", lambda _e: self._clear_cells())
        tr.bind("<Left>", lambda _e: self._shift_col(-1))
        tr.bind("<Right>", lambda _e: self._shift_col(1))
        tr.bind("<Tab>", lambda _e: self._shift_col(1))
        tr.bind("<Control-KeyPress>", self._on_ctrl)
        tr.bind("<KeyPress>", self._on_key)
        tr.bind("<Button-3>", self._on_menu)
        tr.bind("<<TreeviewSelect>>", lambda _e: self._show_pos())
        self._menu = tk.Menu(tr, tearoff=0)

    def _build_bottom(self):
        bot = ctk.CTkFrame(self, fg_color="transparent")
        bot.grid(row=3, column=0, sticky="ew", padx=12, pady=(6, 10))
        bot.grid_columnconfigure(1, weight=1)
        self._pos_lbl = ctk.CTkLabel(bot, text="", font=FONT_SMALL,
                                     text_color=TEXT_SECONDARY, width=200, anchor="w")
        self._pos_lbl.grid(row=0, column=0, sticky="w")
        self._sum_lbl = ctk.CTkLabel(bot, text="", font=FONT_SMALL,
                                     text_color=TEXT_PRIMARY, anchor="e")
        self._sum_lbl.grid(row=0, column=1, sticky="e")

    def _set_headings(self):
        for c in COLS:
            if c == "row":
                text = "стр."
            elif c == "status":
                text = "Результат проверки"
            else:
                text = ("● " if KEYS.index(c) == self._col else "") + TITLE[c]
            self.tree.heading(c, text=text)

    # ── Строки ───────────────────────────────────────────────────────────────

    def _segment(self) -> str:
        for c, t in SEGMENTS:
            if t == self._seg_var.get():
                return c
        return "ss"

    def _new_iid(self) -> str:
        self._seq += 1
        return f"r{self._seq}"

    def _add_rows(self, n: int):
        start = len(self.tree.get_children())
        for i in range(n):
            iid = self._new_iid()
            self._data[iid] = {}
            self.tree.insert("", "end", iid=iid, values=[start + i + 1] + [""] * (len(COLS) - 1))

    def _render(self, iid: str):
        d = self._data.get(iid, {})
        st, text = self._status.get(iid, ("", ""))
        n = self.tree.index(iid) + 1
        self.tree.item(iid, values=[n, text] + [d.get(k, "") for k in KEYS],
                       tags=(st,) if st in TAGS else ())

    def _renumber(self):
        for n, iid in enumerate(self.tree.get_children(), 1):
            self.tree.set(iid, "row", n)

    def _is_empty(self, iid: str) -> bool:
        return not any((v or "").strip() for v in self._data.get(iid, {}).values())

    def _touch(self, iid: str):
        """Строку поменяли — прежний результат проверки больше не верен."""
        self._status.pop(iid, None)
        self._sum_lbl.configure(text="")

    def _invalidate_all(self):
        for iid in list(self._status):
            self._status.pop(iid)
            if self.tree.exists(iid):
                self._render(iid)
        self._sum_lbl.configure(text="")

    def _delete_rows(self):
        sel = self.tree.selection()
        if not sel:
            messagebox.showinfo("Удалить строки", "Выделите строки в таблице.", parent=self)
            return
        for iid in sel:
            self._data.pop(iid, None)
            self._status.pop(iid, None)
            self.tree.delete(iid)
        if not self.tree.get_children():
            self._add_rows(START_ROWS)
        self._renumber()

    def _clear(self):
        if any(not self._is_empty(i) for i in self._data) and not messagebox.askyesno(
                "Очистить", "Очистить всю таблицу ввода?", parent=self):
            return
        self.tree.delete(*self.tree.get_children())
        self._data.clear()
        self._status.clear()
        self._sum_lbl.configure(text="")
        self._add_rows(START_ROWS)

    # ── Навигация и правка ───────────────────────────────────────────────────

    def _current_row(self) -> str:
        iid = self.tree.focus()
        if not iid:
            kids = self.tree.get_children()
            iid = kids[0] if kids else ""
        return iid

    def _show_pos(self):
        iid = self.tree.focus()
        if iid and self.tree.exists(iid):
            # Полный текст проверки — в колонке он может не уместиться
            st = self._status.get(iid, ("", ""))[1]
            self._pos_lbl.configure(
                text=f"Строка {self.tree.index(iid) + 1}, «{TITLE[KEYS[self._col]]}»"
                     + (f"  —  {st}" if st else ""))

    def _set_col(self, i: int):
        self._col = max(0, min(len(KEYS) - 1, i))
        self._set_headings()
        self._show_pos()

    def _shift_col(self, step: int):
        self._set_col(self._col + step)
        return "break"

    def _on_click(self, event):
        if self.tree.identify_region(event.x, event.y) != "cell":
            return
        col = COLS[int(self.tree.identify_column(event.x)[1:]) - 1]
        if col in KEYS:
            self._set_col(KEYS.index(col))

    def _on_double(self, event):
        if self.tree.identify_region(event.x, event.y) != "cell":
            return
        iid = self.tree.identify_row(event.y)
        if iid and COLS[int(self.tree.identify_column(event.x)[1:]) - 1] in KEYS:
            self._edit(iid, self._col)

    def _edit_current(self):
        iid = self._current_row()
        if iid:
            self._edit(iid, self._col)
        return "break"

    def _on_key(self, event):
        ch = event.char
        if ch and len(ch) == 1 and ch.isprintable() and not (event.state & 0x4):
            iid = self._current_row()
            if iid:
                self._edit(iid, self._col, typed=ch)
            return "break"
        return None

    def _edit(self, iid: str, col: int, typed=None):
        key = KEYS[col]
        self.tree.selection_set(iid)
        self.tree.focus(iid)
        self._set_col(col)
        self._editor.open(iid, f"#{COLS.index(key) + 1}",
                          self._data.get(iid, {}).get(key, ""),
                          lambda text, move, iid=iid, col=col: self._commit(iid, col, text, move),
                          typed=typed)

    def _commit(self, iid: str, col: int, text: str, move: str):
        if not self.tree.exists(iid):
            return
        key = KEYS[col]
        text = " ".join((text or "").split()) if KIND[key] == "text" else (text or "").strip()
        if self._data[iid].get(key, "") != text:
            self._data[iid][key] = text
            self._touch(iid)
            self._render(iid)
        if move == "down":
            nxt = self.tree.next(iid)
            if not nxt:
                self._add_rows(5)
                nxt = self.tree.next(iid)
            self.tree.selection_set(nxt)
            self.tree.focus(nxt)
            self.tree.see(nxt)
            self._show_pos()
        elif move in ("right", "left"):
            c = col + (1 if move == "right" else -1)
            if 0 <= c < len(KEYS):
                self._edit(iid, c)

    def _clear_cells(self):
        key = KEYS[self._col]
        for iid in self.tree.selection():
            if self._data.get(iid, {}).get(key):
                self._data[iid][key] = ""
                self._touch(iid)
                self._render(iid)
        return "break"

    # ── Буфер обмена ─────────────────────────────────────────────────────────

    def _on_ctrl(self, event):
        if event.keycode == KEY_V:
            self._paste()
            return "break"
        if event.keycode == KEY_C:
            self._copy()
            return "break"
        if event.keycode == KEY_A:
            self.tree.selection_set(self.tree.get_children())
            return "break"
        return None

    def _paste(self):
        self._editor.commit()
        try:
            text = self.clipboard_get()
        except tk.TclError:
            messagebox.showinfo("Вставка", "В буфере обмена нет текста. "
                                "Скопируйте ячейки в Excel и повторите.", parent=self)
            return
        rows, how = parse_clipboard(text, start_col=self._col)
        if not rows:
            return
        iid = self.tree.focus()
        if not iid or not self.tree.exists(iid):
            iid = next((i for i in self.tree.get_children() if self._is_empty(i)), "")
        if not iid:
            self._add_rows(1)
            iid = self.tree.get_children()[-1]
        kids = list(self.tree.get_children())
        pos = kids.index(iid)
        need = pos + len(rows) - len(kids)
        if need > 0:
            self._add_rows(need + 5)
            kids = list(self.tree.get_children())
        for r, target in zip(rows, kids[pos:pos + len(rows)]):
            d = self._data[target]
            for k, v in r.items():
                d[k] = " ".join(v.split()) if KIND[k] == "text" else v
            self._touch(target)
            self._render(target)
        self.tree.selection_set(kids[pos])
        self.tree.focus(kids[pos])
        self.tree.see(kids[pos])
        self._pos_lbl.configure(text=f"Вставлено строк: {len(rows)} ({how})")

    def _copy(self):
        sel = [i for i in self.tree.selection() if not self._is_empty(i)]
        if not sel:
            return
        rows = [{k: self._data[i].get(k) or None for k in KEYS} for i in sel]
        # Текстом, как набрано — без преобразования цен
        text = "\n".join(["\t".join(TITLE[k] for k in KEYS)] +
                         ["\t".join(r[k] or "" for k in KEYS) for r in rows])
        self.clipboard_clear()
        self.clipboard_append(text)
        self._pos_lbl.configure(text=f"Скопировано строк: {len(sel)}")

    def _on_menu(self, event):
        iid = self.tree.identify_row(event.y)
        if iid:
            if iid not in self.tree.selection():
                self.tree.selection_set(iid)
            self.tree.focus(iid)
            col = COLS[int(self.tree.identify_column(event.x)[1:]) - 1]
            if col in KEYS:
                self._set_col(KEYS.index(col))
        m = self._menu
        m.delete(0, "end")
        m.add_command(label="📋 Вставить из Excel (Ctrl+V)", command=self._paste)
        m.add_command(label="Копировать строки (Ctrl+C)", command=self._copy)
        m.add_command(label="Очистить ячейки (Del)", command=self._clear_cells)
        m.add_separator()
        m.add_command(label="✕ Удалить строки", command=self._delete_rows)
        m.tk_popup(event.x_root, event.y_root)

    # ── Проверка и запись ────────────────────────────────────────────────────

    def _collect(self):
        """Строки для сервера. Ошибки разбора помечаются на месте и не уходят."""
        payload, iids, errors = [], [], 0
        for iid in self.tree.get_children():
            if self._is_empty(iid):
                self._status.pop(iid, None)
                self._render(iid)
                continue
            d, row, problems = self._data[iid], {}, []
            for k in KEYS:
                try:
                    row[k] = parse_cell(k, d.get(k))
                except ValueError as e:
                    problems.append(str(e))
            if not problems and not (row.get("article") or row.get("name")):
                problems.append("нет ни артикула, ни наименования")
            if problems:
                errors += 1
                self._status[iid] = ("error", "⚠ " + "; ".join(problems))
                self._render(iid)
                continue
            payload.append(row)
            iids.append(iid)
        return payload, iids, errors

    def _set_busy(self, busy: bool):
        self._busy = busy
        st = "disabled" if busy else "normal"
        self._check_btn.configure(state=st)
        self._save_btn.configure(state=st)

    def _check(self, then=None):
        if self._busy:
            return
        self._editor.commit()
        payload, iids, errors = self._collect()
        if not payload:
            messagebox.showinfo(
                "Проверка",
                "Нет строк для записи." + (f"\nСтрок с ошибками: {errors} — они отмечены красным."
                                           if errors else ""),
                parent=self)
            return
        seg = self._segment()
        self._set_busy(True)
        self._sum_lbl.configure(text="Проверка…")

        def work():
            try:
                res = self.api.upsert_products(seg, payload, dry_run=True)
            except Exception as e:
                self.after(0, lambda e=e: self._failed(e, "Проверка"))
                return
            self.after(0, lambda: self._checked(res, iids, errors, seg, then))
        threading.Thread(target=work, daemon=True).start()

    def _failed(self, exc, ctx):
        self._set_busy(False)
        self._sum_lbl.configure(text="")
        self.on_error(exc, ctx)

    def _apply_results(self, res: dict, iids: list, saved: bool):
        for r in res.get("results", []):
            i = r.get("index")
            if i is None or i >= len(iids):
                continue
            iid = iids[i]
            if self.tree.exists(iid):
                self._status[iid] = (r.get("status", ""), _status_text(r, saved))
                self._render(iid)

    def _summary(self, st: dict, errors: int, saved: bool) -> str:
        parts = [
            (("Создано" if saved else "Новых"), st.get("new", 0)),
            (("Цены обновлены" if saved else "Цены изменятся"), st.get("updated", 0)),
            (("Возвращено" if saved else "Вернутся"), st.get("restored", 0)),
            ("Без изменений", st.get("same", 0)),
            ("В корзине", st.get("deleted", 0)),
            ("Повторы", st.get("duplicate", 0)),
            ("Ошибки", errors + st.get("error", 0)),
        ]
        return "  ·  ".join(f"{n}: {v}" for n, v in parts if v)

    def _checked(self, res, iids, errors, seg, then):
        self._set_busy(False)
        self.tree.selection_set(())         # иначе выделение закрывает цвета строк
        self._apply_results(res, iids, saved=False)
        st = res.get("stats", {})
        self._sum_lbl.configure(text=self._summary(st, errors, saved=False) or "—")
        if then:
            then(st, errors, seg, iids)

    def _save(self):
        self._check(then=self._confirm_save)

    def _confirm_save(self, st, errors, seg, _iids):
        work = st.get("new", 0) + st.get("updated", 0) + st.get("restored", 0)
        if not work:
            messagebox.showinfo("Сохранение",
                                "Записывать нечего: все позиции уже есть в базе с теми же "
                                "ценами или пропущены. Подробности — в колонке "
                                "«Результат проверки».", parent=self)
            return
        lines = [f"Сегмент: {SEGMENT_TITLE[seg]}", ""]
        if st.get("new"):
            lines.append(f"• будет создано позиций: {st['new']}")
        if st.get("updated"):
            lines.append(f"• изменятся цены у позиций: {st['updated']}")
        if st.get("restored"):
            lines.append(f"• вернутся в базу: {st['restored']}")
        skipped = st.get("deleted", 0) + st.get("duplicate", 0) + errors
        if skipped:
            lines.append(f"• пропущено (корзина, повторы, ошибки): {skipped}")
        lines += ["", "Сохранить?"]
        if not messagebox.askyesno("Сохранить в базу", "\n".join(lines), parent=self):
            return

        payload, iids, errors = self._collect()
        self._set_busy(True)
        self._sum_lbl.configure(text="Сохранение…")

        def run():
            try:
                res = self.api.upsert_products(seg, payload, dry_run=False)
            except Exception as e:
                self.after(0, lambda e=e: self._failed(e, "Сохранение позиций"))
                return
            self.after(0, lambda: self._saved(res, iids, errors))
        threading.Thread(target=run, daemon=True).start()

    def _saved(self, res, iids, errors):
        self._set_busy(False)
        self.tree.selection_set(())
        self._apply_results(res, iids, saved=True)
        st = res.get("stats", {})
        self._sum_lbl.configure(text="✓ " + self._summary(st, errors, saved=True))
        self.on_changed()
        done = [iid for iid in iids if self._status.get(iid, ("",))[0]
                in ("new", "updated", "restored", "same")]
        # Записанное уже в базе — из таблицы ввода убираем без вопросов.
        # Остаются только строки, которые требуют внимания.
        left = sum(1 for i in self.tree.get_children()
                   if i not in done and not self._is_empty(i))
        msg = "Позиции сохранены в базу.\n\n" + \
            self._summary(st, errors, saved=True).replace("  ·  ", "\n")
        if left:
            msg += (f"\n\nВ таблице осталось строк: {left} — они не записаны, "
                    "причина указана в колонке «Результат проверки».")
        messagebox.showinfo("Сохранено", msg, parent=self)
        if done:
            for iid in done:
                self._data.pop(iid, None)
                self._status.pop(iid, None)
                self.tree.delete(iid)
            if len(self.tree.get_children()) < START_ROWS:
                self._add_rows(START_ROWS - len(self.tree.get_children()))
            self._renumber()
