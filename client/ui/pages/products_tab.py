"""
Вкладка «Продукция» раздела «База данных».

Просмотр позиций базы с поиском, фильтрами и сортировкой; правка ячейки
двойным щелчком; удаление в корзину и возврат из неё. Колонки — как в
листе «БД» шаблона, плюс сегмент и кто/когда менял.

Все запросы к серверу идут в фоне; ответ на устаревший запрос (менеджер
уже сменил фильтр) отбрасывается.
"""
from __future__ import annotations

import threading
import tkinter as tk
from datetime import datetime
from tkinter import messagebox, ttk

import customtkinter as ctk

from assets.theme import *  # noqa: F401,F403
from services.product_fields import (FIELDS, KEYS, KIND, SEGMENTS, SEGMENT_TITLE,
                                     TITLE, fmt_cell, parse_cell, rows_to_tsv)
from ui.pages.grid_cell_editor import CellEditor

ALL_LABEL = "Все сегменты"
STATUS_ACTIVE = "Активные"
STATUS_TRASH = "Корзина"
PAGE_SIZES = ["100", "200", "500"]

COLS = ["segment"] + KEYS + ["changed_at", "changed_by"]
EDITABLE = set(KEYS)
SHORT_SEG = {"ss": "СС", "os": "ОС", "sil": "Сил", "gen": "Общая"}

KEY_C, KEY_V = 67, 86


def _fmt_dt(s) -> str:
    if not s:
        return ""
    try:
        return datetime.fromisoformat(str(s)).astimezone().strftime("%d.%m.%Y %H:%M")
    except ValueError:
        return str(s)[:16]


class ProductsTab(ctk.CTkFrame):
    def __init__(self, parent, api, on_error, on_changed=None,
                 default_segment: str = ""):
        super().__init__(parent, fg_color="transparent")
        self.api = api
        self.on_error = on_error
        self.on_changed = on_changed or (lambda: None)

        self._items: dict = {}           # id → позиция с сервера
        self._total = 0
        self._offset = 0
        self._sort, self._desc = "article", False
        self._req = 0                    # номер последнего запроса списка
        self._search_job = None
        self._loaded = False

        self._seg_labels = [ALL_LABEL] + [t for _c, t in SEGMENTS]
        self._seg_codes = [""] + [c for c, _t in SEGMENTS]
        start = SEGMENT_TITLE.get(default_segment, ALL_LABEL)

        self.grid_columnconfigure(0, weight=1)
        self.grid_rowconfigure(2, weight=1)
        self._build_toolbar(start)
        self._build_tree()
        self._build_bottom()

    # ── Построение ───────────────────────────────────────────────────────────

    def _build_toolbar(self, start_seg: str):
        bar = ctk.CTkFrame(self, fg_color="transparent")
        bar.grid(row=0, column=0, sticky="ew", padx=12, pady=(10, 4))
        bar.grid_columnconfigure(3, weight=1)

        self._seg_var = ctk.StringVar(value=start_seg)
        ctk.CTkOptionMenu(
            bar, values=self._seg_labels, variable=self._seg_var, width=170,
            font=FONT_SMALL, fg_color=NAVY_LIGHT, button_color=NAVY,
            command=lambda _v: self._on_segment(),
        ).grid(row=0, column=0, padx=(0, 8))

        self._status_var = ctk.StringVar(value=STATUS_ACTIVE)
        ctk.CTkSegmentedButton(
            bar, values=[STATUS_ACTIVE, STATUS_TRASH], variable=self._status_var,
            font=FONT_SMALL, selected_color=NAVY_LIGHT, unselected_color=BLUE_MID,
            command=lambda _v: self._on_status(),
        ).grid(row=0, column=1, padx=(0, 8))

        self._brand_var = ctk.StringVar(value="Все бренды")
        self._brand_menu = ctk.CTkOptionMenu(
            bar, values=["Все бренды"], variable=self._brand_var, width=170,
            font=FONT_SMALL, fg_color=NAVY_LIGHT, button_color=NAVY,
            dynamic_resizing=False, command=lambda _v: self.reload(reset=True),
        )
        self._brand_menu.grid(row=0, column=2, padx=(0, 8))

        self._q_var = ctk.StringVar()
        q = ctk.CTkEntry(bar, textvariable=self._q_var, font=FONT_NORMAL, height=32,
                         placeholder_text="🔍  Поиск: артикул, наименование, код КазНИИСА, бренд")
        q.grid(row=0, column=3, sticky="ew", padx=(0, 8))
        self._q_var.trace_add("write", lambda *_: self._on_search())
        q.bind("<Return>", lambda _e: self.reload(reset=True))

        ctk.CTkButton(
            bar, text="⟳ Обновить", font=FONT_SMALL, width=100, height=32,
            fg_color=NAVY_LIGHT, hover_color=NAVY, corner_radius=RADIUS_SM,
            command=lambda: (self._load_brands(), self.reload()),
        ).grid(row=0, column=4)

        self._hint = ctk.CTkLabel(
            self, font=FONT_SMALL, text_color=TEXT_SECONDARY, anchor="w",
            text="Двойной щелчок по ячейке — правка (Enter — сохранить, Esc — отмена). "
                 "Del — в корзину, Ctrl+C — копировать строки для Excel. "
                 "Щелчок по заголовку — сортировка.")
        self._hint.grid(row=1, column=0, sticky="ew", padx=14, pady=(0, 4))

    def _build_tree(self):
        wrap = tk.Frame(self, bg=BG_CARD)
        wrap.grid(row=2, column=0, sticky="nsew", padx=12)
        wrap.grid_rowconfigure(0, weight=1)
        wrap.grid_columnconfigure(0, weight=1)

        st = ttk.Style()
        st.configure("Prod.Treeview", rowheight=24, font=("Calibri", 11))
        st.configure("Prod.Treeview.Heading", background=NAVY, foreground="white",
                     font=("Calibri", 11, "bold"))
        st.map("Prod.Treeview", background=[("selected", NAVY_LIGHT)],
               foreground=[("selected", "white")])

        self.tree = ttk.Treeview(wrap, columns=COLS, show="headings",
                                 style="Prod.Treeview", selectmode="extended")
        widths = {"segment": 60, "changed_at": 115, "changed_by": 130}
        widths.update({k: w for k, _t, w, _k in FIELDS})
        for c in COLS:
            self.tree.column(c, width=widths[c], minwidth=40,
                             stretch=(c == "name"),
                             anchor=("e" if KIND.get(c) == "price" else "w"))
        self._set_headings()

        vsb = ttk.Scrollbar(wrap, orient="vertical", command=self.tree.yview)
        hsb = ttk.Scrollbar(wrap, orient="horizontal", command=self.tree.xview)
        self.tree.configure(yscrollcommand=lambda *a: (vsb.set(*a), self._editor.commit()),
                            xscrollcommand=lambda *a: (hsb.set(*a), self._editor.commit()))
        self.tree.grid(row=0, column=0, sticky="nsew")
        vsb.grid(row=0, column=1, sticky="ns")
        hsb.grid(row=1, column=0, sticky="ew")
        self.tree.tag_configure("odd", background="#F7F9FC")

        self._empty = tk.Label(wrap, text="", bg="white", fg=TEXT_SECONDARY,
                               font=("Calibri", 12))

        self._editor = CellEditor(self.tree, font=("Calibri", 11))
        self.tree.bind("<Double-1>", self._on_double)
        self.tree.bind("<Return>", lambda _e: self._edit_focused())
        self.tree.bind("<Delete>", lambda _e: self._delete_selected())
        self.tree.bind("<Control-KeyPress>", self._on_ctrl)
        self.tree.bind("<Button-3>", self._on_menu)

        self._menu = tk.Menu(self.tree, tearoff=0)

    def _build_bottom(self):
        bot = ctk.CTkFrame(self, fg_color="transparent")
        bot.grid(row=3, column=0, sticky="ew", padx=12, pady=(6, 10))
        bot.grid_columnconfigure(3, weight=1)

        self._prev_btn = ctk.CTkButton(bot, text="◀", width=36, height=30,
                                       fg_color=NAVY_LIGHT, hover_color=NAVY,
                                       command=lambda: self._page(-1))
        self._prev_btn.grid(row=0, column=0)
        self._page_lbl = ctk.CTkLabel(bot, text="", font=FONT_SMALL,
                                      text_color=TEXT_PRIMARY, width=170)
        self._page_lbl.grid(row=0, column=1, padx=6)
        self._next_btn = ctk.CTkButton(bot, text="▶", width=36, height=30,
                                       fg_color=NAVY_LIGHT, hover_color=NAVY,
                                       command=lambda: self._page(1))
        self._next_btn.grid(row=0, column=2)

        self._size_var = ctk.StringVar(value=PAGE_SIZES[0])
        ctk.CTkOptionMenu(bot, values=PAGE_SIZES, variable=self._size_var, width=80,
                          font=FONT_SMALL, fg_color=NAVY_LIGHT, button_color=NAVY,
                          command=lambda _v: self.reload(reset=True),
                          ).grid(row=0, column=3, sticky="w", padx=(10, 0))

        self._info_lbl = ctk.CTkLabel(bot, text="", font=FONT_SMALL,
                                      text_color=TEXT_SECONDARY)
        self._info_lbl.grid(row=0, column=4, padx=10)

        self._copy_btn = ctk.CTkButton(
            bot, text="📋 Копировать", width=120, height=30, font=FONT_SMALL,
            fg_color=NAVY_LIGHT, hover_color=NAVY, command=self._copy_selected)
        self._copy_btn.grid(row=0, column=5, padx=(0, 8))
        self._action_btn = ctk.CTkButton(
            bot, text="🗑 В корзину", width=140, height=30, font=FONT_SMALL,
            fg_color="#C0392B", hover_color="#922B21", command=self._action)
        self._action_btn.grid(row=0, column=6)

    def _set_headings(self):
        trash = self._trash()
        names = {"segment": "Сегмент",
                 "changed_at": "Удалено" if trash else "Изменено",
                 "changed_by": "Кем удалено" if trash else "Кем"}
        names.update(TITLE)
        for c in COLS:
            text = names[c]
            if self._sort_key(c) == self._sort:
                text += " ▼" if self._desc else " ▲"
            cmd = (lambda c=c: self._on_heading(c)) if self._sort_key(c) else ""
            self.tree.heading(c, text=text, command=cmd)

    # ── Состояние фильтров ───────────────────────────────────────────────────

    def _trash(self) -> bool:
        return getattr(self, "_status_var", None) is not None and \
            self._status_var.get() == STATUS_TRASH

    def _segment(self) -> str:
        try:
            return self._seg_codes[self._seg_labels.index(self._seg_var.get())]
        except ValueError:
            return ""

    def _sort_key(self, col: str) -> str:
        if col == "changed_at":
            return "deleted_at" if self._trash() else "updated_at"
        if col == "changed_by":
            return ""
        return col

    def _brand(self) -> str:
        b = self._brand_var.get()
        return "" if b == "Все бренды" else b

    # ── Загрузка ─────────────────────────────────────────────────────────────

    def on_show(self):
        """Вкладку открыли: первый раз — загрузить."""
        if not self._loaded:
            self._loaded = True
            self._load_brands()
            self.reload(reset=True)

    def mark_dirty(self):
        """База изменилась в другой вкладке — перечитать при следующем показе."""
        self._loaded = False

    def _on_segment(self):
        self._brand_var.set("Все бренды")
        self._load_brands()
        self.reload(reset=True)

    def _on_status(self):
        if self._trash() and self._sort == "updated_at":
            self._sort = "deleted_at"
        elif not self._trash() and self._sort == "deleted_at":
            self._sort = "updated_at"
        self._action_btn.configure(
            text="↩ Восстановить" if self._trash() else "🗑 В корзину",
            fg_color="#27AE60" if self._trash() else "#C0392B",
            hover_color="#1E8449" if self._trash() else "#922B21")
        self._set_headings()
        self.reload(reset=True)

    def _on_search(self):
        if self._search_job:
            self.after_cancel(self._search_job)
        self._search_job = self.after(450, lambda: self.reload(reset=True))

    def _on_heading(self, col: str):
        key = self._sort_key(col)
        if key == self._sort:
            self._desc = not self._desc
        else:
            self._sort, self._desc = key, key in ("updated_at", "deleted_at")
        self._set_headings()
        self.reload(reset=True)

    def _page(self, step: int):
        size = int(self._size_var.get())
        new = self._offset + step * size
        if 0 <= new < max(self._total, 1):
            self._offset = new
            self.reload()

    def _load_brands(self):
        seg = self._segment()

        def work():
            try:
                brands = self.api.product_brands(seg)
            except Exception:
                return
            self.after(0, lambda: self._brand_menu.configure(
                values=["Все бренды"] + list(brands)))
        threading.Thread(target=work, daemon=True).start()

    def reload(self, reset: bool = False):
        self._search_job = None
        self._editor.commit()
        if reset:
            self._offset = 0
        self._req += 1
        req = self._req
        params = dict(segment=self._segment(),
                      status="deleted" if self._trash() else "active",
                      brand=self._brand(), q=self._q_var.get().strip(),
                      sort=self._sort, desc=self._desc,
                      limit=int(self._size_var.get()), offset=self._offset)
        self._info_lbl.configure(text="Загрузка…")

        def work():
            try:
                data = self.api.browse_products(**params)
            except Exception as e:
                self.after(0, lambda e=e: self._load_failed(req, e))
                return
            self.after(0, lambda: self._show(req, data))
        threading.Thread(target=work, daemon=True).start()

    def _load_failed(self, req, exc):
        if req != self._req:
            return
        self._info_lbl.configure(text="")
        self.on_error(exc, "Продукция")

    def _show(self, req: int, data: dict):
        if req != self._req:
            return                       # пока шёл ответ, фильтр сменили
        items = data.get("items", [])
        self._total = int(data.get("total", 0))
        self.tree.delete(*self.tree.get_children())
        self._items = {}
        for i, p in enumerate(items):
            self._items[str(p["id"])] = p
            self.tree.insert("", "end", iid=str(p["id"]), values=self._row_values(p),
                             tags=("odd",) if i % 2 else ())
        size = int(self._size_var.get())
        if self._total:
            a, b = self._offset + 1, self._offset + len(items)
            self._page_lbl.configure(text=f"{a:,}–{b:,} из {self._total:,}".replace(",", " "))
            self._empty.place_forget()
        else:
            self._page_lbl.configure(text="0 из 0")
            q = self._q_var.get().strip()
            self._empty.configure(
                text=("Ничего не найдено — измените запрос или фильтры" if q or self._brand()
                      else "Корзина пуста" if self._trash() else "В этом сегменте нет позиций"))
            self._empty.place(relx=0.5, rely=0.3, anchor="center")
        self._prev_btn.configure(state="normal" if self._offset > 0 else "disabled")
        self._next_btn.configure(
            state="normal" if self._offset + size < self._total else "disabled")
        self._info_lbl.configure(text="")

    def _row_values(self, p: dict) -> list:
        trash = self._trash()
        vals = [SHORT_SEG.get(p.get("segment"), p.get("segment") or "")]
        vals += [fmt_cell(k, p.get(k)) for k in KEYS]
        vals.append(_fmt_dt(p.get("deleted_at" if trash else "updated_at")))
        vals.append((p.get("deleted_by" if trash else "updated_by")) or "")
        return vals

    # ── Правка ячейки ────────────────────────────────────────────────────────

    def _on_double(self, event):
        if self.tree.identify_region(event.x, event.y) != "cell":
            return
        iid = self.tree.identify_row(event.y)
        col = self.tree.identify_column(event.x)
        if iid:
            self._edit(iid, col)

    def _edit_focused(self):
        iid = self.tree.focus()
        if iid:
            self._edit(iid, f"#{COLS.index('article') + 1}")
        return "break"

    def _edit(self, iid: str, col_id: str):
        key = COLS[int(col_id[1:]) - 1]
        if key not in EDITABLE:
            return
        if self._trash():
            messagebox.showinfo("Корзина",
                                "Позиция в корзине. Чтобы её править, сначала восстановите.",
                                parent=self)
            return
        p = self._items.get(iid)
        if p is None:
            return
        text = "" if p.get(key) is None else (
            f"{p[key]:.2f}".rstrip("0").rstrip(".") if KIND[key] == "price" else str(p[key]))
        self._editor.open(iid, col_id, text,
                          lambda t, move, iid=iid, key=key: self._commit(iid, key, t, move))

    def _commit(self, iid: str, key: str, text: str, move: str):
        p = self._items.get(iid)
        if p is None:
            return
        try:
            value = parse_cell(key, text)
        except ValueError as e:
            messagebox.showwarning("Неверное значение", str(e), parent=self)
            return
        old = p.get(key)
        same = (value == old) or (KIND[key] == "price" and value is not None and
                                  old is not None and abs(float(old) - value) < 0.005)
        if not same:
            if key in ("article", "name") and not value and not (
                    p.get("name" if key == "article" else "article")):
                messagebox.showwarning("Нельзя",
                                       "У позиции должен остаться артикул или наименование.",
                                       parent=self)
                return
            self._save_field(iid, key, value)
        self._move(iid, key, move)

    def _move(self, iid: str, key: str, move: str):
        if move == "none":
            return
        if move == "down":
            nxt = self.tree.next(iid)
            if nxt:
                self.tree.selection_set(nxt)
                self.tree.focus(nxt)
                self.tree.see(nxt)
            return
        i = KEYS.index(key) + (1 if move == "right" else -1)
        if 0 <= i < len(KEYS):
            self._edit(iid, f"#{COLS.index(KEYS[i]) + 1}")

    def _save_field(self, iid: str, key: str, value):
        pid = int(iid)
        self._info_lbl.configure(text="Сохранение…")

        def work():
            try:
                res = self.api.update_product(pid, {key: value})
            except Exception as e:
                self.after(0, lambda e=e: (self._info_lbl.configure(text=""),
                                           self.on_error(e, "Правка позиции")))
                return
            self.after(0, lambda: self._saved(iid, res))
        threading.Thread(target=work, daemon=True).start()

    def _saved(self, iid: str, res: dict):
        p = res.get("product")
        if p and self.tree.exists(iid):
            self._items[iid] = p
            self.tree.item(iid, values=self._row_values(p))
        self._info_lbl.configure(text="✓ Сохранено")
        self.after(2500, lambda: self._info_lbl.configure(text=""))
        if res.get("status") == "updated":
            self.on_changed()

    # ── Корзина ──────────────────────────────────────────────────────────────

    def _action(self):
        if self._trash():
            self._restore_selected()
        else:
            self._delete_selected()

    def _selected_ids(self):
        return [int(i) for i in self.tree.selection() if i in self._items]

    def _delete_selected(self):
        if self._trash():
            return "break"
        ids = self._selected_ids()
        if not ids:
            messagebox.showinfo("В корзину", "Выделите позиции в таблице.", parent=self)
            return "break"
        p = self._items[str(ids[0])]
        what = (f"«{p.get('article') or p.get('name')}»" if len(ids) == 1
                else f"{len(ids)} позиций")
        if not messagebox.askyesno(
                "В корзину",
                f"Переместить в корзину {what}?\n\n"
                "Позиции перестанут подбираться в КП. Их можно вернуть "
                "из вкладки «Корзина». Повторный импорт их не вернёт.",
                parent=self):
            return "break"
        self._run(lambda: self.api.delete_products(ids),
                  lambda r: f"В корзину: {r.get('deleted', 0)}")
        return "break"

    def _restore_selected(self):
        ids = self._selected_ids()
        if not ids:
            messagebox.showinfo("Восстановить", "Выделите позиции в таблице.", parent=self)
            return

        def done(r):
            skipped = r.get("skipped") or []
            if skipped:
                names = "\n".join(f"• {s.get('article') or s.get('name')}"
                                  for s in skipped[:10])
                more = f"\n…и ещё {len(skipped) - 10}" if len(skipped) > 10 else ""
                messagebox.showwarning(
                    "Не все восстановлены",
                    f"В сегменте уже есть такие позиции — эти остались в корзине:\n"
                    f"{names}{more}", parent=self)
            return f"Восстановлено: {r.get('restored', 0)}"
        self._run(lambda: self.api.restore_products(ids), done)

    def _run(self, call, describe):
        self._info_lbl.configure(text="…")

        def work():
            try:
                r = call()
            except Exception as e:
                self.after(0, lambda e=e: (self._info_lbl.configure(text=""),
                                           self.on_error(e, "Продукция")))
                return

            def ok():
                self.reload()
                self._load_brands()
                self.on_changed()
                self.after(300, lambda: self._info_lbl.configure(text=describe(r)))
            self.after(0, ok)
        threading.Thread(target=work, daemon=True).start()

    # ── Буфер и меню ─────────────────────────────────────────────────────────

    def _on_ctrl(self, event):
        if event.keycode == KEY_C:
            self._copy_selected()
            return "break"
        if event.keycode == 65:                    # Ctrl+A
            self.tree.selection_set(self.tree.get_children())
            return "break"
        return None

    def _copy_selected(self):
        sel = [i for i in self.tree.selection() if i in self._items]
        if not sel:
            sel = list(self.tree.get_children())
        if not sel:
            return
        text = rows_to_tsv([self._items[i] for i in sel], KEYS, header=True)
        self.clipboard_clear()
        self.clipboard_append(text)
        self._info_lbl.configure(text=f"Скопировано строк: {len(sel)}")

    def _on_menu(self, event):
        iid = self.tree.identify_row(event.y)
        if not iid:
            return
        if iid not in self.tree.selection():
            self.tree.selection_set(iid)
        self.tree.focus(iid)
        col = self.tree.identify_column(event.x)
        m = self._menu
        m.delete(0, "end")
        if not self._trash():
            m.add_command(label="✎ Править ячейку", command=lambda: self._edit(iid, col))
        m.add_command(label="📋 Копировать", command=self._copy_selected)
        m.add_separator()
        if self._trash():
            m.add_command(label="↩ Восстановить", command=self._restore_selected)
        else:
            m.add_command(label="🗑 В корзину", command=self._delete_selected)
        m.tk_popup(event.x_root, event.y_root)
