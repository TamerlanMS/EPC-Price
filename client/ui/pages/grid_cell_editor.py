"""
Правка ячейки ttk.Treeview на месте — как в Excel.

Поверх ячейки кладётся поле ввода. Enter — сохранить и вниз, Tab — вправо,
Shift+Tab — влево, Esc — отмена, щелчок в сторону — сохранить.
"""
from __future__ import annotations

import tkinter as tk
from typing import Callable, Optional


class CellEditor:
    """Одно поле ввода на таблицу; открытое закрывается перед новым."""

    def __init__(self, tree, font=("Calibri", 12)):
        self.tree = tree
        self.font = font
        self._entry: Optional[tk.Entry] = None
        self._done = True
        self._on_commit: Optional[Callable] = None

    @property
    def is_open(self) -> bool:
        return self._entry is not None

    def open(self, iid: str, col: str, text: str,
             on_commit: Callable[[str, str], None],
             typed: Optional[str] = None) -> bool:
        """col — «#N». on_commit(text, move): move — down|right|left|none.

        typed — первый набранный символ: правка начинается с него, как в
        Excel, когда начинают печатать по выделенной ячейке.
        """
        self.commit("none")
        tree = self.tree
        tree.see(iid)
        tree.update_idletasks()
        box = tree.bbox(iid, col)
        if not box:
            return False
        x, y, w, h = box
        e = tk.Entry(tree, font=self.font, relief="solid", bd=1,
                     highlightthickness=1, highlightcolor="#2E75B6")
        e.place(x=x, y=y, width=max(w, 60), height=h)
        if typed is not None:
            e.insert(0, typed)
            e.icursor("end")
        else:
            e.insert(0, text or "")
            e.select_range(0, "end")
            e.icursor("end")
        e.focus_set()
        self._entry, self._done, self._on_commit = e, False, on_commit

        e.bind("<Return>",    lambda _e: self._finish("down"))
        e.bind("<KP_Enter>",  lambda _e: self._finish("down"))
        e.bind("<Tab>",       lambda _e: self._finish("right"))
        e.bind("<Shift-Tab>", lambda _e: self._finish("left"))
        e.bind("<ISO_Left_Tab>", lambda _e: self._finish("left"))
        e.bind("<Escape>",    lambda _e: self._cancel())
        e.bind("<FocusOut>",  lambda _e: self.commit("none"))
        return True

    def _finish(self, move: str):
        self.commit(move)
        return "break"

    def commit(self, move: str = "none") -> None:
        if self._done or self._entry is None:
            return
        self._done = True
        text = self._entry.get()
        cb = self._on_commit
        self._close()
        if cb:
            cb(text, move)

    def _cancel(self):
        self._done = True
        self._close()
        return "break"

    def _close(self):
        e, self._entry = self._entry, None
        if e is not None:
            try:
                e.destroy()
            except tk.TclError:
                pass
        try:
            self.tree.focus_set()
        except tk.TclError:
            pass
