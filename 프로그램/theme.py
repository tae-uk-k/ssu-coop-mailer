"""
theme.py — 색·글꼴 토큰과 공용 위젯

토스 디자인 시스템의 그레이 스케일과 파랑 하나를 기준으로 삼는다.
강조는 파랑 한 곳에만 쓰고 나머지는 회색으로 눌러 둔다.
"""
from __future__ import annotations

import re
import tkinter as tk
from tkinter import font as tkfont

import customtkinter as ctk

# 드래그앤드롭 — tkinterdnd2 만 있으면 된다 (CTkinterDnD 불필요)
try:
    from tkinterdnd2 import DND_FILES, TkinterDnD
    DND_OK = True
except Exception:      # 없으면 파일 고르기 버튼으로만 쓴다
    DND_FILES = None
    TkinterDnD = None
    DND_OK = False

# ──────────────────────────────────────────────────────────
# 색
# ──────────────────────────────────────────────────────────

BLUE     = "#3182F6"
BLUE_D   = "#2272EB"
BLUE_BG  = "#EBF3FE"
BLUE_BG2 = "#DCEAFD"

G900 = "#191F28"
G800 = "#333D4B"
G700 = "#4E5968"
G600 = "#6B7684"
G500 = "#8B95A1"
G400 = "#B0B8C1"
G300 = "#D1D6DB"
G200 = "#E5E8EB"
G100 = "#F2F4F6"
G50  = "#F9FAFB"
WHITE = "#FFFFFF"

RED      = "#F04452"
RED_BG   = "#FEF0F1"
GREEN    = "#0E9B54"
GREEN_BG = "#EAF9F1"
AMBER    = "#B87400"
AMBER_BG = "#FFF6E5"
PURPLE   = "#6941C6"
PURPLE_BG = "#F3F0FE"

# 자주 쓰는 조합
BG      = G50       # 작업 영역 바탕
CARD    = WHITE
LINE    = G100

_FAMILY = None
_fonts: dict = {}


def family() -> str:
    """설치돼 있는 한글 글꼴 중 가장 나은 것을 고른다."""
    global _FAMILY
    if _FAMILY:
        return _FAMILY
    try:
        avail = set(tkfont.families())
    except Exception:
        avail = set()
    for f in ("Pretendard Variable", "Pretendard", "Toss Product Sans",
              "맑은 고딕", "Malgun Gothic", "Apple SD Gothic Neo", "Segoe UI"):
        if f in avail:
            _FAMILY = f
            return f
    _FAMILY = "Malgun Gothic"
    return _FAMILY


def mono_family() -> str:
    try:
        avail = set(tkfont.families())
    except Exception:
        avail = set()
    for f in ("Cascadia Mono", "D2Coding", "Consolas", "Courier New"):
        if f in avail:
            return f
    return "Consolas"


def font(size: int = 15, bold: bool = False, mono: bool = False) -> ctk.CTkFont:
    """CTkFont 는 루트 창이 생긴 뒤에만 만들 수 있어 만들어 두고 돌려쓴다."""
    key = (size, bold, mono)
    f = _fonts.get(key)
    if f is None:
        f = ctk.CTkFont(family=mono_family() if mono else family(),
                        size=size, weight="bold" if bold else "normal")
        _fonts[key] = f
    return f


# ──────────────────────────────────────────────────────────
# 기본 위젯
# ──────────────────────────────────────────────────────────

class Card(ctk.CTkFrame):
    """흰 카드. 그림자 대신 바탕색 차이로만 구분한다."""

    def __init__(self, master, pad: int = 20, **kw):
        kw.setdefault("fg_color", CARD)
        kw.setdefault("corner_radius", 16)
        super().__init__(master, **kw)
        self._pad = pad

    def inner(self) -> ctk.CTkFrame:
        f = ctk.CTkFrame(self, fg_color="transparent")
        f.pack(fill="both", expand=True, padx=self._pad, pady=self._pad)
        return f


def title(master, text: str, **kw) -> ctk.CTkLabel:
    return ctk.CTkLabel(master, text=text, font=font(23, True),
                        text_color=G900, anchor="w", justify="left", **kw)


def subtitle(master, text: str, **kw) -> ctk.CTkLabel:
    return ctk.CTkLabel(master, text=text, font=font(15, True),
                        text_color=G900, anchor="w", justify="left", **kw)


def desc(master, text: str, **kw) -> ctk.CTkLabel:
    kw.setdefault("wraplength", 620)
    return ctk.CTkLabel(master, text=text, font=font(13),
                        text_color=G600, anchor="w", justify="left", **kw)


def label(master, text: str, color: str = G600, size: int = 13,
          bold: bool = False, **kw) -> ctk.CTkLabel:
    return ctk.CTkLabel(master, text=text, font=font(size, bold),
                        text_color=color, anchor="w", justify="left", **kw)


class Btn(ctk.CTkButton):
    """variant: primary | secondary | ghost | danger"""

    _STYLE = {
        "primary":   (BLUE,  BLUE_D,  WHITE),
        "secondary": (G100,  G200,    G700),
        "ghost":     ("transparent", BLUE_BG, BLUE),
        "danger":    (RED_BG, "#FDE3E5", RED),
    }

    def __init__(self, master, text: str, variant: str = "primary",
                 big: bool = False, small: bool = False, **kw):
        fg, hover, txt = self._STYLE.get(variant, self._STYLE["primary"])
        size = 15 if big else (12 if small else 13)
        kw.setdefault("fg_color", fg)
        kw.setdefault("hover_color", hover)
        kw.setdefault("text_color", txt)
        kw.setdefault("corner_radius", 14 if big else 10)
        kw.setdefault("height", 48 if big else (30 if small else 38))
        kw.setdefault("font", font(size, True))
        if variant == "ghost":
            kw.setdefault("border_width", 0)
        super().__init__(master, text=text, **kw)


class Tag(ctk.CTkLabel):
    """작은 상태 알약. 남발하지 않고 정말 필요할 때만 쓴다."""

    _STYLE = {
        "blue":  (BLUE_BG, BLUE_D),
        "green": (GREEN_BG, GREEN),
        "red":   (RED_BG, RED),
        "amber": (AMBER_BG, AMBER),
        "grey":  (G100, G500),
        "purple": (PURPLE_BG, PURPLE),
    }

    def __init__(self, master, text: str, kind: str = "grey", **kw):
        bg, fg = self._STYLE.get(kind, self._STYLE["grey"])
        kw.setdefault("fg_color", bg)
        kw.setdefault("text_color", fg)
        kw.setdefault("corner_radius", 7)
        kw.setdefault("font", font(12, True))
        kw.setdefault("padx", 9)
        kw.setdefault("pady", 3)
        super().__init__(master, text=text, **kw)


class Note(ctk.CTkFrame):
    """한 줄 안내. kind: blue | green | amber | red | grey"""

    _STYLE = {
        "blue":  (BLUE_BG, "#1B5FBF"),
        "green": (GREEN_BG, "#0B7F45"),
        "amber": (AMBER_BG, "#96590A"),
        "red":   (RED_BG, "#C3323F"),
        "grey":  (G100, G600),
    }

    def __init__(self, master, text: str, kind: str = "blue",
                 mark: str = "", wrap: int = 560, **kw):
        bg, fg = self._STYLE.get(kind, self._STYLE["blue"])
        kw.setdefault("fg_color", bg)
        kw.setdefault("corner_radius", 12)
        super().__init__(master, **kw)
        row = ctk.CTkFrame(self, fg_color="transparent")
        row.pack(fill="x", padx=14, pady=11)
        if mark:
            ctk.CTkLabel(row, text=mark, font=font(13, True), text_color=fg,
                         width=16).pack(side="left", anchor="n", padx=(0, 8))
        self._lbl = ctk.CTkLabel(row, text=text, font=font(13), text_color=fg,
                                 wraplength=wrap, justify="left", anchor="w")
        self._lbl.pack(side="left", fill="x", expand=True)

    def set_text(self, text: str) -> None:
        self._lbl.configure(text=text)


class Stat(Card):
    """큰 숫자 하나와 이름."""

    def __init__(self, master, value: str, key: str, color: str = G900, **kw):
        super().__init__(master, **kw)
        box = ctk.CTkFrame(self, fg_color="transparent")
        box.pack(fill="both", expand=True, padx=16, pady=9)
        self._n = ctk.CTkLabel(box, text=str(value), font=font(21, True),
                               text_color=color, anchor="w")
        self._n.pack(anchor="w")
        self._k = ctk.CTkLabel(box, text=key, font=font(12), text_color=G500,
                               anchor="w")
        self._k.pack(anchor="w")

    def set_value(self, v, color: str | None = None) -> None:
        self._n.configure(text=str(v))
        if color:
            self._n.configure(text_color=color)

    def set_key(self, k: str) -> None:
        self._k.configure(text=k)


class Select(ctk.CTkOptionMenu):
    """열 고르기 등에 쓰는 드롭다운."""

    def __init__(self, master, values: list[str], variable=None,
                 highlight: bool = False, width: int = 190, **kw):
        kw.setdefault("fg_color", BLUE_BG if highlight else G100)
        kw.setdefault("button_color", BLUE_BG if highlight else G100)
        kw.setdefault("button_hover_color", BLUE_BG2 if highlight else G200)
        kw.setdefault("text_color", BLUE_D if highlight else G800)
        kw.setdefault("dropdown_fg_color", WHITE)
        kw.setdefault("dropdown_text_color", G800)
        kw.setdefault("dropdown_hover_color", BLUE_BG)
        kw.setdefault("corner_radius", 10)
        kw.setdefault("height", 38)
        kw.setdefault("width", width)
        kw.setdefault("font", font(13, True))
        kw.setdefault("dropdown_font", font(13))
        super().__init__(master, values=values or [""], variable=variable, **kw)

    def set_highlight(self, on: bool) -> None:
        self.configure(fg_color=BLUE_BG if on else G100,
                       button_color=BLUE_BG if on else G100,
                       button_hover_color=BLUE_BG2 if on else G200,
                       text_color=BLUE_D if on else G800)


class Entry(ctk.CTkEntry):
    def __init__(self, master, **kw):
        kw.setdefault("fg_color", G50)
        kw.setdefault("border_width", 0)
        kw.setdefault("corner_radius", 10)
        kw.setdefault("height", 40)
        kw.setdefault("font", font(13))
        kw.setdefault("text_color", G900)
        super().__init__(master, **kw)


class TextBox(ctk.CTkTextbox):
    def __init__(self, master, **kw):
        kw.setdefault("fg_color", G50)
        kw.setdefault("border_width", 0)
        kw.setdefault("corner_radius", 10)
        kw.setdefault("font", font(13))
        kw.setdefault("text_color", G800)
        kw.setdefault("wrap", "word")
        super().__init__(master, **kw)


class Bar(ctk.CTkProgressBar):
    def __init__(self, master, **kw):
        kw.setdefault("fg_color", G100)
        kw.setdefault("progress_color", BLUE)
        kw.setdefault("corner_radius", 4)
        kw.setdefault("height", 8)
        super().__init__(master, **kw)
        self.set(0)


class Steps4(ctk.CTkFrame):
    """제안서 → PDF → 명단 → 메일 네 단계를 작은 막대로."""

    COLORS = {"": G200, "run": AMBER, "ok": GREEN, "bad": RED}

    def __init__(self, master, **kw):
        kw.setdefault("fg_color", "transparent")
        super().__init__(master, **kw)
        self._bars = []
        for i in range(4):
            b = ctk.CTkFrame(self, width=20, height=6, corner_radius=3, fg_color=G200)
            b.pack(side="left", padx=(0 if i == 0 else 3, 0))
            b.pack_propagate(False)
            self._bars.append(b)

    def set_step(self, step: int, state: str) -> None:
        if 1 <= step <= 4:
            self._bars[step - 1].configure(fg_color=self.COLORS.get(state, G200))

    def reset(self) -> None:
        for b in self._bars:
            b.configure(fg_color=G200)


# ──────────────────────────────────────────────────────────
# 표
# ──────────────────────────────────────────────────────────

class Table(ctk.CTkFrame):
    """스크롤되는 표. 셀을 눌러 그 자리에서 고칠 수 있다.

    columns: [{"key","title","width","edit"(선택),"color"(선택)}]
    """

    ROW_H = 34

    def __init__(self, master, columns: list[dict], height: int = 250,
                 on_edit=None, **kw):
        kw.setdefault("fg_color", "transparent")
        super().__init__(master, **kw)
        self.columns = columns
        self.on_edit = on_edit
        self._rows: list[dict] = []
        self._widgets: list[list] = []
        self._editing = None

        head = ctk.CTkFrame(self, fg_color=G50, corner_radius=10, height=34)
        head.pack(fill="x")
        head.pack_propagate(False)
        for c in columns:
            ctk.CTkLabel(head, text=c["title"], font=font(12, True),
                         text_color=G500, width=c.get("width", 120),
                         anchor="w").pack(side="left", padx=(12, 0), pady=6)

        self._body = ctk.CTkScrollableFrame(self, fg_color="transparent",
                                            height=height, corner_radius=0)
        self._body.pack(fill="both", expand=True, pady=(2, 0))
        # 행이 늘어도 표가 창을 밀어내지 않게 한다 (남는 공간은 expand 로 받는다)
        for attr in ("_parent_frame", "_parent_canvas"):
            box = getattr(self._body, attr, None)
            if box is not None:
                try:
                    box.pack_propagate(False)
                    box.grid_propagate(False)
                    box.configure(height=height)
                except Exception:
                    pass

    def set_rows(self, rows: list[dict]) -> None:
        for w in self._body.winfo_children():
            w.destroy()
        self._rows = rows
        self._widgets = []

        for i, r in enumerate(rows):
            line = ctk.CTkFrame(self._body, fg_color="transparent", height=self.ROW_H)
            line.pack(fill="x")
            line.pack_propagate(False)
            cells = []
            for c in self.columns:
                key = c["key"]
                val = r.get(key, "")
                color = r.get("_color_" + key) or c.get("color") or G800
                if r.get("_muted"):
                    color = G400
                lbl = ctk.CTkLabel(line, text=str(val), font=font(13),
                                   text_color=color, width=c.get("width", 120),
                                   anchor="w")
                lbl.pack(side="left", padx=(12, 0))
                if c.get("edit") and self.on_edit:
                    lbl.configure(cursor="xterm")
                    lbl.bind("<Button-1>",
                             lambda _e, ri=i, ck=key, wl=lbl: self._edit(ri, ck, wl))
                cells.append(lbl)
            sep = ctk.CTkFrame(self._body, height=1, fg_color=LINE)
            sep.pack(fill="x")
            self._widgets.append(cells)

    def _edit(self, row_i: int, key: str, lbl) -> None:
        if self._editing:
            return
        parent = lbl.master
        width = lbl.cget("width")
        ent = ctk.CTkEntry(parent, width=width, height=26, font=font(13),
                           fg_color=WHITE, border_color=BLUE, border_width=2,
                           corner_radius=7, text_color=G900)
        ent.insert(0, str(self._rows[row_i].get(key, "")))
        lbl.pack_forget()
        ent.pack(side="left", padx=(12, 0), before=None)
        # 원래 자리로 되돌리기 위해 순서를 다시 맞춘다
        for c in parent.winfo_children():
            if c is not ent:
                c.pack_forget()
        for c_def, w in zip(self.columns, self._widgets[row_i]):
            if c_def["key"] == key:
                ent.pack(side="left", padx=(12, 0))
            else:
                w.pack(side="left", padx=(12, 0))
        ent.focus_set()
        self._editing = (row_i, key, ent, lbl)

        def finish(_e=None):
            if not self._editing:
                return
            ri, ck, e, l = self._editing
            new = e.get().strip()
            self._editing = None
            try:
                e.destroy()
            except Exception:
                pass
            self._rows[ri][ck] = new
            l.configure(text=new)
            for c_def, w in zip(self.columns, self._widgets[ri]):
                w.pack_forget()
            for c_def, w in zip(self.columns, self._widgets[ri]):
                w.pack(side="left", padx=(12, 0))
            if self.on_edit:
                self.on_edit(ri, ck, new)

        ent.bind("<Return>", finish)
        ent.bind("<FocusOut>", finish)
        ent.bind("<Escape>", lambda _e: finish())


def hsep(master, pad: int = 0) -> ctk.CTkFrame:
    f = ctk.CTkFrame(master, height=1, fg_color=LINE)
    f.pack(fill="x", pady=pad)
    return f


def row(master, **kw) -> ctk.CTkFrame:
    kw.setdefault("fg_color", "transparent")
    f = ctk.CTkFrame(master, **kw)
    return f


# ──────────────────────────────────────────────────────────
# 드래그앤드롭
# ──────────────────────────────────────────────────────────

def enable_dnd_root(root) -> bool:
    """CTk 루트 창에 tkdnd 를 붙인다. 성공하면 True."""
    if not DND_OK:
        return False
    try:
        root.TkdndVersion = TkinterDnD._require(root)
        return True
    except Exception:
        return False


def parse_drop(data: str) -> list[str]:
    """드롭된 문자열을 경로 목록으로. 공백이 든 경로는 {중괄호}로 온다."""
    out = []
    for brace, plain in re.findall(r"\{([^}]*)\}|(\S+)", str(data or "")):
        p = (brace or plain).strip()
        if p:
            out.append(p)
    return out


def enable_drop(widget, on_files, hover_color: str | None = None,
                base_color: str | None = None) -> bool:
    """widget 과 그 자식들을 드롭 대상으로 만든다.

    on_files(paths: list[str]) 로 넘겨준다.
    """
    if not DND_OK:
        return False

    target = widget

    def _paint(color):
        if color is None:
            return
        try:
            target.configure(fg_color=color)
        except Exception:
            pass

    def _drop(event):
        _paint(base_color)
        paths = parse_drop(event.data)
        if paths:
            on_files(paths)
        return event.action

    def _enter(event):
        _paint(hover_color)
        return event.action

    def _leave(event):
        _paint(base_color)
        return event.action

    def _register(w):
        try:
            w.drop_target_register(DND_FILES)
            w.dnd_bind("<<Drop>>", _drop)
            w.dnd_bind("<<DropEnter>>", _enter)
            w.dnd_bind("<<DropLeave>>", _leave)
        except Exception:
            pass
        for c in w.winfo_children():
            _register(c)

    _register(widget)
    return True


# ──────────────────────────────────────────────────────────
# 반응형 도우미
# ──────────────────────────────────────────────────────────

class Flow(ctk.CTkFrame):
    """가로 폭이 좁아지면 자식을 다음 줄로 넘긴다 (칩 묶음용)."""

    def __init__(self, master, gap: int = 5, **kw):
        kw.setdefault("fg_color", "transparent")
        super().__init__(master, **kw)
        self._items: list = []
        self._gap = gap
        self._last = -1
        self.bind("<Configure>", self._on_configure)

    def add(self, widget) -> None:
        self._items.append(widget)
        self._last = -1
        self.after_idle(lambda: self._reflow(self.winfo_width()))

    def clear(self) -> None:
        for w in self._items:
            try:
                w.destroy()
            except Exception:
                pass
        self._items = []
        self._last = -1
        self.configure(height=1)

    def _on_configure(self, event):
        if event.width != self._last:
            self._reflow(event.width)

    def _scaling(self) -> float:
        """winfo_* 는 배율이 적용된 값을 주는데 place(x=) 는 적용 전 값을 받는다.

        고해상도 화면(125%·150%·200%)에서 이걸 맞추지 않으면 간격이 배율만큼 벌어진다.
        """
        try:
            return float(self._get_widget_scaling()) or 1.0
        except Exception:
            return 1.0

    def _reflow(self, width: int) -> None:
        if width <= 1 or not self._items:
            return
        self._last = width
        s = self._scaling()
        avail = width / s
        x = y = row_h = 0
        for w in self._items:
            ww = w.winfo_reqwidth() / s
            wh = w.winfo_reqheight() / s
            if x + ww > avail and x > 0:
                x = 0
                y += row_h + self._gap
                row_h = 0
            w.place(x=x, y=y)
            x += ww + self._gap
            row_h = max(row_h, wh)
        self.configure(height=max(1, int(y + row_h)))


def bind_wrap(widget, source=None, margin: int = 44, minimum: int = 220) -> None:
    """창 폭에 맞춰 줄바꿈 폭을 다시 잡는다."""
    src = source if source is not None else widget.master

    def _on(event):
        try:
            widget.configure(wraplength=max(minimum, event.width - margin))
        except Exception:
            pass

    src.bind("<Configure>", _on, add="+")


# ──────────────────────────────────────────────────────────
# 오류 알림 창
# ──────────────────────────────────────────────────────────

def show_error(master, title: str, what: str, how: str = "",
               detail: str = "", kind: str = "red") -> None:
    """무슨 일인지 / 어떻게 하면 되는지를 보여 준다.

    파이썬 원문은 '자세한 내용'에 접어 둔다. 국원분들께는 필요 없고,
    고칠 사람에게 보여 줄 때만 펴면 된다.
    """
    win = ctk.CTkToplevel(master)
    win.title(title)
    win.configure(fg_color=BG)
    win.resizable(False, False)
    win.transient(master.winfo_toplevel())

    body = ctk.CTkFrame(win, fg_color="transparent")
    body.pack(fill="both", expand=True, padx=24, pady=22)

    head = ctk.CTkFrame(body, fg_color="transparent")
    head.pack(fill="x", pady=(0, 12))
    mark, color, bg = {
        "red":   ("!", RED, RED_BG),
        "amber": ("!", AMBER, AMBER_BG),
        "blue":  ("i", BLUE_D, BLUE_BG),
    }.get(kind, ("!", RED, RED_BG))
    ctk.CTkLabel(head, text=mark, font=font(16, True), text_color=color,
                 width=32, height=32, corner_radius=16,
                 fg_color=bg).pack(side="left", padx=(0, 12))
    ctk.CTkLabel(head, text=what, font=font(16, True), text_color=G900,
                 wraplength=380, justify="left", anchor="w").pack(side="left", fill="x", expand=True)

    if how:
        box = ctk.CTkFrame(body, fg_color=G100, corner_radius=12)
        box.pack(fill="x", pady=(0, 14))
        ctk.CTkLabel(box, text=how, font=font(13), text_color=G700,
                     wraplength=400, justify="left", anchor="w"
                     ).pack(fill="x", padx=16, pady=13)

    if detail:
        holder = {"open": False}
        toggle = ctk.CTkButton(
            body, text="자세한 내용 보기", font=font(12), height=26,
            fg_color="transparent", hover_color=G100, text_color=G500,
            corner_radius=8, anchor="w")
        toggle.pack(fill="x", pady=(0, 6))

        det = TextBox(body, height=110, font=font(11, mono=True))
        det.insert("1.0", detail)
        det.configure(state="disabled")

        def flip():
            holder["open"] = not holder["open"]
            if holder["open"]:
                det.pack(fill="x", pady=(0, 10))
                toggle.configure(text="자세한 내용 접기")
            else:
                det.pack_forget()
                toggle.configure(text="자세한 내용 보기")
            win.update_idletasks()

        toggle.configure(command=flip)

    Btn(body, "확인", width=90, command=win.destroy).pack(anchor="e", pady=(4, 0))

    win.update_idletasks()
    m = master.winfo_toplevel()
    x = m.winfo_rootx() + (m.winfo_width() - win.winfo_width()) // 2
    y = m.winfo_rooty() + (m.winfo_height() - win.winfo_height()) // 3
    win.geometry(f"+{max(0, x)}+{max(0, y)}")
    win.grab_set()
    win.focus_force()
