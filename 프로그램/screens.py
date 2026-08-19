"""
screens.py — 단계별 화면

1 파일 → 2 명단 → 3 제안서 → 4 메일 → 5 보내기, 그리고 보낸 뒤 반송 확인.
"""
from __future__ import annotations

import threading
from datetime import datetime
from pathlib import Path
from tkinter import filedialog, messagebox

import customtkinter as ctk

import core
import engine
import theme as T
from theme import font

NONE_LABEL = "— 안 씀 —"


class Screen(ctk.CTkFrame):
    """화면 공통. app 은 셸(App) 이다."""

    def __init__(self, app, master):
        super().__init__(master, fg_color=T.BG)
        self.app = app
        self.build()

    def build(self) -> None:      # 자식이 채운다
        pass

    def on_show(self) -> None:    # 화면에 나타날 때마다
        pass

    def on_leave(self) -> None:   # 다른 화면으로 갈 때
        pass

    # 자주 쓰는 조각
    def head(self, title: str, desc: str = "") -> None:
        T.title(self, title).pack(anchor="w", pady=(0, 4))
        if desc:
            lbl = T.desc(self, desc)
            lbl.pack(anchor="w", pady=(0, 12))
            T.bind_wrap(lbl, self)   # 창 폭에 맞춰 줄바꿈
        else:
            ctk.CTkFrame(self, height=10, fg_color="transparent").pack()


# ══════════════════════════════════════════════════════════
# 첫 화면 — 행사 고르기
# ══════════════════════════════════════════════════════════

class WelcomeScreen(Screen):

    def build(self):
        wrap = ctk.CTkFrame(self, fg_color="transparent")
        wrap.pack(expand=True)

        ctk.CTkLabel(wrap, text="어떤 행사로 보낼까요?", font=font(25, True),
                     text_color=T.G900).pack(pady=(0, 8))
        ctk.CTkLabel(wrap,
                     text="행사마다 작업 공간을 따로 만들어요.\n"
                          "제안서·명단·메일 문구가 각각 저장돼서 섞이지 않아요.",
                     font=font(13), text_color=T.G600, justify="center").pack(pady=(0, 22))

        T.Btn(wrap, "새 행사 만들기", big=True, width=180,
              command=self.app.new_workspace).pack()

        self._recent_box = ctk.CTkFrame(wrap, fg_color="transparent")
        self._recent_box.pack(fill="x", pady=(34, 0))

        self._note = T.Note(wrap, "", "amber", mark="!", wrap=420)

    def on_show(self):
        for w in self._recent_box.winfo_children():
            w.destroy()

        names = core.list_workspaces()
        if names:
            T.label(self._recent_box, "이어서 하기", T.G600, 13, True).pack(anchor="w", pady=(0, 10))
        for n in names[:5]:
            try:
                s = core.workspace_summary(n)
            except Exception:
                s = {"name": n, "pdfs": 0, "sent": 0, "step": 0}
            self._recent_row(s)

        self._note.pack_forget()
        if not engine.gmail_connected():
            self._note.set_text("아직 Gmail이 연결되지 않았어요. "
                                "오른쪽 위 계정 칸을 눌러 연결해 주세요.")
            self._note.pack(fill="x", pady=(24, 0))

    def _recent_row(self, s: dict):
        card = T.Card(self._recent_box, corner_radius=14)
        card.pack(fill="x", pady=(0, 8))
        inner = ctk.CTkFrame(card, fg_color="transparent")
        inner.pack(fill="x", padx=16, pady=13)

        left = ctk.CTkFrame(inner, fg_color="transparent")
        left.pack(side="left", fill="x", expand=True)
        T.label(left, s["name"], T.G900, 14, True).pack(anchor="w")
        bits = []
        if s["pdfs"]:
            bits.append(f"제안서 {s['pdfs']}개")
        if s["sent"]:
            bits.append(f"보낸 곳 {s['sent']}곳")
        bits.append(f"{s['step']}단계까지 채움")
        T.label(left, " · ".join(bits), T.G500, 12).pack(anchor="w")

        T.Btn(inner, "열기", "secondary", small=True, width=60,
              command=lambda n=s["name"]: self.app.switch_workspace(n)).pack(side="right")


# ══════════════════════════════════════════════════════════
# 1 파일 — 앱이 보관한다
# ══════════════════════════════════════════════════════════

class FilesScreen(Screen):

    def build(self):
        self.head("제안서와 명단을 올려주세요",
                  "올린 파일은 앱이 보관해요. 원본을 다른 곳으로 옮기거나 "
                  "이름을 바꿔도 계속 쓸 수 있어요.")

        self._ppt_card = self._file_card("제안서 PPT", "template", ".pptx")
        self._xls_card = self._file_card("기업 명단", "recipients", ".xlsx")

        bottom = ctk.CTkFrame(self, fg_color="transparent")
        bottom.pack(fill="x", pady=(12, 0))
        T.Btn(bottom, "다음", width=90,
              command=lambda: self.app.go("list")).pack(side="right")

    def _file_card(self, title: str, kind: str, ext: str) -> dict:
        card = T.Card(self)
        card.pack(fill="x", pady=(0, 10))
        box = ctk.CTkFrame(card, fg_color="transparent")
        box.pack(fill="x", padx=18, pady=14)

        top = ctk.CTkFrame(box, fg_color="transparent")
        top.pack(fill="x", pady=(0, 12))
        T.subtitle(top, title).pack(side="left")
        ver = T.Tag(top, "", "grey")

        body = ctk.CTkFrame(box, fg_color="transparent")
        body.pack(fill="x")

        icon = ctk.CTkLabel(body, text=ext.upper().strip("."), font=font(11, True),
                            width=44, height=44, corner_radius=12,
                            fg_color=T.G100, text_color=T.G500)
        icon.pack(side="left", padx=(0, 14))

        text = ctk.CTkFrame(body, fg_color="transparent")
        text.pack(side="left", fill="x", expand=True)
        name = T.label(text, "아직 없어요", T.G900, 14, True)
        name.pack(anchor="w")
        meta = T.label(text, "파일을 골라 주세요", T.G500, 12)
        meta.pack(anchor="w")

        btns = ctk.CTkFrame(body, fg_color="transparent")
        btns.pack(side="right")
        hist = T.Btn(btns, "이전 버전", "secondary", small=True, width=76,
                     command=lambda k=kind: self._history(k))
        pick = T.Btn(btns, "파일 고르기", "secondary", small=True, width=86,
                     command=lambda k=kind, e=ext: self._pick(k, e))
        pick.pack(side="right")

        drop_hint = T.label(box, "", T.G400, 12)

        note = T.Note(box, "", "blue", mark="i", wrap=580)

        item = {"card": card, "icon": icon, "name": name, "meta": meta,
                "ver": ver, "hist": hist, "pick": pick, "note": note,
                "hint": drop_hint, "kind": kind}

        # 카드 어디에 놓아도 받는다
        item["dnd"] = T.enable_drop(
            card, lambda paths, k=kind: self._dropped(k, paths),
            hover_color=T.BLUE_BG, base_color=T.CARD)

        return item

    def _dropped(self, kind: str, paths: list[str]):
        want = ".pptx" if kind == "template" else (".xlsx", ".xlsm")
        want = (want,) if isinstance(want, str) else want
        hit = next((p for p in paths if Path(p).suffix.lower() in want), None)
        if not hit:
            messagebox.showwarning(
                "이 파일은 받을 수 없어요",
                "제안서는 .pptx, 명단은 .xlsx 파일만 올릴 수 있어요.\n"
                f"놓은 파일: {Path(paths[0]).name if paths else '(없음)'}")
            return
        self._accept(kind, Path(hit))

    def _pick(self, kind: str, ext: str):
        if not self.app.ws:
            return
        types = [("PowerPoint", "*.pptx")] if kind == "template" else [("엑셀", "*.xlsx *.xlsm")]
        path = filedialog.askopenfilename(title="파일 고르기", filetypes=types + [("모든 파일", "*.*")])
        if not path:
            return
        self._accept(kind, Path(path))

    def _accept(self, kind: str, path: Path):
        if not self.app.ws:
            messagebox.showinfo("행사를 먼저 만들어 주세요",
                                "위쪽 행사 칸에서 새 행사를 만들어 주세요.")
            return
        try:
            rel = core.import_file(self.app.ws, path, kind)
        except Exception as e:
            messagebox.showerror("올리지 못했어요", str(e))
            return

        key = "template_pptx" if kind == "template" else "recipients_xlsx"
        self.app.cfg[key] = rel
        if kind == "template":
            self.app.slots = None          # 다시 훑어야 한다
            self.app.cfg["slot_map"] = self.app.cfg.get("slot_map") or {}
        else:
            self.app.reload_sheet(force=True)
            self._autoguess_columns()
        self.app.save_cfg()
        self.on_show()
        self.app.refresh_rail()

    def _autoguess_columns(self):
        sheet = self.app.sheet
        if not sheet:
            return
        cfg = self.app.cfg
        if not cfg.get("col_email"):
            cfg["col_email"] = core.guess_column(sheet.columns, *core.EMAIL_HINTS)
        if not cfg.get("col_category"):
            cfg["col_category"] = core.guess_column(sheet.columns, *core.CATEGORY_HINTS)
            if cfg["col_category"] and not cfg.get("skip_categories"):
                cfg["skip_categories"] = ["기타"]

    def _history(self, kind: str):
        vers = core.file_versions(self.app.ws, kind)
        if len(vers) <= 1:
            messagebox.showinfo("이전 버전", "아직 이전 버전이 없어요.")
            return
        lines = "\n".join(
            f"  v{v['version']}  {v['original']}  ({v['imported']})" for v in vers)
        messagebox.showinfo("이전 버전", f"보관 중인 파일\n\n{lines}\n\n"
                                     f"폴더: {core.files_dir(self.app.ws)}")

    def on_show(self):
        cfg = self.app.cfg
        self._fill(self._ppt_card, cfg.get("template_pptx", ""), "template")
        self._fill(self._xls_card, cfg.get("recipients_xlsx", ""), "recipients")

    def _fill(self, c: dict, rel: str, kind: str):
        c["note"].pack_forget()
        c["ver"].pack_forget()
        c["hist"].pack_forget()
        c["hint"].pack_forget()

        # 파일이 아직 없을 때만 끌어다 놓기 안내를 한 줄로 보여 준다.
        # 파일이 올라오면 그 자리는 아래 안내문이 대신 쓴다.
        if c.get("dnd") and not rel:
            c["hint"].configure(text="여기로 파일을 끌어다 놓아도 돼요")
            c["hint"].pack(anchor="w", pady=(10, 0))

        if not rel or not self.app.ws:
            c["name"].configure(text="아직 없어요", text_color=T.G500)
            c["meta"].configure(text="파일을 골라 주세요")
            c["icon"].configure(fg_color=T.G100, text_color=T.G500)
            return

        path = core.resolve(self.app.ws, rel)
        if not path.exists():
            c["name"].configure(text="파일이 사라졌어요", text_color=T.RED)
            c["meta"].configure(text="다시 올려 주세요")
            return

        vers = core.file_versions(self.app.ws, kind)
        cur = next((v for v in vers if v["rel"] == rel), None)
        orig = cur["original"] if cur else path.name
        when = cur["imported"] if cur else ""
        size = core.human_size(path.stat().st_size)

        c["name"].configure(text=orig, text_color=T.G900)
        c["icon"].configure(fg_color=T.BLUE_BG if kind == "template" else T.GREEN_BG,
                            text_color=T.BLUE_D if kind == "template" else T.GREEN)

        if kind == "template":
            slots = self.app.get_slots()
            c["meta"].configure(text=f"{size} · 올린 날 {when}")
            if slots:
                c["note"].configure(fg_color=T.GREEN_BG)
                c["note"].set_text(
                    f"바꿀 수 있는 자리 {len(slots)}개를 찾았어요 — "
                    + ", ".join(s.name for s in slots[:4])
                    + (" 외" if len(slots) > 4 else "")
                    + ". 3단계에서 명단과 짝지어 주세요."
                    + ("  ·  새 파일을 끌어다 놓으면 바뀌어요." if c.get("dnd") else ""))
            else:
                c["note"].configure(fg_color=T.AMBER_BG)
                c["note"].set_text("괄호로 묶인 자리를 찾지 못했어요. "
                                   "PPT에서 바꿀 곳을 (기업명)처럼 괄호로 감싸 주세요.")
            c["note"].pack(fill="x", pady=(14, 0))
        else:
            sheet = self.app.reload_sheet()
            n = len(sheet.rows) if sheet else 0
            cols = len(sheet.columns) if sheet else 0
            c["meta"].configure(text=f"{size} · {n}행 · 열 {cols}개 · 올린 날 {when}")
            c["note"].configure(fg_color=T.BLUE_BG)
            c["note"].set_text("명단은 앱 안에서 바로 고칠 수 있어요. 엑셀을 따로 열어둘 필요가 없어요."
                               + ("  ·  새 파일을 끌어다 놓으면 바뀌어요." if c.get("dnd") else ""))
            c["note"].pack(fill="x", pady=(14, 0))

        if cur:
            c["ver"].configure(text=f"버전 {cur['version']}")
            c["ver"].pack(side="left", padx=(10, 0))
        if len(vers) > 1:
            c["hist"].pack(side="right", padx=(0, 8))


# ══════════════════════════════════════════════════════════
# 2 명단
# ══════════════════════════════════════════════════════════

class ListScreen(Screen):

    def build(self):
        self.head("명단을 확인해 주세요",
                  "표를 눌러 바로 고칠 수 있어요. 앱이 알아야 하는 건 "
                  "이메일이 어느 열인지 하나뿐이에요.")

        pick = T.Card(self)
        pick.pack(fill="x", pady=(0, 10))
        pr = ctk.CTkFrame(pick, fg_color="transparent")
        pr.pack(fill="x", padx=18, pady=10)

        T.label(pr, "이메일 열", T.G600, 13, True).pack(side="left", padx=(0, 10))
        self._v_email = ctk.StringVar()
        self._sel_email = T.Select(pr, [NONE_LABEL], self._v_email, highlight=True,
                                   width=150, command=self._on_email_col)
        self._sel_email.pack(side="left")

        T.label(pr, "제외 열", T.G600, 13, True).pack(side="left", padx=(20, 10))
        self._v_cat = ctk.StringVar()
        self._sel_cat = T.Select(pr, [NONE_LABEL], self._v_cat, width=140,
                                 command=self._on_cat_col)
        self._sel_cat.pack(side="left")
        self._e_skip = T.Entry(pr, width=110, placeholder_text="기타")
        self._e_skip.pack(side="left", padx=(8, 0))
        self._e_skip.bind("<FocusOut>", lambda _e: self._on_skip_vals())
        self._e_skip.bind("<Return>", lambda _e: self._on_skip_vals())

        self._sample = T.label(pr, "", T.G400, 12)
        self._sample.pack(side="left", padx=(14, 0))


        stats = ctk.CTkFrame(self, fg_color="transparent")
        stats.pack(fill="x", pady=(0, 10))
        self._st_send = T.Stat(stats, "0", "보낼 곳", T.BLUE)
        self._st_send.pack(side="left", fill="x", expand=True, padx=(0, 6))
        self._st_skip = T.Stat(stats, "0", "빠지는 곳", T.G400)
        self._st_skip.pack(side="left", fill="x", expand=True, padx=6)
        self._st_all = T.Stat(stats, "0", "전체", T.G400)
        self._st_all.pack(side="left", fill="x", expand=True, padx=(6, 0))

        tcard = T.Card(self)
        tcard.pack(fill="both", expand=True)
        th = ctk.CTkFrame(tcard, fg_color="transparent")
        th.pack(fill="x", padx=20, pady=(14, 6))
        T.subtitle(th, "명단").pack(side="left")
        self._breakdown = T.label(th, "", T.G500, 12)
        self._breakdown.pack(side="left", padx=(12, 0))
        T.Btn(th, "행 추가", "secondary", small=True, width=66,
              command=self._add_row).pack(side="right")

        self._table = T.Table(
            tcard,
            columns=[
                {"key": "name",   "title": "업체명",   "width": 150},
                {"key": "email",  "title": "이메일",   "width": 240, "edit": True},
                {"key": "cat",    "title": "분류",     "width": 120},
                {"key": "status", "title": "상태",     "width": 110},
            ],
            height=60, on_edit=self._cell_edited)
        self._table.pack(fill="both", expand=True, padx=16, pady=(0, 12))

        bottom = ctk.CTkFrame(self, fg_color="transparent")
        bottom.pack(fill="x", pady=(12, 0))
        T.Btn(bottom, "다음", width=90,
              command=lambda: self.app.go("slots")).pack(side="right")

    # ── 동작 ──
    def _on_email_col(self, v):
        self.app.cfg["col_email"] = "" if v == NONE_LABEL else v
        self.app.save_cfg()
        self.refresh()
        self.app.refresh_rail()

    def _on_cat_col(self, v):
        self.app.cfg["col_category"] = "" if v == NONE_LABEL else v
        self.app.save_cfg()
        self.refresh()

    def _on_skip_vals(self):
        raw = self._e_skip.get()
        vals = [s.strip() for s in raw.replace(",", " ").split() if s.strip()]
        self.app.cfg["skip_categories"] = vals
        self.app.save_cfg()
        self.refresh()

    def _cell_edited(self, row_i: int, key: str, value: str):
        sheet = self.app.sheet
        if not sheet or key != "email":
            return
        target = self._view[row_i]
        sheet.set(target["index"], self.app.cfg.get("col_email", ""), value)
        try:
            sheet.save()
        except Exception as e:
            messagebox.showwarning("저장하지 못했어요", str(e))
        self.refresh()

    def _add_row(self):
        sheet = self.app.sheet
        if not sheet:
            return
        sheet.add_row()
        try:
            sheet.save()
        except Exception as e:
            messagebox.showwarning("저장하지 못했어요", str(e))
        self.refresh()

    def on_show(self):
        sheet = self.app.reload_sheet()
        cols = sheet.columns if sheet else []
        vals = [NONE_LABEL] + cols

        self._sel_email.configure(values=vals)
        self._sel_cat.configure(values=vals)
        self._v_email.set(self.app.cfg.get("col_email") or NONE_LABEL)
        self._v_cat.set(self.app.cfg.get("col_category") or NONE_LABEL)
        self._sel_email.set_highlight(bool(self.app.cfg.get("col_email")))

        self._e_skip.delete(0, "end")
        self._e_skip.insert(0, " ".join(self.app.cfg.get("skip_categories") or []))
        self.refresh()

    def refresh(self):
        sheet = self.app.sheet
        if not sheet:
            self._table.set_rows([])
            return

        col_em = self.app.cfg.get("col_email", "")
        if col_em:
            samples = [core.first_email(r.get(col_em, "")) for r in sheet.rows[:3]]
            self._sample.configure(text=" · ".join([s for s in samples if s][:2]))
        else:
            self._sample.configure(text="열을 골라야 보낼 수 있어요")

        marks = core.classify(sheet, self.app.cfg)
        self._view = marks

        n_send = sum(1 for m in marks if m["kind"] == core.SEND)
        self._st_send.set_value(n_send)
        self._st_skip.set_value(len(marks) - n_send)
        self._st_all.set_value(len(marks))

        counts = {}
        for m in marks:
            if m["kind"] != core.SEND:
                counts[m["reason"]] = counts.get(m["reason"], 0) + 1
        self._breakdown.configure(
            text=" · ".join(f"{k} {v}" for k, v in counts.items()) if counts else "")

        rows = []
        for m in marks:
            ok = m["kind"] == core.SEND
            rows.append({
                "name":   m["name"] or "-",
                "email":  m["email"] or "미확인",
                "cat":    m["category"] or "-",
                "status": "보낼 예정" if ok else m["reason"],
                "_muted": not ok,
                "_color_status": T.G500 if ok else T.AMBER,
            })
        self._table.set_rows(rows)


# ══════════════════════════════════════════════════════════
# 3 제안서 — PPT 자리 ↔ 명단 열
# ══════════════════════════════════════════════════════════

class SlotsScreen(Screen):

    def build(self):
        self.head("제안서에서 바꿀 자리를 찾았어요",
                  "각 자리에 명단의 어느 열을 넣을지 골라 주세요. "
                  "자리 이름은 (기업명)이든 (회사명)이든 상관없어요.")

        opt = ctk.CTkFrame(self, fg_color="transparent")
        opt.pack(fill="x", pady=(0, 12))
        self._v_josa = ctk.BooleanVar(value=True)
        ctk.CTkCheckBox(opt, text="뒤에 오는 조사를 알아서 맞추기  (농심을 / 오뚜기를)",
                        variable=self._v_josa, font=font(13),
                        text_color=T.G700, fg_color=T.BLUE,
                        hover_color=T.BLUE_D, corner_radius=6,
                        checkbox_width=20, checkbox_height=20,
                        command=self._on_josa).pack(side="left")

        self._list = ctk.CTkScrollableFrame(self, fg_color="transparent")
        self._list.pack(fill="both", expand=True)

        bottom = ctk.CTkFrame(self, fg_color="transparent")
        bottom.pack(fill="x", pady=(14, 0))
        T.Btn(bottom, "다음", width=90,
              command=lambda: self.app.go("mail")).pack(side="right")
        T.Btn(bottom, "미리보기 1장 만들기", "secondary", width=150,
              command=self._preview).pack(side="right", padx=(0, 8))
        self._prev_state = T.label(bottom, "", T.G500, 12)
        self._prev_state.pack(side="right", padx=(0, 12))

    def on_leave(self):
        try:
            self._save_name()
        except Exception:
            pass

    def _on_josa(self):
        self.app.cfg["slot_josa"] = bool(self._v_josa.get())
        self.app.save_cfg()
        self.on_show()

    def on_show(self):
        for w in self._list.winfo_children():
            w.destroy()

        self._v_josa.set(bool(self.app.cfg.get("slot_josa", True)))
        slots = self.app.get_slots()
        sheet = self.app.reload_sheet()
        cols = sheet.columns if sheet else []

        # 만들어질 파일 이름 — 명단 열을 쓰므로 여기가 제자리
        self._name_card(self._list)
        self._fill_name_card()

        if not slots:
            T.Note(self._list,
                   "PPT에서 괄호로 묶인 자리를 찾지 못했어요. "
                   "제안서에서 바꿀 곳을 (기업명)처럼 괄호로 감싼 뒤 1단계에서 다시 올려 주세요.",
                   "amber", mark="!", wrap=640).pack(fill="x")
            return

        smap = self.app.cfg.get("slot_map") or {}
        sample = self._sample_row()

        for slot in slots:
            picked = smap.get(slot.name, "")
            if not picked:
                guess = core.guess_column(cols, slot.inner)
                if guess:
                    picked = guess
                    smap[slot.name] = guess
        self.app.cfg["slot_map"] = smap
        self.app.save_cfg()

        for slot in slots:
            self._slot_card(slot, cols, smap.get(slot.name, ""), sample)

    # ── 만들어질 파일 이름 ──
    def _name_card(self, parent):
        card = T.Card(parent)
        card.pack(fill="x", pady=(0, 10))
        box = ctk.CTkFrame(card, fg_color="transparent")
        box.pack(fill="x", padx=18, pady=14)

        top = ctk.CTkFrame(box, fg_color="transparent")
        top.pack(fill="x", pady=(0, 10))
        T.subtitle(top, "만들어질 파일 이름").pack(side="left")
        T.label(top, "기업마다 이 규칙으로 제안서가 만들어져요",
                T.G500, 12).pack(side="left", padx=(10, 0))

        row = ctk.CTkFrame(box, fg_color="transparent")
        row.pack(fill="x")
        self._e_name = T.Entry(row)
        self._e_name.pack(side="left", fill="x", expand=True)
        self._e_name.bind("<KeyRelease>", lambda _e: self._name_changed())
        self._e_name.bind("<FocusOut>", lambda _e: self._save_name())
        self._e_name.bind("<Return>", lambda _e: self._save_name())

        self._name_chips = T.Flow(box, gap=5)
        self._name_chips.pack(fill="x", pady=(8, 0))

        self._name_preview = T.Note(box, "", "grey", mark="→", wrap=560)
        self._name_preview.pack(fill="x", pady=(10, 0))

    def _name_changed(self):
        """입력할 때마다 실제 만들어질 이름을 보여 준다."""
        pattern = self._e_name.get()
        sheet = self.app.sheet
        rows = sheet.rows if sheet else []
        if not rows:
            self._name_preview.configure(fg_color=T.G100)
            self._name_preview.set_text("명단을 올리면 실제 이름을 보여 드릴게요.")
            return

        info = core.preview_filenames(pattern, rows)
        first = info["first"] + ".pdf"

        if info["leftover"]:
            self._name_preview.configure(fg_color=T.AMBER_BG)
            self._name_preview.set_text(
                f"{first}\n"
                f"명단에 없는 열이에요: {', '.join('{{'+c+'}}' for c in info['leftover'])}"
                "  ·  아래 파란 칸을 눌러 넣어 주세요.")
        elif info["unique"] < info["total"]:
            self._name_preview.configure(fg_color=T.AMBER_BG)
            self._name_preview.set_text(
                f"{first}\n"
                f"{info['total']}곳 중 {info['unique']}가지 이름만 나와요. "
                "기업마다 달라지는 값을 넣지 않으면 파일이 서로 덮어써져요.")
        else:
            self._name_preview.configure(fg_color=T.GREEN_BG)
            self._name_preview.set_text(f"{first}   ·  {info['total']}곳 모두 다른 이름이에요.")

    def _save_name(self):
        self.app.cfg["file_name_pattern"] = self._e_name.get().strip() or "제안서"
        self.app.save_cfg()
        self._name_changed()

    def _insert_name_token(self, token: str):
        self._e_name.insert("insert", token)
        self._e_name.focus_set()
        self._name_changed()

    def _fill_name_card(self):
        self._e_name.delete(0, "end")
        self._e_name.insert(0, self.app.cfg.get("file_name_pattern") or "")

        self._name_chips.clear()
        sheet = self.app.sheet
        for col in (sheet.columns if sheet else [])[:6]:
            token = "{{" + col + "}}"
            units = sum(2 if ord(ch) > 0x7F else 1 for ch in token)
            b = ctk.CTkButton(self._name_chips, text=token,
                              font=font(11, True, mono=True),
                              fg_color=T.BLUE_BG, hover_color=T.BLUE_BG2,
                              text_color=T.BLUE_D, corner_radius=8,
                              height=26, width=units * 5 + 16,
                              command=lambda t=token: self._insert_name_token(t))
            self._name_chips.add(b)
        self._name_changed()

    def _sample_row(self) -> dict:
        sheet = self.app.sheet
        if not sheet:
            return {}
        marks = core.classify(sheet, self.app.cfg)
        for m in marks:
            if m["kind"] == core.SEND:
                return m["row"]
        return sheet.rows[0] if sheet.rows else {}

    def _slot_card(self, slot, cols: list[str], picked: str, sample: dict):
        unset = not picked
        card = T.Card(self._list)
        card.pack(fill="x", pady=(0, 10))
        box = ctk.CTkFrame(card, fg_color="transparent")
        box.pack(fill="x", padx=20, pady=18)

        top = ctk.CTkFrame(box, fg_color="transparent")
        top.pack(fill="x", pady=(0, 14))
        ctk.CTkLabel(top, text=slot.name, font=font(14, True, mono=True),
                     fg_color=T.AMBER_BG if unset else T.G100,
                     text_color=T.AMBER if unset else T.G800,
                     corner_radius=9, padx=11, pady=4).pack(side="left")
        if unset:
            T.Tag(top, "아직 안 골랐어요", "amber").pack(side="left", padx=(8, 0))
        elif picked == slot.inner:
            T.Tag(top, "자동으로 짝지었어요", "green").pack(side="left", padx=(8, 0))
        T.label(top, slot.where, T.G500, 12).pack(side="right")

        mid = ctk.CTkFrame(box, fg_color="transparent")
        mid.pack(fill="x")
        T.label(mid, "여기에", T.G600, 13, True).pack(side="left", padx=(0, 10))

        var = ctk.StringVar(value=picked or NONE_LABEL)
        sel = T.Select(mid, [NONE_LABEL] + cols, var, highlight=bool(picked),
                       width=180,
                       command=lambda v, s=slot: self._pick(s, v))
        sel.pack(side="left")
        T.label(mid, "를 넣어요", T.G500, 13).pack(side="left", padx=(10, 0))

        # 어떤 문장에 쓰이는지 — 원문과 채운 결과를 나란히
        ctx = ctk.CTkFrame(box, fg_color="transparent")
        ctx.pack(fill="x", pady=(14, 0))
        T.hsep(ctx)
        before = slot.context
        T.label(ctx, before, T.G400, 12).pack(anchor="w", pady=(10, 2))

        smap_one = {slot.name: picked} if picked else {}
        after = engine.fill_text(before, smap_one, sample,
                                 bool(self.app.cfg.get("slot_josa", True)))
        if picked:
            T.label(ctx, after, T.G800, 12, True).pack(anchor="w")
        else:
            T.label(ctx, f"고르지 않으면 {slot.name} 글자가 그대로 남아요.",
                    T.AMBER, 12).pack(anchor="w")

    def _pick(self, slot, value: str):
        smap = self.app.cfg.get("slot_map") or {}
        if value == NONE_LABEL:
            smap.pop(slot.name, None)
        else:
            smap[slot.name] = value
        self.app.cfg["slot_map"] = smap
        self.app.save_cfg()
        self.on_show()
        self.app.refresh_rail()

    def _preview(self):
        cfg = self.app.cfg
        if not cfg.get("template_pptx"):
            messagebox.showwarning("먼저 파일을 올려 주세요", "1단계에서 제안서 PPT를 올려 주세요.")
            return
        row = self._sample_row()
        if not row:
            messagebox.showwarning("명단이 비어 있어요", "2단계에서 명단을 확인해 주세요.")
            return

        self._prev_state.configure(text="만드는 중…")
        self.update_idletasks()

        def work():
            try:
                with engine.PowerPointSession() as ppt:
                    pptx, pdf = engine.build_one(self.app.ws, cfg, row, ppt)
                made = pdf or pptx
                self.after(0, lambda: self._preview_done(made))
            except Exception as e:
                self.after(0, lambda: self._preview_fail(e))

        threading.Thread(target=work, daemon=True).start()

    def _preview_done(self, path):
        self._prev_state.configure(text=f"만들었어요 — {path.name}")
        try:
            import os
            os.startfile(str(path))
        except Exception:
            messagebox.showinfo("만들었어요", str(path))

    def _preview_fail(self, e):
        self._prev_state.configure(text="")
        messagebox.showerror("만들지 못했어요", str(e))


# ══════════════════════════════════════════════════════════
# 4 메일
# ══════════════════════════════════════════════════════════

class MailScreen(Screen):

    def build(self):
        self.head("메일 문구를 쓰면 오른쪽에 바로 보여요",
                  "기업을 바꿔 가며 실제로 어떻게 나갈지 확인할 수 있어요.")

        cols = ctk.CTkFrame(self, fg_color="transparent")
        cols.pack(fill="both", expand=True)
        cols.rowconfigure(0, weight=1)
        self._cols = cols
        self._stacked = None
        cols.bind("<Configure>", self._relayout)

        # 왼쪽 — 편집
        left = ctk.CTkFrame(cols, fg_color="transparent")
        self._left = left

        T.label(left, "제목", T.G600, 13, True).pack(anchor="w", pady=(0, 6))
        self._e_subject = T.Entry(left)
        self._e_subject.pack(fill="x", pady=(0, 14))
        self._e_subject.bind("<KeyRelease>", self._queue_preview)

        hdr = ctk.CTkFrame(left, fg_color="transparent")
        hdr.pack(fill="x")
        T.label(hdr, "본문", T.G600, 13, True).pack(side="left")
        T.label(hdr, "누르면 커서 자리에 들어가요", T.G500, 12).pack(side="right")

        self._chips = T.Flow(left, gap=5)
        self._chips.pack(fill="x", pady=(6, 8))

        self._t_body = T.TextBox(left, height=128, width=240)
        self._t_body.pack(fill="both", expand=True)
        self._t_body.bind("<KeyRelease>", self._queue_preview)

        hint = T.Note(left, "{{업체명은}} 처럼 조사를 붙이면 받침에 따라 은/는을 알아서 골라요.",
                      "blue", mark="가", wrap=300)
        hint.pack(fill="x", pady=(10, 0))
        T.bind_wrap(hint._lbl, left, margin=64, minimum=200)

        # 오른쪽 — 미리보기
        right = ctk.CTkFrame(cols, fg_color="transparent")
        self._right = right

        ph = ctk.CTkFrame(right, fg_color="transparent")
        ph.pack(fill="x", pady=(0, 6))
        T.label(ph, "이렇게 나가요", T.G600, 13, True).pack(side="left")
        self._v_who = ctk.StringVar()
        self._sel_who = T.Select(ph, [NONE_LABEL], self._v_who, width=150,
                                 command=lambda _v: self._render_preview())
        self._sel_who.pack(side="right")

        card = T.Card(right, corner_radius=14)
        card.pack(fill="both", expand=True)
        head = ctk.CTkFrame(card, fg_color="transparent")
        head.pack(fill="x", padx=16, pady=(14, 10))
        self._p_to = T.label(head, "", T.G700, 12)
        self._p_to.pack(anchor="w")
        self._p_subject = T.label(head, "", T.G900, 13, True, wraplength=300)
        self._p_subject.pack(anchor="w", pady=(4, 0))
        T.bind_wrap(self._p_subject, right, margin=56, minimum=200)
        T.hsep(card)
        self._p_body = T.TextBox(card, fg_color=T.WHITE, height=120, width=240)
        self._p_body.pack(fill="both", expand=True, padx=10, pady=8)
        self._p_body.configure(state="disabled")

        att = ctk.CTkFrame(card, fg_color="transparent")
        att.pack(fill="x", padx=16, pady=(0, 14))
        ctk.CTkLabel(att, text="PDF", font=font(10, True), width=34, height=34,
                     corner_radius=9, fg_color=T.RED_BG,
                     text_color=T.RED).pack(side="left", padx=(0, 10))
        self._p_att = T.label(att, "", T.G800, 13, True)
        self._p_att.pack(anchor="w")

        bottom = ctk.CTkFrame(self, fg_color="transparent")
        bottom.pack(fill="x", pady=(14, 0))
        T.Btn(bottom, "다음", width=90,
              command=lambda: self.app.go("send")).pack(side="right")
        T.Btn(bottom, "저장", "secondary", width=76,
              command=self._save).pack(side="right", padx=(0, 8))

        self._job = None
        self._relayout()

    def _relayout(self, _event=None):
        """좁으면 편집/미리보기를 위아래로, 넓으면 좌우로 놓는다."""
        width = self._cols.winfo_width()
        if width <= 1:
            width = self.winfo_toplevel().winfo_width() - 300
        stacked = width < 820
        if stacked == self._stacked:
            return
        self._stacked = stacked

        self._left.grid_forget()
        self._right.grid_forget()
        for i in (0, 1):
            self._cols.columnconfigure(i, weight=0, uniform="")
            self._cols.rowconfigure(i, weight=0)

        if stacked:
            self._cols.columnconfigure(0, weight=1)
            self._cols.rowconfigure(0, weight=1)
            self._cols.rowconfigure(1, weight=1)
            self._left.grid(row=0, column=0, sticky="nsew", pady=(0, 8))
            self._right.grid(row=1, column=0, sticky="nsew")
        else:
            self._cols.columnconfigure(0, weight=1, uniform="c")
            self._cols.columnconfigure(1, weight=1, uniform="c")
            self._cols.rowconfigure(0, weight=1)
            self._left.grid(row=0, column=0, sticky="nsew", padx=(0, 6))
            self._right.grid(row=0, column=1, sticky="nsew", padx=(6, 0))

    def on_show(self):
        cfg = self.app.cfg
        self._e_subject.delete(0, "end")
        self._e_subject.insert(0, cfg.get("email_subject", ""))
        self._t_body.delete("1.0", "end")
        self._t_body.insert("1.0", cfg.get("email_body", ""))

        sheet = self.app.reload_sheet()
        self._chips.clear()
        cols = (sheet.columns if sheet else [])[:6]
        for c in cols:
            self._chip("{{" + c + "}}", T.BLUE_BG, T.BLUE_D)
        for c in cols[:2]:
            self._chip("{{" + c + "은}}", T.PURPLE_BG, T.PURPLE)

        self._targets = []
        if sheet:
            self._targets = [m for m in core.classify(sheet, cfg)
                             if m["kind"] == core.SEND] or core.classify(sheet, cfg)
        names = [m["name"] or f"{i+1}행" for i, m in enumerate(self._targets)][:50]
        self._sel_who.configure(values=names or [NONE_LABEL])
        if names:
            self._v_who.set(names[0])
        self._render_preview()

    def _chip(self, text: str, bg: str, fg: str):
        # width 는 최소값일 뿐 — CTkButton 이 글자에 맞춰 알아서 늘린다.
        # 한글은 폭이 두 배라 대충 세어 최소값만 잡아 준다.
        units = sum(2 if ord(c) > 0x7F else 1 for c in text)
        b = ctk.CTkButton(self._chips, text=text, font=font(11, True, mono=True),
                          fg_color=bg, hover_color=bg, text_color=fg,
                          corner_radius=8, height=26, width=units * 5 + 16,
                          command=lambda t=text: self._insert(t))
        self._chips.add(b)

    def _insert(self, token: str):
        self._t_body.insert("insert", token)
        self._queue_preview()

    def _queue_preview(self, _e=None):
        if self._job:
            self.after_cancel(self._job)
        self._job = self.after(180, self._render_preview)

    def _save(self):
        self.app.cfg["email_subject"] = self._e_subject.get().strip()
        self.app.cfg["email_body"] = self._t_body.get("1.0", "end").rstrip()
        self.app.save_cfg()
        self.app.refresh_rail()
        messagebox.showinfo("저장했어요", "메일 문구를 저장했어요.")

    def on_leave(self):
        self.app.cfg["email_subject"] = self._e_subject.get().strip()
        self.app.cfg["email_body"] = self._t_body.get("1.0", "end").rstrip()
        self.app.save_cfg()

    def _render_preview(self):
        self._job = None
        subject = self._e_subject.get()
        body = self._t_body.get("1.0", "end").rstrip()

        who = self._v_who.get()
        target = next((m for m in getattr(self, "_targets", []) if m["name"] == who), None)
        row = target["row"] if target else {}

        self._p_to.configure(text="받는이  " + (core.first_email(target["email"])
                                             if target else "명단이 없어요"))
        self._p_subject.configure(text=core.replace_placeholders(subject, row) or "(제목 없음)")

        self._p_body.configure(state="normal")
        self._p_body.delete("1.0", "end")
        self._p_body.insert("1.0", core.replace_placeholders(body, row) or "(본문 없음)")
        self._p_body.configure(state="disabled")

        base = core.safe_filename(core.replace_placeholders(
            self.app.cfg.get("file_name_pattern") or "제안서", row))
        self._p_att.configure(text=f"{base}.pdf")


# ══════════════════════════════════════════════════════════
# 5 보내기 (준비 → 진행 → 결과)
# ══════════════════════════════════════════════════════════

class SendScreen(Screen):

    def build(self):
        self._stop = threading.Event()
        self._rows_widgets = {}

        self._title = T.title(self, "보낼 준비를 확인할게요")
        self._title.pack(anchor="w", pady=(0, 6))
        self._sub = T.desc(self, "")
        self._sub.pack(anchor="w", pady=(0, 18))

        # 점검 결과
        self._check = T.Note(self, "", "green", mark="✓", wrap=640)

        # 숫자
        self._stats = ctk.CTkFrame(self, fg_color="transparent")
        self._s1 = T.Stat(self._stats, "0", "보낼 곳", T.BLUE)
        self._s1.pack(side="left", fill="x", expand=True, padx=(0, 6))
        self._s2 = T.Stat(self._stats, "0", "만들 제안서", T.G900)
        self._s2.pack(side="left", fill="x", expand=True, padx=6)
        self._s3 = T.Stat(self._stats, "-", "예상 시간", T.G400)
        self._s3.pack(side="left", fill="x", expand=True, padx=(6, 0))

        # 설정 (준비 상태에서만)
        self._opts = T.Card(self)
        orow = ctk.CTkFrame(self._opts, fg_color="transparent")
        orow.pack(fill="x", padx=20, pady=16)
        T.subtitle(orow, "보내는 간격").pack(side="left")
        T.label(orow, "한꺼번에 몰아 보내면 스팸으로 분류될 수 있어요",
                T.G500, 12).pack(side="left", padx=(10, 0))
        self._v_gap = ctk.StringVar(value="5초")
        T.Select(orow, ["0초", "3초", "5초", "10초", "30초"], self._v_gap,
                 width=100, command=self._on_gap).pack(side="right")

        # 진행
        self._prog_card = T.Card(self)
        pb = ctk.CTkFrame(self._prog_card, fg_color="transparent")
        pb.pack(fill="x", padx=20, pady=16)
        prow = ctk.CTkFrame(pb, fg_color="transparent")
        prow.pack(fill="x", pady=(0, 10))
        self._p_count = ctk.CTkLabel(prow, text="0 / 0", font=font(22, True),
                                     text_color=T.G900)
        self._p_count.pack(side="left")
        self._p_pct = T.label(prow, "0%", T.G500, 12)
        self._p_pct.pack(side="right")
        self._bar = T.Bar(pb)
        self._bar.pack(fill="x")
        crow = ctk.CTkFrame(pb, fg_color="transparent")
        crow.pack(fill="x", pady=(14, 0))
        self._p_now = T.label(crow, "", T.G800, 13, True)
        self._p_now.pack(side="left")
        self._btn_stop = T.Btn(crow, "중지", "danger", small=True, width=60,
                               command=self._on_stop)
        self._btn_stop.pack(side="right")

        # 기업별 진행 목록
        self._list_card = T.Card(self)
        lh = ctk.CTkFrame(self._list_card, fg_color="transparent")
        lh.pack(fill="x", padx=20, pady=(16, 6))
        T.subtitle(lh, "기업별 진행").pack(side="left")
        T.label(lh, "제안서 → PDF → 명단 → 메일", T.G500, 12).pack(side="right")
        self._rows = ctk.CTkScrollableFrame(self._list_card, fg_color="transparent",
                                            height=130)
        self._rows.pack(fill="both", expand=True, padx=16, pady=(0, 14))

        # 결과
        self._after_card = T.Card(self, fg_color=T.BLUE_BG)
        ab = ctk.CTkFrame(self._after_card, fg_color="transparent")
        ab.pack(fill="x", padx=20, pady=16)
        T.label(ab, "이제 반송을 확인할 차례예요", "#1B5FBF", 14, True).pack(anchor="w")
        T.label(ab, "주소가 틀린 메일은 30분에서 1시간 뒤에 되돌아와요.",
                "#2C6BC4", 12).pack(anchor="w", pady=(2, 10))
        T.Btn(ab, "반송 확인하러 가기", width=150,
              command=lambda: self.app.go("bounce")).pack(anchor="w")

        # 아래 버튼
        self._bottom = ctk.CTkFrame(self, fg_color="transparent")
        self._bottom.pack(fill="x", side="bottom", pady=(14, 0))
        self._btn_send = T.Btn(self._bottom, "보내기", big=True, width=200,
                               command=self._start)
        self._btn_hint = T.label(self._bottom, "", T.G500, 12)

        self._state = "ready"

    # ── 상태 전환 ──
    def _layout(self, state: str):
        self._state = state
        for w in (self._check, self._stats, self._opts, self._prog_card,
                  self._list_card, self._after_card):
            w.pack_forget()
        self._btn_send.pack_forget()
        self._btn_hint.pack_forget()

        if state == "ready":
            self._title.configure(text=self._ready_title)
            self._check.pack(fill="x", pady=(0, 12))
            self._stats.pack(fill="x", pady=(0, 12))
            self._opts.pack(fill="x")
            self._btn_send.pack(side="left")
            self._btn_hint.pack(side="left", padx=(14, 0))
        elif state == "running":
            self._prog_card.pack(fill="x", pady=(0, 12))
            self._list_card.pack(fill="both", expand=True)
        else:  # done
            self._stats.pack(fill="x", pady=(0, 12))
            self._list_card.pack(fill="both", expand=True, pady=(0, 12))
            self._after_card.pack(fill="x")

    def _on_gap(self, v):
        self.app.cfg["send_interval"] = float(str(v).replace("초", ""))
        self.app.save_cfg()

    def on_show(self):
        if self._state == "running":
            return
        self._prepare()

    def _prepare(self):
        cfg = self.app.cfg
        sheet = self.app.reload_sheet()
        problems = self._check_all(cfg, sheet)

        self._targets = []
        if sheet and cfg.get("col_email"):
            self._targets = [m for m in core.classify(sheet, cfg)
                             if m["kind"] == core.SEND]

        n = len(self._targets)
        self._ready_title = f"{n}곳에 보낼 준비가 됐어요" if n and not problems \
            else ("보내기 전에 확인할 게 있어요" if problems else "보낼 곳이 없어요")
        self._sub.configure(text="기업마다 제안서를 만들고 PDF로 바꿔서 첨부해 보내요.")

        self._s1.set_value(n)
        self._s2.set_value(n)
        secs = int(n * (12 + float(cfg.get("send_interval") or 0)))
        self._s3.set_value(f"{max(1, secs // 60)}분" if n else "-")

        if problems:
            self._check.configure(fg_color=T.AMBER_BG)
            self._check.set_text("· " + "\n· ".join(problems))
            self._btn_send.configure(state="disabled")
            self._btn_hint.configure(text="위 내용을 채우면 보낼 수 있어요.")
        else:
            self._check.configure(fg_color=T.GREEN_BG)
            self._check.set_text("확인해야 할 것을 모두 통과했어요.")
            self._btn_send.configure(state="normal" if n else "disabled")
            self._btn_hint.configure(
                text="보내는 중에 멈출 수 있고, 다시 시작하면 보낸 곳은 건너뛰어요.")

        self._btn_send.configure(text=f"{n}곳에 보내기" if n else "보내기")
        self._v_gap.set(f"{int(float(cfg.get('send_interval') or 0))}초")
        self._layout("ready")

    def _check_all(self, cfg: dict, sheet) -> list[str]:
        out = []
        if not engine.gmail_connected():
            out.append("Gmail이 연결되지 않았어요 — 오른쪽 위 계정 칸을 눌러 주세요.")
        if not cfg.get("template_pptx"):
            out.append("제안서 PPT가 없어요 — 1단계에서 올려 주세요.")
        if not sheet:
            out.append("명단이 없어요 — 1단계에서 올려 주세요.")
        if not cfg.get("col_email"):
            out.append("이메일 열을 고르지 않았어요 — 2단계에서 골라 주세요.")
        smap = {k: v for k, v in (cfg.get("slot_map") or {}).items() if v}
        if not smap:
            out.append("제안서에서 바꿀 자리를 한 개도 짝짓지 않았어요 — 3단계를 확인해 주세요.")
        if not cfg.get("email_subject", "").strip():
            out.append("메일 제목이 비어 있어요 — 4단계에서 써 주세요.")
        if not cfg.get("email_body", "").strip():
            out.append("메일 본문이 비어 있어요 — 4단계에서 써 주세요.")
        if not engine.powerpoint_available():
            out.append("PowerPoint를 찾지 못했어요 — PDF로 바꿀 수 없어요.")

        if sheet and sheet.rows:
            info = core.preview_filenames(
                cfg.get("file_name_pattern") or "제안서", sheet.rows)
            if info["leftover"]:
                out.append(
                    "파일 이름 규칙에 명단에 없는 열이 있어요: "
                    + ", ".join("{{" + c + "}}" for c in info["leftover"])
                    + " — 1단계에서 고쳐 주세요.")
            elif info["unique"] < info["total"]:
                out.append(
                    f"파일 이름이 {info['total']}곳 중 {info['unique']}가지뿐이에요. "
                    "1단계에서 기업마다 달라지는 값을 넣어 주세요.")
        return out

    # ── 발송 ──
    def _start(self):
        n = len(self._targets)
        if not n:
            return
        if not messagebox.askyesno("보낼까요?", f"{n}곳에 제안서를 보냅니다.\n계속할까요?"):
            return

        self._stop.clear()
        self._btn_stop.configure(state="normal")
        self._layout("running")
        self._build_rows()
        self._bar.set(0)
        self._p_count.configure(text=f"0 / {n}")
        self._p_pct.configure(text="0%")
        self.app.set_busy(True)

        cfg = dict(self.app.cfg)
        sheet = self.app.sheet
        targets = list(self._targets)

        def log(msg):
            self.after(0, lambda m=msg: self.app.log(m))

        def progress(i, total, name, step, state):
            self.after(0, lambda: self._on_progress(i, total, name, step, state))

        def done(sent, failed):
            self.after(0, lambda: self._on_done(sent, failed))

        threading.Thread(
            target=engine.run_send,
            args=(self.app.ws, cfg, targets, sheet, log, progress, done),
            kwargs={"stop_event": self._stop},
            daemon=True).start()

    def _build_rows(self):
        for w in self._rows.winfo_children():
            w.destroy()
        self._rows_widgets = {}
        for i, t in enumerate(self._targets):
            line = ctk.CTkFrame(self._rows, fg_color="transparent", height=32)
            line.pack(fill="x")
            line.pack_propagate(False)
            name = T.label(line, t["name"] or f"{i+1}행", T.G800, 13)
            name.configure(width=170)
            name.pack(side="left", padx=(8, 0))
            s4 = T.Steps4(line)
            s4.pack(side="left", padx=(8, 0))
            res = T.label(line, "", T.G500, 12)
            res.pack(side="left", padx=(14, 0))
            T.hsep(self._rows)
            self._rows_widgets[i] = (s4, res, name)

    def _on_progress(self, i, total, name, step, state):
        w = self._rows_widgets.get(i)
        if w:
            s4, res, _ = w
            s4.set_step(step, state)
            if state == "bad":
                res.configure(text={1: "제안서 실패", 2: "PDF 실패",
                                    3: "기록 실패", 4: "메일 실패"}.get(step, "실패"),
                              text_color=T.RED)
            elif step == engine.STEP_MAIL and state == "ok":
                res.configure(text="보냄 · " + datetime.now().strftime("%H:%M"),
                              text_color=T.G500)
        if step == engine.STEP_PPT and state == "run":
            self._p_now.configure(text=f"{name} 처리 중")
            pct = i / max(total, 1)
            self._bar.set(pct)
            self._p_count.configure(text=f"{i} / {total}")
            self._p_pct.configure(text=f"{int(pct*100)}%")
            try:
                self._rows._parent_canvas.yview_moveto(
                    max(0, (i - 4) / max(total, 1)))
            except Exception:
                pass

    def _on_stop(self):
        self._stop.set()
        self._btn_stop.configure(state="disabled")
        self._p_now.configure(text="멈추는 중…")

    def _on_done(self, sent, failed):
        self.app.set_busy(False)
        self._bar.set(1)
        self._title.configure(text=f"{sent}곳에 보냈어요" if sent else "보내지 못했어요")
        self._sub.configure(text=datetime.now().strftime("%m월 %d일 %H:%M에 끝났어요"))
        self._s1.set_value(sent, T.GREEN);          self._s1.set_key("보냄")
        self._s2.set_value(failed, T.RED if failed else T.G400); self._s2.set_key("실패")
        self._s3.set_value(max(0, len(self._targets) - sent - failed), T.G400)
        self._s3.set_key("안 보냄")
        self._layout("done")
        self.app.refresh_rail()
        self.app.set_last_sent(datetime.now())


# ══════════════════════════════════════════════════════════
# 반송 확인
# ══════════════════════════════════════════════════════════

class BounceScreen(Screen):

    def build(self):
        self.head("되돌아온 메일을 확인해요",
                  "주소가 틀린 메일은 보낸 뒤 30분에서 1시간쯤 지나 되돌아와요. "
                  "찾은 주소는 명단에 '반송됨'으로 표시할게요.")

        top = ctk.CTkFrame(self, fg_color="transparent")
        top.pack(fill="x", pady=(0, 12))
        self._btn = T.Btn(top, "지금 확인하기", width=140, command=self._run)
        self._btn.pack(side="left")
        self._state = T.label(top, "", T.G500, 12)
        self._state.pack(side="left", padx=(14, 0))

        self._stats = ctk.CTkFrame(self, fg_color="transparent")
        self._stats.pack(fill="x", pady=(0, 12))
        self._s_ok = T.Stat(self._stats, "-", "잘 도착", T.GREEN)
        self._s_ok.pack(side="left", fill="x", expand=True, padx=(0, 6))
        self._s_bad = T.Stat(self._stats, "-", "되돌아옴", T.RED)
        self._s_bad.pack(side="left", fill="x", expand=True, padx=6)
        self._s_when = T.Stat(self._stats, "-", "마지막 확인", T.G400)
        self._s_when.pack(side="left", fill="x", expand=True, padx=(6, 0))

        card = T.Card(self)
        card.pack(fill="both", expand=True)
        T.subtitle(card, "되돌아온 주소").pack(anchor="w", padx=20, pady=(16, 8))
        self._table = T.Table(card, columns=[
            {"key": "name",   "title": "업체명", "width": 150},
            {"key": "email",  "title": "이메일", "width": 250},
            {"key": "reason", "title": "이유",   "width": 190},
        ], height=130)
        self._table.pack(fill="both", expand=True, padx=16, pady=(0, 12))

    def on_show(self):
        if not engine.gmail_connected():
            self._state.configure(text="Gmail을 먼저 연결해 주세요.")
            self._btn.configure(state="disabled")
        else:
            self._btn.configure(state="normal")
            self._state.configure(text="")

    def _run(self):
        sheet = self.app.reload_sheet()
        if not sheet:
            messagebox.showwarning("명단이 없어요", "1단계에서 명단을 올려 주세요.")
            return
        self._btn.configure(state="disabled")
        self._state.configure(text="확인하는 중…")
        cfg = dict(self.app.cfg)

        def work():
            try:
                hits = engine.check_bounces(
                    cfg, sheet, lambda m: self.after(0, lambda mm=m: self.app.log(mm)))
                self.after(0, lambda: self._done(hits, sheet))
            except Exception as e:
                self.after(0, lambda: self._fail(e))

        threading.Thread(target=work, daemon=True).start()

    def _done(self, hits, sheet):
        self._btn.configure(state="normal")
        self._state.configure(text="")
        marks = core.classify(sheet, self.app.cfg)
        sent_ok = sum(1 for m in marks if m["kind"] == core.ALREADY)
        self._s_ok.set_value(sent_ok)
        self._s_bad.set_value(len(hits))
        self._s_when.set_value(datetime.now().strftime("%H:%M"))
        self._table.set_rows([{**h, "_color_reason": T.RED} for h in hits])
        if not hits:
            self._state.configure(text="되돌아온 주소가 없어요.")

    def _fail(self, e):
        self._btn.configure(state="normal")
        self._state.configure(text="")
        messagebox.showerror("확인하지 못했어요", str(e))
