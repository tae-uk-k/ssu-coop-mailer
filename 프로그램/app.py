"""
대외협력국 메일 자동발송

화면 구조
    1 파일 → 2 명단 → 3 제안서 → 4 메일 → 5 보내기 → (보낸 뒤) 반송 확인

모듈
    core.py     설정 · 워크스페이스 · 파일 보관함 · 조사 · 명단
    engine.py   PPT 자리 찾기/채우기 · PDF · Gmail · 발송
    theme.py    색 · 글꼴 · 공용 위젯
    screens.py  화면들
"""
from __future__ import annotations

import threading
from datetime import datetime
from tkinter import messagebox, simpledialog

import customtkinter as ctk

import core
import engine
import explain
import screens
import theme as T
from theme import font

ctk.set_appearance_mode("light")
ctk.set_default_color_theme("blue")

# 레일에 놓을 항목 — (화면 이름, 번호, 이름표)
RAIL = [
    ("files", 1, "파일"),
    ("list",  2, "명단"),
    ("slots", 3, "제안서"),
    ("mail",  4, "메일"),
    ("send",  5, "보내기"),
]


class App(ctk.CTk):

    def __init__(self, startup_note: str = ""):
        super().__init__()
        self._startup_note = startup_note
        self.title("메일 자동화")
        self.minsize(1180, 700)
        self.configure(fg_color=T.BG)

        w, h = 1120, 760
        x = max(0, (self.winfo_screenwidth() - w) // 2)
        y = max(0, (self.winfo_screenheight() - h) // 2 - 20)
        self.geometry(f"{w}x{h}+{x}+{y}")

        self.root_cfg: dict = core.load_root_cfg()
        self.cfg: dict = dict(core.WS_CFG_DEFAULTS)
        self.ws: str = ""
        self.sheet = None
        self.slots = None
        self._sheet_key = None
        self._log_lines: list[str] = []
        self._busy = False
        self._current = ""

        # 드래그앤드롭 (tkinterdnd2 가 있을 때만)
        self.dnd_ready = T.enable_dnd_root(self)

        self._build()

        last = self.root_cfg.get("last_workspace", "")
        if last in core.list_workspaces():
            self.switch_workspace(last, remember=False)
        else:
            self.go("welcome")

        self.after(300, self._check_powerpoint)
        self.after(400, self._after_started)

    # ══════════════════════════════════════════════════════
    # 틀
    # ══════════════════════════════════════════════════════

    def _build(self):
        self._build_topbar()
        self._build_statusbar()

        body = ctk.CTkFrame(self, fg_color="transparent")
        body.pack(fill="both", expand=True)
        body.columnconfigure(1, weight=1)
        body.rowconfigure(0, weight=1)

        self._build_rail(body)

        self._stage = ctk.CTkFrame(body, fg_color=T.BG)
        self._stage.grid(row=0, column=1, sticky="nsew")

        self._screens: dict[str, screens.Screen] = {}
        for name, cls in (("welcome", screens.WelcomeScreen),
                          ("files",   screens.FilesScreen),
                          ("list",    screens.ListScreen),
                          ("slots",   screens.SlotsScreen),
                          ("mail",    screens.MailScreen),
                          ("send",    screens.SendScreen),
                          ("bounce",  screens.BounceScreen)):
            self._screens[name] = cls(self, self._stage)

    def _build_topbar(self):
        bar = ctk.CTkFrame(self, fg_color=T.WHITE, corner_radius=0, height=62)
        bar.pack(fill="x")
        bar.pack_propagate(False)

        inner = ctk.CTkFrame(bar, fg_color="transparent")
        inner.pack(fill="both", expand=True, padx=20)

        ctk.CTkLabel(inner, text="메일 자동화", font=font(16, True),
                     text_color=T.G900).pack(side="left")

        self._ws_btn = ctk.CTkButton(
            inner, text="행사를 골라 주세요", font=font(13, True),
            fg_color=T.G50, hover_color=T.G100, text_color=T.G800,
            corner_radius=10, height=36, width=210, anchor="w",
            command=self._workspace_menu)
        self._ws_btn.pack(side="left", padx=(12, 0))

        ctk.CTkButton(inner, text="?", font=font(13, True), width=32, height=32,
                      corner_radius=10, fg_color=T.G50, hover_color=T.G100,
                      text_color=T.G600, command=self._help).pack(side="right")

        self._acct_btn = ctk.CTkButton(
            inner, text="Gmail 연결하기", font=font(12), anchor="w",
            fg_color=T.G50, hover_color=T.G100, text_color=T.G700,
            corner_radius=10, height=34, width=230,
            command=self._gmail_clicked)
        self._acct_btn.pack(side="right", padx=(0, 8))

        ctk.CTkFrame(self, height=1, fg_color=T.G100).pack(fill="x")

    def _build_rail(self, parent):
        rail = ctk.CTkFrame(parent, fg_color=T.WHITE, corner_radius=0, width=232)
        rail.grid(row=0, column=0, sticky="nsw")
        rail.grid_propagate(False)

        ctk.CTkLabel(rail, text="보내기까지", font=font(12, True),
                     text_color=T.G400, anchor="w").pack(fill="x", padx=24, pady=(22, 10))

        self._rail_items: dict[str, dict] = {}
        for name, num, text in RAIL:
            self._rail_items[name] = self._rail_row(rail, name, num, text)

        ctk.CTkFrame(rail, height=1, fg_color=T.G100).pack(fill="x", padx=22, pady=16)
        ctk.CTkLabel(rail, text="보낸 뒤", font=font(12, True),
                     text_color=T.G400, anchor="w").pack(fill="x", padx=24, pady=(0, 10))
        self._rail_items["bounce"] = self._rail_row(rail, "bounce", 0, "반송 확인")

        foot = ctk.CTkFrame(rail, fg_color="transparent")
        foot.pack(side="bottom", fill="x", padx=22, pady=20)
        line = ctk.CTkFrame(foot, fg_color="transparent")
        line.pack(fill="x", pady=(0, 7))
        self._prog_txt = ctk.CTkLabel(line, text="0 / 5 단계", font=font(12),
                                      text_color=T.G500, anchor="w")
        self._prog_txt.pack(side="left")
        self._prog_pct = ctk.CTkLabel(line, text="0%", font=font(12),
                                      text_color=T.G500, anchor="e")
        self._prog_pct.pack(side="right")
        self._prog = T.Bar(foot, height=4)
        self._prog.pack(fill="x")

    def _rail_row(self, parent, name: str, num: int, text: str) -> dict:
        holder = ctk.CTkFrame(parent, fg_color="transparent", corner_radius=12,
                              height=44)
        holder.pack(fill="x", padx=14, pady=1)
        holder.pack_propagate(False)

        badge = ctk.CTkLabel(holder, text=str(num) if num else "·",
                             font=font(12, True), width=24, height=24,
                             corner_radius=12, fg_color=T.G100, text_color=T.G500)
        badge.pack(side="left", padx=(12, 10))
        lbl = ctk.CTkLabel(holder, text=text, font=font(13), text_color=T.G600,
                           anchor="w")
        lbl.pack(side="left")
        tag = ctk.CTkLabel(holder, text="", font=font(12), text_color=T.G400,
                           anchor="e")
        tag.pack(side="right", padx=(0, 12))

        item = {"holder": holder, "badge": badge, "label": lbl, "tag": tag,
                "num": num, "text": text}
        for w in (holder, badge, lbl, tag):
            w.bind("<Button-1>", lambda _e, n=name: self._rail_click(n))
            w.configure(cursor="hand2")
        return item

    def _build_statusbar(self):
        ctk.CTkFrame(self, height=1, fg_color=T.G100).pack(fill="x", side="bottom")
        bar = ctk.CTkFrame(self, fg_color=T.WHITE, corner_radius=0, height=38)
        bar.pack(fill="x", side="bottom")
        bar.pack_propagate(False)
        inner = ctk.CTkFrame(bar, fg_color="transparent")
        inner.pack(fill="both", expand=True, padx=20)

        self._sb_msg = ctk.CTkLabel(inner, text="", font=font(12),
                                    text_color=T.G500, anchor="w")
        self._sb_msg.pack(side="left")

        self._sb_time = ctk.CTkLabel(inner, text="아직 보낸 적 없어요", font=font(12),
                                     text_color=T.G500, anchor="e")
        self._sb_time.pack(side="right")

        ctk.CTkButton(inner, text="기록", font=font(11), width=44, height=24,
                      corner_radius=8, fg_color=T.G50, hover_color=T.G100,
                      text_color=T.G600, command=self._show_log).pack(side="right", padx=(0, 10))

        self._sb_ppt = ctk.CTkLabel(inner, text="", font=font(12),
                                    text_color=T.G500, anchor="e")
        self._sb_ppt.pack(side="right", padx=(0, 14))

    # ══════════════════════════════════════════════════════
    # 화면 이동
    # ══════════════════════════════════════════════════════

    def go(self, name: str) -> None:
        if name not in self._screens:
            return
        if self._current and self._current in self._screens:
            try:
                self._screens[self._current].on_leave()
            except Exception:
                pass
            self._screens[self._current].pack_forget()

        self._current = name
        scr = self._screens[name]
        scr.pack(fill="both", expand=True, padx=30, pady=26)
        try:
            scr.on_show()
        except Exception as e:
            self.log("화면을 여는 중 문제가 생겼어요 — " + explain.short(e))
        self.refresh_rail()

    def _rail_click(self, name: str):
        if not self.ws:
            messagebox.showinfo("행사를 먼저 만들어 주세요",
                                "위쪽 행사 칸에서 새 행사를 만들거나 골라 주세요.")
            return
        if self._busy:
            return
        self.go(name)

    def refresh_rail(self) -> None:
        done = core.ready_step(self.ws, self.cfg) if self.ws else 0

        for name, item in self._rail_items.items():
            num = item["num"]
            active = (name == self._current)
            complete = bool(num) and num <= done and not active
            locked = not self.ws

            if locked:
                item["holder"].configure(fg_color="transparent")
                item["badge"].configure(fg_color=T.G100, text_color=T.G300,
                                        text=str(num) if num else "·")
                item["label"].configure(text_color=T.G300)
                item["tag"].configure(text="")
                continue

            if active:
                item["holder"].configure(fg_color=T.BLUE_BG)
                item["badge"].configure(fg_color=T.BLUE, text_color=T.WHITE,
                                        text=str(num) if num else "·")
                item["label"].configure(text_color=T.BLUE_D, font=font(13, True))
            elif complete:
                item["holder"].configure(fg_color="transparent")
                item["badge"].configure(fg_color=T.BLUE, text_color=T.WHITE, text="✓")
                item["label"].configure(text_color=T.G800, font=font(13))
            else:
                item["holder"].configure(fg_color="transparent")
                item["badge"].configure(fg_color=T.G100, text_color=T.G500,
                                        text=str(num) if num else "·")
                item["label"].configure(text_color=T.G600, font=font(13))

        # 곁들임 숫자
        self._rail_items["list"]["tag"].configure(text=self._count_label())
        slots = self.slots or []
        mapped = sum(1 for v in (self.cfg.get("slot_map") or {}).values() if v)
        self._rail_items["slots"]["tag"].configure(
            text=f"{mapped}/{len(slots)}" if slots else "")

        pct = int(done / 5 * 100)
        self._prog_txt.configure(text=f"{done} / 5 단계" if self.ws else "행사를 먼저 만들어요")
        self._prog_pct.configure(text=f"{pct}%")
        self._prog.set(pct / 100)

    def _count_label(self) -> str:
        if not self.sheet or not self.cfg.get("col_email"):
            return ""
        try:
            marks = core.classify(self.sheet, self.cfg)
            n = sum(1 for m in marks if m["kind"] == core.SEND)
            return f"{n}곳"
        except Exception:
            return ""

    # ══════════════════════════════════════════════════════
    # 워크스페이스
    # ══════════════════════════════════════════════════════

    def _workspace_menu(self):
        names = core.list_workspaces()
        win = ctk.CTkToplevel(self)
        win.title("행사 고르기")
        win.configure(fg_color=T.BG)
        win.geometry("380x420")
        win.transient(self)
        win.grab_set()

        ctk.CTkLabel(win, text="행사", font=font(17, True),
                     text_color=T.G900).pack(anchor="w", padx=22, pady=(20, 12))

        box = ctk.CTkScrollableFrame(win, fg_color="transparent")
        box.pack(fill="both", expand=True, padx=16)

        for n in names:
            rowf = T.Card(box, corner_radius=12)
            rowf.pack(fill="x", pady=(0, 7))
            ins = ctk.CTkFrame(rowf, fg_color="transparent")
            ins.pack(fill="x", padx=14, pady=11)
            T.label(ins, n, T.G900 if n != self.ws else T.BLUE_D, 13, True).pack(side="left")
            if n == self.ws:
                T.Tag(ins, "지금 열림", "blue").pack(side="left", padx=(8, 0))
            ctk.CTkButton(ins, text="삭제", font=font(11), width=42, height=24,
                          corner_radius=8, fg_color=T.G100, hover_color=T.RED_BG,
                          text_color=T.G500,
                          command=lambda x=n, w=win: self._delete_workspace(x, w)
                          ).pack(side="right")
            ctk.CTkButton(ins, text="열기", font=font(11), width=42, height=24,
                          corner_radius=8, fg_color=T.G100, hover_color=T.BLUE_BG,
                          text_color=T.G700,
                          command=lambda x=n, w=win: (w.destroy(), self.switch_workspace(x))
                          ).pack(side="right", padx=(0, 6))

        T.Btn(win, "새 행사 만들기", width=200,
              command=lambda: (win.destroy(), self.new_workspace())).pack(pady=16)

    def new_workspace(self):
        name = simpledialog.askstring("새 행사", "행사 이름을 적어 주세요\n(예: 키움 히어로즈 DAY)",
                                      parent=self)
        if not name:
            return
        name = name.strip()
        if not name or any(c in name for c in '\\/:*?"<>|'):
            messagebox.showwarning("쓸 수 없는 이름", '\\ / : * ? " < > | 는 쓸 수 없어요.')
            return
        if not core.create_workspace(name):
            messagebox.showwarning("이미 있어요", f"'{name}' 행사가 이미 있어요.")
            return
        self.switch_workspace(name)

    def switch_workspace(self, name: str, remember: bool = True):
        self.ws = name
        self.cfg = core.load_ws_cfg(name)
        self.sheet = None
        self.slots = None
        self._sheet_key = None
        self._ws_btn.configure(text=f"  {name}   ▾")

        if remember:
            self.root_cfg["last_workspace"] = name
            core.save_root_cfg(self.root_cfg)

        self.reload_sheet(force=True)
        self._refresh_account()

        step = core.ready_step(name, self.cfg)
        self.go(RAIL[min(step, 4)][0])

    def _delete_workspace(self, name: str, win):
        if not messagebox.askyesno(
                "지울까요?",
                f"'{name}' 행사를 지웁니다.\n"
                "보관 중인 파일과 만들어 둔 제안서가 함께 지워져요.\n계속할까요?",
                parent=win):
            return
        core.delete_workspace(name)
        win.destroy()
        if name == self.ws:
            self.ws = ""
            self.cfg = dict(core.WS_CFG_DEFAULTS)
            self.sheet = None
            self.slots = None
            self._ws_btn.configure(text="행사를 골라 주세요")
            self.root_cfg["last_workspace"] = ""
            core.save_root_cfg(self.root_cfg)
            self.go("welcome")

    # ══════════════════════════════════════════════════════
    # 데이터
    # ══════════════════════════════════════════════════════

    def save_cfg(self):
        if self.ws:
            core.save_ws_cfg(self.ws, self.cfg)

    def reload_sheet(self, force: bool = False):
        rel = self.cfg.get("recipients_xlsx", "")
        if not rel or not self.ws:
            self.sheet = None
            self._sheet_key = None
            return None
        path = core.resolve(self.ws, rel)
        if not path.exists():
            self.sheet = None
            return None
        key = (str(path), path.stat().st_mtime)
        if not force and self.sheet is not None and self._sheet_key == key:
            return self.sheet
        try:
            self.sheet = core.load_sheet(path)
            self._sheet_key = key
        except Exception as e:
            self.sheet = None
            self.log("명단을 읽지 못했어요 — " + explain.short(e))
        return self.sheet

    def get_slots(self):
        """PPT 자리는 한 번만 훑고 기억해 둔다."""
        if self.slots is not None:
            return self.slots
        rel = self.cfg.get("template_pptx", "")
        if not rel or not self.ws:
            self.slots = []
            return self.slots
        path = core.resolve(self.ws, rel)
        if not path.exists():
            self.slots = []
            return self.slots
        try:
            self.slots = engine.scan_slots(path)
        except Exception as e:
            self.slots = []
            self.log("제안서를 훑지 못했어요 — " + explain.short(e))
        return self.slots

    # ══════════════════════════════════════════════════════
    # Gmail
    # ══════════════════════════════════════════════════════

    def _refresh_account(self):
        if engine.gmail_connected():
            addr = self.root_cfg.get("gmail_user", "")
            self._acct_btn.configure(text=f"  ● {addr or '연결됨'}",
                                     text_color=T.G700)
        else:
            self._acct_btn.configure(text="  ○ Gmail 연결하기", text_color=T.G500)

    def _gmail_clicked(self):
        if engine.gmail_connected():
            addr = self.root_cfg.get("gmail_user", "") or "(주소 확인 중)"
            if messagebox.askyesno(
                    "Gmail 연결",
                    f"지금 계정: {addr}\n\n연결을 끊고 다른 계정으로 바꿀까요?"):
                engine.gmail_disconnect()
                self.root_cfg["gmail_user"] = ""
                core.save_root_cfg(self.root_cfg)
                self._refresh_account()
                self.refresh_rail()
            return
        self._connect_gmail()

    def _connect_gmail(self):
        if not core.CREDENTIALS_PATH.exists():
            messagebox.showerror(
                "credentials.json 이 없어요",
                "구글에서 받은 credentials.json 파일을\n"
                f"{core.BASE_DIR} 폴더에 넣어 주세요.")
            return
        self.log("브라우저에서 구글 로그인을 진행해 주세요…")

        def work():
            try:
                engine.gmail_service(force=True)
                addr = engine.gmail_address()
                self.after(0, lambda: self._gmail_ok(addr))
            except Exception as e:
                self.after(0, lambda: self._gmail_fail(e))

        threading.Thread(target=work, daemon=True).start()

    def _gmail_ok(self, addr: str):
        self.root_cfg["gmail_user"] = addr
        core.save_root_cfg(self.root_cfg)
        self._refresh_account()
        self.log(f"{addr} 계정을 연결했어요.")
        if self._current in ("welcome", "send", "bounce"):
            self.go(self._current)

    def _gmail_fail(self, e):
        self.log("")
        what, how = explain.explain(e)
        T.show_error(self, "Gmail에 연결하지 못했어요", what, how, explain.detail(e))

    # ══════════════════════════════════════════════════════
    # 기타
    # ══════════════════════════════════════════════════════

    def log(self, msg: str):
        if not msg:
            self._sb_msg.configure(text="")
            return
        stamp = datetime.now().strftime("%H:%M:%S")
        self._log_lines.append(f"[{stamp}] {msg}")
        self._sb_msg.configure(text=msg[:70])

    def _show_log(self):
        win = ctk.CTkToplevel(self)
        win.title("작업 기록")
        win.geometry("640x430")
        win.configure(fg_color=T.BG)
        win.transient(self)
        box = T.TextBox(win, fg_color=T.WHITE)
        box.pack(fill="both", expand=True, padx=16, pady=16)
        box.insert("1.0", "\n".join(self._log_lines) or "아직 기록이 없어요.")
        box.configure(state="disabled")

    def set_busy(self, busy: bool):
        self._busy = busy
        self._ws_btn.configure(state="disabled" if busy else "normal")

    def set_last_sent(self, when: datetime):
        self._sb_time.configure(text=when.strftime("마지막 발송 %m월 %d일 %H:%M"))

    def _after_started(self):
        """창이 무사히 떴다는 표시. 이걸 남겨야 다음에 되돌리지 않는다."""
        try:
            import updater
            updater.mark_started(core.BASE_DIR)
        except Exception:
            pass
        if self._startup_note:
            self.log(" ".join(self._startup_note.split()))

    def _check_powerpoint(self):
        if engine.powerpoint_available():
            self._sb_ppt.configure(text="PowerPoint 사용할 수 있어요")
        else:
            self._sb_ppt.configure(text="PowerPoint 없음 — PDF 변환 불가",
                                   text_color=T.RED)

    def _help(self):
        messagebox.showinfo(
            "도움말",
            "왼쪽 단계를 위에서부터 채우면 보낼 준비가 끝나요.\n\n"
            "1 파일     제안서 PPT와 기업 명단을 올려요. 앱이 보관해요.\n"
            "2 명단     이메일이 어느 열인지 고르고, 표에서 바로 고쳐요.\n"
            "3 제안서   PPT에서 찾은 (자리)에 명단의 어느 열을 넣을지 골라요.\n"
            "4 메일     제목과 본문을 쓰면 오른쪽에 실제 모습이 보여요.\n"
            "5 보내기   확인 후 보내요. 중간에 멈춰도 보낸 곳은 건너뛰어요.\n\n"
            "보낸 뒤 30분~1시간 지나 '반송 확인'을 눌러 주세요.")


def main(startup_note: str = "") -> None:
    core.ensure_dirs()
    App(startup_note=startup_note).mainloop()


if __name__ == "__main__":
    main()
