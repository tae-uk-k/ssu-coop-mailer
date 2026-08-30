"""
boot.py — 켜지는 과정 전부

run.py 는 exe 안에 박혀 원격으로 못 고친다. 그래서 run.py 에는
'프로그램 폴더를 찾아 여기로 넘긴다' 만 남기고, 실제로 하는 일은
전부 이 파일에 둔다. 이 파일은 자동 갱신으로 고칠 수 있다.

  하는 일
    1. 필요한 라이브러리가 있는지 확인
    2. 새 코드가 있으면 받아서 갈아끼움
    3. 앱을 켬
"""
from __future__ import annotations

import importlib
import json
import sys
from pathlib import Path

# 없으면 프로그램이 안 켜지는 것들
REQUIRED = [
    ("customtkinter",        "화면",         "customtkinter"),
    ("openpyxl",             "엑셀 명단",     "openpyxl"),
    ("googleapiclient",      "Gmail 발송",   "google-api-python-client"),
    ("google_auth_oauthlib", "Gmail 로그인", "google-auth-oauthlib"),
]

# 없어도 켜지지만 기능이 빠지는 것들
OPTIONAL = [
    ("win32com",    "PDF 변환",     "pywin32"),
    ("tkinterdnd2", "드래그앤드롭",  "tkinterdnd2"),
]


def tell(title: str, body: str) -> None:
    """켜지기 전에 생긴 문제를 알린다.

    pythonw 나 exe 로 켜면 검은 콘솔이 없어서 print 가 아무 데도 안 보인다.
    그래서 창으로 띄우고, 그마저 안 되면 콘솔로 떨어진다.
    """
    try:
        import tkinter as tk
        from tkinter import messagebox
        root = tk.Tk()
        root.withdraw()
        messagebox.showerror(title, body)
        root.destroy()
        return
    except Exception:
        pass

    print()
    print("=" * 56)
    print("  " + title)
    print("=" * 56)
    print(body)
    print()
    try:
        input("  엔터를 누르면 닫힙니다... ")
    except Exception:
        pass


def _check():
    missing, weak = [], []
    for mod, why, pkg in REQUIRED:
        try:
            importlib.import_module(mod)
        except Exception:
            missing.append((why, pkg))
    for mod, why, pkg in OPTIONAL:
        try:
            importlib.import_module(mod)
        except Exception:
            weak.append((why, pkg))
    return missing, weak


def _update_url(root: Path) -> str:
    try:
        cfg = json.loads((root / "설정" / "config.json").read_text(encoding="utf-8"))
        return str(cfg.get("update_url", "") or "")
    except Exception:
        return ""


def _save_update_url(root: Path, url: str) -> None:
    try:
        p = root / "설정" / "config.json"
        cfg = json.loads(p.read_text(encoding="utf-8"))
        cfg["update_url"] = url
        p.write_text(json.dumps(cfg, ensure_ascii=False, indent=2), encoding="utf-8")
    except Exception:
        pass


def _is_dev(root: Path) -> bool:
    """코드를 고치는 폴더인가.

    개발 폴더에서 갱신기가 돌면 고치던 코드가 릴리스 버전으로 덮어써진다.
    exe 로 쓰는 국원 PC 에는 .git 이 없으므로 정상 갱신된다.
    """
    return not getattr(sys, "frozen", False) and (root / ".git").exists()


def _self_update(root: Path, code_dir: Path) -> str:
    """무슨 일이 있어도 프로그램이 켜지는 걸 막지 않는다."""
    if _is_dev(root):
        return ""
    try:
        import updater
    except Exception:
        return ""

    note = ""
    try:
        note = updater.rollback_if_broken(root, code_dir)
    except Exception:
        pass

    url = _update_url(root)
    if not url:
        return note
    try:
        r = updater.check_and_apply(root, code_dir, url)
        if r.get("next_url"):
            _save_update_url(root, r["next_url"])      # 저장소를 옮겼을 때
        if r.get("updated"):
            note = "새 버전 " + r["version"] + " 을 적용했어요."
            if r.get("notes"):
                note += "  " + r["notes"]
    except Exception:
        pass
    return note


def main(root: Path, code_dir: Path, fixed: str = "") -> int:
    """run.py 가 부른다. fixed 는 코드를 되살렸을 때의 사연."""
    missing, weak = _check()

    if missing:
        body = ["아래가 설치되어 있지 않습니다.", ""]
        body += ["  · " + why + " → " + pkg for why, pkg in missing]
        body += [
            "",
            "아래 한 줄을 복사해 실행하면 한 번에 설치됩니다.",
            "",
            "pip install " + " ".join(pkg for _, pkg in missing + weak),
        ]
        tell("프로그램을 켤 수 없습니다", "\n".join(body))
        return 1

    # 갱신은 화면을 만들기 전에 — 그래야 이번 실행부터 새 코드가 쓰인다
    note = _self_update(root, code_dir)
    if fixed:
        note = (fixed + "  " + note).strip()
    if weak:
        note = (note + "  " + "  ".join(
            why + " 기능이 꺼집니다 (pip install " + pkg + ")"
            for why, pkg in weak)).strip()

    try:
        import app
    except Exception as e:
        tell("프로그램을 켜다가 문제가 생겼습니다",
             type(e).__name__ + ": " + str(e) + "\n\n"
             "받으신 폴더를 통째로 다시 받으면 대개 해결됩니다.")
        return 1

    app.main(startup_note=note)
    return 0
