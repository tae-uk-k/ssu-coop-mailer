"""
메일 자동화 — 실행 파일

  실행하기.bat 을 더블클릭하거나,  pythonw run.py

────────────────────────────────────────────────────────────
이 파일은 exe 안에 박혀서 나가므로 원격으로 고칠 수 없습니다.
그래서 여기에는 '프로그램 폴더를 찾아 넘긴다' 만 두고,
실제로 하는 일은 전부 프로그램/boot.py 에 있습니다.
boot.py 는 자동 갱신으로 고칠 수 있습니다.

  이 파일을 고칠 일은 사실상 없습니다. 고치면 exe 를 다시 만들어야 합니다.
────────────────────────────────────────────────────────────

설정은 '설정' 폴더, 행사별 자료는 '행사별 작업' 폴더에 저장됩니다.
"""
import sys
from pathlib import Path

if getattr(sys, "frozen", False):
    ROOT = Path(sys.executable).resolve().parent
else:
    ROOT = Path(__file__).resolve().parent

CODE_DIR = ROOT / "프로그램"
sys.path.insert(0, str(CODE_DIR))

# 이 파일들이 없으면 아무것도 못 한다
NEEDED = ("boot.py", "app.py", "core.py", "engine.py", "theme.py", "screens.py")


def _tell(title: str, body: str) -> None:
    """boot.py 조차 못 불렀을 때만 쓰는 최소한의 알림."""
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
    print("\n" + title + "\n" + body + "\n")
    try:
        input("엔터를 누르면 닫힙니다... ")
    except Exception:
        pass


def _ensure_code() -> str:
    """'프로그램' 폴더가 성하지 않으면 직전 백업으로 되살린다."""
    if CODE_DIR.exists() and all((CODE_DIR / n).exists() for n in NEEDED):
        return ""

    backup = ROOT / "프로그램_직전"
    if backup.exists() and all((backup / n).exists() for n in NEEDED):
        try:
            import shutil
            CODE_DIR.mkdir(parents=True, exist_ok=True)
            for f in backup.glob("*.py"):
                shutil.copy2(f, CODE_DIR / f.name)
            v = backup / "version.json"
            if v.exists():
                shutil.copy2(v, CODE_DIR / "version.json")
            return "코드가 손상돼 직전 버전으로 되살렸어요."
        except Exception:
            pass

    lost = [n for n in NEEDED if not (CODE_DIR / n).exists()]
    body = ["'" + CODE_DIR.name + "' 폴더에 있어야 할 파일이 없습니다.", ""]
    body += ["  · " + n for n in lost]
    body += [
        "",
        "위치: " + str(CODE_DIR),
        "",
        "받으신 폴더를 통째로 다시 받아 주세요.",
        "(exe 만 옮기면 안 되고 '프로그램' 폴더가 옆에 있어야 합니다)",
    ]
    _tell("프로그램을 켤 수 없습니다", "\n".join(body))
    raise SystemExit(1)


def main() -> int:
    fixed = _ensure_code()
    try:
        import boot
    except Exception as e:
        _tell("프로그램을 켜다가 문제가 생겼습니다",
              type(e).__name__ + ": " + str(e) + "\n\n"
              "받으신 폴더를 통째로 다시 받으면 대개 해결됩니다.")
        return 1
    return boot.main(ROOT, CODE_DIR, fixed)


if __name__ == "__main__":
    sys.exit(main())
