"""
메일 자동화 — 실행 파일

  실행하기.bat 을 더블클릭하거나,  python run.py

이 파일 이름만 영어인 이유가 있습니다. 배치 파일(.bat)은 한글이 들어가면
쓰는 PC의 언어 설정에 따라 깨져서 실행이 안 되는 경우가 있습니다.
그래서 배치 파일은 영어로만 두고, 한글 안내는 여기(파이썬)에서 합니다.

실제 코드는 '프로그램' 폴더 안에 있습니다.
설정은 '설정' 폴더, 행사별 자료는 '행사별 작업' 폴더에 저장됩니다.

켤 때마다 새 코드가 있는지 조용히 확인해 갈아끼웁니다. 자세한 것은
프로그램/updater.py 를 보세요. 인터넷이 안 되면 그냥 넘어갑니다.
"""
import importlib
import json
import sys
from pathlib import Path

if getattr(sys, "frozen", False):
    ROOT = Path(sys.executable).resolve().parent
else:
    ROOT = Path(__file__).resolve().parent

CODE_DIR = ROOT / "프로그램"
sys.path.insert(0, str(CODE_DIR))

# 없으면 프로그램이 안 켜지는 것들
REQUIRED = [
    ("customtkinter",   "화면",            "customtkinter"),
    ("openpyxl",        "엑셀 명단",        "openpyxl"),
    ("googleapiclient", "Gmail 발송",       "google-api-python-client"),
    ("google_auth_oauthlib", "Gmail 로그인", "google-auth-oauthlib"),
]

# 없어도 켜지지만 기능이 빠지는 것들
OPTIONAL = [
    ("win32com",     "PDF 변환",      "pywin32"),
    ("tkinterdnd2",  "드래그앤드롭",   "tkinterdnd2"),
]


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


def _update_url() -> str:
    """설정/config.json 의 update_url. 비어 있으면 갱신을 아예 하지 않는다."""
    try:
        cfg = json.loads((ROOT / "설정" / "config.json").read_text(encoding="utf-8"))
        return str(cfg.get("update_url", "") or "")
    except Exception:
        return ""


def _self_update() -> str:
    """켤 때 한 번. 무슨 일이 있어도 프로그램이 켜지는 걸 막지 않는다."""
    try:
        import updater
    except Exception:
        return ""

    note = ""
    try:
        note = updater.rollback_if_broken(ROOT, CODE_DIR)
    except Exception:
        pass

    url = _update_url()
    if not url:
        return note
    try:
        r = updater.check_and_apply(ROOT, CODE_DIR, url)
        if r.get("updated"):
            note = f"새 버전 {r['version']} 을 적용했어요." + (
                f"\n{r['notes']}" if r.get("notes") else "")
    except Exception:
        pass
    return note


NEEDED = ("app.py", "core.py", "engine.py", "theme.py", "screens.py")


def _ensure_code() -> str:
    """'프로그램' 폴더가 성하지 않으면 직전 백업으로 되살린다.

    이 파일(run.py)은 exe 안에 박혀 있어 원격으로 못 고친다. 그래서
    코드가 통째로 사라지는 최악의 경우까지 여기서 스스로 수습해야 한다.
    """
    have = CODE_DIR.exists() and all((CODE_DIR / n).exists() for n in NEEDED)
    if have:
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
    print()
    print("=" * 56)
    print("  프로그램을 켤 수 없습니다.")
    print("=" * 56)
    print(f"  '{CODE_DIR.name}' 폴더에 있어야 할 파일이 없습니다.")
    for n in lost:
        print(f"    · {n}")
    print()
    print(f"  위치: {CODE_DIR}")
    print()
    print("  받으신 폴더를 통째로 다시 받아 주세요.")
    print("  (exe 만 옮기면 안 되고, '프로그램' 폴더가 옆에 있어야 합니다)")
    print()
    input("  엔터를 누르면 닫힙니다... ")
    raise SystemExit(1)


def main() -> int:
    fixed = _ensure_code()
    missing, weak = _check()

    if missing:
        print()
        print("=" * 56)
        print("  프로그램을 켤 수 없습니다. 아래가 설치되어 있지 않습니다.")
        print("=" * 56)
        for why, pkg in missing:
            print(f"    · {why:14s} → {pkg}")
        print()
        print("  아래 한 줄을 복사해서 실행하면 한 번에 설치됩니다.")
        print()
        print("    pip install " + " ".join(pkg for _, pkg in missing + weak))
        print()
        input("  엔터를 누르면 닫힙니다... ")
        return 1

    if weak:
        for why, pkg in weak:
            print(f"[알림] {why} 기능이 꺼집니다. 쓰려면:  pip install {pkg}")

    # 코드 갱신은 화면을 만들기 전에 — 그래야 이번 실행부터 새 코드가 쓰인다
    note = _self_update()
    if fixed:
        note = (fixed + "  " + note).strip()

    try:
        import app
    except Exception as e:
        print()
        print("=" * 56)
        print("  프로그램을 켜다가 문제가 생겼습니다.")
        print("=" * 56)
        print(f"  {type(e).__name__}: {e}")
        print()
        print("  받으신 폴더를 통째로 다시 받으면 대개 해결됩니다.")
        print()
        input("  엔터를 누르면 닫힙니다... ")
        return 1

    app.main(startup_note=note)
    return 0


if __name__ == "__main__":
    sys.exit(main())
