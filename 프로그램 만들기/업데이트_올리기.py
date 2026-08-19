"""
업데이트 올릴 파일 만들기

  python "프로그램 만들기\업데이트_올리기.py"

'업로드' 폴더가 만들어집니다. 그 안의 파일을 전부 GitHub Release 에 올리면
국원들이 프로그램을 켤 때 자동으로 받아 갑니다.

exe 는 다시 만들 필요 없습니다. 코드(.py)만 바뀌니까요.
"""
import hashlib
import json
import shutil
import sys
from datetime import datetime
from pathlib import Path

HERE = Path(__file__).resolve().parent
ROOT = HERE.parent
CODE = ROOT / "프로그램"
OUT  = ROOT / "업로드"

# 올리지 않을 것 — updater 자신도 올린다(고칠 수 있어야 하므로)
SKIP = {"__pycache__"}


def sha256(p: Path) -> str:
    return hashlib.sha256(p.read_bytes()).hexdigest()


def main() -> int:
    if not CODE.exists():
        print("'프로그램' 폴더를 찾지 못했습니다.")
        return 1

    version = sys.argv[1] if len(sys.argv) > 1 else datetime.now().strftime("%Y.%m.%d.%H%M")
    notes = sys.argv[2] if len(sys.argv) > 2 else ""

    shutil.rmtree(OUT, ignore_errors=True)
    OUT.mkdir(parents=True)

    files = {}
    for p in sorted(CODE.glob("*.py")):
        if p.name in SKIP:
            continue
        shutil.copy2(p, OUT / p.name)
        files[p.name] = {"sha256": sha256(p), "size": p.stat().st_size}

    manifest = {"version": version, "notes": notes, "files": files}

    # 저장소를 옮겼다면 새 주소를 함께 알려 준다 (국원 앱이 따라간다)
    nxt = sys.argv[3] if len(sys.argv) > 3 else ""
    if nxt:
        manifest["next_url"] = nxt
        print(f"[알림] 다음부터는 이 주소를 보게 합니다: {nxt}")
    (OUT / "manifest.json").write_text(
        json.dumps(manifest, ensure_ascii=False, indent=2), encoding="utf-8")

    # 지금 폴더의 코드도 같은 버전으로 표시해 둔다 (개발 PC 가 스스로 갱신하지 않게)
    (CODE / "version.json").write_text(
        json.dumps({"version": version, "notes": notes},
                   ensure_ascii=False, indent=2), encoding="utf-8")

    total = sum(f["size"] for f in files.values())
    print()
    print(f"버전 {version}")
    if notes:
        print(f"바뀐 점: {notes}")
    print(f"파일 {len(files)}개 · 합쳐서 {total/1024:.0f}KB")
    for n in files:
        print(f"   · {n}")
    print()
    print(f"올릴 폴더 → {OUT}")
    print()
    print("  1. GitHub 저장소 → Releases → Draft a new release")
    print(f"  2. 태그를 '{version}' 로 적고, 위 폴더의 파일을 전부 끌어다 놓기")
    print("  3. Publish release")
    print()
    print("  국원들은 다음에 프로그램을 켤 때 자동으로 받아 갑니다.")

    try:
        import subprocess
        subprocess.Popen(["explorer", str(OUT)])
    except Exception:
        pass
    return 0


if __name__ == "__main__":
    sys.exit(main())
