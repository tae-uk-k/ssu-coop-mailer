"""
exe 만들기

  배포파일_만들기.bat 을 더블클릭하면 이 파일이 실행됩니다.

끝나면 프로젝트 폴더에 '배포' 폴더가 생깁니다. 그 폴더를 통째로 전달하면 됩니다.
계정 정보가 든 token.json 과 config.json 은 일부러 넣지 않습니다.
"""
import json
import shutil
import subprocess
import sys
from pathlib import Path

HERE = Path(__file__).resolve().parent      # 프로그램 만들기/
ROOT = HERE.parent                          # 프로젝트 폴더


def run(*args) -> int:
    print(">", " ".join(str(a) for a in args))
    return subprocess.call(list(args), cwd=str(ROOT))


def main() -> int:
    print()
    print("[1/3] 필요한 것 설치 중…")
    run(sys.executable, "-m", "pip", "install", "pyinstaller", "--quiet")
    run(sys.executable, "-m", "pip", "install", "-r", str(HERE / "requirements.txt"), "--quiet")

    print()
    print("[2/3] exe 만드는 중 (3~10분 걸려요)…")
    rc = run(sys.executable, "-m", "PyInstaller", str(HERE / "메일자동화.spec"),
             "--noconfirm", "--clean",
             "--distpath", str(ROOT / "dist"), "--workpath", str(ROOT / "build"))

    exe = ROOT / "dist" / "메일자동화.exe"
    if rc != 0 or not exe.exists():
        print()
        print("만들지 못했습니다. 위에 나온 오류를 확인해 주세요.")
        return 1

    print()
    print("[3/3] 배포 폴더 만드는 중…")
    out = ROOT / "배포"
    (out / "설정").mkdir(parents=True, exist_ok=True)
    shutil.copy2(exe, out / "메일자동화.exe")

    # 앱 코드는 exe 안이 아니라 옆 폴더에 둔다. 그래야 원격 갱신이 적용된다.
    code_out = out / "프로그램"
    shutil.rmtree(code_out, ignore_errors=True)
    code_out.mkdir(parents=True)
    n = 0
    for f in sorted((ROOT / "프로그램").glob("*.py")):
        shutil.copy2(f, code_out / f.name)
        n += 1
    ver = ROOT / "프로그램" / "version.json"
    if ver.exists():
        shutil.copy2(ver, code_out / "version.json")
    print(f"  코드 {n}개를 배포/프로그램/ 에 넣었습니다 (갱신 대상)")

    cred = ROOT / "설정" / "credentials.json"
    if cred.exists():
        shutil.copy2(cred, out / "설정" / "credentials.json")
    else:
        print("  [알림] 설정/credentials.json 이 없어 넣지 못했습니다.")

    # 업데이트 주소만 담은 설정을 넣는다.
    # 이게 없으면 국원 PC 마다 주소를 손으로 넣어야 해서 자동 갱신이 무의미해진다.
    # 계정 정보(gmail_user)와 작업 이력(last_workspace)은 일부러 넣지 않는다.
    url = ""
    try:
        url = json.loads((ROOT / "설정" / "config.json").read_text(
            encoding="utf-8")).get("update_url", "")
    except Exception:
        pass
    (out / "설정" / "config.json").write_text(
        json.dumps({"gmail_user": "", "last_workspace": "", "update_url": url},
                   ensure_ascii=False, indent=2), encoding="utf-8")
    if url:
        print(f"  업데이트 주소를 넣었습니다: {url}")
    else:
        print("  [알림] 설정/config.json 에 update_url 이 없어 자동 갱신이 꺼진 채 나갑니다.")

    guide = ROOT / "사용설명서.md"
    if guide.exists():
        shutil.copy2(guide, out / "사용설명서.md")

    print()
    print(f"다 됐습니다 → {out}")
    print("이 '배포' 폴더를 통째로 전달하세요.")
    print("(계정 정보가 든 token.json 과 config.json 은 넣지 않았습니다)")

    try:
        subprocess.Popen(["explorer", str(out)])
    except Exception:
        pass
    return 0


if __name__ == "__main__":
    sys.exit(main())
