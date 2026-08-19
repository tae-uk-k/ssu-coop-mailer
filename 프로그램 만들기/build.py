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
import zipfile
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

    # 국원들에게 나눠 줄 zip. 구글 드라이브에 올려 링크를 공유한다.
    #
    # credentials.json 은 넣는다 — 개인정보가 아니라 프로그램 자체의 신분증이고,
    # 이게 없으면 구글 로그인 창조차 안 뜬다. 대신 GitHub 같은 공개된 곳에는
    # 올리지 않는다. 공개되면 구글이 키를 회수해 전원이 로그인을 못 하게 된다.
    share_dir = ROOT / "나눠줄 파일"
    share_dir.mkdir(exist_ok=True)
    zip_path = share_dir / "메일자동화_나눠주기.zip"
    if zip_path.exists():
        zip_path.unlink()
    has_cred = False
    with zipfile.ZipFile(zip_path, "w", zipfile.ZIP_DEFLATED) as z:
        for f in sorted(out.rglob("*")):
            if not f.is_file():
                continue
            rel = f.relative_to(out)
            if rel.name == "token.json":
                continue            # 로그인 열쇠 — 절대 공유 금지
            if rel.name == "credentials.json":
                has_cred = True
            z.write(f, Path("메일자동화") / rel)

    print(f"  나눠 줄 파일: {zip_path.name} ({zip_path.stat().st_size / 1e6:.0f}MB)")
    if has_cred:
        print("  구글 인증 파일이 들어 있습니다 → 구글 드라이브에 올려 링크를 공유하세요.")
        print("  ※ GitHub 같은 공개된 곳에는 올리지 마세요.")
    else:
        print("  [알림] 설정/credentials.json 이 없어 넣지 못했습니다.")

    # 중간 산물은 지운다. 다시 만들면 되고, 놔두면 200MB 넘게 쌓인다.
    for junk in (ROOT / "dist", ROOT / "build", out):
        shutil.rmtree(junk, ignore_errors=True)

    print()
    print("다 됐습니다.")
    print(f"이 파일을 구글 드라이브에 올려 링크를 공유하세요 → {zip_path}")
    print("(로그인 열쇠 token.json 과 계정 정보는 빠져 있습니다)")

    try:
        subprocess.Popen(["explorer", str(share_dir)])
    except Exception:
        pass
    return 0


if __name__ == "__main__":
    sys.exit(main())
