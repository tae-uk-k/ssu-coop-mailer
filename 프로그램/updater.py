"""
updater.py — 켤 때 조용히 코드만 갱신한다

프로그램을 켜면 인터넷에서 목록 파일(manifest.json)을 받아 지금 코드와 견주고,
바뀐 .py 파일만 내려받아 갈아끼운다. exe(67MB)는 그대로 두고 코드(154KB)만 바꾼다.

지켜야 할 것 세 가지
  1. 인터넷이 안 되거나 주소가 틀려도 프로그램은 반드시 켜져야 한다  → 모두 무시하고 진행
  2. 받다 만 파일로 갈아끼우면 안 된다                              → 전부 받고 검사한 뒤 한 번에 교체
  3. 잘못된 코드가 올라가면 모두가 못 켠다                           → 직전 버전을 남겨 두고 되돌린다
"""
from __future__ import annotations

import hashlib
import json
import shutil
import ssl
import urllib.request
from pathlib import Path

TIMEOUT = 4.0          # 초 — 이보다 오래 걸리면 그냥 포기하고 켠다
MAX_FILE = 2_000_000   # 코드 파일 하나가 2MB 를 넘을 리 없다

VERSION_FILE = "version.json"
BACKUP_DIR   = "프로그램_직전"
TEMP_DIR     = ".받는중"
FLAG_FILE    = ".적용중"


# ──────────────────────────────────────────────────────────
# 지금 버전
# ──────────────────────────────────────────────────────────

def local_version(code_dir: Path) -> str:
    try:
        data = json.loads((code_dir / VERSION_FILE).read_text(encoding="utf-8"))
        return str(data.get("version", ""))
    except Exception:
        return ""


def _write_version(code_dir: Path, version: str, notes: str = "") -> None:
    (code_dir / VERSION_FILE).write_text(
        json.dumps({"version": version, "notes": notes},
                   ensure_ascii=False, indent=2), encoding="utf-8")


# ──────────────────────────────────────────────────────────
# 되돌리기 — 지난번에 켜다가 죽었으면 직전 코드로 복구
# ──────────────────────────────────────────────────────────

def rollback_if_broken(root: Path, code_dir: Path) -> str:
    """켜기 직전에 부른다. 되돌렸으면 그 사연을 문자열로 돌려준다."""
    flag = root / TEMP_DIR / FLAG_FILE
    backup = root / BACKUP_DIR
    if not flag.exists():
        return ""
    # 표시가 남아 있다 = 갈아끼운 뒤 제대로 못 켜졌다
    try:
        if backup.exists():
            for f in backup.glob("*.py"):
                shutil.copy2(f, code_dir / f.name)
            v = backup / VERSION_FILE
            if v.exists():
                shutil.copy2(v, code_dir / VERSION_FILE)
            flag.unlink(missing_ok=True)
            return f"새 버전이 제대로 켜지지 않아 직전 버전({local_version(backup)})으로 되돌렸어요."
    except Exception:
        pass
    flag.unlink(missing_ok=True)
    return ""


def mark_started(root: Path) -> None:
    """창이 무사히 뜨면 부른다. 이걸 지워야 다음에 되돌리지 않는다."""
    try:
        (root / TEMP_DIR / FLAG_FILE).unlink(missing_ok=True)
    except Exception:
        pass


# ──────────────────────────────────────────────────────────
# 내려받기
# ──────────────────────────────────────────────────────────

def _allowed(url: str) -> bool:
    """https 만 받는다. 남이 중간에서 코드를 바꿔치기하면 그대로 실행되기 때문이다.

    예외는 내 컴퓨터(127.0.0.1)뿐 — 배포 전에 시험할 때 쓴다.
    """
    u = (url or "").lower()
    if u.startswith("https://"):
        return True
    return u.startswith("http://127.0.0.1") or u.startswith("http://localhost")


def _fetch(url: str, limit: int = MAX_FILE) -> bytes:
    ctx = ssl.create_default_context()
    req = urllib.request.Request(url, headers={"User-Agent": "mail-automation-updater"})
    with urllib.request.urlopen(req, timeout=TIMEOUT, context=ctx) as r:
        return r.read(limit + 1)[:limit]


def _sha256(data: bytes) -> str:
    return hashlib.sha256(data).hexdigest()


def check_and_apply(root: Path, code_dir: Path, url: str) -> dict:
    """새 코드가 있으면 받아서 갈아끼운다.

    돌려주는 것: {"updated": bool, "version": str, "notes": str, "error": str}
    무슨 일이 있어도 예외를 밖으로 내보내지 않는다.
    """
    out = {"updated": False, "version": "", "notes": "", "error": "", "next_url": ""}
    if not _allowed(url):
        return out          # 주소가 없거나 https 가 아니면 아무것도 하지 않는다

    try:
        manifest = json.loads(_fetch(url).decode("utf-8"))
    except Exception as e:
        out["error"] = f"목록을 받지 못했어요: {type(e).__name__}"
        return out

    new_ver = str(manifest.get("version", "")).strip()
    files   = manifest.get("files") or {}
    if not new_ver or not isinstance(files, dict):
        out["error"] = "목록 형식이 올바르지 않아요."
        return out

    out["version"] = new_ver
    out["notes"] = str(manifest.get("notes", ""))

    # 저장소를 옮겼을 때 앱이 새 주소를 따라가게 한다.
    # 이게 없으면 주소가 바뀌는 순간 갱신이 영영 끊긴다.
    nxt = str(manifest.get("next_url", "") or "").strip()
    if nxt and nxt != url and _allowed(nxt):
        out["next_url"] = nxt

    if new_ver == local_version(code_dir):
        return out          # 이미 최신

    base = url.rsplit("/", 1)[0]
    tmp = root / TEMP_DIR
    shutil.rmtree(tmp, ignore_errors=True)
    tmp.mkdir(parents=True, exist_ok=True)

    # 1) 전부 받아서 검사부터 한다 (하나라도 어긋나면 아무것도 바꾸지 않는다)
    staged: dict[str, Path] = {}
    try:
        for name, info in files.items():
            if not name.endswith(".py") or "/" in name or "\\" in name or ".." in name:
                raise ValueError(f"이상한 파일 이름: {name}")
            want = str((info or {}).get("sha256", "")).lower()
            if len(want) != 64:
                raise ValueError(f"검사값이 없어요: {name}")
            data = _fetch(f"{base}/{name}")
            if _sha256(data) != want:
                raise ValueError(f"내려받은 파일이 손상됐어요: {name}")
            p = tmp / name
            p.write_bytes(data)
            staged[name] = p
    except Exception as e:
        shutil.rmtree(tmp, ignore_errors=True)
        out["error"] = f"내려받기 실패: {e}"
        return out

    # 2) 지금 코드를 백업하고
    try:
        backup = root / BACKUP_DIR
        shutil.rmtree(backup, ignore_errors=True)
        backup.mkdir(parents=True, exist_ok=True)
        for f in code_dir.glob("*.py"):
            shutil.copy2(f, backup / f.name)
        v = code_dir / VERSION_FILE
        if v.exists():
            shutil.copy2(v, backup / VERSION_FILE)
    except Exception as e:
        shutil.rmtree(tmp, ignore_errors=True)
        out["error"] = f"백업 실패: {e}"
        return out

    # 3) 한 번에 갈아끼운다
    try:
        tmp.mkdir(parents=True, exist_ok=True)
        (tmp / FLAG_FILE).write_text(new_ver, encoding="utf-8")   # 켜지면 지워진다
        for name, p in staged.items():
            shutil.copy2(p, code_dir / name)
        _write_version(code_dir, new_ver, out["notes"])
        out["updated"] = True
    except Exception as e:
        out["error"] = f"적용 실패: {e}"
        try:                                    # 되돌린다
            for f in (root / BACKUP_DIR).glob("*.py"):
                shutil.copy2(f, code_dir / f.name)
        except Exception:
            pass
        (tmp / FLAG_FILE).unlink(missing_ok=True)
    finally:
        for p in staged.values():
            try:
                p.unlink(missing_ok=True)
            except Exception:
                pass
    return out
