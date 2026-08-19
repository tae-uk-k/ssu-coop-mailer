"""
explain.py — 오류를 쉬운 말로 바꿔 준다

파이썬이 내는 오류는 이렇게 생겼다.
    [Errno 13] Permission denied: 'C:\\...\\명단.xlsx'
    <HttpError 400 ... returned "Invalid To header">

이걸 그대로 보여주면 국원분들은 무슨 일인지도, 뭘 해야 하는지도 알 수 없다.
그래서 여기서 '무슨 일이 있었는지 / 어떻게 하면 되는지' 두 줄로 바꿔 준다.
"""
from __future__ import annotations

import re

# (무슨 일, 어떻게 하면 되는지)
Explained = tuple[str, str]

_UNKNOWN = ("알 수 없는 문제가 생겼어요.",
            "다시 한 번 해 보시고, 계속 그러면 아래 '자세한 내용'을 대외협력국장에게 보여 주세요.")


def _text(e: BaseException) -> str:
    return f"{type(e).__name__} {e}"


def explain(e: BaseException, doing: str = "") -> Explained:
    """예외 하나를 사람 말로. doing 은 '메일 보내기' 처럼 뭘 하다 났는지."""
    name = type(e).__name__
    msg = str(e)
    low = msg.lower()

    # ── 파일 ──
    if isinstance(e, PermissionError) or "permission denied" in low:
        target = _filename(msg)
        what = f"{target or '파일'}을(를) 쓸 수 없어요."
        return what, ("그 파일이 다른 프로그램에서 열려 있는 것 같아요.\n"
                      "엑셀이나 PowerPoint 창을 모두 닫고 다시 해 주세요.")

    if isinstance(e, FileNotFoundError) or "no such file" in low:
        target = _filename(msg)
        return (f"{target or '파일'}을(를) 찾지 못했어요.",
                "1단계에서 파일을 다시 올려 주세요.\n"
                "원본을 지우거나 옮기셨다면 새로 올리시면 됩니다.")

    if name in ("BadZipFile", "InvalidFileException") or "not a zip file" in low:
        return ("파일이 손상됐거나 형식이 달라요.",
                "제안서는 .pptx, 명단은 .xlsx 여야 해요.\n"
                "예전 형식(.ppt, .xls)이면 PowerPoint·엑셀에서 열어 새 형식으로 저장해 주세요.")

    if isinstance(e, OSError) and ("No space" in msg or "28" in msg[:12]):
        return ("저장할 공간이 부족해요.",
                "디스크 여유 공간을 확보한 뒤 다시 해 주세요.")

    # ── 인터넷 ──
    if name in ("URLError", "HTTPError", "TimeoutError", "socket.timeout",
                "ConnectionError", "ConnectionResetError", "ConnectionRefusedError") \
            or "urlopen error" in low or "timed out" in low or "getaddrinfo" in low:
        return ("인터넷에 연결하지 못했어요.",
                "와이파이가 연결돼 있는지 확인하고 다시 해 주세요.\n"
                "학교 와이파이에서 막히면 휴대폰 핫스팟으로 해 보세요.")

    if "certificate" in low or name == "SSLError":
        return ("보안 연결에 실패했어요.",
                "컴퓨터의 날짜·시간이 맞는지 확인해 주세요.\n"
                "회사·학교 보안 프로그램이 막는 경우도 있어요.")

    # ── Gmail ──
    g = _gmail(e, msg, low)
    if g:
        return g

    # ── PowerPoint ──
    if name == "com_error" or "powerpoint" in low or "presentations" in low:
        return ("PowerPoint를 다루지 못했어요.",
                "열려 있는 PowerPoint 창을 모두 닫고 다시 해 주세요.\n"
                "그래도 안 되면 컴퓨터를 다시 켠 뒤 시도해 보세요.")

    if isinstance(e, ImportError) and "win32" in low:
        return ("PowerPoint를 쓸 준비가 안 됐어요.",
                "PDF로 바꾸려면 PowerPoint가 설치돼 있어야 해요.")

    # ── 명단 ──
    if isinstance(e, KeyError):
        col = msg.strip("'\"")
        return (f"명단에서 '{col}' 열을 찾지 못했어요.",
                "2단계에서 열을 다시 골라 주세요.\n"
                "엑셀의 열 이름을 바꾸셨다면 새로 지정해야 해요.")

    if isinstance(e, MemoryError):
        return ("파일이 너무 커서 처리하지 못했어요.",
                "제안서 PPT의 사진 용량을 줄인 뒤 다시 올려 주세요.")

    return _UNKNOWN


def _filename(msg: str) -> str:
    m = re.search(r"['\"]([^'\"]*[\\/])?([^'\"\\/]+\.\w{2,5})['\"]", msg)
    return m.group(2) if m else ""


def _gmail(e: BaseException, msg: str, low: str) -> Explained | None:
    """Gmail API 가 내는 오류. HttpError 는 상태코드가 안에 들어 있다."""
    status = None
    m = re.search(r"HttpError\s+(\d{3})", msg)
    if m:
        status = int(m.group(1))
    elif hasattr(e, "resp") and hasattr(e.resp, "status"):
        try:
            status = int(e.resp.status)
        except Exception:
            pass

    if status is None and "credentials" not in low and "token" not in low:
        return None

    if status == 400 and ("to header" in low or "invalid to" in low or "recipient" in low):
        return ("받는 사람 주소가 올바르지 않아요.",
                "2단계 명단에서 그 기업의 이메일을 확인해 주세요.\n"
                "빈칸이거나 @ 가 빠졌을 수 있어요.")
    if status in (401, 403) or "invalid_grant" in low or "unauthorized" in low:
        return ("Gmail 로그인이 풀렸어요.",
                "오른쪽 위 계정 칸을 눌러 다시 연결해 주세요.")
    if status == 429 or "rate" in low or "quota" in low or "limit" in low:
        return ("Gmail이 잠시 발송을 막았어요.",
                "너무 빨리 많이 보내면 이렇게 돼요.\n"
                "5단계에서 '보내는 간격'을 10초 이상으로 늘리고,\n"
                "1~2시간 뒤에 남은 곳만 다시 보내 주세요.\n"
                "(무료 Gmail은 하루 500통까지예요)")
    if status and 500 <= status < 600:
        return ("Gmail 쪽에 일시적인 문제가 있어요.",
                "잠시 뒤 다시 해 주세요. 이미 보낸 곳은 건너뜁니다.")
    if "credentials.json" in low or "client_secrets" in low:
        return ("구글 인증 파일이 없어요.",
                "'설정' 폴더에 credentials.json 을 넣어 주세요.\n"
                "대외협력국장에게 파일을 요청하시면 됩니다.")
    if status:
        return (f"Gmail이 요청을 거절했어요. (코드 {status})",
                "잠시 뒤 다시 해 주세요.\n"
                "계속 그러면 오른쪽 위 계정을 다시 연결해 보세요.")
    return None


# ──────────────────────────────────────────────────────────
# 한 줄로 쓰고 싶을 때 (로그·표)
# ──────────────────────────────────────────────────────────

def short(e: BaseException) -> str:
    what, _how = explain(e)
    return what


def detail(e: BaseException) -> str:
    """'자세한 내용'에 넣을 원문. 필요할 때만 보여 준다."""
    return _text(e)
