"""
engine.py — PPT 자리 찾기/채우기 · PDF 변환 · Gmail · 발송 파이프라인

PPT에 어떤 단어가 들어 있든 대응한다. 코드에 자리 이름을 박아 두지 않는다.
"""
from __future__ import annotations

import base64
import re
import tempfile
import threading
import time
import zipfile
from dataclasses import dataclass, field
from email import encoders, message_from_bytes
from email.mime.base import MIMEBase
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from pathlib import Path
from xml.sax.saxutils import escape as _xml_escape, unescape as _xml_unescape

import core
import explain
from core import (STATUS_COL, first_email, pick_josa, replace_placeholders,
                  resolve, safe_filename)

# ──────────────────────────────────────────────────────────
# PPT 자리 찾기
# ──────────────────────────────────────────────────────────

# <a:t>글자</a:t> 하나가 '런'. 한 문단이 여러 런으로 쪼개져 있을 수 있어
# 문단 단위로 이어붙여 찾는다. (예전 코드가 정규식으로 때우던 부분)
_T_RE = re.compile(r"(<a:t>)(.*?)(</a:t>)", re.DOTALL)
_P_RE = re.compile(r"<a:p\b.*?</a:p>", re.DOTALL)

# 괄호로 묶인 자리. 두 글자 이상이어야 (주)·(토) 같은 것에 걸리지 않는다
_SLOT_RE = re.compile(r"\(([^()\n\r]{2,20})\)")

# 자리 뒤에 붙는 조사 (긴 것 먼저)
_JOSA_ALT = "으로|을|를|이|가|은|는|과|와|아|야|로"
_BOUNDARY = r'(?=[\s.,!?)\]"\'”’]|$)'

_SLOT_STOP = {"주", "토", "일", "월", "화", "수", "목", "금"}


@dataclass
class Slot:
    """PPT에서 찾은 바꿀 자리 하나."""
    name: str                                  # "(업체명)"
    slides: list[int] = field(default_factory=list)
    count: int = 0
    context: str = ""                          # 그 자리가 들어간 실제 문장

    @property
    def label(self) -> str:
        return self.name

    @property
    def inner(self) -> str:
        return self.name[1:-1]

    @property
    def where(self) -> str:
        s = ", ".join(f"{n}장" for n in self.slides[:4])
        if len(self.slides) > 4:
            s += " 외"
        return f"{self.count}곳 · {s}" if s else f"{self.count}곳"


def _slide_no(name: str) -> int:
    m = re.search(r"slide(\d+)\.xml", name)
    return int(m.group(1)) if m else 0


def _paragraph_text(para_xml: str) -> str:
    return _xml_unescape("".join(m.group(2) for m in _T_RE.finditer(para_xml)))


def scan_slots(pptx: Path) -> list[Slot]:
    """PPT를 훑어 괄호로 묶인 자리를 모두 찾는다.

    자리 이름이 (업체명)이든 (기업)이든 (회사명)이든 상관없다.
    """
    pptx = Path(pptx)
    found: dict[str, Slot] = {}
    if not pptx.exists():
        return []

    with zipfile.ZipFile(pptx, "r") as zf:
        names = sorted((n for n in zf.namelist()
                        if re.fullmatch(r"ppt/slides/slide\d+\.xml", n)),
                       key=_slide_no)
        for n in names:
            no = _slide_no(n)
            try:
                xml = zf.read(n).decode("utf-8")
            except Exception:
                continue
            for pm in _P_RE.finditer(xml):
                text = _paragraph_text(pm.group(0))
                if "(" not in text:
                    continue
                for sm in _SLOT_RE.finditer(text):
                    inner = sm.group(1).strip()
                    if not inner or inner in _SLOT_STOP:
                        continue
                    if not re.search(r"[가-힣A-Za-z]", inner):
                        continue
                    if re.fullmatch(r"[\d\s.\-/~]+", inner):
                        continue
                    key = f"({inner})"
                    slot = found.get(key)
                    if slot is None:
                        slot = Slot(name=key, context=_trim(text, sm.start(), sm.end()))
                        found[key] = slot
                    slot.count += 1
                    if no not in slot.slides:
                        slot.slides.append(no)

    return sorted(found.values(), key=lambda s: (-s.count, s.slides[0] if s.slides else 0))


def _trim(text: str, a: int, b: int, span: int = 42) -> str:
    """자리 주변 문장만 잘라 보여준다."""
    text = re.sub(r"\s+", " ", text).strip()
    # 공백 정리로 위치가 밀렸을 수 있으니 다시 찾는다
    frag = re.sub(r"\s+", " ", text[a:b]) if b <= len(text) else ""
    i = text.find(frag) if frag else -1
    if i < 0:
        return text[:110] + ("…" if len(text) > 110 else "")
    s = max(0, i - span)
    e = min(len(text), i + len(frag) + span)
    out = text[s:e]
    return ("… " if s > 0 else "") + out + (" …" if e < len(text) else "")


# ──────────────────────────────────────────────────────────
# PPT 채우기
# ──────────────────────────────────────────────────────────

def _slot_pattern(slot: str) -> re.Pattern:
    """(업체명) 또는 (업체명)+조사 를 한 번에 잡는다."""
    return re.compile(
        re.escape(slot) + rf"(?:({_JOSA_ALT}){_BOUNDARY})?"
    )


def fill_text(text: str, slot_map: dict, row: dict, use_josa: bool = True) -> str:
    """문자열 하나에 자리 채우기를 적용한다."""
    for slot, col in slot_map.items():
        if not col:
            continue
        val = str(row.get(col, "") or "").strip()
        rx = _slot_pattern(slot)

        def _sub(m, _v=val):
            josa = m.group(1)
            if josa and use_josa:
                return _v + pick_josa(_v, josa)
            return _v + (josa or "")

        text = rx.sub(_sub, text)
    return text


def _fill_paragraph(para_xml: str, slot_map: dict, row: dict, use_josa: bool) -> str:
    """한 문단의 런들을 이어붙여 바꾸고, 서식은 최대한 지킨 채 되돌린다."""
    runs = [(m.start(2), m.end(2), _xml_unescape(m.group(2)))
            for m in _T_RE.finditer(para_xml)]
    if not runs:
        return para_xml

    texts = [t for _, _, t in runs]
    joined = "".join(texts)
    if "(" not in joined:
        return para_xml

    # 런 경계표 — 이어붙인 위치를 (런번호, 런 안 위치)로 되돌리는 데 쓴다
    bounds = []
    acc = 0
    for t in texts:
        bounds.append((acc, acc + len(t)))
        acc += len(t)

    def locate(pos: int) -> tuple[int, int]:
        for i, (s, e) in enumerate(bounds):
            if s <= pos < e:
                return i, pos - s
            if pos == e and i == len(bounds) - 1:
                return i, pos - s
        return len(bounds) - 1, len(texts[-1])

    edits = []
    for slot, col in slot_map.items():
        if not col:
            continue
        val = str(row.get(col, "") or "").strip()
        for m in _slot_pattern(slot).finditer(joined):
            josa = m.group(1)
            repl = val + (pick_josa(val, josa) if (josa and use_josa) else (josa or ""))
            edits.append((m.start(), m.end(), repl))

    if not edits:
        return para_xml

    # 겹치는 자리는 앞선 것만 남기고, 뒤에서부터 고쳐야 위치가 안 밀린다
    edits.sort(key=lambda e: e[0])
    merged = []
    last_end = -1
    for a, b, r in edits:
        if a >= last_end:
            merged.append((a, b, r))
            last_end = b
    for a, b, repl in reversed(merged):
        ia, oa = locate(a)
        ib, ob = locate(b)
        if ia == ib:
            texts[ia] = texts[ia][:oa] + repl + texts[ia][ob:]
        else:
            texts[ia] = texts[ia][:oa] + repl
            for k in range(ia + 1, ib):
                texts[k] = ""
            texts[ib] = texts[ib][ob:]

    # XML로 되돌린다 (뒤에서부터 갈아끼워야 위치가 유지된다)
    out = para_xml
    for (s, e, _), new in reversed(list(zip(runs, texts))):
        out = out[:s] + _xml_escape(new) + out[e:]
    return out


def fill_pptx(src: Path, dst: Path, slot_map: dict, row: dict, use_josa: bool = True) -> None:
    """PPT를 복사하면서 자리를 채운다."""
    src, dst = Path(src), Path(dst)
    with tempfile.TemporaryDirectory() as tmp:
        ex = Path(tmp) / "ex"
        with zipfile.ZipFile(src, "r") as zf:
            zf.extractall(ex)

        for sp in sorted((ex / "ppt" / "slides").glob("slide*.xml")):
            xml = sp.read_text(encoding="utf-8")
            new = _P_RE.sub(
                lambda m: _fill_paragraph(m.group(0), slot_map, row, use_josa), xml)
            if new != xml:
                sp.write_text(new, encoding="utf-8")

        dst.parent.mkdir(parents=True, exist_ok=True)
        with zipfile.ZipFile(dst, "w", zipfile.ZIP_DEFLATED) as zf:
            for item in sorted(ex.rglob("*")):
                if item.is_file():
                    zf.write(item, item.relative_to(ex))


# ──────────────────────────────────────────────────────────
# PDF 변환 — PowerPoint를 한 번만 켠다
# ──────────────────────────────────────────────────────────

def powerpoint_available() -> bool:
    try:
        import win32com.client  # noqa: F401
        return True
    except ImportError:
        return False


class PowerPointSession:
    """PowerPoint를 한 번 켜서 계속 쓴다. 기업마다 켰다 끄면 훨씬 느리다."""

    def __init__(self):
        self._app = None
        self._co = False

    def __enter__(self):
        try:
            import pythoncom
            pythoncom.CoInitialize()
            self._co = True
        except Exception:
            pass
        return self

    def _ensure(self):
        if self._app is None:
            import win32com.client
            self._app = win32com.client.Dispatch("PowerPoint.Application")
        return self._app

    def convert(self, pptx: Path, pdf: Path, retries: int = 3) -> bool:
        pptx, pdf = Path(pptx), Path(pdf)
        for attempt in range(1, retries + 1):
            try:
                app = self._ensure()
                prs = app.Presentations.Open(str(pptx.resolve()), WithWindow=False)
                try:
                    prs.SaveAs(str(pdf.resolve()), 32)  # 32 = PDF
                finally:
                    prs.Close()
                if pdf.exists() and pdf.stat().st_size >= 1024:
                    return True
            except Exception:
                # 파워포인트가 죽었을 수 있으니 다음 시도에 다시 켠다
                self._quit()
            if attempt < retries:
                time.sleep(2)
        return False

    def _quit(self):
        try:
            if self._app is not None:
                self._app.Quit()
        except Exception:
            pass
        self._app = None

    def __exit__(self, *exc):
        self._quit()
        if self._co:
            try:
                import pythoncom
                pythoncom.CoUninitialize()
            except Exception:
                pass
        return False


# ──────────────────────────────────────────────────────────
# Gmail
# ──────────────────────────────────────────────────────────

_GMAIL_SCOPES = [
    "https://www.googleapis.com/auth/gmail.send",
    "https://www.googleapis.com/auth/gmail.readonly",
]

_svc_lock = threading.Lock()
_svc_cache = None


def gmail_connected() -> bool:
    return core.TOKEN_PATH.exists()


def gmail_service(force: bool = False):
    """서비스 객체를 한 번만 만들어 돌려쓴다."""
    global _svc_cache
    with _svc_lock:
        if _svc_cache is not None and not force:
            return _svc_cache

        from googleapiclient.discovery import build
        from google.oauth2.credentials import Credentials
        from google_auth_oauthlib.flow import InstalledAppFlow
        from google.auth.transport.requests import Request

        creds = None
        if core.TOKEN_PATH.exists():
            try:
                creds = Credentials.from_authorized_user_file(
                    str(core.TOKEN_PATH), _GMAIL_SCOPES)
            except Exception:
                creds = None
        if not creds or not creds.valid:
            if creds and creds.expired and creds.refresh_token:
                creds.refresh(Request())
            else:
                if not core.CREDENTIALS_PATH.exists():
                    raise FileNotFoundError(
                        "credentials.json 이 없습니다. 프로그램 폴더에 넣어 주세요.")
                flow = InstalledAppFlow.from_client_secrets_file(
                    str(core.CREDENTIALS_PATH), _GMAIL_SCOPES)
                creds = flow.run_local_server(port=0)
            core.SETTINGS_DIR.mkdir(parents=True, exist_ok=True)
            core.TOKEN_PATH.write_text(creds.to_json(), encoding="utf-8")

        _svc_cache = build("gmail", "v1", credentials=creds)
        return _svc_cache


def gmail_address() -> str:
    try:
        svc = gmail_service()
        return svc.users().getProfile(userId="me").execute().get("emailAddress", "")
    except Exception:
        return ""


def gmail_disconnect() -> None:
    global _svc_cache
    with _svc_lock:
        _svc_cache = None
    try:
        core.TOKEN_PATH.unlink()
    except Exception:
        pass


def send_mail(svc, to: str, subject: str, body: str, attach: Path | None) -> None:
    to = first_email(to)
    if not to or "@" not in to:
        raise ValueError(f"보낼 수 없는 주소입니다: {to!r}")

    msg = MIMEMultipart()
    msg["To"] = to
    msg["Subject"] = subject
    msg.attach(MIMEText(body, "plain", "utf-8"))

    if attach:
        attach = Path(attach)
        with open(attach, "rb") as f:
            part = MIMEBase("application", "octet-stream")
            part.set_payload(f.read())
        encoders.encode_base64(part)
        part.add_header("Content-Disposition", "attachment",
                        filename=("utf-8", "", attach.name))
        msg.attach(part)

    raw = base64.urlsafe_b64encode(msg.as_bytes()).decode("utf-8")
    svc.users().messages().send(userId="me", body={"raw": raw}).execute()


# ──────────────────────────────────────────────────────────
# 발송 파이프라인
# ──────────────────────────────────────────────────────────

STEP_PPT, STEP_PDF, STEP_SHEET, STEP_MAIL = 1, 2, 3, 4


def build_one(ws: str, cfg: dict, row: dict, ppt: PowerPointSession | None,
              make_pdf: bool = True) -> tuple[Path, Path | None]:
    """한 기업 몫의 제안서를 만든다. (pptx, pdf) 를 돌려준다."""
    template = resolve(ws, cfg["template_pptx"])
    out = core.output_dir(ws)
    out.mkdir(parents=True, exist_ok=True)

    base = safe_filename(replace_placeholders(
        cfg.get("file_name_pattern") or "제안서", row))
    pptx_out = out / f"{base}.pptx"
    fill_pptx(template, pptx_out, cfg.get("slot_map") or {}, row,
              bool(cfg.get("slot_josa", True)))

    if not make_pdf or ppt is None:
        return pptx_out, None

    pdf_out = out / f"{base}.pdf"
    ok = ppt.convert(pptx_out, pdf_out)
    return pptx_out, (pdf_out if ok else None)


def run_send(ws: str, cfg: dict, targets: list[dict], sheet,
             log_fn, progress_fn, done_fn, fail_fn=None,
             stop_event: threading.Event | None = None,
             sent_log: "core.SendLog | None" = None) -> None:
    """targets 는 core.classify() 결과에서 보낼 것만 걸러낸 목록.

    fail_fn(기업이름, 단계, 예외) 는 실패한 곳을 화면이 모아 두는 데 쓴다.
    나중에 결과 화면에서 '왜 안 됐고 어떻게 하면 되는지' 를 보여 준다.
    """
    if fail_fn is None:
        def fail_fn(*_a, **_k):
            pass
    sent = failed = 0
    try:
        total = len(targets)
        if total == 0:
            log_fn("보낼 곳이 없습니다.")
            return

        interval = float(cfg.get("send_interval") or 0)
        slot_map = cfg.get("slot_map") or {}
        used_names: set[str] = set()      # 파일이 서로 덮어써지지 않게
        sheet.ensure_column(STATUS_COL)
        if sent_log is None:
            sent_log = core.SendLog(ws)

        log_fn(f"{total}곳에 보내기 시작합니다.")
        svc = gmail_service()

        with PowerPointSession() as ppt:
            for i, t in enumerate(targets):
                if stop_event and stop_event.is_set():
                    log_fn("중지했습니다.")
                    break

                row  = t["row"]
                name = t["name"] or f"{i + 1}번째"
                idx  = t["index"]

                # 1. 제안서 만들기
                progress_fn(i, total, name, STEP_PPT, "run")
                try:
                    base = core.unique_name(
                        safe_filename(replace_placeholders(
                            cfg.get("file_name_pattern") or "제안서", row)),
                        used_names)
                    pptx_out = core.output_dir(ws) / f"{base}.pptx"
                    core.output_dir(ws).mkdir(parents=True, exist_ok=True)
                    fill_pptx(resolve(ws, cfg["template_pptx"]), pptx_out,
                              slot_map, row, bool(cfg.get("slot_josa", True)))
                    progress_fn(i, total, name, STEP_PPT, "ok")
                except Exception as e:
                    progress_fn(i, total, name, STEP_PPT, "bad")
                    log_fn(f"[{name}] 제안서를 만들지 못했어요 — {explain.short(e)}")
                    fail_fn(name, STEP_PPT, e)
                    failed += 1
                    continue

                # 2. PDF로 바꾸기
                progress_fn(i, total, name, STEP_PDF, "run")
                pdf_out = core.output_dir(ws) / f"{base}.pdf"
                if ppt.convert(pptx_out, pdf_out):
                    progress_fn(i, total, name, STEP_PDF, "ok")
                    # 중간 산물은 지운다. 템플릿이 크면 기업마다 수십 MB씩 쌓인다.
                    # (변환에 실패한 것만 남겨 두고 직접 열어볼 수 있게 한다)
                    try:
                        pptx_out.unlink()
                    except Exception:
                        pass
                else:
                    progress_fn(i, total, name, STEP_PDF, "bad")
                    log_fn(f"[{name}] PDF로 바꾸지 못했어요.")
                    fail_fn(name, STEP_PDF, None)
                    failed += 1
                    continue

                # 3. 명단에 기록
                progress_fn(i, total, name, STEP_SHEET, "run")
                try:
                    sheet.set(idx, "제안서파일", pdf_out.name)
                    progress_fn(i, total, name, STEP_SHEET, "ok")
                except Exception as e:
                    progress_fn(i, total, name, STEP_SHEET, "bad")
                    log_fn(f"[{name}] 명단에 기록하지 못했어요 — {explain.short(e)}")

                # 4. 메일 보내기
                progress_fn(i, total, name, STEP_MAIL, "run")
                try:
                    subject = replace_placeholders(cfg.get("email_subject", ""), row)
                    body    = replace_placeholders(cfg.get("email_body", ""), row)
                    send_mail(svc, t["email"], subject, body, pdf_out)
                    progress_fn(i, total, name, STEP_MAIL, "ok")
                    sheet.set(idx, STATUS_COL, "발송완료")
                    sent_log.mark(t["email"], name, "발송완료")
                    sent_log.save()
                    sent += 1
                    _save_quietly(sheet, log_fn)
                except Exception as e:
                    progress_fn(i, total, name, STEP_MAIL, "bad")
                    log_fn(f"[{name}] 메일을 보내지 못했어요 — {explain.short(e)}")
                    fail_fn(name, STEP_MAIL, e)
                    sheet.set(idx, STATUS_COL, "발송실패")
                    sent_log.mark(t["email"], name, "발송실패")
                    sent_log.save()
                    failed += 1
                    _save_quietly(sheet, log_fn)

                if interval > 0 and i < total - 1:
                    if stop_event and stop_event.wait(interval):
                        log_fn("중지했습니다.")
                        break

        _save_quietly(sheet, log_fn, loud=True)

    except Exception as e:
        log_fn(f"[문제가 생겼어요] {explain.short(e)}")
        fail_fn("", 0, e)
    finally:
        done_fn(sent, failed)


def _save_quietly(sheet, log_fn, loud: bool = False) -> None:
    try:
        sheet.save()
        if loud:
            log_fn("명단에 결과를 기록했습니다.")
    except PermissionError:
        log_fn("명단 파일이 다른 프로그램에서 열려 있어 저장하지 못했습니다.")
    except Exception as e:
        log_fn("명단을 저장하지 못했어요 — " + explain.short(e))


# ──────────────────────────────────────────────────────────
# 반송 확인
# ──────────────────────────────────────────────────────────

_ADDR_RE = re.compile(r"[\w.+\-]+@[\w.\-]+\.\w+")
_STATUS_RE = re.compile(r"^\s*Status:\s*([245]\.\d+\.\d+)", re.MULTILINE | re.IGNORECASE)

_REASON = {
    "5.1.1": "없는 계정이에요",
    "5.1.2": "없는 주소예요",
    "5.1.3": "주소 형식이 잘못됐어요",
    "5.2.1": "받지 않는 계정이에요",
    "5.2.2": "받은편지함이 가득 찼어요",
    "5.4.1": "서버에 닿지 못했어요",
    "5.7.1": "수신이 거부됐어요",
    "5.7.26": "보낸 사람 인증에 실패했어요",
}

_IGNORE_SENDERS = {"mailer-daemon@googlemail.com", "mailer-daemon@google.com",
                   "postmaster@gmail.com", "mailer-daemon@gmail.com"}


def _bounced_from(raw: bytes, me: str) -> dict[str, str]:
    """DSN 한 통에서 (주소 → 이유) 를 뽑는다."""
    ignore = set(_IGNORE_SENDERS) | ({me.lower()} if me else set())
    out: dict[str, str] = {}

    try:
        msg = message_from_bytes(raw)
    except Exception:
        text = raw.decode("utf-8", errors="replace")
        for a in _ADDR_RE.findall(text):
            if a.lower() not in ignore:
                out[a.lower()] = "되돌아왔어요"
        return out

    for a in _ADDR_RE.findall(msg.get("X-Failed-Recipients", "") or ""):
        if a.lower() not in ignore:
            out.setdefault(a.lower(), "되돌아왔어요")

    for part in msg.walk():
        ctype = part.get_content_type()

        if ctype == "message/delivery-status":
            payload = part.get_payload()
            chunks = []
            if isinstance(payload, list):
                for p in payload:
                    try:
                        chunks.append(str(p))
                    except Exception:
                        pass
            else:
                chunks.append(str(payload))
            blob = "\n".join(chunks)

            code = ""
            m = _STATUS_RE.search(blob)
            if m:
                code = m.group(1)
            reason = _REASON.get(code, f"되돌아왔어요 ({code})" if code else "되돌아왔어요")

            for m2 in re.finditer(
                    r"(?:Final|Original)-Recipient\s*:.*?rfc822\s*;\s*"
                    r"([\w.+\-]+@[\w.\-]+\.\w+)", blob, re.IGNORECASE):
                a = m2.group(1).lower()
                if a not in ignore:
                    out[a] = reason

        elif ctype.startswith("text/"):
            data = part.get_payload(decode=True)
            if not data:
                continue
            text = data.decode(part.get_content_charset() or "utf-8", errors="replace")
            code = ""
            m = re.search(r"\b([45]\.\d+\.\d+)\b", text)
            if m:
                code = m.group(1)
            reason = _REASON.get(code, "되돌아왔어요")
            for a in _ADDR_RE.findall(text):
                if a.lower() not in ignore:
                    out.setdefault(a.lower(), reason)

    return out


def check_bounces(cfg: dict, sheet, log_fn, max_messages: int = 200) -> list[dict]:
    """반송 메일을 찾아 명단에 표시한다. 표에 보여줄 목록을 돌려준다."""
    log_fn("반송 메일을 확인하고 있어요…")
    svc = gmail_service()
    me = cfg.get("gmail_user", "") or gmail_address()

    query = ("from:mailer-daemon@googlemail.com OR from:mailer-daemon@google.com "
             "OR from:postmaster@gmail.com")
    res = svc.users().messages().list(
        userId="me", q=query, maxResults=max_messages).execute()
    messages = res.get("messages", [])
    log_fn(f"반송으로 보이는 메일 {len(messages)}통을 찾았어요.")

    bounced: dict[str, str] = {}
    for m in messages:
        try:
            data = svc.users().messages().get(
                userId="me", id=m["id"], format="raw").execute()
            raw = base64.urlsafe_b64decode(data["raw"].encode("utf-8") + b"==")
            for addr, reason in _bounced_from(raw, me).items():
                bounced.setdefault(addr, reason)
        except Exception:
            continue

    if not bounced:
        log_fn("되돌아온 주소를 찾지 못했어요.")
        return []

    col_em = cfg.get("col_email", "")
    name_col = core._display_name_column(sheet, cfg)
    sheet.ensure_column(STATUS_COL)

    hits = []
    for i, row in enumerate(sheet.rows):
        addr = first_email(row.get(col_em, "")).lower()
        if addr and addr in bounced:
            sheet.set(i, STATUS_COL, "반송됨")
            hits.append({
                "name":   (row.get(name_col, "") or "").strip(),
                "email":  addr,
                "reason": bounced[addr],
            })

    try:
        sheet.save()
        log_fn(f"{len(hits)}곳을 명단에 '반송됨'으로 표시했어요.")
    except Exception as e:
        log_fn("명단을 저장하지 못했어요 — " + explain.short(e))

    return hits
