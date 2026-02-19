import os
import re
import sys
import time
import json
import csv
import random
import threading
import tkinter as tk
from tkinter import messagebox, filedialog
from tkinter import ttk

import PyPDF2
from docx import Document

from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import TimeoutException, WebDriverException
from webdriver_manager.chrome import ChromeDriverManager


# =========================
#   CONFIG / REGEX
# =========================
YT_RE = re.compile(r"\b\d{7}-\d\b|\b\d{8}\b")
EMAIL_RE = re.compile(r"[A-Za-z0-9_.+-]+@[A-Za-z0-9-]+\.[A-Za-z0-9-.]+")
EMAIL_A_RE = re.compile(r"[A-Za-z0-9_.+-]+\s*\(a\)\s*[A-Za-z0-9-]+\.[A-Za-z0-9-.]+", re.I)

YTJ_HOME = "https://tietopalvelu.ytj.fi/"
YTJ_COMPANY_URL = "https://tietopalvelu.ytj.fi/yritys/{}"

BAD_LINES = {
    "yritys", "sijainti", "summa", "häiriöpäivä", "tyyppi", "lähde",
    "viimeisimmät protestit", "protestilista",
    "y-tunnus", "julkaisupäivä", "alue", "velkoja",
}
BAD_CONTAINS = [
    "velkomustuomiot", "ulosotto", "konkurssi", "dun & bradstreet", "bisnode",
    "protestit", "protestia",
]

GENERIC_EMAIL_PREFIXES = (
    "info@", "office@", "support@", "asiakaspalvelu@", "sales@", "myynti@",
    "noreply@", "no-reply@", "donotreply@", "admin@", "contact@", "hello@"
)


# =========================
#   PATHS / LOG
# =========================
def get_exe_dir():
    if getattr(sys, "frozen", False):
        return os.path.dirname(sys.executable)
    return os.path.dirname(os.path.abspath(__file__))


def get_output_dir():
    base = get_exe_dir()
    try:
        p = os.path.join(base, "_write_test.tmp")
        with open(p, "w", encoding="utf-8") as f:
            f.write("ok")
        os.remove(p)
    except Exception:
        home = os.path.expanduser("~")
        docs = os.path.join(home, "Documents")
        base = os.path.join(docs, "ProtestiBotti")

    date_folder = time.strftime("%Y-%m-%d")
    out = os.path.join(base, date_folder)
    os.makedirs(out, exist_ok=True)
    return out


OUT_DIR = get_output_dir()
LOG_PATH = os.path.join(OUT_DIR, "log.txt")

CACHE_DIR = os.path.join(os.path.expanduser("~"), "Documents", "ProtestiBotti")
os.makedirs(CACHE_DIR, exist_ok=True)
CACHE_PATH = os.path.join(CACHE_DIR, "ytj_cache.json")


def log_to_file(msg: str):
    ts = time.strftime("%H:%M:%S")
    line = f"[{ts}] {msg}"
    try:
        with open(LOG_PATH, "a", encoding="utf-8") as f:
            f.write(line + "\n")
    except Exception:
        pass
    return line


def reset_log():
    try:
        with open(LOG_PATH, "w", encoding="utf-8") as f:
            f.write("=== BOTTI KÄYNNISTETTY ===\n")
    except Exception:
        pass
    log_to_file(f"Output: {OUT_DIR}")
    log_to_file(f"Logi: {LOG_PATH}")
    log_to_file(f"Cache: {CACHE_PATH}")
    log_to_file("PDF: sisäinen minipdf-writer (ei reportlabia)")


# =========================
#   CACHE
# =========================
def _load_cache():
    try:
        if os.path.exists(CACHE_PATH):
            with open(CACHE_PATH, "r", encoding="utf-8") as f:
                data = json.load(f)
                if isinstance(data, dict):
                    data.setdefault("yt", {})
                    data.setdefault("name", {})
                    return data
    except Exception:
        pass
    return {"yt": {}, "name": {}}


def _save_cache(cache):
    try:
        tmp = CACHE_PATH + ".tmp"
        with open(tmp, "w", encoding="utf-8") as f:
            json.dump(cache, f, ensure_ascii=False, indent=2)
        os.replace(tmp, CACHE_PATH)
    except Exception:
        pass


def cache_stats(cache):
    return len(cache.get("yt", {})), len(cache.get("name", {}))


def cache_clear():
    try:
        if os.path.exists(CACHE_PATH):
            os.remove(CACHE_PATH)
    except Exception:
        pass


def cache_get_by_yt(cache, yt: str):
    return cache.get("yt", {}).get((yt or "").strip())


def cache_put_yt(cache, yt: str, email: str, name: str = ""):
    yt = (yt or "").strip()
    if not yt:
        return
    cache.setdefault("yt", {})
    cache["yt"][yt] = {"email": (email or "").strip(), "name": (name or "").strip(), "ts": int(time.time())}


def normalize_company_name(name: str) -> str:
    s = (name or "").strip().casefold()
    s = re.sub(r"\s+", " ", s)
    s = re.sub(r"[.,;:()\[\]{}\"'´`]", "", s)
    s = s.replace("&", " and ")
    s = re.sub(r"\s+", " ", s).strip()
    s = re.sub(r"\b(osakeyhtiö|osakeyhtio)\b", "oy", s)
    return s


def cache_get_by_name(cache, name: str):
    return cache.get("name", {}).get(normalize_company_name(name))


def cache_put_name(cache, name: str, email: str, yt: str = ""):
    key = normalize_company_name(name)
    if not key:
        return
    cache.setdefault("name", {})
    cache["name"][key] = {"email": (email or "").strip(), "yt": (yt or "").strip(), "ts": int(time.time()), "raw": (name or "").strip()}


# =========================
#   UTIL
# =========================
def normalize_yt(yt: str):
    yt = (yt or "").strip().replace(" ", "")
    if re.fullmatch(r"\d{7}-\d", yt):
        return yt
    if re.fullmatch(r"\d{8}", yt):
        return yt[:7] + "-" + yt[7]
    return None


def pick_email_from_text(text: str) -> str:
    if not text:
        return ""
    m = EMAIL_RE.search(text)
    if m:
        return m.group(0).strip().replace(" ", "")
    m2 = EMAIL_A_RE.search(text)
    if m2:
        return m2.group(0).replace(" ", "").replace("(a)", "@")
    return ""


def classify_email(email: str) -> str:
    e = (email or "").strip().lower()
    if not e:
        return ""
    if e.startswith(GENERIC_EMAIL_PREFIXES):
        return "GENEERINEN"
    return "OK"


def save_word_plain_lines(lines, filename):
    path = os.path.join(OUT_DIR, filename)
    doc = Document()
    for line in lines:
        if line and str(line).strip():
            doc.add_paragraph(str(line).strip())
    doc.save(path)
    return path


def save_txt_lines(lines, filename):
    path = os.path.join(OUT_DIR, filename)
    with open(path, "w", encoding="utf-8") as f:
        for line in lines:
            if line and str(line).strip():
                f.write(str(line).strip() + "\n")
    return path


def save_csv_rows(rows, headers, filename):
    path = os.path.join(OUT_DIR, filename)
    with open(path, "w", encoding="utf-8", newline="") as f:
        w = csv.writer(f, delimiter=";")
        w.writerow(headers)
        for r in rows:
            w.writerow(r)
    return path


# =========================
#   PDF WITHOUT REPORTLAB
# =========================
def _pdf_escape_text(s: str) -> str:
    s = s.replace("\\", "\\\\").replace("(", "\\(").replace(")", "\\)")
    s = s.replace("\t", "    ")
    return s


def save_pdf_lines(lines, filename, title=None):
    path = os.path.join(OUT_DIR, filename)
    PAGE_W, PAGE_H = 595.28, 841.89
    LEFT = 50
    TOP = PAGE_H - 60
    BOTTOM = 60
    FONT_SIZE = 11
    LEADING = 14

    out_lines = []
    if title:
        out_lines.append(title)
        out_lines.append("")
    for ln in lines:
        if ln is None:
            continue
        t = str(ln).strip()
        if t:
            out_lines.append(t)

    max_lines_per_page = int((TOP - BOTTOM) // LEADING)
    pages, cur = [], []
    for ln in out_lines:
        cur.append(ln)
        if len(cur) >= max_lines_per_page:
            pages.append(cur)
            cur = []
    if cur:
        pages.append(cur)
    if not pages:
        pages = [[""]]

    objects = []

    def add_obj(data_bytes: bytes) -> int:
        objects.append(data_bytes)
        return len(objects)

    font_obj = add_obj(b"<< /Type /Font /Subtype /Type1 /BaseFont /Helvetica /Encoding /WinAnsiEncoding >>")
    page_objs = []
    pages_obj_id = add_obj(b"<< /Type /Pages /Kids [] /Count 0 >>")

    for page_lines in pages:
        y = TOP
        chunks = ["BT\n", f"/F1 {FONT_SIZE} Tf\n", f"{LEFT} {y:.2f} Td\n"]
        first = True
        for ln in page_lines:
            txt = _pdf_escape_text(ln)
            if not first:
                chunks.append(f"0 -{LEADING} Td\n")
            first = False
            chunks.append(f"({txt}) Tj\n")
        chunks.append("ET\n")
        content_str = "".join(chunks)
        content_bytes = content_str.encode("latin-1", errors="replace")
        stream = b"<< /Length %d >>\nstream\n%s\nendstream" % (len(content_bytes), content_bytes)
        content_obj = add_obj(stream)

        page_dict = (
            b"<< /Type /Page /Parent %d 0 R "
            b"/MediaBox [0 0 %.2f %.2f] "
            b"/Resources << /Font << /F1 %d 0 R >> >> "
            b"/Contents %d 0 R >>"
            % (pages_obj_id, PAGE_W, PAGE_H, font_obj, content_obj)
        )
        page_obj_id = add_obj(page_dict)
        page_objs.append(page_obj_id)

    kids = " ".join([f"{pid} 0 R" for pid in page_objs]).encode("ascii")
    objects[pages_obj_id - 1] = b"<< /Type /Pages /Kids [%s] /Count %d >>" % (kids, len(page_objs))
    catalog_obj_id = add_obj(b"<< /Type /Catalog /Pages %d 0 R >>" % pages_obj_id)

    header = b"%PDF-1.4\n%\xe2\xe3\xcf\xd3\n"
    offsets = [0]
    body = b""
    for i, obj in enumerate(objects, start=1):
        offsets.append(len(header) + len(body))
        body += f"{i} 0 obj\n".encode("ascii") + obj + b"\nendobj\n"

    xref_start = len(header) + len(body)
    xref = [b"xref\n", f"0 {len(objects)+1}\n".encode("ascii")]
    xref.append(b"0000000000 65535 f \n")
    for off in offsets[1:]:
        xref.append(f"{off:010d} 00000 n \n".encode("ascii"))
    xref_bytes = b"".join(xref)

    trailer = (
        b"trailer\n"
        b"<< /Size %d /Root %d 0 R >>\n"
        b"startxref\n%d\n%%%%EOF\n"
        % (len(objects) + 1, catalog_obj_id, xref_start)
    )

    with open(path, "wb") as f:
        f.write(header + body + xref_bytes + trailer)
    return path


# =========================
#   PARSERS
# =========================
def parse_yts_only(raw: str):
    raw = raw or ""
    found, seen = [], set()

    for m in YT_RE.findall(raw):
        n = normalize_yt(m)
        if n and n not in seen:
            seen.add(n)
            found.append(n)

    tokens = re.split(r"[,\s;]+", raw.strip())
    for t in tokens:
        n = normalize_yt(t)
        if n and n not in seen:
            seen.add(n)
            found.append(n)

    return sorted(found)


def _is_money_line(s: str) -> bool:
    return bool(re.fullmatch(r"\d{1,3}(\s?\d{3})*\s?€", s.strip()))


def _is_date_line(s: str) -> bool:
    return bool(re.fullmatch(r"\d{1,2}\.\d{1,2}\.\d{4}", s.strip()))


def _looks_like_location(s: str) -> bool:
    s = s.strip()
    if not s or any(ch.isdigit() for ch in s):
        return False
    parts = s.split()
    if len(parts) != 1:
        return False
    if not parts[0][:1].isupper():
        return False
    low = s.lower()
    if any(x in low for x in [" oy", " ab", " ky", " ay", " ry", " tmi", " oyj", " lkv", " osakeyhtiö"]):
        return False
    return True


def _looks_like_company_name(s: str) -> bool:
    s = s.strip()
    if len(s) < 3:
        return False
    if not re.search(r"[A-Za-zÅÄÖåäö]", s):
        return False

    low = s.lower()
    if low in BAD_LINES:
        return False
    if any(b in low for b in BAD_CONTAINS):
        return False
    if _is_money_line(s) or _is_date_line(s):
        return False
    if "y-tunnus" in low or YT_RE.search(s):
        return False

    suffix_ok = any(x in low for x in [" oy", " ab", " ky", " ay", " ry", " tmi", " oyj", " lkv"])
    many_words = len(s.split()) >= 2

    if len(s.split()) == 1 and _looks_like_location(s):
        return False

    return suffix_ok or many_words


def _normalize_name(s: str) -> str:
    s = re.sub(r"\s+", " ", s.strip())
    s = s.strip(" -•\u2022")
    return s


def parse_names_and_yts_from_copied_text(raw: str):
    if not raw:
        return [], []

    yts = parse_yts_only(raw)

    lines = [ln.strip() for ln in raw.splitlines() if ln and ln.strip()]
    lines = [ln for ln in lines if len(ln) >= 2]

    names, seen = [], set()

    money_idxs = [i for i, ln in enumerate(lines) if _is_money_line(ln)]
    for i in money_idxs:
        candidates = []
        for back in (1, 2, 3):
            j = i - back
            if j >= 0:
                candidates.append(lines[j])

        best = ""
        for c in candidates:
            c = _normalize_name(c)
            if _looks_like_company_name(c):
                best = c
                break

        if best:
            key = best.casefold()
            if key not in seen:
                seen.add(key)
                names.append(best)

    if not names:
        for ln in lines:
            ln = _normalize_name(ln)
            if _looks_like_company_name(ln):
                key = ln.casefold()
                if key not in seen:
                    seen.add(key)
                    names.append(ln)

    return names, yts


def extract_names_and_yts_from_pdf(pdf_path: str):
    text_all = ""
    with open(pdf_path, "rb") as f:
        reader = PyPDF2.PdfReader(f)
        for page in reader.pages:
            text_all += (page.extract_text() or "") + "\n"
    return parse_names_and_yts_from_copied_text(text_all)


# =========================
#   YTJ: Selenium + THROTTLE/RETRY
# =========================
def start_new_driver():
    options = webdriver.ChromeOptions()
    options.add_argument("--start-maximized")
    options.add_argument("--disable-gpu")
    options.add_argument("--disable-dev-shm-usage")
    driver_path = ChromeDriverManager().install()
    return webdriver.Chrome(service=Service(driver_path), options=options)


def wait_body(driver, timeout=25):
    WebDriverWait(driver, timeout).until(EC.presence_of_element_located((By.TAG_NAME, "body")))


class Throttle:
    def __init__(self):
        self.base_delay = 0.35
        self.max_delay = 6.0
        self.fail_streak = 0

    def sleep_between(self):
        d = min(self.max_delay, self.base_delay + self.fail_streak * 0.6)
        d += random.uniform(0.0, 0.35)
        time.sleep(d)

    def on_success(self):
        if self.fail_streak > 0:
            self.fail_streak -= 1

    def on_fail(self):
        self.fail_streak = min(10, self.fail_streak + 1)

    def backoff_sleep(self, attempt: int):
        d = min(self.max_delay, (0.9 * (1.7 ** attempt)) + random.uniform(0.0, 0.6))
        time.sleep(d)


def page_looks_blocked_or_captcha(driver) -> bool:
    try:
        t = (driver.find_element(By.TAG_NAME, "body").text or "").lower()
        bad = [
            "unusual traffic", "epätavallista liikennettä", "captcha",
            "varmista että et ole robotti", "verify you are human",
            "liian monta pyyntöä", "too many requests", "429"
        ]
        return any(x in t for x in bad)
    except Exception:
        return False


def safe_get(driver, url: str, throttle: Throttle, stop_evt, status_cb, max_attempts=5):
    for attempt in range(max_attempts):
        if stop_evt.is_set():
            return False
        try:
            throttle.sleep_between()
            driver.get(url)
            wait_body(driver, 25)

            if page_looks_blocked_or_captcha(driver):
                status_cb("YTJ näyttää estolta/CAPTCHA:lta → hidastan…")
                throttle.on_fail()
                throttle.backoff_sleep(attempt)
                continue

            throttle.on_success()
            return True

        except TimeoutException:
            status_cb("YTJ timeout → retry/hidastus…")
            throttle.on_fail()
            throttle.backoff_sleep(attempt)
        except WebDriverException:
            status_cb("YTJ webdriver error → retry/hidastus…")
            throttle.on_fail()
            throttle.backoff_sleep(attempt)

    return False


def extract_email_from_ytj_page(driver):
    try:
        for a in driver.find_elements(By.TAG_NAME, "a"):
            href = (a.get_attribute("href") or "")
            if href.lower().startswith("mailto:"):
                return href.split(":", 1)[1].strip()
    except Exception:
        pass

    try:
        body = driver.find_element(By.TAG_NAME, "body").text or ""
        return pick_email_from_text(body)
    except Exception:
        return ""


# =========================
#   FIX: Find correct YTJ name field by label text
# =========================
def find_input_by_label_text(driver, label_text: str):
    """
    Etsii input/textarea-elementin, joka kuuluu labeliin/otsikkoon jonka teksti sisältää label_text.
    """
    # 1) label[for] -> id
    try:
        labels = driver.find_elements(By.XPATH, f"//label[contains(normalize-space(.), '{label_text}')]")
        for lab in labels:
            fid = (lab.get_attribute("for") or "").strip()
            if fid:
                el = driver.find_element(By.ID, fid)
                if el.is_displayed() and el.is_enabled():
                    return el
    except Exception:
        pass

    # 2) otsikko + input samassa blokkissa
    xpaths = [
        f"//*[self::label or self::div or self::span or self::p or self::h1 or self::h2 or self::h3]"
        f"[contains(normalize-space(.), '{label_text}')]"
        f"/ancestor::*[self::div or self::section or self::fieldset][1]"
        f"//input[not(@type='hidden')][1]",
        f"(//*[contains(normalize-space(.), '{label_text}')])[1]/following::input[not(@type='hidden')][1]",
    ]

    for xp in xpaths:
        try:
            el = driver.find_element(By.XPATH, xp)
            if el.is_displayed() and el.is_enabled():
                return el
        except Exception:
            continue

    return None


# ----- SMART MATCH for name search -----
def _candidate_container_text(a):
    for xp in ["ancestor::li[1]", "ancestor::div[1]", "ancestor::div[2]", "ancestor::tr[1]"]:
        try:
            c = a.find_element(By.XPATH, xp)
            t = (c.text or "").strip()
            if t:
                return t
        except Exception:
            continue
    try:
        return (a.text or "").strip()
    except Exception:
        return ""


def _score_candidate(target_norm: str, cand_name: str, context_text: str):
    cand_norm = normalize_company_name(cand_name)
    ctx = (context_text or "").casefold()
    score = 0

    if cand_norm == target_norm and cand_norm:
        score += 120
    if cand_norm and target_norm and cand_norm in target_norm:
        score += 60
    if cand_norm and target_norm and target_norm in cand_norm:
        score += 50

    tset = set(target_norm.split())
    cset = set(cand_norm.split())
    if tset and cset:
        inter = len(tset & cset)
        score += min(40, inter * 8)

    if "aktiivinen" in ctx:
        score += 8
    if "konkurss" in ctx:
        score -= 10
    if "poistettu" in ctx or "lakan" in ctx:
        score -= 8

    return score


def ytj_find_best_result_link(driver, company_name: str):
    target_norm = normalize_company_name(company_name)

    links = []
    end = time.time() + 10
    while time.time() < end:
        try:
            links = driver.find_elements(By.XPATH, "//a[contains(@href,'/yritys/')]")
            links = [a for a in links if "/yritys/" in (a.get_attribute("href") or "")]
            if links:
                break
        except Exception:
            pass
        time.sleep(0.25)

    if not links:
        return ("", "")

    best = None
    best_score = -10**9

    for a in links[:30]:
        try:
            href = a.get_attribute("href") or ""
            if "/yritys/" not in href:
                continue

            txt = (a.text or "").strip()
            ctx = _candidate_container_text(a)
            cand_name = txt
            if not cand_name:
                lines = [ln.strip() for ln in (ctx or "").splitlines() if ln.strip()]
                cand_name = lines[0] if lines else ""

            score = _score_candidate(target_norm, cand_name, ctx)
            if score > best_score:
                best_score = score
                best = (href, cand_name, score)
        except Exception:
            continue

    if not best:
        return ("", "")

    href, _, _ = best
    yt = ""
    m = re.search(r"/yritys/([^/?#]+)", href)
    if m:
        yt = m.group(1).strip()
    return (yt, href)


def ytj_open_company_by_name_smart(driver, company_name: str, throttle: Throttle, stop_evt, status_cb):
    """
    ✅ KORJATTU: kirjoittaa aina kenttään "Yrityksen tai yhteisön nimi"
    """
    ok = safe_get(driver, YTJ_HOME, throttle, stop_evt, status_cb, max_attempts=5)
    if not ok:
        return (False, "")

    time.sleep(0.6)

    name_input = find_input_by_label_text(driver, "Yrityksen tai yhteisön nimi")
    if not name_input:
        status_cb("YTJ: en löytänyt kenttää 'Yrityksen tai yhteisön nimi'.")
        return (False, "")

    # debug log: mihin kenttään kirjoitetaan
    try:
        iid = name_input.get_attribute("id")
        ph = name_input.get_attribute("placeholder")
        status_cb(f"YTJ: käytän nimikenttää (id={iid}, placeholder={ph})")
    except Exception:
        pass

    try:
        name_input.click()
        name_input.clear()
    except Exception:
        pass

    status_cb(f"YTJ: haku nimellä: {company_name}")
    name_input.send_keys(company_name)
    name_input.send_keys(Keys.ENTER)
    time.sleep(1.2)

    if "/yritys/" in (driver.current_url or ""):
        m = re.search(r"/yritys/([^/?#]+)", driver.current_url or "")
        yt = m.group(1).strip() if m else ""
        return (True, yt)

    yt, url = ytj_find_best_result_link(driver, company_name)
    if not url:
        return (False, "")

    ok2 = safe_get(driver, url, throttle, stop_evt, status_cb, max_attempts=5)
    if not ok2:
        return (False, "")

    return ("/yritys/" in (driver.current_url or ""), yt)


def fetch_emails_from_ytj_by_yts_and_names(driver, yts, names, status_cb, progress_cb, log_cb, stop_evt):
    cache = _load_cache()
    throttle = Throttle()

    emails_out = []
    seen_emails = set()

    total = len(yts) + len(names)
    progress_cb(0, max(1, total))
    done = 0

    # 1) YT ensin
    for yt in yts:
        if stop_evt.is_set():
            break
        done += 1
        progress_cb(done, total)

        yt = normalize_yt(yt) or yt
        status_cb(f"YTJ (Y-tunnus): {yt} ({done}/{total})")

        hit = cache_get_by_yt(cache, yt)
        if hit and hit.get("email"):
            email = hit["email"]
            tag = classify_email(email)
            log_cb(f"[CACHE YT] {yt} → {email} ({tag})")
            if email.lower() not in seen_emails:
                seen_emails.add(email.lower())
                emails_out.append(email)
            continue

        ok = safe_get(driver, YTJ_COMPANY_URL.format(yt), throttle, stop_evt, status_cb, max_attempts=6)
        if not ok:
            log_cb(f"[FAIL YT] {yt} → (ei saatu ladattua)")
            continue

        email = ""
        for _ in range(10):
            if stop_evt.is_set():
                break
            email = extract_email_from_ytj_page(driver)
            if email:
                break
            time.sleep(0.2)

        if email:
            cache_put_yt(cache, yt, email)
            _save_cache(cache)
            tag = classify_email(email)
            log_cb(f"[LIVE YT] {yt} → {email} ({tag})")
            if email.lower() not in seen_emails:
                seen_emails.add(email.lower())
                emails_out.append(email)
        else:
            log_cb(f"[LIVE YT] {yt} → (ei sähköpostia)")

    # 2) nimet
    for nm in names:
        if stop_evt.is_set():
            break
        done += 1
        progress_cb(done, total)

        nm_clean = (nm or "").strip()
        status_cb(f"YTJ (nimi): {nm_clean} ({done}/{total})")

        hitn = cache_get_by_name(cache, nm_clean)
        if hitn and hitn.get("email"):
            email = hitn["email"]
            tag = classify_email(email)
            log_cb(f"[CACHE NAME] {nm_clean} → {email} ({tag})")
            if email.lower() not in seen_emails:
                seen_emails.add(email.lower())
                emails_out.append(email)
            continue

        if hitn and hitn.get("yt"):
            yt_cached = hitn["yt"]
            hity = cache_get_by_yt(cache, yt_cached)
            if hity and hity.get("email"):
                email = hity["email"]
                cache_put_name(cache, nm_clean, email, yt_cached)
                _save_cache(cache)
                tag = classify_email(email)
                log_cb(f"[CACHE NAME->YT] {nm_clean} ({yt_cached}) → {email} ({tag})")
                if email.lower() not in seen_emails:
                    seen_emails.add(email.lower())
                    emails_out.append(email)
                continue

        ok, yt_found = ytj_open_company_by_name_smart(driver, nm_clean, throttle, stop_evt, status_cb)
        if not ok:
            log_cb(f"[LIVE NAME] ei osumaa: {nm_clean}")
            continue

        email = ""
        for _ in range(10):
            if stop_evt.is_set():
                break
            email = extract_email_from_ytj_page(driver)
            if email:
                break
            time.sleep(0.2)

        if email:
            yt_found = normalize_yt(yt_found) or yt_found
            if yt_found:
                cache_put_yt(cache, yt_found, email, nm_clean)
            cache_put_name(cache, nm_clean, email, yt_found)
            _save_cache(cache)

            tag = classify_email(email)
            log_cb(f"[LIVE NAME] {nm_clean} ({yt_found or 'yt?'}) → {email} ({tag})")
            if email.lower() not in seen_emails:
                seen_emails.add(email.lower())
                emails_out.append(email)
        else:
            cache_put_name(cache, nm_clean, "", yt_found)
            if yt_found:
                cache_put_yt(cache, yt_found, "", nm_clean)
            _save_cache(cache)
            log_cb(f"[LIVE NAME] {nm_clean} ({yt_found or 'yt?'}) → (ei sähköpostia)")

    progress_cb(total, max(1, total))
    return emails_out


# =========================
#   SCROLLABLE TKINTER LAYOUT
# =========================
class ScrollableFrame(ttk.Frame):
    def __init__(self, container):
        super().__init__(container)
        self.canvas = tk.Canvas(self, highlightthickness=0)
        self.scrollbar = ttk.Scrollbar(self, orient="vertical", command=self.canvas.yview)

        self.inner = ttk.Frame(self.canvas)
        self.inner.bind("<Configure>", lambda e: self.canvas.configure(scrollregion=self.canvas.bbox("all")))

        self.canvas_window = self.canvas.create_window((0, 0), window=self.inner, anchor="nw")
        self.canvas.configure(yscrollcommand=self.scrollbar.set)

        self.canvas.pack(side="left", fill="both", expand=True)
        self.scrollbar.pack(side="right", fill="y")

        self.canvas.bind_all("<MouseWheel>", self._on_mousewheel)
        self.canvas.bind("<Configure>", self._on_canvas_configure)

    def _on_canvas_configure(self, event):
        try:
            self.canvas.itemconfig(self.canvas_window, width=event.width)
        except Exception:
            pass

    def _on_mousewheel(self, event):
        try:
            self.canvas.yview_scroll(int(-1 * (event.delta / 120)), "units")
        except Exception:
            pass


# =========================
#   GUI APP
# =========================
class App(tk.Tk):
    def __init__(self):
        super().__init__()
        reset_log()

        self.stop_evt = threading.Event()
        self.last_emails = []
        self.last_yts = []
        self.last_names = []

        self.title("ProtestiBotti v4 (YTJ name-field FIX)")
        self.geometry("1040x860")

        root = ScrollableFrame(self)
        root.pack(fill="both", expand=True)
        self.ui = root.inner

        tk.Label(self.ui, text="ProtestiBotti v4", font=("Arial", 18, "bold")).pack(pady=8)
        tk.Label(
            self.ui,
            text="Uutta:\n• YTJ throttle+retry\n• KL copy/paste: poimi yritysnimet + y-tunnukset → YTJ\n• FIX: nimihaku kirjoittaa oikeaan kenttään 'Yrityksen tai yhteisön nimi'\n",
            justify="center"
        ).pack(pady=4)

        ctrl = tk.Frame(self.ui)
        ctrl.pack(fill="x", padx=12, pady=6)
        tk.Button(ctrl, text="STOP (keskeytä YTJ)", fg="white", bg="#aa0000", command=self.stop_now).pack(side="left", padx=6)
        tk.Button(ctrl, text="Cache stats", command=self.show_cache_stats).pack(side="left", padx=6)
        tk.Button(ctrl, text="Tyhjennä cache", command=self.clear_cache_ui).pack(side="left", padx=6)
        tk.Button(ctrl, text="Kopioi viimeiset emailit", command=self.copy_last_emails).pack(side="left", padx=6)
        tk.Button(ctrl, text="Kopioi viimeiset YT", command=self.copy_last_yts).pack(side="left", padx=6)

        # YT-only
        box_yt = tk.LabelFrame(self.ui, text="NOPEA: Liitä Y-tunnukset → tallenna + hae YTJ sähköpostit", padx=10, pady=10)
        box_yt.pack(fill="x", padx=12, pady=10)

        tk.Label(box_yt, text="Liitä Y-tunnukset (rivit/pilkut/välit käy).", justify="left").pack(anchor="w")
        self.yt_text = tk.Text(box_yt, height=7)
        self.yt_text.pack(fill="x", pady=6)

        row_yt = tk.Frame(box_yt)
        row_yt.pack(fill="x")
        tk.Button(row_yt, text="Tallenna YT (PDF+DOCX+TXT+CSV)", font=("Arial", 11, "bold"), command=self.save_yts_only).pack(side="left", padx=6)
        tk.Button(row_yt, text="Tallenna + Hae YTJ emailit (YT)", font=("Arial", 11, "bold"), command=self.save_and_fetch_yts_only).pack(side="left", padx=6)
        tk.Button(row_yt, text="Tyhjennä", command=lambda: self.yt_text.delete("1.0", tk.END)).pack(side="left", padx=6)

        # KL copy/paste
        box_kl = tk.LabelFrame(self.ui, text="Kauppalehti: liitä koko sivu → poimi nimet + YT → hae emailit YTJ:stä", padx=10, pady=10)
        box_kl.pack(fill="x", padx=12, pady=10)

        tk.Label(box_kl, text="Ctrl+A → Ctrl+C Kauppalehti protestilistasta → liitä tähän:", justify="left").pack(anchor="w")
        self.kl_text = tk.Text(box_kl, height=9)
        self.kl_text.pack(fill="x", pady=6)

        row_kl = tk.Frame(box_kl)
        row_kl.pack(fill="x")
        tk.Button(row_kl, text="Poimi + tee tiedostot (nimet+YT)", font=("Arial", 11, "bold"), command=self.kl_make_files).pack(side="left", padx=6)
        tk.Button(row_kl, text="Poimi + Hae YTJ emailit (nimillä + YT)", font=("Arial", 11, "bold"), command=self.kl_fetch_ytj).pack(side="left", padx=6)
        tk.Button(row_kl, text="Tyhjennä", command=lambda: self.kl_text.delete("1.0", tk.END)).pack(side="left", padx=6)

        # PDF
        box_pdf = tk.LabelFrame(self.ui, text="PDF → (nimet + YT) → YTJ sähköpostit", padx=10, pady=10)
        box_pdf.pack(fill="x", padx=12, pady=10)
        tk.Button(box_pdf, text="Valitse PDF ja hae sähköpostit YTJ:stä", font=("Arial", 11, "bold"), command=self.start_pdf_to_ytj).pack(anchor="w", padx=6, pady=2)

        self.status = tk.Label(self.ui, text="Valmiina.", font=("Arial", 11))
        self.status.pack(pady=6)

        self.progress = ttk.Progressbar(self.ui, orient="horizontal", mode="determinate", length=940)
        self.progress.pack(pady=6)

        frame = tk.Frame(self.ui)
        frame.pack(fill="both", expand=True, padx=14, pady=10)
        tk.Label(frame, text="Live-logi (uusimmat alimmaisena):").pack(anchor="w")
        self.listbox = tk.Listbox(frame, height=14)
        self.listbox.pack(side="left", fill="both", expand=True)
        sb = tk.Scrollbar(frame, orient="vertical", command=self.listbox.yview)
        sb.pack(side="right", fill="y")
        self.listbox.configure(yscrollcommand=sb.set)

        tk.Label(self.ui, text=f"Tallennus: {OUT_DIR}\nCache: {CACHE_PATH}", wraplength=980, justify="center").pack(pady=10)

        self.protocol("WM_DELETE_WINDOW", self.on_close)

    def ui_log(self, msg):
        line = log_to_file(msg)
        self.listbox.insert(tk.END, line)
        self.listbox.yview_moveto(1.0)
        self.update_idletasks()

    def set_status(self, s):
        self.status.config(text=s)
        self.update_idletasks()
        self.ui_log(s)

    def set_progress(self, value, maximum):
        self.progress["maximum"] = max(1, maximum)
        self.progress["value"] = value
        self.update_idletasks()

    def stop_now(self):
        self.stop_evt.set()
        self.set_status("STOP pyydetty. Odota että nykyinen vaihe loppuu…")

    def show_cache_stats(self):
        c = _load_cache()
        yt_n, nm_n = cache_stats(c)
        self.ui_log(f"Cache stats: yt={yt_n}, name={nm_n}")
        messagebox.showinfo("Cache stats", f"Cache:\nYT: {yt_n}\nName: {nm_n}\n\nTiedosto:\n{CACHE_PATH}")

    def clear_cache_ui(self):
        if messagebox.askyesno("Tyhjennä cache", "Haluatko varmasti tyhjentää cachet?"):
            cache_clear()
            self.ui_log("Cache tyhjennetty.")
            messagebox.showinfo("OK", "Cache tyhjennetty.")

    def copy_last_emails(self):
        if not self.last_emails:
            messagebox.showwarning("Ei mitään", "Ei vielä sähköposteja muistissa.")
            return
        s = "\n".join(self.last_emails)
        self.clipboard_clear()
        self.clipboard_append(s)
        self.ui_log("Kopioitu emailit leikepöydälle.")

    def copy_last_yts(self):
        if not self.last_yts:
            messagebox.showwarning("Ei mitään", "Ei vielä Y-tunnuksia muistissa.")
            return
        s = "\n".join(self.last_yts)
        self.clipboard_clear()
        self.clipboard_append(s)
        self.ui_log("Kopioitu Y-tunnukset leikepöydälle.")

    # ---- YT-only ----
    def _read_yts_from_box(self):
        return parse_yts_only(self.yt_text.get("1.0", tk.END))

    def save_yts_only(self):
        yts = self._read_yts_from_box()
        if not yts:
            messagebox.showwarning("Ei löytynyt", "Liitä Y-tunnukset ensin.")
            return

        self.last_yts = yts[:]
        save_word_plain_lines(yts, "ytunnukset_liitetty.docx")
        save_txt_lines(yts, "ytunnukset_liitetty.txt")
        save_pdf_lines(yts, "ytunnukset_liitetty.pdf", title="Y-tunnukset")
        save_csv_rows([[yt] for yt in yts], ["ytunnus"], "ytunnukset_liitetty.csv")

        self.set_status(f"Tallennettu Y-tunnukset: {len(yts)}")
        messagebox.showinfo("Tallennettu", f"Tallennettu {len(yts)} Y-tunnusta.\n\nKansio:\n{OUT_DIR}")

    def save_and_fetch_yts_only(self):
        yts = self._read_yts_from_box()
        if not yts:
            messagebox.showwarning("Ei löytynyt", "Liitä Y-tunnukset ensin.")
            return
        self.stop_evt.clear()
        self.save_yts_only()
        threading.Thread(target=self._run_fetch_ytj, args=(yts, []), daemon=True).start()

    # ---- KL ----
    def _read_kl(self):
        raw = self.kl_text.get("1.0", tk.END)
        return parse_names_and_yts_from_copied_text(raw)

    def kl_make_files(self):
        names, yts = self._read_kl()
        if not names and not yts:
            messagebox.showwarning("Ei löytynyt", "Liitä KL-sivun teksti ja yritä uudelleen.")
            return

        self.last_names = names[:]
        self.last_yts = yts[:]

        if names:
            save_word_plain_lines(names, "yritysnimet_kauppalehti.docx")
            save_txt_lines(names, "yritysnimet_kauppalehti.txt")
            save_pdf_lines(names, "yritysnimet_kauppalehti.pdf", title="Yritysnimet (Kauppalehti)")
            save_csv_rows([[n] for n in names], ["yritysnimi"], "yritysnimet_kauppalehti.csv")

        if yts:
            save_word_plain_lines(yts, "ytunnukset_kauppalehti.docx")
            save_txt_lines(yts, "ytunnukset_kauppalehti.txt")
            save_csv_rows([[yt] for yt in yts], ["ytunnus"], "ytunnukset_kauppalehti.csv")

        combined = []
        if names:
            combined += ["YRITYSNIMET", ""] + names + [""]
        if yts:
            combined += ["Y-TUNNUKSET", ""] + yts
        save_pdf_lines(combined, "kauppalehti_poimitut_tiedot.pdf", title="Kauppalehti → poimitut tiedot")

        self.set_status(f"Poimittu: nimet={len(names)} | y-tunnukset={len(yts)}")
        messagebox.showinfo("Valmis", f"Poimittu:\nNimiä: {len(names)}\nY-tunnuksia: {len(yts)}\n\nKansio:\n{OUT_DIR}")

    def kl_fetch_ytj(self):
        names, yts = self._read_kl()
        if not names and not yts:
            messagebox.showwarning("Ei löytynyt", "Liitä KL-sivun teksti ja yritä uudelleen.")
            return
        self.stop_evt.clear()
        self.kl_make_files()
        threading.Thread(target=self._run_fetch_ytj, args=(yts, names), daemon=True).start()

    # ---- PDF ----
    def start_pdf_to_ytj(self):
        pdf_path = filedialog.askopenfilename(filetypes=[("PDF files", "*.pdf")])
        if not pdf_path:
            return
        self.stop_evt.clear()
        threading.Thread(target=self._run_pdf_to_ytj, args=(pdf_path,), daemon=True).start()

    def _run_pdf_to_ytj(self, pdf_path):
        try:
            self.set_status("Luetaan PDF: nimet + Y-tunnukset…")
            names, yts = extract_names_and_yts_from_pdf(pdf_path)
            if not names and not yts:
                messagebox.showwarning("Ei löytynyt", "PDF:stä ei löytynyt nimiä tai Y-tunnuksia.")
                return

            self.last_names = names[:]
            self.last_yts = yts[:]

            if names:
                save_word_plain_lines(names, "pdf_poimitut_yritysnimet.docx")
                save_pdf_lines(names, "pdf_poimitut_yritysnimet.pdf", title="Yritysnimet (PDF)")
            if yts:
                save_word_plain_lines(yts, "pdf_poimitut_ytunnukset.docx")

            self.set_status("Haetaan sähköpostit YTJ:stä (throttle+retry)…")
            self._run_fetch_ytj(yts, names)
        except Exception as e:
            self.ui_log(f"VIRHE: {e}")
            messagebox.showerror("Virhe", f"Tuli virhe.\n{e}")

    # ---- shared fetch ----
    def _run_fetch_ytj(self, yts, names):
        driver = None
        try:
            self.set_status("Käynnistetään Chrome ja haetaan sähköpostit YTJ:stä…")
            driver = start_new_driver()

            emails = fetch_emails_from_ytj_by_yts_and_names(
                driver,
                yts=yts or [],
                names=names or [],
                status_cb=self.set_status,
                progress_cb=self.set_progress,
                log_cb=self.ui_log,
                stop_evt=self.stop_evt
            )

            emails_sorted = sorted(set([e.strip() for e in emails if e.strip()]), key=lambda x: x.lower())
            self.last_emails = emails_sorted[:]

            save_word_plain_lines(emails_sorted, "sahkopostit_ytj.docx")
            save_txt_lines(emails_sorted, "sahkopostit_ytj.txt")
            save_csv_rows([[e, classify_email(e)] for e in emails_sorted], ["email", "tag"], "sahkopostit_ytj.csv")

            self.set_status("Valmis!")
            messagebox.showinfo("Valmis", f"Sähköposteja: {len(emails_sorted)}\n\nKansio:\n{OUT_DIR}")

        except Exception as e:
            self.ui_log(f"VIRHE: {e}")
            self.set_status("Virhe. Katso log.txt")
            messagebox.showerror("Virhe", f"Tuli virhe.\nKatso log.txt:\n{LOG_PATH}\n\n{e}")
        finally:
            if driver:
                try:
                    driver.quit()
                except Exception:
                    pass

    def on_close(self):
        try:
            self.stop_evt.set()
        except Exception:
            pass
        self.destroy()


if __name__ == "__main__":
    App().mainloop()
