import os
import re
import sys
import time
import json
import csv
import random
import threading
import difflib
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
from selenium.common.exceptions import TimeoutException, WebDriverException, StaleElementReferenceException
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
#   CACHE (negatiivinen ok)
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


def cache_put_yt(cache, yt: str, email: str, name: str = "", status: str = "DONE"):
    yt = (yt or "").strip()
    if not yt:
        return
    cache.setdefault("yt", {})
    cache["yt"][yt] = {"email": (email or "").strip(), "name": (name or "").strip(), "status": status, "ts": int(time.time())}


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


def cache_put_name(cache, name: str, email: str, yt: str = "", status: str = "DONE"):
    key = normalize_company_name(name)
    if not key:
        return
    cache.setdefault("name", {})
    cache["name"][key] = {"email": (email or "").strip(), "yt": (yt or "").strip(), "status": status, "ts": int(time.time()), "raw": (name or "").strip()}


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


def similarity(a: str, b: str) -> float:
    a = normalize_company_name(a)
    b = normalize_company_name(b)
    if not a or not b:
        return 0.0
    return difflib.SequenceMatcher(None, a, b).ratio()


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
#   PDF WITHOUT REPORTLAB (minipdf)
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
#   PARSERS (KL copy/paste)
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


def _normalize_name_line(s: str) -> str:
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
            c = _normalize_name_line(c)
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
            ln = _normalize_name_line(ln)
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
#   SKIP FILTERS (optional)
# =========================
def should_skip_name(name: str, skip_asoy: bool, skip_ry: bool, skip_kiinteisto: bool, skip_julkinen: bool) -> bool:
    n = (name or "").strip()
    low = n.casefold()

    if skip_asoy and (low.startswith("as oy ") or low.startswith("asunto oy ")):
        return True
    if skip_ry and low.endswith(" ry"):
        return True
    if skip_kiinteisto and ("kiinteistö" in low or "kiinteisto" in low):
        return True
    if skip_julkinen:
        public_words = ["kunta", "kaupunki", "seurakunta", "hyvinvointialue", "valtion", "virasto", "oppilaitos"]
        if any(w in low for w in public_words):
            return True

    return False


# =========================
#   YTJ: Selenium + Throttle + Retry
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
        self.max_delay = 7.0
        self.fail_streak = 0
        self.safe_mode = False

    def sleep_between(self):
        base = self.base_delay
        if self.safe_mode:
            base = max(base, 1.5)
        d = min(self.max_delay, base + self.fail_streak * 0.7)
        d += random.uniform(0.0, 0.45)
        time.sleep(d)

    def on_success(self):
        if self.fail_streak > 0:
            self.fail_streak -= 1
        if self.fail_streak == 0:
            self.safe_mode = False

    def on_fail(self):
        self.fail_streak = min(12, self.fail_streak + 1)
        if self.fail_streak >= 3:
            self.safe_mode = True

    def backoff_sleep(self, attempt: int):
        d = min(self.max_delay, (1.0 * (1.7 ** attempt)) + random.uniform(0.0, 0.8))
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


def safe_get(driver, url: str, throttle: Throttle, stop_evt, status_cb, max_attempts=6):
    for attempt in range(max_attempts):
        if stop_evt.is_set():
            return False
        try:
            throttle.sleep_between()
            driver.get(url)
            wait_body(driver, 25)

            if page_looks_blocked_or_captcha(driver):
                status_cb("YTJ näyttää estolta/CAPTCHA:lta → hidastus…")
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


# =========================
#   YTJ: ROBUST "SÄHKÖPOSTI -> NÄYTÄ" CLICK (FIX)
# =========================
def _js_click(driver, el) -> bool:
    try:
        driver.execute_script("arguments[0].scrollIntoView({block:'center'});", el)
        time.sleep(0.05)
    except Exception:
        pass
    try:
        el.click()
        return True
    except Exception:
        pass
    try:
        driver.execute_script("arguments[0].click();", el)
        return True
    except Exception:
        return False


def ytj_click_email_nayta(driver) -> bool:
    """
    Klikkaa nimenomaan Sähköposti-rivin "Näytä".
    Tämä on se mitä sulla jää usein painamatta.
    """
    xpaths = [
        # Table row where left cell contains "Sähköposti", then button/role=button with text "Näytä"
        "//tr[.//*[contains(normalize-space(.),'Sähköposti')]]"
        "//*[self::button or self::a or @role='button'][normalize-space()='Näytä']",
        # Sometimes not inside <tr> (responsive), use any container containing "Sähköposti"
        "//*[contains(normalize-space(.),'Sähköposti')]/following::*"
        "[self::button or self::a or @role='button'][normalize-space()='Näytä'][1]",
    ]
    for xp in xpaths:
        try:
            btns = driver.find_elements(By.XPATH, xp)
            for b in btns:
                try:
                    if b.is_displayed() and b.is_enabled():
                        if _js_click(driver, b):
                            return True
                except Exception:
                    continue
        except Exception:
            continue
    return False


def ytj_click_all_nayta(driver):
    """
    1) Klikkaa ensin varmasti Sähköposti->Näytä
    2) Klikkaa sitten kaikki muut "Näytä" jos näkyy (puhelin jne)
    """
    # 1) EMAIL first
    ytj_click_email_nayta(driver)
    time.sleep(0.15)

    # 2) others
    for _ in range(4):
        clicked = False
        try:
            buttons = driver.find_elements(By.XPATH, "//button|//a|//*[@role='button']")
        except Exception:
            buttons = []

        for b in buttons:
            try:
                if not b.is_displayed() or not b.is_enabled():
                    continue
                txt = (b.text or "").strip()
                if txt.casefold() == "näytä":
                    if _js_click(driver, b):
                        clicked = True
                        time.sleep(0.12)
            except StaleElementReferenceException:
                continue
            except Exception:
                continue

        if not clicked:
            break


def extract_email_from_ytj_page(driver):
    # 1) mailto
    try:
        for a in driver.find_elements(By.TAG_NAME, "a"):
            href = (a.get_attribute("href") or "")
            if href.lower().startswith("mailto:"):
                return href.split(":", 1)[1].strip()
    except Exception:
        pass

    # 2) body regex
    try:
        body = driver.find_element(By.TAG_NAME, "body").text or ""
        return pick_email_from_text(body)
    except Exception:
        return ""


def wait_email_appears(driver, timeout=6.0) -> str:
    """
    Odota että email ilmestyy sivulle (yleensä Sähköposti-riville) klikkauksen jälkeen.
    """
    end = time.time() + timeout
    last = ""
    while time.time() < end:
        try:
            body = driver.find_element(By.TAG_NAME, "body").text or ""
            last = pick_email_from_text(body)
            if last:
                return last
        except Exception:
            pass
        time.sleep(0.25)
    return last


def extract_company_name_from_ytj_page(driver) -> str:
    for tag in ("h1", "h2"):
        try:
            el = driver.find_element(By.TAG_NAME, tag)
            txt = (el.text or "").strip()
            if txt and len(txt) >= 3:
                return txt
        except Exception:
            pass
    try:
        t = driver.title or ""
        t = t.replace("YTJ", "").strip(" -|")
        return t.strip()
    except Exception:
        return ""


# =========================
#   FIX: Find correct YTJ name field
# =========================
def find_input_by_label_text(driver, label_text: str):
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

    xpaths = [
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


# =========================
#   SMART SEARCH RESULTS
# =========================
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


def _score_candidate(target_name: str, cand_name: str, context_text: str):
    target_norm = normalize_company_name(target_name)
    cand_norm = normalize_company_name(cand_name)
    ctx = (context_text or "").casefold()
    score = 0

    if cand_norm == target_norm and cand_norm:
        score += 120

    tset = set(target_norm.split())
    cset = set(cand_norm.split())
    if tset and cset:
        score += min(55, len(tset & cset) * 9)

    score += int(similarity(target_name, cand_name) * 40)

    if "aktiivinen" in ctx:
        score += 10
    if "konkurss" in ctx:
        score -= 10
    if "poistettu" in ctx or "lakan" in ctx:
        score -= 8

    return score


def ytj_rank_result_links(driver, company_name: str, limit=12):
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

    cand_list = []
    for a in links[:60]:
        try:
            href = a.get_attribute("href") or ""
            if "/yritys/" not in href:
                continue
            shown = (a.text or "").strip()
            ctx = _candidate_container_text(a)
            cand_name = shown
            if not cand_name:
                lines = [ln.strip() for ln in (ctx or "").splitlines() if ln.strip()]
                cand_name = lines[0] if lines else ""
            score = _score_candidate(company_name, cand_name, ctx)
            yt = ""
            m = re.search(r"/yritys/([^/?#]+)", href)
            if m:
                yt = m.group(1).strip()
            cand_list.append((yt, href, cand_name, score))
        except Exception:
            continue

    cand_list.sort(key=lambda x: x[3], reverse=True)
    return cand_list[:limit]


def ytj_open_best_company_by_name_verified(driver, company_name: str, throttle: Throttle, stop_evt, status_cb):
    ok = safe_get(driver, YTJ_HOME, throttle, stop_evt, status_cb, max_attempts=6)
    if not ok:
        return (False, "", "")

    time.sleep(0.6)

    name_input = find_input_by_label_text(driver, "Yrityksen tai yhteisön nimi")
    if not name_input:
        status_cb("YTJ: en löytänyt kenttää 'Yrityksen tai yhteisön nimi'.")
        return (False, "", "")

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
        page_name = extract_company_name_from_ytj_page(driver)
        return (True, yt, page_name)

    cands = ytj_rank_result_links(driver, company_name, limit=12)
    if not cands:
        return (False, "", "")

    best_seen = None
    for idx, (yt, url, shown_name, score) in enumerate(cands[:8], start=1):
        if stop_evt.is_set():
            return (False, "", "")

        target_url = YTJ_COMPANY_URL.format(yt) if yt else url
        status_cb(f"YTJ: kokeilen tulosta {idx}/{min(8, len(cands))} (score={score})")
        ok2 = safe_get(driver, target_url, throttle, stop_evt, status_cb, max_attempts=6)
        if not ok2:
            continue

        page_name = extract_company_name_from_ytj_page(driver)
        sim = similarity(company_name, page_name or shown_name)
        if sim >= 0.72:
            return (True, yt, page_name)

        if best_seen is None or sim > best_seen[0]:
            best_seen = (sim, yt, page_name)

    if best_seen and best_seen[0] >= 0.60:
        return (True, best_seen[1], best_seen[2])

    return (False, "", "")


# =========================
#   FETCH (YT + NAME) + MASTER.CSV
#   - NEGATIVE CACHE OK
#   - OUTPUT EMAIL DOCX ALWAYS FILTERED (no empty)
#   - YTJ clicks "Sähköposti -> Näytä" robustly + waits for reveal
# =========================
def fetch_emails_from_ytj_by_yts_and_names(
    driver,
    yts,
    names,
    status_cb,
    progress_cb,
    log_cb,
    stop_evt,
    skip_asoy=False,
    skip_ry=False,
    skip_kiinteisto=False,
    skip_julkinen=False
):
    cache = _load_cache()
    throttle = Throttle()

    master_rows = []  # [yritysnimi, yt, email, tag, status, source, ts]
    emails_out = []
    seen_emails = set()

    def add_email(email: str):
        if not email:
            return
        e = email.strip()
        if not e:
            return
        k = e.lower()
        if k in seen_emails:
            return
        seen_emails.add(k)
        emails_out.append(e)

    # filter names
    names_in = []
    for nm in (names or []):
        nm = (nm or "").strip()
        if not nm:
            continue
        if should_skip_name(nm, skip_asoy, skip_ry, skip_kiinteisto, skip_julkinen):
            master_rows.append([nm, "", "", "", "SKIPPED", "NAME", int(time.time())])
            log_cb(f"[SKIP] {nm}")
            continue
        names_in.append(nm)

    total = len(yts or []) + len(names_in)
    progress_cb(0, max(1, total))
    done = 0

    # 1) YT
    for yt in (yts or []):
        if stop_evt.is_set():
            break
        done += 1
        progress_cb(done, total)

        yt = normalize_yt(yt) or (yt or "").strip()
        if not yt:
            continue

        status_cb(f"YTJ (Y-tunnus): {yt} ({done}/{total})")

        hit = cache_get_by_yt(cache, yt)
        if hit is not None:
            email = (hit.get("email") or "").strip()
            tag = classify_email(email) if email else ""
            st = hit.get("status") or ("FOUND" if email else "NO_EMAIL")
            master_rows.append(["", yt, email, tag, st, "YT(CACHE)", int(time.time())])
            if email:
                log_cb(f"[CACHE YT] {yt} → {email} ({tag})")
                add_email(email)
            else:
                log_cb(f"[CACHE YT] {yt} → (ei emailia)")
            continue

        ok = safe_get(driver, YTJ_COMPANY_URL.format(yt), throttle, stop_evt, status_cb, max_attempts=6)
        if not ok:
            cache_put_yt(cache, yt, "", "", status="ERROR")
            _save_cache(cache)
            master_rows.append(["", yt, "", "", "ERROR", "YT", int(time.time())])
            log_cb(f"[FAIL YT] {yt} → (ei saatu ladattua)")
            continue

        # ✅ FIX: click email reveal robustly
        clicked_email = ytj_click_email_nayta(driver)
        ytj_click_all_nayta(driver)

        email = extract_email_from_ytj_page(driver)
        if not email and clicked_email:
            email = wait_email_appears(driver, timeout=7.0)
        if not email:
            # try again (some pages need second click due to sticky header overlays)
            ytj_click_email_nayta(driver)
            email = wait_email_appears(driver, timeout=5.5)

        if email:
            cache_put_yt(cache, yt, email, "", status="FOUND")
            _save_cache(cache)
            tag = classify_email(email)
            master_rows.append(["", yt, email, tag, "FOUND", "YT", int(time.time())])
            log_cb(f"[LIVE YT] {yt} → {email} ({tag})")
            add_email(email)
        else:
            cache_put_yt(cache, yt, "", "", status="NO_EMAIL")
            _save_cache(cache)
            master_rows.append(["", yt, "", "", "NO_EMAIL", "YT", int(time.time())])
            log_cb(f"[LIVE YT] {yt} → (ei sähköpostia)")

    # 2) NAMES
    for nm in names_in:
        if stop_evt.is_set():
            break
        done += 1
        progress_cb(done, total)

        status_cb(f"YTJ (nimi): {nm} ({done}/{total})")

        hitn = cache_get_by_name(cache, nm)
        if hitn is not None:
            email = (hitn.get("email") or "").strip()
            yt_cached = (hitn.get("yt") or "").strip()
            tag = classify_email(email) if email else ""
            st = hitn.get("status") or ("FOUND" if email else "NO_EMAIL")
            master_rows.append([nm, yt_cached, email, tag, st, "NAME(CACHE)", int(time.time())])
            if email:
                log_cb(f"[CACHE NAME] {nm} → {email} ({tag})")
                add_email(email)
            else:
                log_cb(f"[CACHE NAME] {nm} → (ei emailia)")
            continue

        ok, yt_found, _page_name = ytj_open_best_company_by_name_verified(driver, nm, throttle, stop_evt, status_cb)
        if not ok:
            cache_put_name(cache, nm, "", "", status="NO_MATCH")
            _save_cache(cache)
            master_rows.append([nm, "", "", "", "NO_MATCH", "NAME", int(time.time())])
            log_cb(f"[LIVE NAME] ei osumaa: {nm}")
            continue

        yt_found = normalize_yt(yt_found) or (yt_found or "").strip()

        if yt_found:
            ok2 = safe_get(driver, YTJ_COMPANY_URL.format(yt_found), throttle, stop_evt, status_cb, max_attempts=6)
            if not ok2:
                cache_put_name(cache, nm, "", yt_found, status="ERROR")
                cache_put_yt(cache, yt_found, "", nm, status="ERROR")
                _save_cache(cache)
                master_rows.append([nm, yt_found, "", "", "ERROR", "NAME", int(time.time())])
                log_cb(f"[LIVE NAME] {nm} ({yt_found}) → (ei saatu yrityssivua)")
                continue

        clicked_email = ytj_click_email_nayta(driver)
        ytj_click_all_nayta(driver)

        email = extract_email_from_ytj_page(driver)
        if not email and clicked_email:
            email = wait_email_appears(driver, timeout=7.0)
        if not email:
            ytj_click_email_nayta(driver)
            email = wait_email_appears(driver, timeout=5.5)

        if email:
            tag = classify_email(email)
            cache_put_name(cache, nm, email, yt_found, status="FOUND")
            if yt_found:
                cache_put_yt(cache, yt_found, email, nm, status="FOUND")
            _save_cache(cache)
            master_rows.append([nm, yt_found, email, tag, "FOUND", "NAME", int(time.time())])
            log_cb(f"[LIVE NAME] {nm} ({yt_found or 'yt?'}) → {email} ({tag})")
            add_email(email)
        else:
            cache_put_name(cache, nm, "", yt_found, status="NO_EMAIL")
            if yt_found:
                cache_put_yt(cache, yt_found, "", nm, status="NO_EMAIL")
            _save_cache(cache)
            master_rows.append([nm, yt_found, "", "", "NO_EMAIL", "NAME", int(time.time())])
            log_cb(f"[LIVE NAME] {nm} ({yt_found or 'yt?'}) → (ei sähköpostia)")

    progress_cb(total, max(1, total))

    # ✅ Final output: no empties
    emails_out = sorted(set([e.strip() for e in emails_out if e and e.strip()]), key=lambda x: x.lower())
    return emails_out, master_rows


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

        self.title("ProtestiBotti ULTIMATE (YTJ Näytä FIXED)")
        self.geometry("1100x940")

        root = ScrollableFrame(self)
        root.pack(fill="both", expand=True)
        self.ui = root.inner

        tk.Label(self.ui, text="ProtestiBotti ULTIMATE", font=("Arial", 18, "bold")).pack(pady=8)
        tk.Label(
            self.ui,
            text=(
                "• KL copy/paste parser → yritysnimet + YT\n"
                "• PDF → (nimet+YT) → YTJ\n"
                "• YTJ: klikataan varmasti Sähköposti → Näytä (robust)\n"
                "• Cache myös NO_EMAIL, mutta lopullinen DOCX sisältää vain oikeat emailit\n"
                "• master.csv statusriveillä\n"
            ),
            justify="center"
        ).pack(pady=4)

        ctrl = tk.Frame(self.ui)
        ctrl.pack(fill="x", padx=12, pady=6)
        tk.Button(ctrl, text="STOP (keskeytä YTJ)", fg="white", bg="#aa0000", command=self.stop_now).pack(side="left", padx=6)
        tk.Button(ctrl, text="Avaa output-kansio", command=self.open_output_folder).pack(side="left", padx=6)
        tk.Button(ctrl, text="Tyhjennä cache", command=self.clear_cache_ui).pack(side="left", padx=6)

        filt = tk.LabelFrame(self.ui, text="Filtterit (nimihaussa)", padx=10, pady=8)
        filt.pack(fill="x", padx=12, pady=8)
        self.var_skip_asoy = tk.BooleanVar(value=False)
        self.var_skip_ry = tk.BooleanVar(value=False)
        self.var_skip_kiinteisto = tk.BooleanVar(value=False)
        self.var_skip_julkinen = tk.BooleanVar(value=False)
        tk.Checkbutton(filt, text="Skip As Oy", variable=self.var_skip_asoy).pack(side="left", padx=8)
        tk.Checkbutton(filt, text="Skip Ry", variable=self.var_skip_ry).pack(side="left", padx=8)
        tk.Checkbutton(filt, text="Skip Kiinteistö*", variable=self.var_skip_kiinteisto).pack(side="left", padx=8)
        tk.Checkbutton(filt, text="Skip Julkinen", variable=self.var_skip_julkinen).pack(side="left", padx=8)

        box_yt = tk.LabelFrame(self.ui, text="Liitä Y-tunnukset → YTJ emailit", padx=10, pady=10)
        box_yt.pack(fill="x", padx=12, pady=10)
        tk.Label(box_yt, text="Liitä Y-tunnukset (rivit/pilkut/välit käy).").pack(anchor="w")
        self.yt_text = tk.Text(box_yt, height=7)
        self.yt_text.pack(fill="x", pady=6)

        row_yt = tk.Frame(box_yt)
        row_yt.pack(fill="x")
        tk.Button(row_yt, text="Hae YTJ emailit (YT)", font=("Arial", 11, "bold"), command=self.save_and_fetch_yts_only).pack(side="left", padx=6)

        box_kl = tk.LabelFrame(self.ui, text="Kauppalehti: liitä koko sivu → poimi nimet/YT → YTJ emailit", padx=10, pady=10)
        box_kl.pack(fill="x", padx=12, pady=10)
        tk.Label(box_kl, text="Ctrl+A → Ctrl+C protestilistasta → liitä tähän:").pack(anchor="w")
        self.kl_text = tk.Text(box_kl, height=10)
        self.kl_text.pack(fill="x", pady=6)
        tk.Button(box_kl, text="Poimi + Hae YTJ emailit (nimillä + YT)", font=("Arial", 11, "bold"), command=self.kl_fetch_ytj).pack(anchor="w", padx=6)

        box_pdf = tk.LabelFrame(self.ui, text="PDF → (nimet+YT) → YTJ emailit", padx=10, pady=10)
        box_pdf.pack(fill="x", padx=12, pady=10)
        tk.Button(box_pdf, text="Valitse PDF ja hae emailit", font=("Arial", 11, "bold"), command=self.start_pdf_to_ytj).pack(anchor="w", padx=6, pady=2)

        self.status = tk.Label(self.ui, text="Valmiina.", font=("Arial", 11))
        self.status.pack(pady=6)

        self.progress = ttk.Progressbar(self.ui, orient="horizontal", mode="determinate", length=1000)
        self.progress.pack(pady=6)

        frame = tk.Frame(self.ui)
        frame.pack(fill="both", expand=True, padx=14, pady=10)
        tk.Label(frame, text="Live-logi:").pack(anchor="w")
        self.listbox = tk.Listbox(frame, height=14)
        self.listbox.pack(side="left", fill="both", expand=True)
        sb = tk.Scrollbar(frame, orient="vertical", command=self.listbox.yview)
        sb.pack(side="right", fill="y")
        self.listbox.configure(yscrollcommand=sb.set)

        tk.Label(self.ui, text=f"Tallennus: {OUT_DIR}\nCache: {CACHE_PATH}", wraplength=1020, justify="center").pack(pady=10)
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
        self.set_status("STOP pyydetty…")

    def open_output_folder(self):
        try:
            if sys.platform.startswith("win"):
                os.startfile(OUT_DIR)  # noqa
            else:
                messagebox.showinfo("Polku", OUT_DIR)
        except Exception as e:
            messagebox.showerror("Virhe", f"Ei voitu avata kansiota.\n{e}")

    def clear_cache_ui(self):
        if messagebox.askyesno("Tyhjennä cache", "Haluatko varmasti tyhjentää cachet?"):
            cache_clear()
            self.ui_log("Cache tyhjennetty.")
            messagebox.showinfo("OK", "Cache tyhjennetty.")

    def _read_yts_from_box(self):
        return parse_yts_only(self.yt_text.get("1.0", tk.END))

    def save_and_fetch_yts_only(self):
        yts = self._read_yts_from_box()
        if not yts:
            messagebox.showwarning("Ei löytynyt", "Liitä Y-tunnukset ensin.")
            return
        self.stop_evt.clear()
        threading.Thread(target=self._run_fetch_ytj, args=(yts, []), daemon=True).start()

    def _read_kl(self):
        raw = self.kl_text.get("1.0", tk.END)
        return parse_names_and_yts_from_copied_text(raw)

    def kl_fetch_ytj(self):
        names, yts = self._read_kl()
        if not names and not yts:
            messagebox.showwarning("Ei löytynyt", "Liitä KL-sivun teksti ensin.")
            return
        self.stop_evt.clear()
        threading.Thread(target=self._run_fetch_ytj, args=(yts, names), daemon=True).start()

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
            self.set_status("Haetaan emailit YTJ:stä…")
            self._run_fetch_ytj(yts, names)
        except Exception as e:
            self.ui_log(f"VIRHE: {e}")
            messagebox.showerror("Virhe", str(e))

    def _run_fetch_ytj(self, yts, names):
        driver = None
        try:
            self.set_status("Käynnistetään Chrome (YTJ)…")
            driver = start_new_driver()

            emails, master_rows = fetch_emails_from_ytj_by_yts_and_names(
                driver=driver,
                yts=yts or [],
                names=names or [],
                status_cb=self.set_status,
                progress_cb=self.set_progress,
                log_cb=self.ui_log,
                stop_evt=self.stop_evt,
                skip_asoy=self.var_skip_asoy.get(),
                skip_ry=self.var_skip_ry.get(),
                skip_kiinteisto=self.var_skip_kiinteisto.get(),
                skip_julkinen=self.var_skip_julkinen.get(),
            )

            emails_sorted = sorted(set([e for e in emails if e and e.strip()]), key=lambda x: x.lower())
            self.last_emails = emails_sorted[:]

            save_word_plain_lines(emails_sorted, "sahkopostit_ytj.docx")
            save_txt_lines(emails_sorted, "sahkopostit_ytj.txt")
            save_csv_rows([[e, classify_email(e)] for e in emails_sorted], ["email", "tag"], "sahkopostit_ytj.csv")
            save_csv_rows(master_rows, ["yritysnimi", "ytunnus", "email", "tag", "status", "source", "ts"], "master.csv")

            self.set_status("Valmis!")
            messagebox.showinfo("Valmis", f"Sähköposteja: {len(emails_sorted)}\n\nKansio:\n{OUT_DIR}")

        except Exception as e:
            self.ui_log(f"VIRHE: {e}")
            self.set_status("Virhe. Katso log.txt")
            messagebox.showerror("Virhe", f"{e}\n\nKatso log.txt:\n{LOG_PATH}")
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
