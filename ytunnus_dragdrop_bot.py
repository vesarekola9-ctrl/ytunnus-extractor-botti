import os
import re
import sys
import time
import json
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

# Cache pidetään pysyvästi samassa "ProtestiBotti" juurikansiossa (ei päiväkohtainen)
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
#   CACHE (YTJ)
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


def cache_get_by_yt(cache, yt: str):
    yt = (yt or "").strip()
    return cache.get("yt", {}).get(yt)


def cache_put_yt(cache, yt: str, email: str, name: str = ""):
    yt = (yt or "").strip()
    if not yt:
        return
    cache.setdefault("yt", {})
    cache["yt"][yt] = {
        "email": (email or "").strip(),
        "name": (name or "").strip(),
        "ts": int(time.time()),
    }


def normalize_company_name(name: str) -> str:
    """
    Normalisoi yritysnimen cache-avaimeksi + matchaukseen:
    - lower/casefold
    - poista erikoismerkit/pisteet/ylimääräiset välit
    - normalisoi yritysmuotoja kevyesti
    """
    s = (name or "").strip().casefold()
    s = re.sub(r"\s+", " ", s)
    s = re.sub(r"[.,;:()\[\]{}\"'´`]", "", s)
    s = s.replace("&", " and ")
    s = re.sub(r"\s+", " ", s).strip()

    # Normalisoi yleisiä yrityspäätteitä (ei täydellinen, mutta auttaa)
    # esim "oy", "oyj", "ab", "ky", "tmi", "ry"
    s = re.sub(r"\b(osakeyhtiö|osakeyhtio)\b", "oy", s)
    return s


def cache_get_by_name(cache, name: str):
    key = normalize_company_name(name)
    return cache.get("name", {}).get(key)


def cache_put_name(cache, name: str, email: str, yt: str = ""):
    key = normalize_company_name(name)
    if not key:
        return
    cache.setdefault("name", {})
    cache["name"][key] = {
        "email": (email or "").strip(),
        "yt": (yt or "").strip(),
        "ts": int(time.time()),
        "raw": (name or "").strip(),
    }


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
    pages = []
    cur = []
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

    font_obj = add_obj(
        b"<< /Type /Font /Subtype /Type1 /BaseFont /Helvetica /Encoding /WinAnsiEncoding >>"
    )

    page_objs = []
    pages_obj_id = add_obj(b"<< /Type /Pages /Kids [] /Count 0 >>")

    for page_lines in pages:
        y = TOP
        chunks = []
        chunks.append("BT\n")
        chunks.append(f"/F1 {FONT_SIZE} Tf\n")
        chunks.append(f"{LEFT} {y:.2f} Td\n")

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
    pages_dict = b"<< /Type /Pages /Kids [%s] /Count %d >>" % (kids, len(page_objs))
    objects[pages_obj_id - 1] = pages_dict

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

    pdf_bytes = header + body + xref_bytes + trailer
    with open(path, "wb") as f:
        f.write(pdf_bytes)
    return path


# =========================
#   PARSER
# =========================
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

    suffix_ok = any(x in low for x in [" oy", " ab", " ky", " ay", " ry", " tmi", " oyj", " ltd", " gmbh", " inc", " lkv"])
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

    yts = []
    seen_yts = set()
    for m in YT_RE.findall(raw):
        n = normalize_yt(m)
        if n and n not in seen_yts:
            seen_yts.add(n)
            yts.append(n)

    lines = [ln.strip() for ln in raw.splitlines() if ln and ln.strip()]
    lines = [ln for ln in lines if len(ln) >= 2]

    names = []
    seen = set()

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
            key = best.lower()
            if key not in seen:
                seen.add(key)
                names.append(best)

    if not names:
        for ln in lines:
            ln = _normalize_name(ln)
            if _looks_like_company_name(ln):
                key = ln.lower()
                if key not in seen:
                    seen.add(key)
                    names.append(ln)

    return names, yts


def parse_yts_only(raw: str):
    yts = []
    seen = set()
    for m in YT_RE.findall(raw or ""):
        n = normalize_yt(m)
        if n and n not in seen:
            seen.add(n)
            yts.append(n)
    return sorted(yts)


def extract_names_and_yts_from_pdf(pdf_path: str):
    text_all = ""
    with open(pdf_path, "rb") as f:
        reader = PyPDF2.PdfReader(f)
        for page in reader.pages:
            text_all += (page.extract_text() or "") + "\n"
    return parse_names_and_yts_from_copied_text(text_all)


# =========================
#   YTJ (selenium)
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


def _candidate_container_text(a):
    """
    YTJ-hakutuloksissa linkki on usein kortissa/listassa.
    Poimitaan ympäriltä tekstiä scorettamista varten.
    """
    try:
        # yleiset: li/div/tr
        for xp in ["ancestor::li[1]", "ancestor::div[1]", "ancestor::div[2]", "ancestor::tr[1]"]:
            try:
                c = a.find_element(By.XPATH, xp)
                t = (c.text or "").strip()
                if t:
                    return t
            except Exception:
                continue
    except Exception:
        pass
    try:
        return (a.text or "").strip()
    except Exception:
        return ""


def _score_candidate(target_norm: str, cand_name: str, context_text: str):
    cand_norm = normalize_company_name(cand_name)
    ctx = (context_text or "").casefold()

    score = 0

    # name similarity
    if cand_norm == target_norm and cand_norm:
        score += 120
    if cand_norm and target_norm and cand_norm in target_norm:
        score += 60
    if cand_norm and target_norm and target_norm in cand_norm:
        score += 50

    # token overlap
    tset = set(target_norm.split())
    cset = set(cand_norm.split())
    if tset and cset:
        inter = len(tset & cset)
        score += min(40, inter * 8)

    # business form bonuses if present in both
    for suf in [" oy", " ab", " ky", " ay", " ry", " tmi", " oyj"]:
        if suf.strip() in target_norm and suf.strip() in cand_norm:
            score += 6

    # status hints (best-effort)
    if "aktiivinen" in ctx:
        score += 8
    if "konkurss" in ctx:
        score -= 10
    if "poistettu" in ctx or "lakan" in ctx:
        score -= 8

    return score


def ytj_find_best_result_link(driver, company_name: str, status_cb):
    """
    SMART MATCH:
    - hakee nimellä
    - kerää /yritys/ linkit hakutuloksista
    - scorettää ja valitsee parhaan
    Palauttaa (yt, url) jos löytyy.
    """
    target_norm = normalize_company_name(company_name)
    status_cb("YTJ: haetaan hakutulokset…")

    # Odota että tuloksia renderöityy
    end = time.time() + 12
    links = []
    while time.time() < end:
        try:
            links = driver.find_elements(By.XPATH, "//a[contains(@href,'/yritys/')]")
            links = [a for a in links if (a.get_attribute("href") or "").count("/yritys/") >= 1]
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

            # YTJ joskus näyttää nimen ympärillä, joskus linkin teksti on tyhjä -> ota rivistä eka "nimi" -rivi
            cand_name = txt
            if not cand_name:
                # yritä ottaa ensimmäinen "nimi"-tyylinen rivi kontekstista
                lines = [ln.strip() for ln in (ctx or "").splitlines() if ln.strip()]
                cand_name = lines[0] if lines else ""

            score = _score_candidate(target_norm, cand_name, ctx)

            if score > best_score:
                best_score = score
                best = (href, cand_name, ctx)
        except Exception:
            continue

    if not best:
        return ("", "")

    href, cand_name, ctx = best
    # Poimi yt urlista jos muoto on /yritys/<YT>
    yt = ""
    m = re.search(r"/yritys/([^/?#]+)", href)
    if m:
        yt = m.group(1).strip()
    status_cb(f"YTJ smart match: '{company_name}' → '{cand_name}' (score {best_score})")
    return (yt, href)


def ytj_open_company_by_name_smart(driver, company_name: str, status_cb):
    """
    Avaa yrityssivu nimellä smart matchilla.
    Palauttaa (ok, yt_found).
    """
    status_cb(f"YTJ: haku nimellä: {company_name}")
    driver.get(YTJ_HOME)
    wait_body(driver, 25)
    time.sleep(0.4)

    # etsi hakukenttä
    candidates = []
    for css in ["input[type='search']", "input[type='text']"]:
        try:
            candidates += driver.find_elements(By.CSS_SELECTOR, css)
        except Exception:
            pass

    search_input = None
    for inp in candidates:
        try:
            if inp.is_displayed() and inp.is_enabled():
                search_input = inp
                ph = (inp.get_attribute("placeholder") or "").lower()
                if "hae" in ph or "y-tunnus" in ph or "toiminimi" in ph:
                    break
        except Exception:
            continue

    if not search_input:
        return (False, "")

    try:
        search_input.clear()
    except Exception:
        pass
    search_input.send_keys(company_name)
    search_input.send_keys(Keys.ENTER)

    time.sleep(1.2)

    # jos suoraan yrityssivulle
    if "/yritys/" in (driver.current_url or ""):
        m = re.search(r"/yritys/([^/?#]+)", driver.current_url or "")
        yt = m.group(1).strip() if m else ""
        return (True, yt)

    yt, url = ytj_find_best_result_link(driver, company_name, status_cb)
    if url:
        driver.get(url)
        time.sleep(0.9)
        return ("/yritys/" in (driver.current_url or ""), yt)
    return (False, "")


def fetch_emails_from_ytj_by_yts_and_names(driver, yts, names, status_cb, progress_cb, log_cb, stop_evt):
    cache = _load_cache()

    emails = []
    seen_emails = set()

    total = len(yts) + len(names)
    progress_cb(0, max(1, total))
    done = 0

    # --- 1) Y-tunnus reitti (cache ensin) ---
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
            if email.lower() not in seen_emails:
                seen_emails.add(email.lower())
                emails.append(email)
                log_cb(f"[CACHE YT] {yt} → {email}")
            continue

        driver.get(YTJ_COMPANY_URL.format(yt))
        wait_body(driver, 25)
        time.sleep(0.6)

        email = ""
        for _ in range(10):
            email = extract_email_from_ytj_page(driver)
            if email:
                break
            time.sleep(0.2)

        if email:
            cache_put_yt(cache, yt, email)
            _save_cache(cache)
            if email.lower() not in seen_emails:
                seen_emails.add(email.lower())
                emails.append(email)
                log_cb(f"[LIVE YT] {yt} → {email}")
        else:
            log_cb(f"[LIVE YT] {yt} → (ei sähköpostia)")

    # --- 2) Nimi reitti (cache → smart match → cache) ---
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
            if email.lower() not in seen_emails:
                seen_emails.add(email.lower())
                emails.append(email)
                log_cb(f"[CACHE NAME] {nm_clean} → {email}")
            continue

        # jos cachessa on YT, käytä sitä suoraan
        if hitn and hitn.get("yt"):
            yt_cached = hitn["yt"]
            hity = cache_get_by_yt(cache, yt_cached)
            if hity and hity.get("email"):
                email = hity["email"]
                cache_put_name(cache, nm_clean, email, yt_cached)
                _save_cache(cache)
                if email.lower() not in seen_emails:
                    seen_emails.add(email.lower())
                    emails.append(email)
                    log_cb(f"[CACHE NAME->YT] {nm_clean} ({yt_cached}) → {email}")
                continue

        ok, yt_found = ytj_open_company_by_name_smart(driver, nm_clean, status_cb)
        if not ok:
            log_cb(f"[LIVE NAME] ei osumaa: {nm_clean}")
            continue

        email = ""
        for _ in range(10):
            email = extract_email_from_ytj_page(driver)
            if email:
                break
            time.sleep(0.2)

        if email:
            # päivitä cacheen sekä nimi että yt (jos saatiin)
            if yt_found:
                yt_found = normalize_yt(yt_found) or yt_found
                cache_put_yt(cache, yt_found, email, nm_clean)
            cache_put_name(cache, nm_clean, email, yt_found)
            _save_cache(cache)

            if email.lower() not in seen_emails:
                seen_emails.add(email.lower())
                emails.append(email)
            log_cb(f"[LIVE NAME] {nm_clean} ({yt_found or 'yt?'}) → {email}")
        else:
            # tallenna ainakin yt/name jotta seuraava kierros voi käyttää suoraa yt:tä jos löytyy myöhemmin
            cache_put_name(cache, nm_clean, "", yt_found)
            if yt_found:
                cache_put_yt(cache, yt_found, "", nm_clean)
            _save_cache(cache)
            log_cb(f"[LIVE NAME] {nm_clean} ({yt_found or 'yt?'}) → (ei sähköpostia)")

    return emails


# =========================
#   GUI APP
# =========================
class App(tk.Tk):
    def __init__(self):
        super().__init__()
        reset_log()
        self.stop_evt = threading.Event()

        self.title("ProtestiBotti v2 (CACHE + Smart YTJ Match)")
        self.geometry("1120x960")

        tk.Label(self, text="ProtestiBotti v2", font=("Arial", 18, "bold")).pack(pady=8)
        tk.Label(
            self,
            text="Uutta:\n• Cache (ytj_cache.json): nopeuttaa ja vähentää virheitä\n• Smart YTJ Match: nimihaku valitsee parhaan osuman\n",
            justify="center"
        ).pack(pady=4)

        # ---------- YT-only ----------
        box_yt = tk.LabelFrame(self, text="NOPEA: Liitä Y-tunnuslista → tallenna + hae YTJ sähköpostit", padx=10, pady=10)
        box_yt.pack(fill="x", padx=12, pady=10)

        tk.Label(
            box_yt,
            text="Liitä tähän Y-tunnukset (yksi per rivi). Muoto voi olla 1234567-8 tai 12345678.",
            justify="left"
        ).pack(anchor="w")

        self.yt_text = tk.Text(box_yt, height=7)
        self.yt_text.pack(fill="x", pady=6)

        row_yt = tk.Frame(box_yt)
        row_yt.pack(fill="x")

        tk.Button(row_yt, text="Tallenna YT (PDF+DOCX+TXT)", font=("Arial", 11, "bold"),
                  command=self.save_yts_only).pack(side="left", padx=6)
        tk.Button(row_yt, text="Tallenna + Hae YTJ sähköpostit", font=("Arial", 11, "bold"),
                  command=self.save_and_fetch_yts_only).pack(side="left", padx=6)
        tk.Button(row_yt, text="Tyhjennä", command=lambda: self.yt_text.delete("1.0", tk.END)).pack(side="left", padx=6)

        # ---------- KL copy/paste ----------
        box = tk.LabelFrame(self, text="Kauppalehti: copy/paste → nimet + YT → tiedostot", padx=10, pady=10)
        box.pack(fill="x", padx=12, pady=10)

        tk.Label(
            box,
            text="Vaihtoehto: Ctrl+A → Ctrl+C protestilistasta ja liitä tähän. Parseri yrittää poimia yritysnimet ja Y-tunnukset.",
            justify="left"
        ).pack(anchor="w")

        self.text = tk.Text(box, height=8)
        self.text.pack(fill="x", pady=6)

        row = tk.Frame(box)
        row.pack(fill="x")
        tk.Button(row, text="Tee tiedostot (PDF + DOCX + TXT)", font=("Arial", 11, "bold"),
                  command=self.make_files).pack(side="left", padx=6)
        tk.Button(row, text="Tyhjennä", command=lambda: self.text.delete("1.0", tk.END)).pack(side="left", padx=6)

        # ---------- PDF → YTJ ----------
        box2 = tk.LabelFrame(self, text="PDF → YTJ sähköpostit", padx=10, pady=10)
        box2.pack(fill="x", padx=12, pady=10)

        tk.Button(box2, text="Valitse PDF ja hae sähköpostit YTJ:stä", font=("Arial", 11, "bold"),
                  command=self.start_pdf_to_ytj).pack(anchor="w", padx=6, pady=2)

        # ---------- STATUS / LOG ----------
        self.status = tk.Label(self, text="Valmiina.", font=("Arial", 11))
        self.status.pack(pady=6)

        self.progress = ttk.Progressbar(self, orient="horizontal", mode="determinate", length=1060)
        self.progress.pack(pady=6)

        frame = tk.Frame(self)
        frame.pack(fill="both", expand=True, padx=14, pady=10)

        tk.Label(frame, text="Live-logi (uusimmat alimmaisena):").pack(anchor="w")
        self.listbox = tk.Listbox(frame, height=16)
        self.listbox.pack(side="left", fill="both", expand=True)

        sb = tk.Scrollbar(frame, orient="vertical", command=self.listbox.yview)
        sb.pack(side="right", fill="y")
        self.listbox.configure(yscrollcommand=sb.set)

        tk.Label(self, text=f"Tallennus: {OUT_DIR}\nCache: {CACHE_PATH}", wraplength=1080, justify="center").pack(pady=6)

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

    # =========================
    #   YT-only actions
    # =========================
    def _read_yts_from_box(self):
        raw = self.yt_text.get("1.0", tk.END)
        yts = parse_yts_only(raw)
        return yts

    def save_yts_only(self):
        yts = self._read_yts_from_box()
        if not yts:
            messagebox.showwarning("Ei löytynyt", "Liitä Y-tunnukset ensin.")
            return

        p_docx = save_word_plain_lines(yts, "ytunnukset_liitetty.docx")
        p_txt = save_txt_lines(yts, "ytunnukset_liitetty.txt")
        p_pdf = save_pdf_lines(yts, "ytunnukset_liitetty.pdf", title="Y-tunnukset")

        self.ui_log(f"Tallennettu: {p_docx}")
        self.ui_log(f"Tallennettu: {p_txt}")
        self.ui_log(f"Tallennettu: {p_pdf}")
        self.set_status(f"Tallennettu Y-tunnukset: {len(yts)} kpl")
        messagebox.showinfo("Tallennettu", f"Tallennettu {len(yts)} Y-tunnusta.\n\nKansio:\n{OUT_DIR}")

    def save_and_fetch_yts_only(self):
        yts = self._read_yts_from_box()
        if not yts:
            messagebox.showwarning("Ei löytynyt", "Liitä Y-tunnukset ensin.")
            return

        self.save_yts_only()
        threading.Thread(target=self._run_fetch_emails_yts_only, args=(yts,), daemon=True).start()

    def _run_fetch_emails_yts_only(self, yts):
        driver = None
        try:
            self.set_status("Käynnistetään Chrome ja haetaan sähköpostit YTJ:stä (CACHE + YT)…")
            driver = start_new_driver()

            emails = fetch_emails_from_ytj_by_yts_and_names(
                driver,
                yts=yts,
                names=[],
                status_cb=self.set_status,
                progress_cb=self.set_progress,
                log_cb=self.ui_log,
                stop_evt=self.stop_evt
            )

            em_path = save_word_plain_lines(emails, "sahkopostit_ytj.docx")
            self.ui_log(f"Tallennettu: {em_path}")
            self.set_status("Valmis!")
            messagebox.showinfo("Valmis", f"Sähköposteja löytyi: {len(emails)}\n\nKansio:\n{OUT_DIR}")

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

    # =========================
    #   KL copy/paste -> files
    # =========================
    def make_files(self):
        raw = self.text.get("1.0", tk.END)
        names, yts = parse_names_and_yts_from_copied_text(raw)

        if not names and not yts:
            self.set_status("Ei löytynyt yritysnimiä tai Y-tunnuksia.")
            messagebox.showwarning("Ei löytynyt", "Liitä protestilistan sisältö (Ctrl+A/Ctrl+C) ja yritä uudestaan.")
            return

        self.set_status(f"Poimittu: nimet={len(names)} | y-tunnukset={len(yts)}")

        if names:
            p1 = save_word_plain_lines(names, "yritysnimet_kauppalehti.docx")
            t1 = save_txt_lines(names, "yritysnimet_kauppalehti.txt")
            pdf1 = save_pdf_lines(names, "yritysnimet_kauppalehti.pdf", title="Yritysnimet (Kauppalehti)")
            self.ui_log(f"Tallennettu: {p1}")
            self.ui_log(f"Tallennettu: {t1}")
            self.ui_log(f"Tallennettu: {pdf1}")

        if yts:
            p2 = save_word_plain_lines(yts, "ytunnukset_kauppalehti.docx")
            t2 = save_txt_lines(yts, "ytunnukset_kauppalehti.txt")
            self.ui_log(f"Tallennettu: {p2}")
            self.ui_log(f"Tallennettu: {t2}")

        combined = []
        if names:
            combined.append("YRITYSNIMET")
            combined += names
            combined.append("")
        if yts:
            combined.append("Y-TUNNUKSET")
            combined += yts

        pdf2 = save_pdf_lines(combined, "yritysnimet_ja_ytunnukset.pdf", title="Kauppalehti -> poimitut tiedot")
        self.ui_log(f"Tallennettu: {pdf2}")

        self.set_status("Valmis (tiedostot luotu).")
        messagebox.showinfo("Valmis", f"Valmis!\n\nNimiä: {len(names)}\nY-tunnuksia: {len(yts)}\n\nKansio:\n{OUT_DIR}")

    # =========================
    #   PDF -> YTJ
    # =========================
    def start_pdf_to_ytj(self):
        pdf_path = filedialog.askopenfilename(filetypes=[("PDF files", "*.pdf")])
        if not pdf_path:
            return
        threading.Thread(target=self.run_pdf_to_ytj, args=(pdf_path,), daemon=True).start()

    def run_pdf_to_ytj(self, pdf_path):
        driver = None
        try:
            self.set_status("Luetaan PDF: yritysnimet + Y-tunnukset…")
            names, yts = extract_names_and_yts_from_pdf(pdf_path)

            if not names and not yts:
                self.set_status("PDF:stä ei löytynyt nimiä tai Y-tunnuksia.")
                messagebox.showwarning("Ei löytynyt", "PDF:stä ei löytynyt nimiä tai Y-tunnuksia.")
                return

            self.set_status("Käynnistetään Chrome ja haetaan sähköpostit YTJ:stä (CACHE + Smart Match)…")
            driver = start_new_driver()

            emails = fetch_emails_from_ytj_by_yts_and_names(
                driver,
                yts=yts,
                names=names,
                status_cb=self.set_status,
                progress_cb=self.set_progress,
                log_cb=self.ui_log,
                stop_evt=self.stop_evt
            )

            em_path = save_word_plain_lines(emails, "sahkopostit_ytj.docx")
            self.ui_log(f"Tallennettu: {em_path}")

            self.set_status("Valmis!")
            messagebox.showinfo(
                "Valmis",
                f"Valmis!\n\nNimiä PDF:stä: {len(names)}\nY-tunnuksia PDF:stä: {len(yts)}\nSähköposteja löytyi: {len(emails)}\n\nKansio:\n{OUT_DIR}"
            )

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


if __name__ == "__main__":
    App().mainloop()
