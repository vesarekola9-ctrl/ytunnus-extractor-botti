# protestibotti.py
# ProtestiBotti:
#   1) PDF -> Y-tunnukset -> YTJ sähköpostit
#   2) Kauppalehti (Chrome debug 9222) -> Y-tunnukset -> YTJ sähköpostit
#   3) Clipboard (Ctrl+C -> Ctrl+V) -> yritysnimet -> YTJ: Y-tunnukset -> YTJ: sähköpostit (Näytä klikataan)
#
# FIX:
# - Sähköpostit tallentuvat aina uuteen Wordiin saman päivän kansiossa:
#     sahkopostit_001.docx, sahkopostit_002.docx, ...
#   eikä ylikirjoita aiempaa.
#
# UI:
# - Musta tausta + punaiset napit
# - Kuvia kulmiin + glow (Pillowlla) + hover nappeihin
#
# Riippuvuudet:
#   pip install selenium webdriver-manager PyPDF2 python-docx tkinterdnd2
#   pip install pillow   (SUOSITUS: jotta .jpg kuvat + glow toimii)
#
# Kuvat samaan kansioon kuin tämä .py / exe:
#   h1.jpg, h.jpg, uo.png, polis.jpg, vero.png
#
# Build:
#   pyinstaller --noconfirm --onefile --windowed --name ProtestiBotti protestibotti.py

import os
import re
import sys
import time
import threading
import tkinter as tk
from tkinter import messagebox, filedialog
from tkinter import ttk
from difflib import SequenceMatcher
import subprocess

import PyPDF2
from docx import Document

from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import (
    WebDriverException,
    TimeoutException,
)
from webdriver_manager.chrome import ChromeDriverManager

# Drag & Drop (PDF)
try:
    from tkinterdnd2 import DND_FILES, TkinterDnD  # type: ignore
    HAS_DND = True
except Exception:
    HAS_DND = False

# Optional: Pillow for JPG/PNG scaling + glow
try:
    from PIL import Image, ImageTk, ImageFilter, ImageOps  # type: ignore
    HAS_PIL = True
except Exception:
    HAS_PIL = False

# =========================
#   TUNING (NOPEUS)
# =========================
YTJ_PAGE_LOAD_TIMEOUT = 18
YTJ_RETRY_READS = 5
YTJ_RETRY_SLEEP = 0.12
YTJ_NAYTA_PASSES = 2
YTJ_PER_COMPANY_SLEEP = 0.03

KL_LOAD_MORE_WAIT = 1.1
KL_COMPANY_PAGE_TIMEOUT = 18
KL_AFTER_OPEN_SLEEP = 0.05

CLIP_YTJ_SEARCH_TIMEOUT = 10.0
CLIP_PER_NAME_SLEEP = 0.02

PARTIAL_SAVE_EVERY_NEW_EMAILS = 25

# =========================
#   CONFIG / REGEX
# =========================
YT_RE = re.compile(r"\b\d{7}-\d\b|\b\d{8}\b")
EMAIL_RE = re.compile(r"[A-Za-z0-9_.+-]+@[A-Za-z0-9-]+\.[A-Za-z0-9-.]+")
EMAIL_A_RE = re.compile(r"[A-Za-z0-9_.+-]+\s*\(a\)\s*[A-Za-z0-9-]+\.[A-Za-z0-9-.]+", re.I)

KAUPPALEHTI_URL = "https://www.kauppalehti.fi/yritykset/protestilista"
KAUPPALEHTI_MATCH = "kauppalehti.fi/yritykset/protestilista"
YTJ_COMPANY_URL = "https://tietopalvelu.ytj.fi/yritys/{}"

STRICT_FORMS_RE = re.compile(
    r"\b(oy|ab|ky|tmi|oyj|osakeyhtiö|kommandiittiyhtiö|toiminimi|as\.|ltd|llc|inc|gmbh)\b",
    re.IGNORECASE,
)

# =========================
#   PATHS
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
EMAILS_TMP_PATH = os.path.join(OUT_DIR, "emails_tmp.txt")
PARTIAL_DOCX_PATH = os.path.join(OUT_DIR, "sahkopostit_partial.docx")


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
    try:
        with open(EMAILS_TMP_PATH, "w", encoding="utf-8") as f:
            f.write("")
    except Exception:
        pass
    log_to_file(f"Output: {OUT_DIR}")
    log_to_file(f"Logi: {LOG_PATH}")


# =========================
#   UNIQUE FILE NAMES (FIX)
# =========================
def next_indexed_docx(prefix: str, start_at: int = 1) -> str:
    """
    Returns a path like OUT_DIR/prefix_001.docx that doesn't exist yet.
    Example: prefix="sahkopostit" -> sahkopostit_001.docx, sahkopostit_002.docx, ...
    """
    i = start_at
    while True:
        name = f"{prefix}_{i:03d}.docx"
        path = os.path.join(OUT_DIR, name)
        if not os.path.exists(path):
            return path
        i += 1


def save_word_unique(lines, prefix: str):
    path = next_indexed_docx(prefix)
    doc = Document()
    for line in lines:
        if line:
            doc.add_paragraph(line)
    doc.save(path)
    return path


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
        if line:
            doc.add_paragraph(line)
    doc.save(path)
    return path


def save_word_plain_lines_to_path(lines, path):
    doc = Document()
    for line in lines:
        if line:
            doc.add_paragraph(line)
    doc.save(path)
    return path


def append_email_tmp(email: str):
    try:
        with open(EMAILS_TMP_PATH, "a", encoding="utf-8") as f:
            f.write(email.strip() + "\n")
    except Exception:
        pass


def safe_scroll_into_view(driver, elem):
    try:
        driver.execute_script("arguments[0].scrollIntoView({block:'center'});", elem)
    except Exception:
        pass


def safe_click(driver, elem) -> bool:
    try:
        safe_scroll_into_view(driver, elem)
        time.sleep(0.01)
        try:
            elem.click()
        except Exception:
            driver.execute_script("arguments[0].click();", elem)
        return True
    except Exception:
        return False


def try_accept_cookies(driver):
    texts = ["Hyväksy", "Hyväksy kaikki", "Salli kaikki", "Accept", "Accept all", "I agree", "OK", "Selvä"]
    for _ in range(2):
        found = False
        for e in driver.find_elements(By.XPATH, "//button|//a|//*[@role='button']"):
            try:
                t = (e.text or "").strip()
                if not t:
                    continue
                low = t.lower()
                if any(x.lower() in low for x in texts):
                    if e.is_displayed() and e.is_enabled():
                        safe_click(driver, e)
                        time.sleep(0.2)
                        found = True
                        break
            except Exception:
                continue
        if not found:
            break


def split_lines(text: str):
    if not text:
        return []
    text = text.replace("\r\n", "\n").replace("\r", "\n")
    return [ln.strip() for ln in text.split("\n") if ln.strip()]


def extract_yts_from_text(text: str):
    yts = set()
    for m in YT_RE.findall(text or ""):
        n = normalize_yt(m)
        if n:
            yts.add(n)
    return sorted(yts)


def extract_names_from_clipboard(text: str, strict: bool, max_names: int):
    lines = split_lines(text)
    out = []
    seen = set()

    bad_contains = [
        "näytä lisää", "protestilista", "kauppalehti", "kirjaudu", "tilaa", "tilaajille",
        "€", "eur", "summa", "viiväst", "päivä", "päivää", "päivämäärä",
        "y-tunnus", "y tunnus", "ytunnus", "osoite", "postinumero",
        "sähköposti", "puhelin", "www.", "http",
    ]

    for ln in lines:
        if len(out) >= max_names:
            break

        low = ln.lower()
        if YT_RE.search(ln):
            continue
        if any(b in low for b in bad_contains):
            continue
        if len(ln) < 3:
            continue
        digits = sum(ch.isdigit() for ch in ln)
        if digits >= 3:
            continue
        if not any(ch.isalpha() for ch in ln):
            continue

        name = re.sub(r"\s{2,}", " ", ln).strip()
        if len(name) > 80:
            continue
        if strict and not STRICT_FORMS_RE.search(name):
            continue

        key = name.lower()
        if key in seen:
            continue

        seen.add(key)
        out.append(name)

    return out


def extract_yt_from_text_anywhere(txt: str) -> str:
    if not txt:
        return ""
    for m in YT_RE.findall(txt):
        n = normalize_yt(m)
        if n:
            return n
    return ""


# =========================
#   PDF -> YTs
# =========================
def extract_ytunnukset_from_pdf(pdf_path: str):
    yt_set = set()
    with open(pdf_path, "rb") as f:
        reader = PyPDF2.PdfReader(f)
        for page in reader.pages:
            text = page.extract_text() or ""
            for m in YT_RE.findall(text):
                n = normalize_yt(m)
                if n:
                    yt_set.add(n)
    return sorted(yt_set)


# =========================
#   SELENIUM START
# =========================
def start_new_driver():
    options = webdriver.ChromeOptions()
    options.add_argument("--start-maximized")
    options.add_argument("--disable-gpu")
    options.add_argument("--disable-dev-shm-usage")
    driver_path = ChromeDriverManager().install()
    drv = webdriver.Chrome(service=Service(driver_path), options=options)
    drv.set_page_load_timeout(YTJ_PAGE_LOAD_TIMEOUT)
    return drv


def attach_to_existing_chrome():
    options = webdriver.ChromeOptions()
    options.add_experimental_option("debuggerAddress", "127.0.0.1:9222")
    driver_path = ChromeDriverManager().install()
    drv = webdriver.Chrome(service=Service(driver_path), options=options)
    drv.set_page_load_timeout(YTJ_PAGE_LOAD_TIMEOUT)
    return drv


def open_new_tab(driver, url="about:blank"):
    driver.execute_script("window.open(arguments[0], '_blank');", url)
    driver.switch_to.window(driver.window_handles[-1])


# =========================
#   KAUPPALEHTI (kerää YT)
# =========================
def focus_kauppalehti_tab(driver) -> bool:
    for handle in driver.window_handles:
        try:
            driver.switch_to.window(handle)
            url = (driver.current_url or "")
            if KAUPPALEHTI_MATCH in url:
                return True
        except Exception:
            continue
    return False


def page_looks_like_protestilista(driver) -> bool:
    try:
        rows = driver.find_elements(By.XPATH, "//table//tbody//tr")
        if rows and len(rows) >= 3:
            return True
    except Exception:
        pass
    try:
        for b in driver.find_elements(By.XPATH, "//button|//*[@role='button']"):
            if (b.text or "").strip().lower() == "näytä lisää":
                return True
    except Exception:
        pass
    return False


def page_looks_like_login_or_paywall(driver) -> bool:
    try:
        text = (driver.find_element(By.TAG_NAME, "body").text or "").lower()
        bad_words = [
            "kirjaudu", "tilaa", "tilaajille", "osta", "vahvista henkilöllisyytesi",
            "sign in", "subscribe", "login", "digitilaus"
        ]
        return any(w in text for w in bad_words)
    except Exception:
        return False


def ensure_protestilista_open_and_ready(driver, status_cb, log_cb, max_wait_seconds=900, stop_flag=None) -> bool:
    if focus_kauppalehti_tab(driver):
        status_cb("Löytyi protestilista-tab.")
    else:
        status_cb("Protestilista-tab ei löytynyt -> avaan protestilistan uuteen tabiin…")
        log_cb("AUTOFIX: opening protestilista in new tab")
        open_new_tab(driver, KAUPPALEHTI_URL)
        try:
            WebDriverWait(driver, 20).until(EC.presence_of_element_located((By.TAG_NAME, "body")))
        except Exception:
            pass
        try_accept_cookies(driver)

    start = time.time()
    warned = False
    while True:
        if stop_flag and stop_flag.is_set():
            status_cb("Pysäytetty.")
            return False

        try:
            try_accept_cookies(driver)
        except Exception:
            pass

        if page_looks_like_protestilista(driver):
            status_cb("Protestilista valmis (taulukko näkyy).")
            return True

        if page_looks_like_login_or_paywall(driver) and not warned:
            warned = True
            status_cb("Kauppalehti vaatii kirjautumisen/tilaajanäkymän. Kirjaudu nyt auki olevaan Chrome-bottiin.")
            log_cb("AUTOFIX: waiting for user to login / unlock paywall…")
            try:
                messagebox.showinfo(
                    "Kirjaudu Kauppalehteen",
                    "Botti avasi protestilistan.\n\n"
                    "Kirjaudu nyt Kauppalehteen AUKI OLEVAAN Chrome-bottiin (9222).\n"
                    "Kun protestilista näkyy (taulukko + Näytä lisää), botti jatkaa automaattisesti."
                )
            except Exception:
                pass

        if time.time() - start > max_wait_seconds:
            status_cb("Aikakatkaisu: protestilista ei tullut näkyviin. Tarkista kirjautuminen.")
            log_cb("ERROR: timeout waiting protestilista")
            return False

        time.sleep(2)


def click_nayta_lisaa(driver) -> bool:
    for b in driver.find_elements(By.XPATH, "//button|//*[@role='button']"):
        try:
            if not b.is_displayed() or not b.is_enabled():
                continue
            if (b.text or "").strip().lower() == "näytä lisää":
                safe_click(driver, b)
                return True
        except Exception:
            continue
    return False


def get_company_hrefs_from_visible_rows(driver):
    hrefs = []
    rows = driver.find_elements(By.XPATH, "//table//tbody//tr")
    for r in rows:
        try:
            if not r.is_displayed():
                continue
            links = r.find_elements(By.XPATH, ".//td[1]//a[contains(@href,'/yritykset/') and normalize-space(.)!='']")
            for a in links:
                href = (a.get_attribute("href") or "").strip()
                if href and "/yritykset/" in href:
                    hrefs.append(href)
        except Exception:
            continue
    out = []
    seen = set()
    for h in hrefs:
        if h not in seen:
            seen.add(h)
            out.append(h)
    return out


def extract_yt_from_company_page_in_new_tab(driver, href: str, stop_flag):
    if stop_flag.is_set():
        return ""
    parent = driver.current_window_handle
    open_new_tab(driver, href)
    yt = ""
    try:
        t0 = time.time()
        while time.time() - t0 < KL_COMPANY_PAGE_TIMEOUT:
            if stop_flag.is_set():
                return ""
            try:
                if driver.find_elements(By.TAG_NAME, "body"):
                    break
            except Exception:
                pass
            time.sleep(0.1)

        try_accept_cookies(driver)
        time.sleep(KL_AFTER_OPEN_SLEEP)

        try:
            body = driver.find_element(By.TAG_NAME, "body").text or ""
        except Exception:
            body = ""
        yt = extract_yt_from_text_anywhere(body)
    finally:
        try:
            driver.close()
        except Exception:
            pass
        try:
            driver.switch_to.window(parent)
        except Exception:
            pass
    return yt


def collect_yts_from_kauppalehti(driver, status_cb, log_cb, stop_flag):
    try:
        WebDriverWait(driver, 25).until(EC.presence_of_element_located((By.TAG_NAME, "body")))
    except Exception:
        pass
    try_accept_cookies(driver)

    if not ensure_protestilista_open_and_ready(driver, status_cb, log_cb, stop_flag=stop_flag):
        return []

    collected = set()
    seen_hrefs = set()

    while True:
        if stop_flag.is_set():
            status_cb("Pysäytetty.")
            break

        hrefs = get_company_hrefs_from_visible_rows(driver)
        if not hrefs:
            status_cb("Kauppalehti: en löydä yrityslinkkejä. Käytä Clipboard-moodia (Ctrl+A Ctrl+C -> Ctrl+V).")
            log_cb("ERROR: no company hrefs found")
            break

        new_hrefs = [h for h in hrefs if h not in seen_hrefs]
        status_cb(f"Kauppalehti: linkkejä {len(hrefs)} | uudet {len(new_hrefs)} | Y-tunnuksia {len(collected)}")

        got = 0
        for href in new_hrefs:
            if stop_flag.is_set():
                break
            seen_hrefs.add(href)
            yt = extract_yt_from_company_page_in_new_tab(driver, href, stop_flag)
            if yt and yt not in collected:
                collected.add(yt)
                got += 1
                log_cb(f"+ {yt} (yht {len(collected)})")

        if stop_flag.is_set():
            status_cb("Pysäytetty.")
            break

        if click_nayta_lisaa(driver):
            status_cb("Kauppalehti: Näytä lisää…")
            time.sleep(KL_LOAD_MORE_WAIT)
            try:
                driver.execute_script("window.scrollTo(0, document.body.scrollHeight);")
            except Exception:
                pass
            time.sleep(0.25)
            continue

        if got == 0 and not new_hrefs:
            status_cb("Kauppalehti: valmis.")
            break

    return sorted(collected)


# =========================
#   YTJ (email)
# =========================
def click_all_nayta_ytj(driver):
    for _ in range(YTJ_NAYTA_PASSES):
        clicked = False
        for b in driver.find_elements(By.TAG_NAME, "button"):
            try:
                if (b.text or "").strip().lower() == "näytä" and b.is_displayed() and b.is_enabled():
                    safe_click(driver, b)
                    clicked = True
                    time.sleep(0.08)
            except Exception:
                continue
        for a in driver.find_elements(By.TAG_NAME, "a"):
            try:
                if (a.text or "").strip().lower() == "näytä" and a.is_displayed():
                    safe_click(driver, a)
                    clicked = True
                    time.sleep(0.08)
            except Exception:
                continue
        if not clicked:
            break


def wait_ytj_loaded(driver):
    wait = WebDriverWait(driver, YTJ_PAGE_LOAD_TIMEOUT)
    wait.until(EC.presence_of_element_located((By.TAG_NAME, "body")))
    try:
        wait.until(EC.presence_of_element_located(
            (By.XPATH, "//*[contains(normalize-space(.), 'Y-tunnus') or contains(normalize-space(.), 'Toiminimi') or contains(normalize-space(.), 'Sähköposti')]")
        ))
    except Exception:
        pass


def extract_email_from_ytj(driver):
    try:
        for a in driver.find_elements(By.TAG_NAME, "a"):
            href = (a.get_attribute("href") or "")
            if href.lower().startswith("mailto:"):
                return href.split(":", 1)[1].strip()
    except Exception:
        pass

    try:
        cells = driver.find_elements(By.XPATH, "//tr//*[self::td or self::th][contains(normalize-space(.), 'Sähköposti')]")
        for c in cells:
            try:
                tr = c.find_element(By.XPATH, "ancestor::tr[1]")
                email = pick_email_from_text(tr.text or "")
                if email:
                    return email
            except Exception:
                continue
    except Exception:
        pass

    try:
        return pick_email_from_text(driver.find_element(By.TAG_NAME, "body").text or "")
    except Exception:
        return ""


def fetch_emails_from_ytj(driver, yt_list, status_cb, progress_cb, log_cb, stop_flag):
    emails = []
    seen = set()
    new_since_partial = 0
    progress_cb(0, max(1, len(yt_list)))

    for i, yt in enumerate(yt_list, start=1):
        if stop_flag.is_set():
            status_cb("Pysäytetty.")
            break

        status_cb(f"YTJ email: {i}/{len(yt_list)} {yt}")
        progress_cb(i - 1, len(yt_list))

        try:
            driver.get(YTJ_COMPANY_URL.format(yt))
        except TimeoutException:
            pass

        if stop_flag.is_set():
            status_cb("Pysäytetty.")
            break

        try:
            wait_ytj_loaded(driver)
        except Exception:
            pass

        try_accept_cookies(driver)
        click_all_nayta_ytj(driver)

        email = ""
        for _ in range(YTJ_RETRY_READS):
            if stop_flag.is_set():
                break
            email = extract_email_from_ytj(driver)
            if email:
                break
            time.sleep(YTJ_RETRY_SLEEP)

        if email:
            k = email.lower()
            if k not in seen:
                seen.add(k)
                emails.append(email)
                append_email_tmp(email)
                log_cb(email)
                new_since_partial += 1

                if new_since_partial >= PARTIAL_SAVE_EVERY_NEW_EMAILS:
                    try:
                        save_word_plain_lines_to_path(emails, PARTIAL_DOCX_PATH)
                        log_cb(f"(partial) Tallennettu: {PARTIAL_DOCX_PATH}")
                    except Exception:
                        pass
                    new_since_partial = 0

        time.sleep(YTJ_PER_COMPANY_SLEEP)

    progress_cb(len(yt_list), max(1, len(yt_list)))
    try:
        if emails:
            save_word_plain_lines_to_path(emails, PARTIAL_DOCX_PATH)
    except Exception:
        pass
    return emails


# =========================
#   YTJ nimi -> YT
# =========================
def ytj_open_search_home(driver):
    driver.get("https://tietopalvelu.ytj.fi/")
    try:
        wait_ytj_loaded(driver)
    except Exception:
        pass
    try_accept_cookies(driver)


def find_ytj_search_input(driver):
    xpaths = [
        "//input[@type='search']",
        "//input[contains(translate(@placeholder,'HAE','hae'),'hae')]",
        "//input[contains(translate(@aria-label,'HAE','hae'),'hae')]",
        "//input[contains(translate(@name,'HAE','hae'),'hae')]",
        "//form//input",
    ]
    cands = []
    for xp in xpaths:
        try:
            cands.extend(driver.find_elements(By.XPATH, xp))
        except Exception:
            pass
    try:
        cands.extend(driver.find_elements(By.XPATH, "//input"))
    except Exception:
        pass

    for inp in cands:
        try:
            if not inp.is_displayed() or not inp.is_enabled():
                continue
            t = (inp.get_attribute("type") or "").lower()
            if t in ("hidden", "password", "checkbox", "radio", "submit", "button"):
                continue
            ph = (inp.get_attribute("placeholder") or "").lower()
            al = (inp.get_attribute("aria-label") or "").lower()
            nm = (inp.get_attribute("name") or "").lower()
            if "hae" in ph or "hae" in al or "hae" in nm or t == "search":
                return inp
        except Exception:
            continue

    for inp in cands:
        try:
            if inp.is_displayed() and inp.is_enabled():
                return inp
        except Exception:
            continue

    return None


def score_result(name_query: str, card_text: str) -> float:
    txt = (card_text or "").strip()
    m = SequenceMatcher(None, (name_query or "").lower(), txt.lower()).ratio()
    score = m * 100.0
    if extract_yt_from_text_anywhere(txt):
        score += 20.0
    return score


def ytj_find_company_and_open_best(driver, name: str, stop_flag):
    ytj_open_search_home(driver)
    if stop_flag.is_set():
        return False

    inp = find_ytj_search_input(driver)
    if not inp:
        return False

    try:
        try:
            inp.clear()
        except Exception:
            pass
        inp.send_keys(name)
        inp.send_keys(u"\ue007")  # ENTER
    except Exception:
        return False

    best_link = None
    best_score = -1.0

    t0 = time.time()
    while time.time() - t0 < CLIP_YTJ_SEARCH_TIMEOUT:
        if stop_flag.is_set():
            return False

        try:
            try_accept_cookies(driver)
        except Exception:
            pass

        candidate_links = []
        for xp in ("//a[contains(@href,'/yritys/')]", "//a[contains(@href,'yritys')]"):
            try:
                candidate_links.extend(driver.find_elements(By.XPATH, xp))
            except Exception:
                pass

        checked = 0
        for a in candidate_links:
            if checked >= 30:
                break
            try:
                if not a.is_displayed():
                    continue
                href = (a.get_attribute("href") or "")
                if not href or "tietopalvelu.ytj.fi" not in href:
                    continue

                try:
                    card = a.find_element(By.XPATH, "ancestor::*[self::li or self::div or self::article][1]")
                    card_text = (card.text or "")
                except Exception:
                    card_text = (a.text or "")

                s = score_result(name, card_text)
                checked += 1
                if s > best_score:
                    best_score = s
                    best_link = a
            except Exception:
                continue

        if best_link:
            break
        time.sleep(0.2)

    if not best_link:
        return False

    try:
        safe_click(driver, best_link)
        try:
            wait_ytj_loaded(driver)
        except Exception:
            pass
        return True
    except Exception:
        return False


def ytj_name_to_yt(driver, name: str, stop_flag) -> str:
    ok = ytj_find_company_and_open_best(driver, name, stop_flag)
    if not ok:
        return ""
    try_accept_cookies(driver)
    try:
        body = driver.find_element(By.TAG_NAME, "body").text or ""
    except Exception:
        body = ""
    return extract_yt_from_text_anywhere(body)


# =========================
#   CHROME BOT LAUNCHER (9222)
# =========================
def build_chrome_bot_args():
    base = get_exe_dir()
    profile_dir = os.path.join(base, "chrome_bot_profile")
    os.makedirs(profile_dir, exist_ok=True)

    chrome_paths = [
        r"C:\Program Files\Google\Chrome\Application\chrome.exe",
        r"C:\Program Files (x86)\Google\Chrome\Application\chrome.exe",
    ]
    chrome_exe = None
    for p in chrome_paths:
        if os.path.exists(p):
            chrome_exe = p
            break
    if chrome_exe is None:
        chrome_exe = "chrome"

    return [
        chrome_exe,
        "--remote-debugging-port=9222",
        f"--user-data-dir={profile_dir}",
    ]


def launch_chrome_bot():
    try:
        args = build_chrome_bot_args()
        subprocess.Popen(args, stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)
        return True
    except Exception:
        return False


# =========================
#   UI
# =========================
class ScrollableFrame(ttk.Frame):
    def __init__(self, container, max_width=980, *args, **kwargs):
        super().__init__(container, *args, **kwargs)
        self.max_width = max_width

        self.canvas = tk.Canvas(self, borderwidth=0, highlightthickness=0)
        scrollbar = ttk.Scrollbar(self, orient="vertical", command=self.canvas.yview)

        self.scrollable_frame = ttk.Frame(self.canvas)
        self._win = self.canvas.create_window((0, 0), window=self.scrollable_frame, anchor="n")

        self.scrollable_frame.bind(
            "<Configure>",
            lambda e: self.canvas.configure(scrollregion=self.canvas.bbox("all"))
        )
        self.canvas.configure(yscrollcommand=scrollbar.set)

        self.canvas.pack(side="left", fill="both", expand=True)
        scrollbar.pack(side="right", fill="y")

        def _on_mousewheel(event):
            self.canvas.yview_scroll(int(-1 * (event.delta / 120)), "units")
        self.canvas.bind_all("<MouseWheel>", _on_mousewheel)

        self.canvas.bind("<Configure>", self._on_canvas_configure)

    def _on_canvas_configure(self, event):
        new_w = min(self.max_width, max(320, event.width))
        self.canvas.itemconfigure(self._win, width=new_w)
        x = event.width // 2
        self.canvas.coords(self._win, (x, 0))


BaseTk = TkinterDnD.Tk if HAS_DND else tk.Tk


class App(BaseTk):
    def __init__(self):
        super().__init__()
        reset_log()
        self.stop_flag = threading.Event()

        self.BG = "#000000"
        self.FG = "#FFFFFF"
        self.RED = "#B00020"
        self.RED_HOVER = "#E0002A"
        self.PANEL = "#0B0B0B"
        self.BORDER = "#222222"
        self.MUTED = "#CFCFCF"

        self.configure(bg=self.BG)
        self.title("ProtestiBotti")
        self.geometry("980x900")
        self.bind_all("<Escape>", lambda e: self.request_stop())

        outer = ScrollableFrame(self, max_width=980)
        outer.pack(fill="both", expand=True)
        try:
            outer.canvas.configure(bg=self.BG)
        except Exception:
            pass

        root = outer.scrollable_frame
        self.main = tk.Frame(root, bg=self.BG)
        self.main.pack(fill="both", expand=True)

        self._img_refs = {}

        # Corner images
        self.corner_tl = tk.Label(self.main, bg=self.BG)
        self.corner_tr = tk.Label(self.main, bg=self.BG)
        self.corner_bl = tk.Label(self.main, bg=self.BG)
        self.corner_br = tk.Label(self.main, bg=self.BG)

        self.corner_tl.place(x=10, y=10, anchor="nw")
        self.corner_tr.place(relx=1.0, x=-10, y=10, anchor="ne")
        self.corner_bl.place(x=10, rely=1.0, y=-10, anchor="sw")
        self.corner_br.place(relx=1.0, x=-10, rely=1.0, y=-10, anchor="se")

        self._set_image_glow(self.corner_tl, "h1.jpg", (160, 95))
        self._set_image_glow(self.corner_tr, "h.jpg", (160, 95))
        self._set_image_glow(self.corner_bl, "uo.png", (220, 60))
        self._set_image_glow(self.corner_br, "vero.png", (220, 60))

        title_wrap = tk.Frame(self.main, bg=self.BG)
        title_wrap.pack(pady=(26, 6))
        tk.Label(title_wrap, text="ProtestiBotti", bg=self.BG, fg=self.FG, font=("Arial", 24, "bold")).pack()
        tk.Label(title_wrap, text="PDF / Kauppalehti / Clipboard → YTJ → sähköpostit", bg=self.BG, fg=self.MUTED, font=("Arial", 10)).pack(pady=(4, 0))

        logo_strip = tk.Frame(self.main, bg=self.BG)
        logo_strip.pack(pady=(6, 10))
        self.logo_mid = tk.Label(logo_strip, bg=self.BG)
        self.logo_mid.pack()
        self._set_image_glow(self.logo_mid, "polis.jpg", (240, 72), glow_alpha=150)

        btn_row = tk.Frame(self.main, bg=self.BG)
        btn_row.pack(pady=8)
        self._mk_btn(btn_row, "Avaa Chrome-botti (9222)", self.open_chrome_bot).grid(row=0, column=0, padx=6, pady=6)
        self._mk_btn(btn_row, "Kauppalehti → YTJ", self.start_kauppalehti_mode).grid(row=0, column=1, padx=6, pady=6)
        self._mk_btn(btn_row, "PDF → YTJ", self.start_pdf_mode).grid(row=0, column=2, padx=6, pady=6)
        self._mk_btn(btn_row, "Pysäytä", self.request_stop).grid(row=0, column=3, padx=6, pady=6)

        self.status = tk.Label(self.main, text="Valmiina.", bg=self.BG, fg=self.FG, font=("Arial", 11))
        self.status.pack(pady=6)

        self.progress = ttk.Progressbar(self.main, orient="horizontal", mode="determinate", length=920)
        self.progress.pack(pady=6)

        self.drop_var = tk.StringVar(value="PDF: Pudota tähän (tai paina PDF → YTJ ja valitse tiedosto)")
        drop = tk.Label(self.main, textvariable=self.drop_var, relief="groove",
                        bg=self.PANEL, fg=self.FG, bd=1, highlightthickness=1, highlightbackground=self.BORDER, height=2)
        drop.pack(fill="x", padx=14, pady=6)
        if HAS_DND:
            drop.drop_target_register(DND_FILES)
            drop.dnd_bind("<<Drop>>", self._on_drop_pdf)

        clip_wrap = tk.Frame(self.main, bg=self.BG)
        clip_wrap.pack(fill="x", padx=14, pady=(10, 2))

        self.strict_var = tk.BooleanVar(value=True)
        tk.Checkbutton(
            clip_wrap,
            text="Tiukka parsinta (vain Oy/Ab/Ky/Tmi/...)",
            variable=self.strict_var,
            bg=self.BG, fg=self.FG,
            selectcolor=self.PANEL,
            activebackground=self.BG, activeforeground=self.FG
        ).pack(side="left")

        tk.Label(clip_wrap, text="Max nimeä:", bg=self.BG, fg=self.FG).pack(side="left", padx=(18, 6))
        self.max_names_var = tk.IntVar(value=400)
        tk.Spinbox(clip_wrap, from_=10, to=5000, textvariable=self.max_names_var, width=7,
                   bg=self.PANEL, fg=self.FG, insertbackground=self.FG).pack(side="left")

        tk.Label(self.main, text="Clipboard (Ctrl+V tähän):", bg=self.BG, fg=self.FG, font=("Arial", 11, "bold")).pack(pady=(10, 4))
        self.clip_text = tk.Text(self.main, height=8, wrap="word", bg=self.PANEL, fg=self.FG, insertbackground=self.FG,
                                 highlightthickness=1, highlightbackground=self.BORDER)
        self.clip_text.pack(fill="x", padx=14)

        clip_btn_row = tk.Frame(self.main, bg=self.BG)
        clip_btn_row.pack(pady=8)
        self._mk_btn(clip_btn_row, "Clipboard → (Nimet → YT → Email)", self.start_clipboard_mode).pack()

        log_frame = tk.Frame(self.main, bg=self.BG)
        log_frame.pack(fill="both", expand=True, padx=14, pady=10)

        tk.Label(log_frame, text="Live-logi (uusimmat alimmaisena):", bg=self.BG, fg=self.FG).pack(anchor="w")

        box_wrap = tk.Frame(log_frame, bg=self.BG)
        box_wrap.pack(fill="both", expand=True)

        self.listbox = tk.Listbox(box_wrap, height=18, bg=self.PANEL, fg=self.FG,
                                  highlightthickness=1, highlightbackground=self.BORDER, selectbackground=self.RED)
        self.listbox.pack(side="left", fill="both", expand=True)

        sb = ttk.Scrollbar(box_wrap, orient="vertical", command=self.listbox.yview)
        sb.pack(side="right", fill="y")
        self.listbox.configure(yscrollcommand=sb.set)

        tk.Label(self.main, text=f"Tallennus: {OUT_DIR}", bg=self.BG, fg=self.MUTED, justify="center").pack(pady=(0, 14))

        if not HAS_PIL:
            self.ui_log("HUOM: Pillow puuttuu -> .jpg ja glow ei toimi. Asenna: pip install pillow")

    def _mk_btn(self, parent, text, cmd):
        b = tk.Button(
            parent,
            text=text,
            command=cmd,
            bg=self.RED,
            fg=self.FG,
            activebackground=self.RED_HOVER,
            activeforeground=self.FG,
            relief="flat",
            padx=14,
            pady=10,
            font=("Arial", 11, "bold"),
        )
        b.bind("<Enter>", lambda e: b.configure(bg=self.RED_HOVER))
        b.bind("<Leave>", lambda e: b.configure(bg=self.RED))
        return b

    def _set_image_glow(self, label: tk.Label, filename: str, size_hw, glow_alpha=180):
        path = os.path.join(get_exe_dir(), filename)
        if not os.path.exists(path):
            self.ui_log(f"HUOM: kuva puuttuu: {filename} (laita samaan kansioon kuin botti)")
            label.configure(text=f"[{filename} puuttuu]", fg="#888888", bg=self.BG)
            return

        w, h = size_hw
        try:
            if HAS_PIL:
                img = Image.open(path)
                if img.mode not in ("RGB", "RGBA"):
                    img = img.convert("RGBA")
                img = ImageOps.contain(img, (w, h))
                if img.mode != "RGBA":
                    img = img.convert("RGBA")

                # if jpg, create alpha to avoid full rectangle glow
                if filename.lower().endswith((".jpg", ".jpeg")):
                    gray = img.convert("L")
                    mask = gray.point(lambda p: 255 if p < 245 else 0)
                    img.putalpha(mask)

                alpha = img.split()[-1]
                glow = Image.new("RGBA", img.size, (255, 0, 40, glow_alpha))
                glow.putalpha(alpha)

                pad = 16
                glow_pad = Image.new("RGBA", (glow.size[0] + pad * 2, glow.size[1] + pad * 2), (0, 0, 0, 0))
                img_pad = Image.new("RGBA", (img.size[0] + pad * 2, img.size[1] + pad * 2), (0, 0, 0, 0))
                glow_pad.paste(glow, (pad, pad))
                img_pad.paste(img, (pad, pad))
                glow_blur = glow_pad.filter(ImageFilter.GaussianBlur(radius=10))
                out = Image.alpha_composite(glow_blur, img_pad)

                photo = ImageTk.PhotoImage(out)
                label.configure(image=photo)
                self._img_refs[filename] = photo
                return

            if filename.lower().endswith(".png"):
                photo = tk.PhotoImage(file=path)
                label.configure(image=photo)
                self._img_refs[filename] = photo
                return

        except Exception as e:
            self.ui_log(f"KUVA ERROR {filename}: {e}")

        label.configure(text=f"[ei voi näyttää: {filename}]", fg="#888888", bg=self.BG)

    # ---------- UI helpers ----------
    def request_stop(self):
        self.stop_flag.set()
        self.ui_log("STOP: käyttäjä pyysi pysäytystä.")
        self.status.config(text="Pysäytetään…")

    def clear_stop(self):
        self.stop_flag.clear()

    def ui_log(self, msg):
        line = log_to_file(msg)
        try:
            self.listbox.insert(tk.END, line)
            self.listbox.yview_moveto(1.0)
            self.update_idletasks()
        except Exception:
            pass

    def set_status(self, s):
        self.status.config(text=s)
        self.update_idletasks()
        self.ui_log(s)

    def set_progress(self, value, maximum):
        self.progress["maximum"] = maximum
        self.progress["value"] = value
        self.update_idletasks()

    # ---------- actions ----------
    def open_chrome_bot(self):
        ok = launch_chrome_bot()
        if ok:
            self.set_status("Chrome-botti avattu (9222). Kirjaudu Kauppalehteen siinä ikkunassa.")
        else:
            self.set_status("Chrome-botin avaus epäonnistui.")
            messagebox.showerror("Virhe", "Chrome-botin avaus epäonnistui. Tarkista Chrome-asennus.")

    def _on_drop_pdf(self, event):
        path = (event.data or "").strip()
        if path.startswith("{") and path.endswith("}"):
            path = path[1:-1]
        if path.lower().endswith(".pdf") and os.path.exists(path):
            self.drop_var.set(f"PDF valittu: {path}")
            threading.Thread(target=self.run_pdf_mode, args=(path,), daemon=True).start()
        else:
            messagebox.showwarning("Ei PDF", "Pudotettu tiedosto ei ollut .pdf")

    # =========================
    #   KAUPPALEHTI MODE
    # =========================
    def start_kauppalehti_mode(self):
        self.clear_stop()
        threading.Thread(target=self.run_kauppalehti_mode, daemon=True).start()

    def run_kauppalehti_mode(self):
        driver = None
        try:
            self.set_status("Liitytään Chrome-bottiin (9222)…")
            driver = attach_to_existing_chrome()

            self.set_status("Kauppalehti: kerätään Y-tunnukset…")
            yt_list = collect_yts_from_kauppalehti(driver, self.set_status, self.ui_log, self.stop_flag)

            if self.stop_flag.is_set():
                self.set_status("Pysäytetty.")
                return

            if not yt_list:
                self.set_status("Ei löytynyt Y-tunnuksia. Vinkki: käytä Clipboard-moodia.")
                messagebox.showwarning("Ei löytynyt", "Y-tunnuksia ei saatu. Kokeile Clipboard-moodia (Ctrl+A Ctrl+C -> Ctrl+V).")
                return

            self.set_status("Avataan YTJ uuteen tabiin…")
            open_new_tab(driver, "about:blank")

            self.set_status("YTJ: haetaan sähköpostit…")
            emails = fetch_emails_from_ytj(driver, yt_list, self.set_status, self.set_progress, self.ui_log, self.stop_flag)

            if self.stop_flag.is_set():
                self.set_status("Pysäytetty.")
                return

            em_path = save_word_unique(emails, "sahkopostit")
            self.ui_log(f"Tallennettu: {em_path}")

            self.set_status("Valmis!")
            messagebox.showinfo("Valmis", f"Valmis!\n\nKansio:\n{OUT_DIR}\n\nSähköposteja: {len(emails)}")

        except WebDriverException as e:
            self.ui_log(f"SELENIUM VIRHE: {e}")
            self.set_status("Virhe. Katso log.txt")
            messagebox.showerror("Virhe", f"Selenium/Chrome virhe.\nKatso log.txt:\n{LOG_PATH}\n\n{e}")
        except Exception as e:
            self.ui_log(f"VIRHE: {e}")
            self.set_status("Virhe. Katso log.txt")
            messagebox.showerror("Virhe", f"Tuli virhe.\nKatso log.txt:\n{LOG_PATH}\n\n{e}")

    # =========================
    #   PDF MODE
    # =========================
    def start_pdf_mode(self):
        self.clear_stop()
        path = filedialog.askopenfilename(filetypes=[("PDF files", "*.pdf")])
        if path:
            self.drop_var.set(f"PDF valittu: {path}")
            threading.Thread(target=self.run_pdf_mode, args=(path,), daemon=True).start()

    def run_pdf_mode(self, pdf_path):
        driver = None
        try:
            self.set_status("Luetaan PDF ja kerätään Y-tunnukset…")
            yt_list = extract_ytunnukset_from_pdf(pdf_path)

            if self.stop_flag.is_set():
                self.set_status("Pysäytetty.")
                return

            if not yt_list:
                self.set_status("Ei löytynyt Y-tunnuksia PDF:stä.")
                messagebox.showwarning("Ei löytynyt", "PDF:stä ei löytynyt yhtään Y-tunnusta.")
                return

            self.set_status("Käynnistetään Chrome ja haetaan sähköpostit YTJ:stä…")
            driver = start_new_driver()

            emails = fetch_emails_from_ytj(driver, yt_list, self.set_status, self.set_progress, self.ui_log, self.stop_flag)

            if self.stop_flag.is_set():
                self.set_status("Pysäytetty.")
                return

            # FIX: always new file
            em_path = save_word_unique(emails, "sahkopostit")
            self.ui_log(f"Tallennettu: {em_path}")

            self.set_status("Valmis!")
            messagebox.showinfo("Valmis", f"Valmis!\n\nKansio:\n{OUT_DIR}\n\nSähköposteja: {len(emails)}")

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
    #   CLIPBOARD MODE
    # =========================
    def start_clipboard_mode(self):
        self.clear_stop()
        text = self.clip_text.get("1.0", tk.END).strip()
        if not text:
            messagebox.showwarning("Tyhjä", "Liitä ensin teksti (Ctrl+V) kenttään.")
            return
        threading.Thread(target=self.run_clipboard_mode, args=(text,), daemon=True).start()

    def run_clipboard_mode(self, text: str):
        driver = None
        try:
            yt_list_direct = extract_yts_from_text(text)
            if yt_list_direct:
                self.set_status(f"Clipboard: löytyi {len(yt_list_direct)} Y-tunnusta → haetaan sähköpostit…")
                driver = start_new_driver()
                emails = fetch_emails_from_ytj(driver, yt_list_direct, self.set_status, self.set_progress, self.ui_log, self.stop_flag)

                if self.stop_flag.is_set():
                    self.set_status("Pysäytetty.")
                    return

                em_path = save_word_unique(emails, "sahkopostit")
                self.ui_log(f"Tallennettu: {em_path}")
                self.set_status("Valmis!")
                messagebox.showinfo("Valmis", f"Valmis!\n\nKansio:\n{OUT_DIR}\n\nSähköposteja: {len(emails)}")
                return

            strict = bool(self.strict_var.get())
            max_names = int(self.max_names_var.get() or 400)
            names = extract_names_from_clipboard(text, strict=strict, max_names=max_names)

            if not names:
                self.set_status("Clipboard: en löytänyt yritysnimiä.")
                messagebox.showwarning("Ei löytynyt", "Tekstistä ei saatu irti yritysnimiä.\nKokeile ottaa Tiukka parsinta pois.")
                return

            self.set_status(f"Clipboard: haetaan Y-tunnukset YTJ:stä nimillä… (nimiä {len(names)})")
            driver = start_new_driver()

            yts = []
            seen_yts = set()
            self.set_progress(0, max(1, len(names)))

            for i, name in enumerate(names, start=1):
                if self.stop_flag.is_set():
                    self.set_status("Pysäytetty.")
                    return

                self.set_status(f"YTJ Y-tunnus: {i}/{len(names)}  {name}")
                self.set_progress(i - 1, len(names))

                yt = ytj_name_to_yt(driver, name, self.stop_flag)
                if yt:
                    if yt not in seen_yts:
                        seen_yts.add(yt)
                        yts.append(yt)
                        self.ui_log(f"+YT {yt}  ({len(yts)})")
                else:
                    self.ui_log(f"YT NOT FOUND: {name}")

                time.sleep(CLIP_PER_NAME_SLEEP)

            self.set_progress(len(names), max(1, len(names)))

            if not yts:
                self.set_status("Ei löytynyt yhtään Y-tunnusta nimillä.")
                messagebox.showwarning("Ei löytynyt", "YTJ-nimihaulla ei saatu yhtään Y-tunnusta.\nKokeile: Tiukka parsinta pois.")
                return

            self.set_status(f"YTJ: haetaan sähköpostit {len(yts)} Y-tunnuksella…")
            emails = fetch_emails_from_ytj(driver, yts, self.set_status, self.set_progress, self.ui_log, self.stop_flag)

            if self.stop_flag.is_set():
                self.set_status("Pysäytetty.")
                return

            em_path = save_word_unique(emails, "sahkopostit")
            self.ui_log(f"Tallennettu: {em_path}")

            self.set_status("Valmis!")
            messagebox.showinfo(
                "Valmis",
                f"Valmis!\n\nKansio:\n{OUT_DIR}\n\nNimiä: {len(names)}\nY-tunnuksia: {len(yts)}\nSähköposteja: {len(emails)}"
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
