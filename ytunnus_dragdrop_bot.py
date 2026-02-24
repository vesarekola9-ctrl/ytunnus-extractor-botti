import os
import re
import sys
import time
import threading
import subprocess
import tkinter as tk
from tkinter import messagebox, filedialog
from tkinter import ttk
from html import unescape

import PyPDF2
from docx import Document

from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import (
    StaleElementReferenceException,
    WebDriverException,
    TimeoutException,
)
from webdriver_manager.chrome import ChromeDriverManager


# =========================
#   CONFIG / REGEX
# =========================
YT_RE = re.compile(r"\b\d{7}-\d\b|\b\d{8}\b")
EMAIL_RE = re.compile(r"[A-Za-z0-9_.+-]+@[A-Za-z0-9-]+\.[A-Za-z0-9-.]+")
EMAIL_A_RE = re.compile(r"[A-Za-z0-9_.+-]+\s*\(a\)\s*[A-Za-z0-9-]+\.[A-Za-z0-9-.]+", re.I)

KAUPPALEHTI_URL = "https://www.kauppalehti.fi/yritykset/protestilista"
KAUPPALEHTI_MATCH = "kauppalehti.fi/yritykset/protestilista"

YTJ_COMPANY_URL = "https://tietopalvelu.ytj.fi/yritys/{}"
YTJ_HOME = "https://tietopalvelu.ytj.fi/"

COMPANY_SUFFIXES = [
    " oyj", " oy", " ky", " tmi", " ay", " osk", " ry", " säätiö", " oy ab",
    " ab", " hb", " kb", " ltd", " limited", " inc", " gmbh", " as"
]


# =========================
#   PATHS + LOG
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


# =========================
#   STOP / SLEEP helpers
# =========================
def safe_sleep(stop_event: threading.Event, seconds: float, step: float = 0.05):
    """Sleep joka keskeytyy STOPista."""
    end = time.time() + seconds
    while time.time() < end:
        if stop_event.is_set():
            return
        time.sleep(step)


def should_stop(stop_event: threading.Event) -> bool:
    return stop_event.is_set()


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


def safe_scroll_into_view(driver, elem):
    try:
        driver.execute_script("arguments[0].scrollIntoView({block:'center'});", elem)
    except Exception:
        pass


def safe_click(driver, elem) -> bool:
    try:
        safe_scroll_into_view(driver, elem)
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
                        found = True
                        break
            except Exception:
                continue
        if not found:
            break


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
#   SELENIUM START (NOPEA)
# =========================
def _fast_chrome_options(normal_visible=True):
    opts = webdriver.ChromeOptions()
    if normal_visible:
        opts.add_argument("--start-maximized")
    try:
        opts.page_load_strategy = "eager"
    except Exception:
        pass

    prefs = {
        "profile.managed_default_content_settings.images": 2,   # kuvat pois
        "profile.default_content_setting_values.notifications": 2,
    }
    opts.add_experimental_option("prefs", prefs)

    opts.add_argument("--disable-gpu")
    opts.add_argument("--disable-dev-shm-usage")
    opts.add_argument("--no-sandbox")
    opts.add_argument("--disable-extensions")
    opts.add_argument("--disable-background-networking")
    opts.add_argument("--disable-sync")
    opts.add_argument("--disable-default-apps")
    opts.add_argument("--disable-popup-blocking")
    opts.add_argument("--disable-features=Translate,BackForwardCache,AcceptCHFrame")
    return opts


def start_new_driver_fast():
    options = _fast_chrome_options(normal_visible=True)
    driver_path = ChromeDriverManager().install()
    driver = webdriver.Chrome(service=Service(driver_path), options=options)
    try:
        driver.set_page_load_timeout(25)
    except Exception:
        pass
    return driver


def attach_to_existing_chrome():
    options = _fast_chrome_options(normal_visible=True)
    options.add_experimental_option("debuggerAddress", "127.0.0.1:9222")
    driver_path = ChromeDriverManager().install()
    driver = webdriver.Chrome(service=Service(driver_path), options=options)
    try:
        driver.set_page_load_timeout(25)
    except Exception:
        pass
    return driver


def open_new_tab(driver, url="about:blank"):
    driver.execute_script("window.open(arguments[0], '_blank');", url)
    driver.switch_to.window(driver.window_handles[-1])


# =========================
#   KAUPPALEHTI
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
        body = (driver.find_element(By.TAG_NAME, "body").text or "")
        if "Protestilista" in body and "Näytä lisää" in body:
            return True
    except Exception:
        pass
    try:
        rows = driver.find_elements(By.XPATH, "//table//tbody//tr")
        if rows and len(rows) >= 3:
            return True
    except Exception:
        pass
    return False


def page_looks_like_login_or_paywall(driver) -> bool:
    try:
        text = (driver.find_element(By.TAG_NAME, "body").text or "").lower()
        bad_words = ["kirjaudu", "tilaa", "tilaajille", "sign in", "subscribe", "login"]
        return any(w in text for w in bad_words)
    except Exception:
        return False


def ensure_protestilista_open_and_ready(driver, stop_event, status_cb, log_cb, max_wait_seconds=900) -> bool:
    if focus_kauppalehti_tab(driver):
        status_cb("Löytyi protestilista-tab.")
    else:
        status_cb("Protestilista-tab ei löytynyt -> avaan protestilistan uuteen tabiin…")
        open_new_tab(driver, KAUPPALEHTI_URL)
        WebDriverWait(driver, 30).until(EC.presence_of_element_located((By.TAG_NAME, "body")))
        try_accept_cookies(driver)

    start = time.time()
    warned = False
    while True:
        if should_stop(stop_event):
            return False

        try_accept_cookies(driver)

        if page_looks_like_protestilista(driver):
            status_cb("Protestilista valmis.")
            return True

        if page_looks_like_login_or_paywall(driver) and not warned:
            warned = True
            status_cb("Kauppalehti vaatii kirjautumisen. Kirjaudu Chrome-bottiin (9222).")
            log_cb("ODOTAN kirjautumista…")
            try:
                messagebox.showinfo(
                    "Kirjaudu Kauppalehteen",
                    "Kirjaudu Kauppalehteen AUKI OLEVAAN Chrome-bottiin (9222).\n"
                    "Kun protestilista näkyy, botti jatkaa."
                )
            except Exception:
                pass

        if time.time() - start > max_wait_seconds:
            status_cb("Aikakatkaisu: protestilista ei tullut näkyviin.")
            return False

        safe_sleep(stop_event, 2.0)


def click_nayta_lisaa(driver) -> bool:
    try:
        driver.execute_script("window.scrollTo(0, document.body.scrollHeight);")
    except Exception:
        pass

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


def get_company_rows_table(driver):
    rows = []
    candidates = driver.find_elements(By.XPATH, "//table//tbody//tr")
    for r in candidates:
        try:
            if not r.is_displayed():
                continue
            txt = (r.text or "")
            if "Y-TUNNUS" in txt:
                continue
            links = r.find_elements(By.XPATH, ".//td[1]//a[contains(@href,'/yritykset/') and normalize-space(.)!='']")
            if not links:
                continue
            rows.append(r)
        except Exception:
            continue
    return rows


def row_fingerprint_table(row):
    try:
        name = row.find_element(By.XPATH, ".//td[1]//a[contains(@href,'/yritykset/') and normalize-space(.)!='']").text.strip()
    except Exception:
        name = ""
    try:
        loc = row.find_element(By.XPATH, ".//td[2]").text.strip()
    except Exception:
        loc = ""
    try:
        amount = row.find_element(By.XPATH, ".//td[3]").text.strip()
    except Exception:
        amount = ""
    return f"{name}|{loc}|{amount}"


def extract_detail_text_from_table_row(row) -> str:
    for k in range(1, 4):
        try:
            detail = row.find_element(By.XPATH, f"following-sibling::tr[{k}]")
            txt = (detail.text or "")
            if "Y-TUNNUS" in txt:
                return txt
            links = detail.find_elements(By.XPATH, ".//td[1]//a[contains(@href,'/yritykset/') and normalize-space(.)!='']")
            if links:
                return ""
        except Exception:
            continue
    return ""


def extract_yt_from_text(txt: str) -> str:
    if not txt:
        return ""
    found = YT_RE.findall(txt)
    for m in found:
        n = normalize_yt(m)
        if n:
            return n
    return ""


def click_summa_cell(row, driver) -> bool:
    try:
        tds = row.find_elements(By.XPATH, ".//td")
        if len(tds) >= 3:
            target = tds[2]  # SUMMA
            safe_scroll_into_view(driver, target)
            try:
                driver.execute_script("arguments[0].click();", target)
            except Exception:
                safe_click(driver, target)
            return True
    except Exception:
        return False
    return False


def ensure_on_protestilista(driver, log_cb):
    url = (driver.current_url or "")
    if KAUPPALEHTI_MATCH not in url:
        log_cb(f"RECOVER: väärä sivu ({url}) -> takaisin protestilistaan")
        driver.get(KAUPPALEHTI_URL)
        WebDriverWait(driver, 30).until(EC.presence_of_element_located((By.TAG_NAME, "body")))
        try_accept_cookies(driver)


def collect_yts_from_kauppalehti(driver, stop_event, status_cb, log_cb, locked_handle=None):
    WebDriverWait(driver, 30).until(EC.presence_of_element_located((By.TAG_NAME, "body")))
    try_accept_cookies(driver)

    if locked_handle:
        try:
            driver.switch_to.window(locked_handle)
        except Exception:
            locked_handle = None

    if not ensure_protestilista_open_and_ready(driver, stop_event, status_cb, log_cb):
        return []

    collected = set()
    processed = set()

    while True:
        if should_stop(stop_event):
            status_cb("STOP: Kauppalehti-keräys keskeytetty.")
            break

        ensure_on_protestilista(driver, log_cb)

        rows = get_company_rows_table(driver)
        if not rows:
            status_cb("KL: en löydä yritysrivejä.")
            break

        status_cb(f"KL: rivejä {len(rows)} | kerätty {len(collected)}")
        new_in_pass = 0

        for idx in range(len(rows)):
            if should_stop(stop_event):
                break

            try:
                ensure_on_protestilista(driver, log_cb)
                rows_now = get_company_rows_table(driver)
                if idx >= len(rows_now):
                    break

                row = rows_now[idx]
                fp = row_fingerprint_table(row)
                if fp in processed:
                    continue
                processed.add(fp)

                if not click_summa_cell(row, driver):
                    continue

                yt = ""
                for _ in range(22):
                    if should_stop(stop_event):
                        break
                    yt = extract_yt_from_text(extract_detail_text_from_table_row(row))
                    if yt:
                        break
                    safe_sleep(stop_event, 0.08, step=0.02)

                if yt and yt not in collected:
                    collected.add(yt)
                    new_in_pass += 1
                    log_cb(f"+ {yt}")

            except StaleElementReferenceException:
                continue
            except Exception:
                continue

        if should_stop(stop_event):
            break

        old_count = len(get_company_rows_table(driver))
        if click_nayta_lisaa(driver):
            status_cb("KL: Näytä lisää…")
            try:
                WebDriverWait(driver, 20).until(lambda d: len(get_company_rows_table(d)) > old_count)
            except Exception:
                safe_sleep(stop_event, 1.2)
            continue

        if new_in_pass == 0:
            status_cb("KL: valmis.")
            break

    return sorted(collected)


# =========================
#   YTJ (NOPEA)
# =========================
def wait_ytj_loaded_fast(driver):
    WebDriverWait(driver, 18).until(EC.presence_of_element_located((By.TAG_NAME, "body")))


def click_show_for_labels(driver, labels=("Sähköposti", "Puhelin", "Puhelinnumero")):
    for lab in labels:
        try:
            containers = driver.find_elements(
                By.XPATH,
                f"//*[self::tr or self::div][.//*[contains(normalize-space(.), '{lab}')]]"
            )
            for c in containers:
                try:
                    btns = c.find_elements(By.XPATH, ".//button[normalize-space(.)='Näytä' or normalize-space(.)='näytä'] | .//a[normalize-space(.)='Näytä' or normalize-space(.)='näytä']")
                    for b in btns:
                        if b.is_displayed() and b.is_enabled():
                            safe_click(driver, b)
                except Exception:
                    continue
        except Exception:
            continue


def extract_email_from_ytj_fast(driver) -> str:
    try:
        mail = driver.find_elements(By.XPATH, "//a[starts-with(translate(@href,'MAILTO','mailto'),'mailto:')]")
        if mail:
            href = mail[0].get_attribute("href") or ""
            return href.split(":", 1)[1].strip()
    except Exception:
        pass

    try:
        rows = driver.find_elements(By.XPATH, "//tr | //div")
        for r in rows:
            t = (r.text or "")
            if "Sähköposti" in t and "@" in t:
                e = pick_email_from_text(t)
                if e:
                    return e
    except Exception:
        pass

    try:
        return pick_email_from_text(driver.find_element(By.TAG_NAME, "body").text or "")
    except Exception:
        return ""


def fetch_emails_from_ytj_by_yt_fast(driver, stop_event, yt_list, status_cb, progress_cb, log_cb):
    emails = []
    seen = set()
    total = max(1, len(yt_list))
    progress_cb(0, total)

    for i, yt in enumerate(yt_list, start=1):
        if should_stop(stop_event):
            status_cb("STOP: YTJ-haku keskeytetty.")
            break

        status_cb(f"YTJ: {i}/{len(yt_list)} {yt}")
        progress_cb(i - 1, total)

        driver.get(YTJ_COMPANY_URL.format(yt))
        wait_ytj_loaded_fast(driver)
        try_accept_cookies(driver)

        click_show_for_labels(driver)

        email = ""
        for _ in range(6):
            if should_stop(stop_event):
                break
            email = extract_email_from_ytj_fast(driver)
            if email:
                break
            safe_sleep(stop_event, 0.12, step=0.03)
            click_show_for_labels(driver)

        if email:
            k = email.lower()
            if k not in seen:
                seen.add(k)
                emails.append(email)
                log_cb(email)

        safe_sleep(stop_event, 0.02, step=0.02)

    progress_cb(min(len(yt_list), total), total)
    return emails


# =========================
#   MODE 3: TEXT/HTML -> NAMES (esikatselu + poisto)
# =========================
def strip_html(text: str) -> str:
    t = text or ""
    t = unescape(t)
    t = re.sub(r"(?is)<script.*?>.*?</script>", " ", t)
    t = re.sub(r"(?is)<style.*?>.*?</style>", " ", t)
    t = re.sub(r"(?is)<[^>]+>", "\n", t)
    t = re.sub(r"[ \t\r\f\v]+", " ", t)
    t = re.sub(r"\n{2,}", "\n", t)
    return t.strip()


def looks_like_company_name(line: str) -> bool:
    s = (line or "").strip()
    if len(s) < 3 or len(s) > 90:
        return False
    if EMAIL_RE.search(s) or YT_RE.search(s):
        return False
    if not re.search(r"[A-Za-zÅÄÖåäö]", s):
        return False

    low = s.lower()
    for suf in COMPANY_SUFFIXES:
        if low.endswith(suf.strip()):
            return True

    if re.search(r"\b(Oy|Oyj|Ky|Tmi|Ry|Osk|Ab|Ltd|GmbH)\b", s):
        return True

    return False


def extract_company_names_from_pasted(text: str):
    if not text:
        return []
    plain = strip_html(text)
    lines = [ln.strip(" -•\t") for ln in plain.split("\n")]
    names = []
    seen = set()

    for ln in lines:
        ln = re.sub(r"\s{2,}", " ", ln).strip()
        if not ln:
            continue
        ln = re.split(r"\s{2,}|\s+\d{1,3}\s*€|\s+\d{1,2}\.\d{1,2}\.\d{4}", ln)[0].strip()

        if looks_like_company_name(ln):
            key = ln.lower()
            if key not in seen:
                seen.add(key)
                names.append(ln)

    return names


def ytj_open_home_and_find_search(driver):
    driver.get(YTJ_HOME)
    WebDriverWait(driver, 18).until(EC.presence_of_element_located((By.TAG_NAME, "body")))
    try_accept_cookies(driver)

    candidates = []
    for sel in [
        "//input[@type='search']",
        "//input[@type='text']",
        "//input[contains(@aria-label,'Y-tunnus') or contains(@aria-label,'yrityksen') or contains(@aria-label,'Kirjoita')]",
        "//input[contains(@placeholder,'Y-tunnus') or contains(@placeholder,'yrityksen') or contains(@placeholder,'Kirjoita')]",
    ]:
        try:
            candidates.extend(driver.find_elements(By.XPATH, sel))
        except Exception:
            pass

    vis = []
    for c in candidates:
        try:
            if c.is_displayed() and c.is_enabled():
                vis.append(c)
        except Exception:
            pass

    hae_btn = None
    try:
        for b in driver.find_elements(By.XPATH, "//button|//*[@role='button']"):
            t = (b.text or "").strip().lower()
            if t == "hae" and b.is_displayed() and b.is_enabled():
                hae_btn = b
                break
    except Exception:
        hae_btn = None

    if not vis or not hae_btn:
        return None, None

    try:
        vis.sort(key=lambda e: e.location.get("y", 999999))
    except Exception:
        pass

    return vis[0], hae_btn


def ytj_search_company_url_by_name(driver, company_name: str) -> str:
    name = (company_name or "").strip()
    if not name:
        return ""

    input_box, hae_btn = ytj_open_home_and_find_search(driver)
    if not input_box or not hae_btn:
        return ""

    try:
        input_box.clear()
    except Exception:
        pass
    try:
        input_box.send_keys(name)
    except Exception:
        try:
            driver.execute_script("arguments[0].value = arguments[1];", input_box, name)
        except Exception:
            return ""

    safe_click(driver, hae_btn)

    try:
        WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH, "//a[contains(@href,'/yritys/')]")))
    except TimeoutException:
        return ""

    links = []
    try:
        links = driver.find_elements(By.XPATH, "//a[contains(@href,'/yritys/')]")
    except Exception:
        links = []

    results = []
    for a in links:
        try:
            if not a.is_displayed():
                continue
            href = (a.get_attribute("href") or "").strip()
            txt = (a.text or "").strip()
            if "/yritys/" in href:
                results.append((txt, href))
        except Exception:
            continue

    if not results:
        return ""

    low = name.lower()
    for txt, href in results:
        if txt and txt.strip().lower() == low:
            return href
    for txt, href in results:
        if txt and low in txt.strip().lower():
            return href

    return results[0][1]


def fetch_emails_from_ytj_by_names_fast(driver, stop_event, name_list, status_cb, progress_cb, log_cb):
    emails = []
    seen = set()
    total = max(1, len(name_list))
    progress_cb(0, total)

    for i, nm in enumerate(name_list, start=1):
        if should_stop(stop_event):
            status_cb("STOP: nimihaku keskeytetty.")
            break

        status_cb(f"YTJ nimihaku: {i}/{len(name_list)} {nm}")
        progress_cb(i - 1, total)

        url = ytj_search_company_url_by_name(driver, nm)
        if not url:
            log_cb(f"NO MATCH: {nm}")
            continue

        driver.get(url)
        wait_ytj_loaded_fast(driver)
        try_accept_cookies(driver)

        click_show_for_labels(driver)

        email = ""
        for _ in range(6):
            if should_stop(stop_event):
                break
            email = extract_email_from_ytj_fast(driver)
            if email:
                break
            safe_sleep(stop_event, 0.12, step=0.03)
            click_show_for_labels(driver)

        if email:
            k = email.lower()
            if k not in seen:
                seen.add(k)
                emails.append(email)
                log_cb(email)
        else:
            log_cb(f"NO EMAIL: {nm}")

        safe_sleep(stop_event, 0.02, step=0.02)

    progress_cb(min(len(name_list), total), total)
    return emails


# =========================
#   CHROME BOT (9222) launcher
# =========================
def launch_chrome_bot_9222():
    try:
        base = get_exe_dir()
        prof_dir = os.path.join(base, "chrome_bot_profile")
        os.makedirs(prof_dir, exist_ok=True)

        candidates = [
            os.path.join(os.environ.get("PROGRAMFILES", r"C:\Program Files"), "Google", "Chrome", "Application", "chrome.exe"),
            os.path.join(os.environ.get("PROGRAMFILES(X86)", r"C:\Program Files (x86)"), "Google", "Chrome", "Application", "chrome.exe"),
            os.path.join(os.environ.get("LOCALAPPDATA", ""), "Google", "Chrome", "Application", "chrome.exe"),
        ]
        chrome_path = next((c for c in candidates if c and os.path.exists(c)), None)
        if not chrome_path:
            raise FileNotFoundError("chrome.exe ei löytynyt.")

        args = [
            chrome_path,
            "--new-window",
            "--remote-debugging-port=9222",
            f"--user-data-dir={prof_dir}",
            KAUPPALEHTI_URL
        ]
        subprocess.Popen(args, stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)
        return True, f"Chrome-botti avattu (9222). Profiili: {prof_dir}"
    except Exception as e:
        return False, f"Chrome-botin avaus epäonnistui: {e}"


# =========================
#   GUI
# =========================
class App(tk.Tk):
    def __init__(self):
        super().__init__()
        reset_log()

        self.title("ProtestiBotti (STOP + scroll + nopea YTJ)")
        self.geometry("1120x820")

        # STOP event
        self.stop_event = threading.Event()
        self.worker_thread = None
        self.running_driver = None  # jos botin oma chrome, suljetaan STOPissa
        self.locked_handle = None
        self.mode3_names = []

        # Hotkey: Ctrl+Shift+Q
        self.bind_all("<Control-Shift-KeyPress-Q>", lambda e: self.emergency_stop())

        tk.Label(self, text="ProtestiBotti", font=("Arial", 18, "bold")).pack(pady=10)

        btn_row = tk.Frame(self)
        btn_row.pack(pady=6)

        tk.Button(btn_row, text="Avaa Chrome-botti (9222)", font=("Arial", 12), command=self.open_chrome_bot).grid(row=0, column=0, padx=6)
        tk.Button(btn_row, text="Lukitse nykyinen välilehti", font=("Arial", 12), command=self.lock_current_tab).grid(row=0, column=1, padx=6)
        tk.Button(btn_row, text="Kauppalehti → YTJ", font=("Arial", 12), command=self.start_kauppalehti_mode).grid(row=0, column=2, padx=6)
        tk.Button(btn_row, text="PDF → YTJ (nopea)", font=("Arial", 12), command=self.start_pdf_mode).grid(row=0, column=3, padx=6)

        # Iso STOP
        tk.Button(btn_row, text="🛑 STOP (Ctrl+Shift+Q)", font=("Arial", 12, "bold"),
                  fg="white", bg="#B00020", activebackground="#8C0019",
                  command=self.emergency_stop).grid(row=0, column=4, padx=10)

        self.status = tk.Label(self, text="Valmiina.", font=("Arial", 11))
        self.status.pack(pady=6)

        self.progress = ttk.Progressbar(self, orient="horizontal", mode="determinate", length=1060)
        self.progress.pack(pady=6)

        # MoodI 3
        box = tk.LabelFrame(self, text="MoodI 3: Liitä sivun teksti/HTML → poimi yritysnimet → poista väärät → hae YTJ", padx=8, pady=8)
        box.pack(fill="both", expand=False, padx=12, pady=8)

        self.paste_box = tk.Text(box, height=8, wrap="word")
        self.paste_box.pack(fill="x", expand=False)

        m3_row = tk.Frame(box)
        m3_row.pack(fill="x", pady=6)

        tk.Button(m3_row, text="Poimi nimet", command=self.mode3_extract_names).pack(side="left", padx=6)
        tk.Button(m3_row, text="Poista valitut", command=self.mode3_remove_selected).pack(side="left", padx=6)
        tk.Button(m3_row, text="Tyhjennä lista", command=self.mode3_clear_list).pack(side="left", padx=6)
        tk.Button(m3_row, text="Hae YTJ:stä nimillä", command=self.start_text_mode).pack(side="left", padx=6)

        self.names_list = tk.Listbox(box, height=8, selectmode=tk.EXTENDED)
        self.names_list.pack(fill="both", expand=True)

        tk.Label(self, text=f"Tallennus: {OUT_DIR}", wraplength=1080, justify="center").pack(pady=6)

        # Live log
        frame = tk.Frame(self)
        frame.pack(fill="both", expand=True, padx=14, pady=8)

        tk.Label(frame, text="Live-logi (rullaa hiirellä):").pack(anchor="w")
        self.listbox = tk.Listbox(frame, height=16)
        self.listbox.pack(side="left", fill="both", expand=True)

        sb = tk.Scrollbar(frame, orient="vertical", command=self.listbox.yview)
        sb.pack(side="right", fill="y")
        self.listbox.configure(yscrollcommand=sb.set)

        # Mouse wheel support (Windows)
        self._enable_mousewheel(self.listbox)
        self._enable_mousewheel(self.paste_box)
        self._enable_mousewheel(self.names_list)

    # -------- Mousewheel helper --------
    def _enable_mousewheel(self, widget):
        def _on_mousewheel(event):
            # Windows: event.delta = 120/-120
            try:
                widget.yview_scroll(int(-1 * (event.delta / 120)), "units")
            except Exception:
                pass
            return "break"
        widget.bind("<MouseWheel>", _on_mousewheel)

    # -------- UI helpers --------
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
        self.progress["maximum"] = maximum
        self.progress["value"] = value
        self.update_idletasks()

    # -------- Emergency STOP --------
    def emergency_stop(self):
        self.stop_event.set()
        self.set_status("STOP pyydetty… (keskeytetään turvallisesti)")

        # Sulje botin oma chrome jos se on käynnistetty tässä prosessissa
        if self.running_driver is not None:
            try:
                self.running_driver.quit()
            except Exception:
                pass
            self.running_driver = None

        try:
            messagebox.showinfo("STOP", "Botti keskeytetty.\n\nVoit käynnistää uudestaan napista.")
        except Exception:
            pass

    def _start_worker(self, target, args=()):
        if self.worker_thread and self.worker_thread.is_alive():
            messagebox.showwarning("Käynnissä", "Botti on jo käynnissä. Paina STOP jos haluat keskeyttää.")
            return
        self.stop_event.clear()
        self.worker_thread = threading.Thread(target=target, args=args, daemon=True)
        self.worker_thread.start()

    # -------- Chrome / lock --------
    def open_chrome_bot(self):
        ok, msg = launch_chrome_bot_9222()
        self.ui_log(msg)
        if ok:
            messagebox.showinfo("Chrome-botti", msg + "\n\nKirjaudu Kauppalehteen tässä ikkunassa.")
        else:
            messagebox.showerror("Chrome-botti", msg)

    def lock_current_tab(self):
        try:
            self.set_status("Liitytään Chrome-bottiin (9222) lukitusta varten…")
            driver = attach_to_existing_chrome()
            self.locked_handle = driver.current_window_handle
            messagebox.showinfo("Lukittu", f"Lukittu välilehti:\n{driver.title}\n{driver.current_url}")
        except Exception as e:
            self.ui_log(f"VIRHE lukituksessa: {e}")
            messagebox.showerror("Virhe", f"Lukitus epäonnistui:\n{e}")

    # -------- Mode 1: KL -> YTJ --------
    def start_kauppalehti_mode(self):
        self._start_worker(self.run_kauppalehti_mode)

    def run_kauppalehti_mode(self):
        driver = None
        try:
            self.set_status("Liitytään Chrome-bottiin (9222)…")
            driver = attach_to_existing_chrome()

            self.set_status("Kauppalehti: kerätään Y-tunnukset…")
            yt_list = collect_yts_from_kauppalehti(driver, self.stop_event, self.set_status, self.ui_log, locked_handle=self.locked_handle)
            if should_stop(self.stop_event):
                return

            if not yt_list:
                messagebox.showwarning("Ei löytynyt", "Y-tunnuksia ei saatu Kauppalehdestä. Katso log.txt.")
                return

            save_word_plain_lines(yt_list, "ytunnukset.docx")

            self.set_status("Avataan YTJ uuteen tabiin…")
            open_new_tab(driver, "about:blank")

            self.set_status("YTJ: haetaan sähköpostit…")
            emails = fetch_emails_from_ytj_by_yt_fast(driver, self.stop_event, yt_list, self.set_status, self.set_progress, self.ui_log)
            if should_stop(self.stop_event):
                return

            em_path = save_word_plain_lines(emails, "sahkopostit.docx")
            self.ui_log(f"Tallennettu: {em_path}")

            self.set_status("Valmis!")
            messagebox.showinfo("Valmis", f"Valmis!\n\nSähköposteja: {len(emails)}\nKansio:\n{OUT_DIR}")

        except Exception as e:
            self.ui_log(f"VIRHE: {e}")
            messagebox.showerror("Virhe", f"Tuli virhe.\nKatso log.txt:\n{LOG_PATH}\n\n{e}")

    # -------- Mode 2: PDF -> YTJ --------
    def start_pdf_mode(self):
        path = filedialog.askopenfilename(filetypes=[("PDF files", "*.pdf")])
        if path:
            self._start_worker(self.run_pdf_mode, args=(path,))

    def run_pdf_mode(self, pdf_path):
        driver = None
        try:
            self.set_status("Luetaan PDF ja kerätään Y-tunnukset…")
            yt_list = extract_ytunnukset_from_pdf(pdf_path)
            if not yt_list:
                messagebox.showwarning("Ei löytynyt", "PDF:stä ei löytynyt Y-tunnuksia.")
                return

            if should_stop(self.stop_event):
                return

            self.set_status("Käynnistetään Chrome (nopea) ja haetaan sähköpostit YTJ:stä…")
            driver = start_new_driver_fast()
            self.running_driver = driver

            emails = fetch_emails_from_ytj_by_yt_fast(driver, self.stop_event, yt_list, self.set_status, self.set_progress, self.ui_log)
            if should_stop(self.stop_event):
                return

            em_path = save_word_plain_lines(emails, "sahkopostit.docx")
            self.ui_log(f"Tallennettu: {em_path}")

            self.set_status("Valmis!")
            messagebox.showinfo("Valmis", f"Valmis!\n\nSähköposteja: {len(emails)}\nKansio:\n{OUT_DIR}")

        except Exception as e:
            self.ui_log(f"VIRHE: {e}")
            messagebox.showerror("Virhe", f"Tuli virhe.\nKatso log.txt:\n{LOG_PATH}\n\n{e}")
        finally:
            if driver:
                try:
                    driver.quit()
                except Exception:
                    pass
            self.running_driver = None

    # -------- Mode 3: Text -> Names -> YTJ --------
    def mode3_extract_names(self):
        raw = self.paste_box.get("1.0", "end").strip()
        if not raw:
            messagebox.showwarning("Tyhjä", "Liitä ensin sivun teksti/HTML.")
            return

        self.set_status("Poimitaan yritysnimiä…")
        names = extract_company_names_from_pasted(raw)

        self.mode3_names = names
        self.names_list.delete(0, tk.END)
        for n in names:
            self.names_list.insert(tk.END, n)

        self.set_status(f"Poimittu yritysnimiä: {len(names)}")
        if not names:
            messagebox.showwarning("Ei löytynyt", "En löytänyt yritysnimiä (tarvitsen esim. Oy/Oyj/Ky/Tmi/Ry/Ab/Ltd).")

    def mode3_remove_selected(self):
        sel = list(self.names_list.curselection())
        if not sel:
            return
        sel.reverse()
        for idx in sel:
            self.names_list.delete(idx)
        self.mode3_names = list(self.names_list.get(0, tk.END))

    def mode3_clear_list(self):
        self.names_list.delete(0, tk.END)
        self.mode3_names = []

    def start_text_mode(self):
        self._start_worker(self.run_text_mode)

    def run_text_mode(self):
        driver = None
        try:
            names = list(self.names_list.get(0, tk.END))
            names = [n.strip() for n in names if n.strip()]
            if not names:
                messagebox.showwarning("Ei nimiä", "Poimi ensin nimet ja poista väärät.")
                return

            if should_stop(self.stop_event):
                return

            self.set_status("Käynnistetään Chrome (nopea) ja haetaan sähköpostit YTJ:stä nimillä…")
            driver = start_new_driver_fast()
            self.running_driver = driver

            emails = fetch_emails_from_ytj_by_names_fast(driver, self.stop_event, names, self.set_status, self.set_progress, self.ui_log)
            if should_stop(self.stop_event):
                return

            em_path = save_word_plain_lines(emails, "sahkopostit.docx")
            self.ui_log(f"Tallennettu: {em_path}")

            self.set_status("Valmis!")
            messagebox.showinfo("Valmis", f"Valmis!\n\nSähköposteja: {len(emails)}\nKansio:\n{OUT_DIR}")

        except Exception as e:
            self.ui_log(f"VIRHE: {e}")
            messagebox.showerror("Virhe", f"Tuli virhe.\nKatso log.txt:\n{LOG_PATH}\n\n{e}")
        finally:
            if driver:
                try:
                    driver.quit()
                except Exception:
                    pass
            self.running_driver = None


if __name__ == "__main__":
    App().mainloop()
