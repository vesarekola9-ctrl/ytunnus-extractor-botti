import os
import re
import sys
import time
import threading
import tkinter as tk
from tkinter import messagebox
from tkinter import ttk
from dataclasses import dataclass
from urllib.parse import urlparse

from docx import Document

from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import WebDriverException
from webdriver_manager.chrome import ChromeDriverManager


# =========================
#   CONFIG / REGEX
# =========================
YT_RE = re.compile(r"\b\d{7}-\d\b|\b\d{8}\b")
EMAIL_RE = re.compile(r"[A-Za-z0-9_.+-]+@[A-Za-z0-9-]+\.[A-Za-z0-9-.]+")
EMAIL_A_RE = re.compile(r"[A-Za-z0-9_.+-]+\s*\(a\)\s*[A-Za-z0-9-]+\.[A-Za-z0-9-.]+", re.I)

KAUPPALEHTI_URL = "https://www.kauppalehti.fi/yritykset/protestilista"
YTJ_HOME = "https://tietopalvelu.ytj.fi/"
ALLOWED_KL_HOST = "www.kauppalehti.fi"


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


def dump_debug(driver, tag="debug"):
    html_path = os.path.join(OUT_DIR, f"{tag}.html")
    png_path = os.path.join(OUT_DIR, f"{tag}.png")
    try:
        with open(html_path, "w", encoding="utf-8") as f:
            f.write(driver.page_source or "")
    except Exception:
        html_path = None
    try:
        driver.save_screenshot(png_path)
    except Exception:
        png_path = None
    return html_path, png_path


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


def normalize_name(s: str) -> str:
    s = (s or "").strip().casefold()
    s = re.sub(r"\s+", " ", s)
    s = s.replace("asunto oy", "as oy").replace("asunto-osakeyhtiö", "as oy")
    return s


def token_set(s: str):
    s = normalize_name(s)
    s = re.sub(r"[^0-9a-zåäö\s-]", " ", s)
    parts = [p for p in re.split(r"\s+", s) if p]
    return set(parts)


def score_match(query_name: str, query_location: str, candidate_text: str) -> int:
    qn = normalize_name(query_name)
    qtoks = token_set(query_name)
    cl = normalize_name(candidate_text)

    score = 0
    if qn and qn == cl:
        score += 120
    if qn and qn in cl:
        score += 60

    ctoks = token_set(candidate_text)
    score += len(qtoks & ctoks) * 8

    loc = normalize_name(query_location)
    if loc and loc in cl:
        score += 35

    for suf in ["oy", "ab", "ky", "oyj", "tmi", "ry"]:
        if suf in qn and suf in cl:
            score += 10
    return score


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


def save_word_table(rows, headers, filename):
    path = os.path.join(OUT_DIR, filename)
    doc = Document()
    table = doc.add_table(rows=1, cols=len(headers))
    hdr = table.rows[0].cells
    for i, h in enumerate(headers):
        hdr[i].text = str(h)
    for r in rows:
        cells = table.add_row().cells
        for i, v in enumerate(r):
            cells[i].text = "" if v is None else str(v)
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
        time.sleep(0.05)
        try:
            elem.click()
        except Exception:
            driver.execute_script("arguments[0].click();", elem)
        return True
    except Exception:
        return False


# =========================
#   CLICK SAFETY (PREVENT EXTERNAL NAV)
# =========================
def current_host(driver) -> str:
    try:
        return urlparse(driver.current_url).netloc.lower()
    except Exception:
        return ""


def is_external_href(href: str) -> bool:
    if not href:
        return False
    try:
        host = urlparse(href).netloc.lower()
        if not host:  # relative
            return False
        return host != ALLOWED_KL_HOST
    except Exception:
        return False


def safe_click_kl_only(driver, elem) -> bool:
    """
    Klikkaa vain jos:
    - elementti ei ole ulkoinen <a href="https://..."> tai
    - klikkauksen jälkeen pysytään kauppalehti.fi domainissa
    Jos mennään ulos, palautetaan takaisin ja palautetaan False.
    """
    before = driver.current_url

    try:
        tag = (elem.tag_name or "").lower()
        if tag == "a":
            href = (elem.get_attribute("href") or "").strip()
            if is_external_href(href):
                return False
    except Exception:
        pass

    ok = safe_click(driver, elem)
    if not ok:
        return False

    time.sleep(0.15)

    host = current_host(driver)
    if host and host != ALLOWED_KL_HOST:
        # palautetaan takaisin
        try:
            driver.get(before)
        except Exception:
            try:
                driver.back()
            except Exception:
                pass
        return False

    return True


# =========================
#   COOKIES (SAFE)
# =========================
def try_accept_cookies(driver):
    texts = {"hyväksy", "hyväksy kaikki", "salli kaikki", "accept", "accept all", "i agree", "ok", "selvä"}

    containers = []
    for xp in [
        "//*[contains(translate(@id,'COOKIE','cookie'),'cookie') or contains(translate(@class,'COOKIE','cookie'),'cookie')]",
        "//*[contains(translate(@id,'CMP','cmp'),'cmp') or contains(translate(@class,'CMP','cmp'),'cmp')]",
        "//*[@role='dialog']",
    ]:
        try:
            containers += driver.find_elements(By.XPATH, xp)
        except Exception:
            pass

    try:
        search_roots = containers[:3] if containers else [driver.find_element(By.TAG_NAME, "body")]
    except Exception:
        return

    for root in search_roots:
        try:
            buttons = root.find_elements(By.XPATH, ".//button|.//*[@role='button']")
        except Exception:
            continue

        for b in buttons:
            try:
                if not b.is_displayed() or not b.is_enabled():
                    continue
                t = (b.text or "").strip().lower()
                if t in texts:
                    # tärkeä: kl-only ettei lähde mihinkään
                    if safe_click_kl_only(driver, b):
                        time.sleep(0.2)
                        return
            except Exception:
                continue


# =========================
#   SELENIUM START (NO 9222, persistent profile)
# =========================
def get_profile_dir():
    base = get_exe_dir()
    prof = os.path.join(base, "chrome_profile")
    try:
        os.makedirs(prof, exist_ok=True)
        test = os.path.join(prof, "_w.tmp")
        with open(test, "w", encoding="utf-8") as f:
            f.write("ok")
        os.remove(test)
        return prof
    except Exception:
        home = os.path.expanduser("~")
        docs = os.path.join(home, "Documents")
        prof = os.path.join(docs, "ProtestiBotti", "chrome_profile")
        os.makedirs(prof, exist_ok=True)
        return prof


def start_persistent_driver():
    options = webdriver.ChromeOptions()
    options.add_argument("--start-maximized")
    options.add_argument("--disable-gpu")
    options.add_argument("--disable-dev-shm-usage")
    options.add_argument("--no-first-run")
    options.add_argument("--no-default-browser-check")
    options.add_experimental_option("excludeSwitches", ["enable-automation"])
    options.add_experimental_option("useAutomationExtension", False)

    profile_dir = get_profile_dir()
    options.add_argument(f"--user-data-dir={profile_dir}")

    driver_path = ChromeDriverManager().install()
    driver = webdriver.Chrome(service=Service(driver_path), options=options)

    # pieni “stealth”
    try:
        driver.execute_cdp_cmd("Page.addScriptToEvaluateOnNewDocument", {
            "source": "Object.defineProperty(navigator, 'webdriver', {get: () => undefined});"
        })
    except Exception:
        pass

    return driver


# =========================
#   TABS
# =========================
def open_new_tab(driver, url="about:blank"):
    driver.execute_script("window.open(arguments[0], '_blank');", url)
    driver.switch_to.window(driver.window_handles[-1])


def ensure_two_tabs(driver):
    # TAB1 KL
    driver.get(KAUPPALEHTI_URL)
    WebDriverWait(driver, 25).until(EC.presence_of_element_located((By.TAG_NAME, "body")))
    try_accept_cookies(driver)

    # TAB2 YTJ
    open_new_tab(driver, YTJ_HOME)
    WebDriverWait(driver, 25).until(EC.presence_of_element_located((By.TAG_NAME, "body")))
    try_accept_cookies(driver)

    # takaisin KL
    driver.switch_to.window(driver.window_handles[0])


# =========================
#   KL READY + TIMEFRAME
# =========================
def page_looks_like_login_or_paywall(driver) -> bool:
    try:
        text = (driver.find_element(By.TAG_NAME, "body").text or "").lower()
        bad_words = [
            "kirjaudu", "tilaa", "tilaajille", "vahvista henkilöll",
            "sign in", "subscribe", "login",
            "pääsy evätty", "access denied",
            "jokin meni pieleen", "something went wrong",
        ]
        return any(w in text for w in bad_words)
    except Exception:
        return False


def ensure_kl_ready(driver, status_cb, log_cb, stop_evt: threading.Event, max_wait_seconds=900):
    # Varmistus: pysytään domainissa
    if current_host(driver) != ALLOWED_KL_HOST or "protestilista" not in (driver.current_url or ""):
        driver.get(KAUPPALEHTI_URL)
        WebDriverWait(driver, 25).until(EC.presence_of_element_located((By.TAG_NAME, "body")))
        try_accept_cookies(driver)

    start = time.time()
    warned = False

    while True:
        if stop_evt.is_set():
            return False

        try_accept_cookies(driver)

        # jos karkasi ulos, palaa
        if current_host(driver) and current_host(driver) != ALLOWED_KL_HOST:
            driver.get(KAUPPALEHTI_URL)
            WebDriverWait(driver, 25).until(EC.presence_of_element_located((By.TAG_NAME, "body")))
            try_accept_cookies(driver)

        body = ""
        try:
            body = driver.find_element(By.TAG_NAME, "body").text or ""
        except Exception:
            pass

        if "Protestilista" in body and not page_looks_like_login_or_paywall(driver):
            status_cb("KL: protestilista näkyy.")
            time.sleep(0.6)
            return True

        if not warned:
            warned = True
            status_cb("Kirjaudu Kauppalehteen tässä botti-Chromessa. Kun protestilista näkyy, botti jatkaa.")
            log_cb("Waiting for user login on Kauppalehti…")
            try:
                messagebox.showinfo(
                    "Kirjaudu Kauppalehteen",
                    "Chrome avattiin botille omalla profiililla.\n\n"
                    "Kirjaudu Kauppalehteen tässä Chromessa.\n"
                    "Kun protestilista näkyy, botti jatkaa automaattisesti."
                )
            except Exception:
                pass

        if time.time() - start > max_wait_seconds:
            status_cb("Aikakatkaisu: protestilista ei tullut näkyviin.")
            html, png = dump_debug(driver, "kl_timeout")
            log_cb(f"DEBUG dump: {html} | {png}")
            return False

        time.sleep(2)


def set_kl_timeframe(driver, timeframe_text: str, log_cb=None):
    if not timeframe_text:
        return False
    tf = timeframe_text.strip()

    try:
        container = None
        labels = driver.find_elements(By.XPATH, "//*[contains(normalize-space(.), 'Aikarajaus')]")
        for lab in labels[:10]:
            try:
                if not lab.is_displayed():
                    continue
                container = lab.find_element(By.XPATH, "ancestor::*[self::div or self::section][1]")
                if container:
                    break
            except Exception:
                continue

        if not container:
            return False

        opener = None
        for xp in [".//button", ".//*[@role='button']", ".//*[@role='combobox']", ".//div[@role='combobox']"]:
            elems = container.find_elements(By.XPATH, xp)
            for e in elems:
                try:
                    if e.is_displayed() and e.is_enabled():
                        opener = e
                        break
                except Exception:
                    continue
            if opener:
                break

        if not opener:
            return False

        safe_click_kl_only(driver, opener)
        time.sleep(0.25)

        candidates = driver.find_elements(By.XPATH, "//*[self::li or self::button or @role='option' or @role='menuitem' or @role='listitem']")
        for c in candidates:
            try:
                if not c.is_displayed() or not c.is_enabled():
                    continue
                txt = (c.text or "").strip()
                if txt.lower() == tf.lower():
                    safe_click_kl_only(driver, c)
                    time.sleep(0.4)
                    return True
            except Exception:
                continue

        for c in candidates:
            try:
                if not c.is_displayed() or not c.is_enabled():
                    continue
                txt = (c.text or "").strip().lower()
                if tf.lower() in txt:
                    safe_click_kl_only(driver, c)
                    time.sleep(0.4)
                    return True
            except Exception:
                continue

        return False
    except Exception as e:
        if log_cb:
            log_cb(f"KL Aikarajaus set error: {e}")
        return False


def click_nayta_lisaa(driver) -> bool:
    """
    Klikkaa VAIN <button> jossa teksti on täsmälleen 'Näytä lisää'.
    Ei koskaan <a>-linkkejä => ei koskaan ulos hyppyä.
    """
    try:
        driver.execute_script("window.scrollTo(0, document.body.scrollHeight);")
    except Exception:
        pass
    time.sleep(0.25)

    for b in driver.find_elements(By.XPATH, "//button"):
        try:
            if not b.is_displayed() or not b.is_enabled():
                continue
            txt = (b.text or "").strip().lower()
            if txt == "näytä lisää":
                return safe_click_kl_only(driver, b)
        except Exception:
            continue

    return False


# =========================
#   HYBRID DATA
# =========================
@dataclass
class KLRow:
    company: str
    location: str
    amount: str
    date: str
    ptype: str
    source: str


def read_kl_visible_rows(driver):
    rows = []
    trs = driver.find_elements(By.XPATH, "//table//tbody//tr")
    for tr in trs:
        try:
            if not tr.is_displayed():
                continue
            txt = (tr.text or "").strip()
            if not txt:
                continue
            if "Y-tunnus" in txt or "Y-TUNNUS" in txt:
                continue

            tds = tr.find_elements(By.XPATH, ".//td")
            if len(tds) < 6:
                continue

            company = (tds[0].text or "").strip()
            location = (tds[1].text or "").strip()
            amount = (tds[2].text or "").strip()
            date = (tds[3].text or "").strip()
            ptype = (tds[4].text or "").strip()
            source = (tds[5].text or "").strip()

            if company:
                rows.append(KLRow(company, location, amount, date, ptype, source))
        except Exception:
            continue
    return rows


def collect_kl_rows_all(driver, status_cb, log_cb, stop_evt: threading.Event, max_pages=999):
    collected = []
    seen = set()
    rounds_no_new = 0
    page_clicks = 0

    while True:
        if stop_evt.is_set():
            break

        try_accept_cookies(driver)

        # jos karkasi jostain syystä, palaa protestilistaan
        if current_host(driver) != ALLOWED_KL_HOST or "protestilista" not in (driver.current_url or ""):
            log_cb(f"KL guard: palataan protestilistaan (url oli {driver.current_url})")
            driver.get(KAUPPALEHTI_URL)
            WebDriverWait(driver, 25).until(EC.presence_of_element_located((By.TAG_NAME, "body")))
            try_accept_cookies(driver)
            time.sleep(0.6)

        if page_looks_like_login_or_paywall(driver):
            status_cb("KL: näyttää blokilta/paywallilta -> debug dump.")
            html, png = dump_debug(driver, "kl_blocked")
            log_cb(f"DEBUG dump: {html} | {png}")
            time.sleep(2)
            continue

        visible = read_kl_visible_rows(driver)
        new = 0
        for r in visible:
            key = (r.company, r.location, r.amount, r.date, r.ptype, r.source)
            if key in seen:
                continue
            seen.add(key)
            collected.append(r)
            new += 1

        status_cb(f"KL: kerätty {len(collected)} (uusii {new})")

        if new == 0:
            rounds_no_new += 1
        else:
            rounds_no_new = 0

        if page_clicks >= max_pages:
            status_cb("KL: max_pages täynnä -> lopetan.")
            break

        if click_nayta_lisaa(driver):
            page_clicks += 1
            time.sleep(1.1)
            continue

        if rounds_no_new >= 2:
            break

    if not collected:
        html, png = dump_debug(driver, "kl_zero_rows")
        log_cb(f"DEBUG dump: {html} | {png}")

    return collected


# =========================
#   YTJ
# =========================
def wait_ytj_loaded(driver):
    WebDriverWait(driver, 25).until(EC.presence_of_element_located((By.TAG_NAME, "body")))
    time.sleep(0.2)


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


def extract_yt_from_ytj_page(driver) -> str:
    try:
        text = driver.find_element(By.TAG_NAME, "body").text or ""
        for m in YT_RE.findall(text):
            n = normalize_yt(m)
            if n:
                return n
    except Exception:
        pass
    return ""


def ytj_open_home(driver):
    driver.get(YTJ_HOME)
    wait_ytj_loaded(driver)
    try_accept_cookies(driver)


def ytj_search_and_pick_best(driver, company_name: str, location: str, log_cb=None):
    ytj_open_home(driver)

    search_input = None
    inputs = driver.find_elements(By.XPATH, "//input")
    for inp in inputs:
        try:
            if not inp.is_displayed() or not inp.is_enabled():
                continue
            typ = (inp.get_attribute("type") or "").lower()
            aria = (inp.get_attribute("aria-label") or "").lower()
            ph = (inp.get_attribute("placeholder") or "").lower()
            if typ in ("search", "text") and ("hae" in aria or "hae" in ph or "yritys" in aria or "yritys" in ph):
                search_input = inp
                break
        except Exception:
            continue

    if not search_input:
        for inp in inputs:
            try:
                if inp.is_displayed() and inp.is_enabled():
                    search_input = inp
                    break
            except Exception:
                continue

    if not search_input:
        if log_cb:
            log_cb("YTJ: hakukenttää ei löytynyt")
        return "", 0, ""

    try:
        search_input.click()
        time.sleep(0.05)
        search_input.clear()
    except Exception:
        pass

    try:
        search_input.send_keys(company_name)
        time.sleep(0.05)
        search_input.send_keys(Keys.ENTER)
    except Exception:
        try:
            driver.find_element(By.TAG_NAME, "body").send_keys(Keys.ENTER)
        except Exception:
            pass

    time.sleep(1.0)
    try_accept_cookies(driver)

    links = driver.find_elements(By.XPATH, "//a[contains(@href, '/yritys/')]")
    best = ("", 0, "")
    seen_urls = set()

    for a in links[:60]:
        try:
            url = (a.get_attribute("href") or "").strip()
            if not url or "/yritys/" not in url:
                continue
            if url in seen_urls:
                continue
            seen_urls.add(url)

            cand_text = ""
            try:
                container = a.find_element(By.XPATH, "ancestor::*[self::li or self::div or self::article][1]")
                cand_text = (container.text or "").strip()
            except Exception:
                cand_text = (a.text or "").strip()

            if not cand_text:
                continue

            sc = score_match(company_name, location, cand_text)
            if sc > best[1]:
                best = (url, sc, cand_text)
        except Exception:
            continue

    return best


def ytj_get_email_by_company_name(driver, company_name: str, location: str, log_cb=None):
    url, score, _txt = ytj_search_and_pick_best(driver, company_name, location, log_cb=log_cb)
    if not url:
        return {"yt": "", "email": "", "url": "", "score": 0}

    driver.get(url)
    wait_ytj_loaded(driver)
    try_accept_cookies(driver)

    # yritä avata "Näytä"
    for _ in range(2):
        clicked = False
        for b in driver.find_elements(By.XPATH, "//button|//*[@role='button']|//a"):
            try:
                t = (b.text or "").strip().lower()
                if t == "näytä" and b.is_displayed() and b.is_enabled():
                    # YTJ:ssä ei tarvita KL-only -klikkiä
                    safe_click(driver, b)
                    clicked = True
                    time.sleep(0.15)
            except Exception:
                continue
        if not clicked:
            break

    yt = extract_yt_from_ytj_page(driver)
    email = extract_email_from_ytj(driver)
    return {"yt": yt, "email": email, "url": url, "score": score}


# =========================
#   HYBRID RUN (TWO TABS)
# =========================
def run_hybrid(driver, timeframe: str, status_cb, progress_cb, log_cb, stop_evt: threading.Event):
    # Varmista 2 tabia
    status_cb("Alustus: avataan KL ja YTJ omiin tabbeihin…")
    ensure_two_tabs(driver)

    kl_handle = driver.window_handles[0]
    ytj_handle = driver.window_handles[1]

    # KL ready
    driver.switch_to.window(kl_handle)
    status_cb("KL: avataan protestilista…")
    if not ensure_kl_ready(driver, status_cb, log_cb, stop_evt):
        return [], []

    # Timeframe
    if timeframe:
        status_cb(f"KL: asetetaan aikarajaus = {timeframe} …")
        ok = set_kl_timeframe(driver, timeframe, log_cb=log_cb)
        log_cb(f"Aikarajaus set: {ok}")
        time.sleep(0.6)

    # Collect KL rows
    status_cb("KL: kerätään yritysrivit (nimi+sijainti)…")
    kl_rows = collect_kl_rows_all(driver, status_cb, log_cb, stop_evt)

    if stop_evt.is_set():
        return kl_rows, []

    if not kl_rows:
        return [], []

    kl_table = [[r.company, r.location, r.date, r.amount, r.ptype, r.source] for r in kl_rows]
    kl_path = save_word_table(
        kl_table,
        headers=["Yritys", "Sijainti", "Häiriöpäivä", "Summa", "Tyyppi", "Lähde"],
        filename="kl_rivit.docx",
    )
    log_cb(f"Tallennettu: {kl_path}")

    # YTJ loop in tab2
    results = []
    progress_cb(0, len(kl_rows))

    driver.switch_to.window(ytj_handle)
    ytj_open_home(driver)

    for i, r in enumerate(kl_rows, start=1):
        if stop_evt.is_set():
            break

        status_cb(f"YTJ: {i}/{len(kl_rows)} — {r.company} ({r.location})")
        progress_cb(i - 1, len(kl_rows))

        try:
            data = ytj_get_email_by_company_name(driver, r.company, r.location, log_cb=log_cb)
        except Exception as e:
            log_cb(f"YTJ ERROR {r.company}: {e}")
            data = {"yt": "", "email": "", "url": "", "score": 0}

        results.append([
            r.company,
            r.location,
            data.get("yt", ""),
            data.get("email", ""),
            data.get("url", ""),
            str(data.get("score", 0)),
        ])
        time.sleep(0.12)

    progress_cb(len(results), len(kl_rows))

    out_path = save_word_table(
        results,
        headers=["Yritys", "Sijainti", "Y-tunnus", "Sähköposti", "YTJ-linkki", "Score"],
        filename="tulokset.docx",
    )
    log_cb(f"Tallennettu: {out_path}")

    return kl_rows, results


# =========================
#   GUI
# =========================
class App(tk.Tk):
    def __init__(self):
        super().__init__()
        reset_log()

        self.stop_evt = threading.Event()
        self.worker = None

        self.title("ProtestiBotti (TOIMIVA: HYBRID, NO 9222 + turvalliset klikkaukset)")
        self.geometry("1040x700")

        tk.Label(self, text="ProtestiBotti (Hybrid, NO 9222)", font=("Arial", 18, "bold")).pack(pady=8)

        top = tk.Frame(self)
        top.pack(pady=6)

        tk.Label(top, text="Aikarajaus:", font=("Arial", 11)).grid(row=0, column=0, padx=6)
        self.timeframe = ttk.Combobox(top, values=["Päivä", "Viikko", "Kuukausi"], state="readonly", width=12)
        self.timeframe.set("Päivä")
        self.timeframe.grid(row=0, column=1, padx=6)

        self.btn_start = tk.Button(top, text="Käynnistä HYBRID", font=("Arial", 12, "bold"), command=self.start)
        self.btn_start.grid(row=0, column=2, padx=10)

        self.btn_stop = tk.Button(top, text="STOP", font=("Arial", 12, "bold"), command=self.stop, state="disabled")
        self.btn_stop.grid(row=0, column=3, padx=6)

        self.status = tk.Label(self, text="Valmiina.", font=("Arial", 11))
        self.status.pack(pady=6)

        self.progress = ttk.Progressbar(self, orient="horizontal", mode="determinate", length=980)
        self.progress.pack(pady=6)

        frame = tk.Frame(self)
        frame.pack(fill="both", expand=True, padx=14, pady=10)

        tk.Label(frame, text="Live-logi:").pack(anchor="w")
        self.listbox = tk.Listbox(frame, height=22)
        self.listbox.pack(side="left", fill="both", expand=True)

        sb = tk.Scrollbar(frame, orient="vertical", command=self.listbox.yview)
        sb.pack(side="right", fill="y")
        self.listbox.configure(yscrollcommand=sb.set)

        tk.Label(self, text=f"Tallennus: {OUT_DIR}", wraplength=1000, justify="center").pack(pady=6)

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

    def start(self):
        if self.worker and self.worker.is_alive():
            return
        self.stop_evt.clear()
        self.btn_start.config(state="disabled")
        self.btn_stop.config(state="normal")
        self.worker = threading.Thread(target=self.run, daemon=True)
        self.worker.start()

    def stop(self):
        self.stop_evt.set()
        self.set_status("STOP pyydetty – lopetetaan hallitusti…")
        self.btn_stop.config(state="disabled")

    def run(self):
        driver = None
        try:
            self.set_status("Käynnistetään Chrome (pysyvä profiili)…")
            driver = start_persistent_driver()

            tf = self.timeframe.get().strip()
            self.set_status(f"Aloitetaan: KL({tf}) → YTJ (nimihaku)…")
            kl_rows, results = run_hybrid(
                driver=driver,
                timeframe=tf,
                status_cb=self.set_status,
                progress_cb=self.set_progress,
                log_cb=self.ui_log,
                stop_evt=self.stop_evt
            )

            if self.stop_evt.is_set():
                self.set_status("Pysäytetty.")
                return

            if not kl_rows:
                self.set_status("KL: Ei saatu rivejä. Katso debug dumpit output-kansiosta.")
                messagebox.showwarning("Ei rivejä", f"Kauppalehdestä ei saatu rivejä.\nKatso: {OUT_DIR}")
                return

            emails = [r[3] for r in results if r and len(r) >= 4 and r[3]]
            uniq_emails = len(set(e.lower() for e in emails))

            self.set_status("Valmis!")
            messagebox.showinfo(
                "Valmis",
                f"Valmis!\n\nKansio:\n{OUT_DIR}\n\n"
                f"KL rivejä: {len(kl_rows)}\n"
                f"Sähköposteja (uniikki): {uniq_emails}\n\n"
                f"Tiedostot:\n- kl_rivit.docx\n- tulokset.docx"
            )

        except WebDriverException as e:
            self.ui_log(f"SELENIUM VIRHE: {e}")
            self.set_status("Virhe. Katso log.txt")
            messagebox.showerror("Virhe", f"Selenium/Chrome virhe.\n\n{e}\n\nLogi: {LOG_PATH}")
        except Exception as e:
            self.ui_log(f"VIRHE: {e}")
            self.set_status("Virhe. Katso log.txt")
            messagebox.showerror("Virhe", f"Tuli virhe.\n\n{e}\n\nLogi: {LOG_PATH}")
        finally:
            try:
                if driver:
                    driver.quit()
            except Exception:
                pass
            self.btn_start.config(state="normal")
            self.btn_stop.config(state="disabled")


if __name__ == "__main__":
    App().mainloop()
