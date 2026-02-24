import os
import re
import sys
import time
import threading
import tkinter as tk
from tkinter import messagebox, filedialog
from tkinter import ttk
import subprocess
import queue

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
    ElementClickInterceptedException,
)
from webdriver_manager.chrome import ChromeDriverManager


# =========================
#   CONFIG / REGEX
# =========================
YT_RE = re.compile(r"\b\d{7}-\d\b|\b\d{8}\b")

EMAIL_RE = re.compile(r"[A-Za-z0-9_.+-]+@[A-Za-z0-9-]+\.[A-Za-z0-9-.]+")
EMAIL_OBF_RE = re.compile(
    r"[A-Za-z0-9_.+-]+\s*(?:\(a\)|\(at\)|\[at\]|\sat\s|\sät\s|\(ät\)|\{at\})\s*[A-Za-z0-9-]+\.[A-Za-z0-9-.]+",
    re.I
)

KAUPPALEHTI_URL = "https://www.kauppalehti.fi/yritykset/protestilista"
YTJ_COMPANY_URL = "https://tietopalvelu.ytj.fi/yritys/{}"

CHROME_DEBUG_PORT = 9222
CHROME_PROFILE_DIR = r"C:\chrome-botti"


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
#   UTIL
# =========================
def normalize_yt(yt: str):
    yt = (yt or "").strip().replace(" ", "")
    if re.fullmatch(r"\d{7}-\d", yt):
        return yt
    if re.fullmatch(r"\d{8}", yt):
        return yt[:7] + "-" + yt[7]
    return None


def normalize_email_obfuscated(s: str) -> str:
    if not s:
        return ""
    x = s.strip()

    # normalize common obfuscations before whitespace removal
    x = x.replace("[at]", "@").replace("[AT]", "@")
    x = x.replace("{at}", "@").replace("{AT}", "@")
    x = x.replace("(at)", "@").replace("(AT)", "@")
    x = x.replace("(a)", "@").replace("(A)", "@")
    x = x.replace("(ät)", "@").replace("ät", "@")

    # " at " patterns (keep safer)
    x = re.sub(r"\s+at\s+", "@", x, flags=re.I)

    # remove all whitespace
    x = re.sub(r"\s+", "", x)

    # handle "nameatdomain.fi" (only if it matches whole)
    x = re.sub(r"^([A-Za-z0-9_.+-]+)at([A-Za-z0-9-]+\.[A-Za-z0-9-.]+)$", r"\1@\2", x, flags=re.I)
    return x


def pick_email_from_text(text: str) -> str:
    if not text:
        return ""
    m = EMAIL_RE.search(text)
    if m:
        return m.group(0).strip()

    m2 = EMAIL_OBF_RE.search(text)
    if m2:
        return normalize_email_obfuscated(m2.group(0))

    # normalize whole chunk and search again
    t2 = normalize_email_obfuscated(text)
    m3 = EMAIL_RE.search(t2)
    if m3:
        return m3.group(0).strip()

    return ""


def save_word_plain_lines(lines, filename):
    path = os.path.join(OUT_DIR, filename)
    doc = Document()
    for line in lines:
        if line and str(line).strip():
            doc.add_paragraph(str(line).strip())
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
        except (ElementClickInterceptedException, Exception):
            driver.execute_script("arguments[0].click();", elem)
        return True
    except Exception:
        return False


def wait_dom_settle(driver, max_wait=2.5):
    """
    SPA settle: wait for readyState complete + small delay.
    """
    end = time.time() + max_wait
    while time.time() < end:
        try:
            rs = driver.execute_script("return document.readyState;")
            if rs == "complete":
                break
        except Exception:
            pass
        time.sleep(0.05)
    time.sleep(0.15)


# =========================
#   Cookie consent (SAFER)
# =========================
def try_accept_cookies(driver):
    """
    Safer consent accept:
    1) try to find banner/dialog containing cookie/eväste/consent terms
    2) click accept within that container only
    3) fallback to minimal global scan (very limited)
    """
    accept_texts = ["hyväksy", "hyväksy kaikki", "salli kaikki", "accept", "accept all", "i agree", "ok", "selvä"]
    container_terms = ["eväste", "cookie", "consent", "käyttöehdot", "privacy", "tietosuoja"]

    def click_accept_within(container):
        btns = []
        try:
            btns = container.find_elements(By.XPATH, ".//button|.//a|.//*[@role='button']")
        except Exception:
            return False
        for b in btns:
            try:
                if not b.is_displayed() or not b.is_enabled():
                    continue
                t = (b.text or "").strip().lower()
                if not t:
                    continue
                if any(a in t for a in accept_texts):
                    if safe_click(driver, b):
                        time.sleep(0.2)
                        return True
            except Exception:
                continue
        return False

    # 1) dialog-like containers first
    containers = []
    for xp in [
        "//*[@role='dialog']",
        "//*[contains(@class,'cookie') or contains(@id,'cookie') or contains(@class,'consent') or contains(@id,'consent')]",
        "//div",
    ]:
        try:
            containers.extend(driver.find_elements(By.XPATH, xp))
        except Exception:
            pass

    # filter containers by text
    cand = []
    for c in containers:
        try:
            if not c.is_displayed():
                continue
            txt = (c.text or "").strip().lower()
            if not txt:
                continue
            if any(term in txt for term in container_terms):
                cand.append(c)
        except Exception:
            continue

    # try click accept in best containers first (shorter text tends to be banners)
    cand.sort(key=lambda e: len((e.text or "")))
    for c in cand[:8]:
        if click_accept_within(c):
            return

    # 2) minimal fallback: only look for very explicit accept buttons
    for _ in range(2):
        try:
            elems = driver.find_elements(By.XPATH, "//button|//a|//*[@role='button']")
        except Exception:
            elems = []
        for e in elems:
            try:
                if not e.is_displayed() or not e.is_enabled():
                    continue
                t = (e.text or "").strip().lower()
                if t in ["hyväksy kaikki", "accept all", "salli kaikki"]:
                    if safe_click(driver, e):
                        time.sleep(0.2)
                        return
            except Exception:
                continue


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
#   CHROME STARTER
# =========================
def find_chrome_exe():
    candidates = [
        "chrome.exe",
        r"C:\Program Files\Google\Chrome\Application\chrome.exe",
        r"C:\Program Files (x86)\Google\Chrome\Application\chrome.exe",
    ]
    for c in candidates:
        try:
            if os.path.isabs(c):
                if os.path.exists(c):
                    return c
            else:
                return c
        except Exception:
            continue
    return "chrome.exe"


def launch_chrome_debug(profile_dir=CHROME_PROFILE_DIR, port=CHROME_DEBUG_PORT, start_url=KAUPPALEHTI_URL):
    os.makedirs(profile_dir, exist_ok=True)
    chrome = find_chrome_exe()
    cmd = [
        chrome,
        f"--remote-debugging-port={port}",
        f"--user-data-dir={profile_dir}",
        "--start-maximized",
        start_url,
    ]
    subprocess.Popen(cmd, shell=False)


# =========================
#   SELENIUM START
# =========================
def start_new_driver():
    options = webdriver.ChromeOptions()
    options.add_argument("--start-maximized")
    options.add_argument("--disable-gpu")
    options.add_argument("--disable-dev-shm-usage")

    # Keep webdriver_manager for your setup; it's fine but can mismatch sometimes.
    driver_path = ChromeDriverManager().install()
    return webdriver.Chrome(service=Service(driver_path), options=options)


def attach_to_existing_chrome():
    options = webdriver.ChromeOptions()
    options.add_experimental_option("debuggerAddress", f"127.0.0.1:{CHROME_DEBUG_PORT}")

    driver_path = ChromeDriverManager().install()
    return webdriver.Chrome(service=Service(driver_path), options=options)


def open_new_tab(driver, url="about:blank"):
    driver.execute_script("window.open(arguments[0], '_blank');", url)
    driver.switch_to.window(driver.window_handles[-1])


def list_tabs(driver):
    tabs = []
    for h in driver.window_handles:
        try:
            driver.switch_to.window(h)
            tabs.append((driver.title or "", driver.current_url or ""))
        except Exception:
            tabs.append(("", ""))
    return tabs


def focus_protestilista_tab(driver) -> bool:
    for handle in driver.window_handles:
        try:
            driver.switch_to.window(handle)
            url = (driver.current_url or "")
            if "kauppalehti.fi" in url and "protestilista" in url:
                return True
        except Exception:
            continue
    return False


def page_looks_like_protestilista(driver) -> bool:
    try:
        body = (driver.find_element(By.TAG_NAME, "body").text or "")
        if "Protestilista" not in body:
            return False
    except Exception:
        return False

    try:
        toggles = driver.find_elements(By.XPATH, "//*[@aria-expanded='false' or @aria-expanded='true']")
        toggles = [t for t in toggles if t.is_displayed()]
        if len(toggles) >= 1:
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
            "kirjaudu", "tilaa", "tilaajille", "vahvista henkilöllisyytesi",
            "sign in", "subscribe", "login"
        ]
        return any(w in text for w in bad_words)
    except Exception:
        return False


def ensure_protestilista_open_and_ready(driver, status_cb, log_cb, max_wait_seconds=900) -> bool:
    if focus_protestilista_tab(driver):
        status_cb("Löytyi protestilista-tab.")
    else:
        status_cb("Protestilista-tab ei löytynyt -> avaan protestilistan uuteen tabiin…")
        log_cb("AUTOFIX: opening protestilista in new tab")
        open_new_tab(driver, KAUPPALEHTI_URL)
        WebDriverWait(driver, 20).until(EC.presence_of_element_located((By.TAG_NAME, "body")))
        try_accept_cookies(driver)

    start = time.time()
    warned = False

    while True:
        try:
            try_accept_cookies(driver)
        except Exception:
            pass

        if page_looks_like_protestilista(driver):
            status_cb("Protestilista valmis (toggle-nuolet löytyvät).")
            return True

        if page_looks_like_login_or_paywall(driver) and not warned:
            warned = True
            status_cb("Kauppalehti vaatii kirjautumisen/tilaajanäkymän. Kirjaudu Chrome-bottiin.")
            log_cb("AUTOFIX: waiting for user to login / unlock paywall…")
            status_cb("Kirjaudu Kauppalehteen AUKI OLEVAAN Chrome-bottiin ja palaa tähän.")

        if time.time() - start > max_wait_seconds:
            status_cb("Aikakatkaisu: protestilista ei tullut näkyviin. Tarkista kirjautuminen.")
            log_cb("ERROR: timeout waiting protestilista")
            return False

        time.sleep(2)


# =========================
#   KAUPPALEHTI SCRAPE
# =========================
def click_nayta_lisaa(driver) -> bool:
    try:
        driver.execute_script("window.scrollTo(0, document.body.scrollHeight);")
    except Exception:
        pass
    time.sleep(0.3)

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


def get_visible_toggles(driver):
    elems = driver.find_elements(By.XPATH, "//*[@aria-expanded='false' or @aria-expanded='true']")
    toggles = []
    for e in elems:
        try:
            if not e.is_displayed():
                continue
            # avoid header thead
            try:
                e.find_element(By.XPATH, "ancestor::thead")
                continue
            except Exception:
                pass
            toggles.append(e)
        except Exception:
            continue
    return toggles


def wait_toggle_state(toggle, want_true: bool, timeout=2.5):
    end = time.time() + timeout
    while time.time() < end:
        try:
            val = (toggle.get_attribute("aria-expanded") or "").strip().lower()
            if want_true and val == "true":
                return True
            if (not want_true) and val == "false":
                return True
        except Exception:
            pass
        time.sleep(0.05)
    return False


def extract_yt_near_toggle(toggle):
    # nearest following rows first
    try:
        tr = toggle.find_element(By.XPATH, "ancestor::tr[1]")
        for k in range(1, 6):
            try:
                sib = tr.find_element(By.XPATH, f"following-sibling::tr[{k}]")
                txt = (sib.text or "")
                if "Y-TUNNUS" in txt:
                    for m in YT_RE.findall(txt):
                        n = normalize_yt(m)
                        if n:
                            return n
            except Exception:
                continue
    except Exception:
        pass

    # fallback wider container
    for xp in ["ancestor::tbody[1]", "ancestor::table[1]", "ancestor::div[2]", "ancestor::div[3]"]:
        try:
            c = toggle.find_element(By.XPATH, xp)
            txt = (c.text or "")
            if "Y-TUNNUS" in txt:
                for m in YT_RE.findall(txt):
                    n = normalize_yt(m)
                    if n:
                        return n
        except Exception:
            continue

    return ""


def collect_yts_from_kauppalehti(driver, status_cb, log_cb):
    WebDriverWait(driver, 25).until(EC.presence_of_element_located((By.TAG_NAME, "body")))
    try_accept_cookies(driver)

    if not ensure_protestilista_open_and_ready(driver, status_cb, log_cb, max_wait_seconds=900):
        return []

    collected = set()
    time.sleep(1.0)

    rounds_without_new = 0
    while True:
        toggles = get_visible_toggles(driver)
        if not toggles:
            status_cb("Kauppalehti: en löydä toggle-nuolia (aria-expanded).")
            log_cb("ERROR: no toggles found")
            break

        status_cb(f"Kauppalehti: toggleja {len(toggles)} näkyvissä | kerätty {len(collected)}")
        new_in_pass = 0

        for idx in range(len(toggles)):
            try:
                toggles = get_visible_toggles(driver)
                if idx >= len(toggles):
                    break
                t = toggles[idx]

                if not safe_click(driver, t):
                    continue
                wait_toggle_state(t, want_true=True, timeout=2.5)

                yt = ""
                for _ in range(30):
                    yt = extract_yt_near_toggle(t)
                    if yt:
                        break
                    time.sleep(0.08)

                if yt and yt not in collected:
                    collected.add(yt)
                    new_in_pass += 1
                    log_cb(f"+ {yt} (yht {len(collected)})")

                # close
                try:
                    safe_click(driver, t)
                    wait_toggle_state(t, want_true=False, timeout=1.2)
                except Exception:
                    pass

                time.sleep(0.03)

            except StaleElementReferenceException:
                continue
            except Exception:
                continue

        if new_in_pass == 0:
            rounds_without_new += 1
        else:
            rounds_without_new = 0

        if click_nayta_lisaa(driver):
            status_cb("Kauppalehti: Näytä lisää…")
            time.sleep(1.2)
            continue

        if rounds_without_new >= 2:
            status_cb("Kauppalehti: ei uusia + ei Näytä lisää -> valmis.")
            break

    return sorted(collected)


# =========================
#   YTJ EMAILS (ROBUST)
# =========================
def wait_ytj_loaded(driver):
    wait = WebDriverWait(driver, 35)
    wait.until(EC.presence_of_element_located((By.TAG_NAME, "body")))
    wait_dom_settle(driver, 2.0)
    try:
        wait.until(
            EC.presence_of_element_located(
                (By.XPATH, "//*[contains(., 'Y-tunnus') or contains(., 'Toiminimi') or contains(., 'Yritysmuoto') or contains(., 'Sähköposti')]")
            )
        )
    except Exception:
        pass
    wait_dom_settle(driver, 1.0)


def _expand_contact_sections_ytj(driver, log_cb=None):
    keywords = ["yhteyst", "asiointi", "lisät", "kontakt", "contact"]
    opened = 0

    for _round in range(3):
        try_accept_cookies(driver)
        wait_dom_settle(driver, 1.0)

        elems = []
        for xp in [
            "//*[@aria-expanded='false' and (self::button or self::a or @role='button')]",
            "//button[@aria-expanded='false']",
            "//*[@role='button' and @aria-expanded='false']",
        ]:
            try:
                elems.extend(driver.find_elements(By.XPATH, xp))
            except Exception:
                continue

        filtered = []
        for e in elems:
            try:
                if not e.is_displayed() or not e.is_enabled():
                    continue
                txt = ((e.text or "") + " " + (e.get_attribute("aria-label") or "")).strip().lower()
                if any(k in txt for k in keywords):
                    filtered.append(e)
            except Exception:
                continue

        # dedup
        uniq = []
        seen = set()
        for e in filtered:
            try:
                key = (e.get_attribute("outerHTML") or "")[:160]
            except Exception:
                key = str(id(e))
            if key in seen:
                continue
            seen.add(key)
            uniq.append(e)

        did = 0
        for e in uniq:
            try:
                if safe_click(driver, e):
                    did += 1
                    opened += 1
                    time.sleep(0.15)
            except Exception:
                continue

        if log_cb and did:
            log_cb(f"YTJ: avattiin accordion-osioita: {did}")

        if did == 0:
            break

    return opened


def _find_contact_container(driver):
    """
    Try to locate a container that likely holds contact details.
    We use it to limit generic "Näytä" clicks.
    """
    terms = ["yhteystied", "asiointi", "contact", "kontakt"]
    # find headings/labels and return a reasonable ancestor container
    for term in terms:
        try:
            anchors = driver.find_elements(By.XPATH, f"//*[contains(translate(normalize-space(.),'ABCDEFGHIJKLMNOPQRSTUVWXYZÄÖÅ','abcdefghijklmnopqrstuvwxyzäöå'), '{term}')]")
        except Exception:
            anchors = []
        for a in anchors[:8]:
            try:
                if not a.is_displayed():
                    continue
                for anc in ["ancestor::section[1]", "ancestor::div[1]", "ancestor::div[2]"]:
                    try:
                        c = a.find_element(By.XPATH, anc)
                        if c and c.is_displayed():
                            return c
                    except Exception:
                        continue
            except Exception:
                continue
    return None


def _find_email_row_show_button(driver):
    xps = [
        "//tr[.//*[contains(normalize-space(.), 'Sähköposti')]]//*[self::button or self::a or @role='button'][normalize-space()='Näytä' or .//span[normalize-space()='Näytä']]",
        "//*[contains(normalize-space(.), 'Sähköposti')]/following::*[self::button or self::a or @role='button'][1]",
        "//*[self::button or self::a or @role='button'][contains(translate(@aria-label,'SÄHKÖPOSTINÄYTÄ','sähköpostinäytä'),'sähköposti') and contains(translate(@aria-label,'NÄYTÄ','näytä'),'näytä')]",
    ]
    for xp in xps:
        try:
            elems = driver.find_elements(By.XPATH, xp)
        except Exception:
            elems = []
        for e in elems:
            try:
                if e.is_displayed() and e.is_enabled():
                    # ensure it really says Näytä if possible
                    t = (e.text or "").strip()
                    if t and "näytä" not in t.lower():
                        # might be wrong next sibling button; allow but lower priority
                        pass
                    return e
            except Exception:
                continue
    return None


def _find_visible_nayta_candidates_in_scope(scope_elem):
    xpaths = [
        ".//button[normalize-space()='Näytä']",
        ".//button[.//span[normalize-space()='Näytä']]",
        ".//*[@role='button' and normalize-space()='Näytä']",
        ".//a[normalize-space()='Näytä']",
        ".//*[self::button or self::a or @role='button'][contains(translate(@aria-label,'NÄYTÄ','näytä'),'näytä')]",
        ".//*[self::button or self::a or @role='button'][contains(translate(@title,'NÄYTÄ','näytä'),'näytä')]",
    ]
    found = []
    for xp in xpaths:
        try:
            elems = scope_elem.find_elements(By.XPATH, xp)
        except Exception:
            elems = []
        for e in elems:
            try:
                if e.is_displayed() and e.is_enabled():
                    found.append(e)
            except Exception:
                continue
    return found


def click_all_nayta_ytj(driver, log_cb=None):
    """
    1) expand contact sections
    2) click email-row "Näytä" first (high value)
    3) only if needed: click remaining "Näytä" inside contact container (NOT whole page)
    """
    total_clicked = 0

    try:
        _expand_contact_sections_ytj(driver, log_cb=log_cb)
    except Exception:
        pass

    try_accept_cookies(driver)
    wait_dom_settle(driver, 1.0)

    # 1) email-row show button
    btn = _find_email_row_show_button(driver)
    if btn:
        try:
            if safe_click(driver, btn):
                total_clicked += 1
                if log_cb:
                    log_cb("YTJ: klikattiin Sähköposti-rivin 'Näytä'")
                time.sleep(0.25)
        except Exception:
            pass

    # 2) scope-limited generic "Näytä"
    scope = _find_contact_container(driver)
    if not scope:
        # if we can't find a contact container, do NOT spam whole page
        if log_cb:
            log_cb("YTJ: contact-container ei löytynyt -> skipataan massaklikkaus (turvallisuus)")
        return total_clicked

    for round_idx in range(1, 6):
        try_accept_cookies(driver)
        wait_dom_settle(driver, 1.0)

        candidates = _find_visible_nayta_candidates_in_scope(scope)

        # dedup
        unique = []
        seen = set()
        for c in candidates:
            try:
                key = (c.get_attribute("outerHTML") or "")[:180]
            except Exception:
                key = str(id(c))
            if key in seen:
                continue
            seen.add(key)
            unique.append(c)

        if not unique:
            break

        clicked_this_round = 0
        for c in unique:
            try:
                if safe_click(driver, c):
                    clicked_this_round += 1
                    total_clicked += 1
                    time.sleep(0.12)
            except Exception:
                continue

        if log_cb:
            log_cb(f"YTJ: Näytä-klikkejä (scope) kierros {round_idx}: {clicked_this_round}")

        if clicked_this_round == 0:
            break

        time.sleep(0.25)

    return total_clicked


def extract_email_from_ytj(driver):
    # 1) mailto
    try:
        for a in driver.find_elements(By.TAG_NAME, "a"):
            href = (a.get_attribute("href") or "")
            if href.lower().startswith("mailto:"):
                return href.split(":", 1)[1].strip()
    except Exception:
        pass

    # 2) anything containing "Sähköposti"
    try:
        candidates = driver.find_elements(By.XPATH, "//*[contains(normalize-space(.), 'Sähköposti')]")
        for c in candidates:
            email = pick_email_from_text(c.text or "")
            if email:
                return email
    except Exception:
        pass

    # 3) whole page
    try:
        body = driver.find_element(By.TAG_NAME, "body").text or ""
        return pick_email_from_text(body)
    except Exception:
        return ""


def fetch_emails_from_ytj(driver, yt_list, status_cb, progress_cb, log_cb):
    emails = []
    seen = set()
    progress_cb(0, max(1, len(yt_list)))

    for i, yt in enumerate(yt_list, start=1):
        status_cb(f"YTJ: {i}/{len(yt_list)} {yt}")
        progress_cb(i - 1, len(yt_list))

        driver.get(YTJ_COMPANY_URL.format(yt))

        try:
            wait_ytj_loaded(driver)
        except TimeoutException:
            log_cb("YTJ: timeout ladatessa, yritän silti klikata/poimia…")

        try_accept_cookies(driver)

        email = ""
        for attempt in range(1, 5):
            try:
                clicks = click_all_nayta_ytj(driver, log_cb=log_cb)
                wait_dom_settle(driver, 1.5)

                for _ in range(10):
                    email = extract_email_from_ytj(driver)
                    if email:
                        break
                    time.sleep(0.2)

                if email:
                    break

                log_cb(f"YTJ: ei emailia {yt} (attempt {attempt}, clicks {clicks})")
                wait_dom_settle(driver, 1.0)

            except StaleElementReferenceException:
                time.sleep(0.2)
                continue
            except Exception as e:
                log_cb(f"YTJ: virhe attempt {attempt}: {e}")
                time.sleep(0.25)
                continue

        if email:
            k = email.lower()
            if k not in seen:
                seen.add(k)
                emails.append(email)
                log_cb(email)

        time.sleep(0.08)

    progress_cb(len(yt_list), max(1, len(yt_list)))
    return emails


# =========================
#   GUI (THREAD-SAFE)
# =========================
class App(tk.Tk):
    def __init__(self):
        super().__init__()
        reset_log()

        self._uiq = queue.Queue()
        self._running_lock = threading.Lock()
        self._is_running = False

        self.title("ProtestiBotti (FULL FIX)")
        self.geometry("980x620")

        tk.Label(self, text="ProtestiBotti", font=("Arial", 18, "bold")).pack(pady=10)
        tk.Label(
            self,
            text="Moodit:\n1) Kauppalehti (Chrome debug 9222) → Y-tunnukset → YTJ sähköpostit\n2) PDF → Y-tunnukset → YTJ sähköpostit",
            justify="center"
        ).pack(pady=4)

        btn_row = tk.Frame(self)
        btn_row.pack(pady=8)

        self.btn_open = tk.Button(btn_row, text="Avaa Chrome-botti (9222)", font=("Arial", 12), command=self.open_chrome_bot)
        self.btn_open.grid(row=0, column=0, padx=8)

        self.btn_kl = tk.Button(btn_row, text="Kauppalehti → YTJ", font=("Arial", 12), command=self.start_kauppalehti_mode)
        self.btn_kl.grid(row=0, column=1, padx=8)

        self.btn_pdf = tk.Button(btn_row, text="PDF → YTJ", font=("Arial", 12), command=self.start_pdf_mode)
        self.btn_pdf.grid(row=0, column=2, padx=8)

        self.status = tk.Label(self, text="Valmiina.", font=("Arial", 11))
        self.status.pack(pady=6)

        self.progress = ttk.Progressbar(self, orient="horizontal", mode="determinate", length=920)
        self.progress.pack(pady=6)

        frame = tk.Frame(self)
        frame.pack(fill="both", expand=True, padx=14, pady=10)

        tk.Label(frame, text="Live-logi (uusimmat alimmaisena):").pack(anchor="w")
        self.listbox = tk.Listbox(frame, height=18)
        self.listbox.pack(side="left", fill="both", expand=True)

        sb = tk.Scrollbar(frame, orient="vertical", command=self.listbox.yview)
        sb.pack(side="right", fill="y")
        self.listbox.configure(yscrollcommand=sb.set)

        tk.Label(self, text=f"Tallennus: {OUT_DIR}", wraplength=960, justify="center").pack(pady=6)

        # UI queue poller
        self.after(50, self._drain_ui_queue)

    # ---------- thread-safe UI API ----------
    def _post(self, kind, *args):
        self._uiq.put((kind, args))

    def _drain_ui_queue(self):
        try:
            while True:
                kind, args = self._uiq.get_nowait()
                if kind == "log":
                    msg = args[0]
                    line = log_to_file(msg)
                    self.listbox.insert(tk.END, line)
                    self.listbox.yview_moveto(1.0)
                elif kind == "status":
                    s = args[0]
                    self.status.config(text=s)
                    line = log_to_file(s)
                    self.listbox.insert(tk.END, line)
                    self.listbox.yview_moveto(1.0)
                elif kind == "progress":
                    value, maximum = args
                    self.progress["maximum"] = maximum
                    self.progress["value"] = value
                elif kind == "msgbox":
                    mtype, title, text = args
                    if mtype == "info":
                        messagebox.showinfo(title, text)
                    elif mtype == "warn":
                        messagebox.showwarning(title, text)
                    else:
                        messagebox.showerror(title, text)
                elif kind == "running":
                    running = args[0]
                    self._set_buttons_enabled(not running)
        except queue.Empty:
            pass
        self.after(50, self._drain_ui_queue)

    def _set_buttons_enabled(self, enabled: bool):
        state = tk.NORMAL if enabled else tk.DISABLED
        self.btn_open.config(state=state)
        self.btn_kl.config(state=state)
        self.btn_pdf.config(state=state)

    def ui_log(self, msg):
        self._post("log", msg)

    def set_status(self, s):
        self._post("status", s)

    def set_progress(self, value, maximum):
        self._post("progress", value, maximum)

    def show_info(self, title, text):
        self._post("msgbox", "info", title, text)

    def show_warn(self, title, text):
        self._post("msgbox", "warn", title, text)

    def show_error(self, title, text):
        self._post("msgbox", "error", title, text)

    def _begin_run(self) -> bool:
        with self._running_lock:
            if self._is_running:
                return False
            self._is_running = True
        self._post("running", True)
        return True

    def _end_run(self):
        with self._running_lock:
            self._is_running = False
        self._post("running", False)

    # ---------- actions ----------
    def open_chrome_bot(self):
        try:
            self.set_status("Avataan Chrome-botti (9222) ja protestilista…")
            launch_chrome_debug(profile_dir=CHROME_PROFILE_DIR, port=CHROME_DEBUG_PORT, start_url=KAUPPALEHTI_URL)
            self.set_status("Chrome auki. Kirjaudu Kauppalehteen tässä Chromessa (jos pyytää), sitten paina 'Kauppalehti → YTJ'.")
        except Exception as e:
            self.ui_log(f"VIRHE Chrome-botin avauksessa: {e}")
            self.show_error("Virhe", f"Chrome-botin avaus epäonnistui:\n{e}")

    def start_kauppalehti_mode(self):
        if not self._begin_run():
            self.show_warn("Kesken", "Ajo on jo käynnissä. Odota että se valmistuu.")
            return
        threading.Thread(target=self.run_kauppalehti_mode, daemon=True).start()

    def run_kauppalehti_mode(self):
        driver = None
        try:
            self.set_status("Liitytään Chrome-bottiin (9222)…")
            driver = attach_to_existing_chrome()

            self.set_status("Kauppalehti: kerätään Y-tunnukset…")
            yt_list = collect_yts_from_kauppalehti(driver, self.set_status, self.ui_log)

            if not yt_list:
                self.set_status("Ei löytynyt Y-tunnuksia.")
                self.show_warn("Ei löytynyt", "Y-tunnuksia ei saatu. Katso log.txt (kirjautuminen/paynwall/DOM).")
                return

            yt_path = save_word_plain_lines(yt_list, "ytunnukset.docx")
            self.ui_log(f"Tallennettu: {yt_path}")

            self.set_status("Avataan YTJ uuteen tabiin…")
            open_new_tab(driver, "about:blank")

            self.set_status("YTJ: haetaan sähköpostit…")
            emails = fetch_emails_from_ytj(driver, yt_list, self.set_status, self.set_progress, self.ui_log)

            em_path = save_word_plain_lines(emails, "sahkopostit.docx")
            self.ui_log(f"Tallennettu: {em_path}")

            self.set_status("Valmis!")
            self.show_info("Valmis", f"Valmis!\n\nKansio:\n{OUT_DIR}\n\nY-tunnuksia: {len(yt_list)}\nSähköposteja: {len(emails)}")

        except WebDriverException as e:
            self.ui_log(f"SELENIUM VIRHE: {e}")
            self.set_status("Virhe. Katso log.txt")
            self.show_error("Virhe", f"Selenium/Chrome virhe.\nKatso log.txt:\n{LOG_PATH}\n\n{e}")
        except Exception as e:
            self.ui_log(f"VIRHE: {e}")
            self.set_status("Virhe. Katso log.txt")
            self.show_error("Virhe", f"Tuli virhe.\nKatso log.txt:\n{LOG_PATH}\n\n{e}")
        finally:
            # attach-mode: don't quit user's chrome
            self._end_run()

    def start_pdf_mode(self):
        path = filedialog.askopenfilename(filetypes=[("PDF files", "*.pdf")])
        if not path:
            return
        if not self._begin_run():
            self.show_warn("Kesken", "Ajo on jo käynnissä. Odota että se valmistuu.")
            return
        threading.Thread(target=self.run_pdf_mode, args=(path,), daemon=True).start()

    def run_pdf_mode(self, pdf_path):
        driver = None
        try:
            self.set_status("Luetaan PDF ja kerätään Y-tunnukset…")
            yt_list = extract_ytunnukset_from_pdf(pdf_path)

            if not yt_list:
                self.set_status("Ei löytynyt Y-tunnuksia PDF:stä.")
                self.show_warn("Ei löytynyt", "PDF:stä ei löytynyt yhtään Y-tunnusta.")
                return

            yt_path = save_word_plain_lines(yt_list, "ytunnukset.docx")
            self.ui_log(f"Tallennettu: {yt_path}")

            self.set_status("Käynnistetään Chrome ja haetaan sähköpostit YTJ:stä…")
            driver = start_new_driver()

            emails = fetch_emails_from_ytj(driver, yt_list, self.set_status, self.set_progress, self.ui_log)
            em_path = save_word_plain_lines(emails, "sahkopostit.docx")
            self.ui_log(f"Tallennettu: {em_path}")

            self.set_status("Valmis!")
            self.show_info("Valmis", f"Valmis!\n\nKansio:\n{OUT_DIR}\n\nY-tunnuksia: {len(yt_list)}\nSähköposteja: {len(emails)}")

        except Exception as e:
            self.ui_log(f"VIRHE: {e}")
            self.set_status("Virhe. Katso log.txt")
            self.show_error("Virhe", f"Tuli virhe.\nKatso log.txt:\n{LOG_PATH}\n\n{e}")
        finally:
            if driver:
                try:
                    driver.quit()
                except Exception:
                    pass
            self._end_run()


if __name__ == "__main__":
    App().mainloop()
