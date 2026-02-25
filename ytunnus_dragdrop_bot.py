# protestibotti.py
# ProtestiBotti:
#   1) PDF -> YTJ sähköpostit (toimiva, ei kosketa)
#   2) Kauppalehti -> (kerää Y-tunnukset) -> YTJ sähköpostit
#   3) Clipboard (Ctrl+C sivulta -> Ctrl+V bottiin) -> YTJ sähköpostit
#
# Riippuvuudet:
#   pip install selenium webdriver-manager PyPDF2 python-docx tkinterdnd2
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

# Drag & Drop (PDF)
try:
    from tkinterdnd2 import DND_FILES, TkinterDnD  # type: ignore
    HAS_DND = True
except Exception:
    HAS_DND = False

# =========================
#   TUNING (NOPEUS)
# =========================
# YTJ sivun odotus / retryt (nopeammaksi)
YTJ_PAGE_LOAD_TIMEOUT = 18
YTJ_RETRY_READS = 5
YTJ_RETRY_SLEEP = 0.12

# “Näytä” klikkien max kierrokset
YTJ_NAYTA_PASSES = 2

# Pieni per-yritys viive
YTJ_PER_COMPANY_SLEEP = 0.03

# Kauppalehti “Näytä lisää” jälkeen odotus
KL_LOAD_MORE_WAIT = 1.1

# Kauppalehti yrityssivu-tabin odotus
KL_COMPANY_PAGE_TIMEOUT = 18
KL_AFTER_OPEN_SLEEP = 0.05

# =========================
#   CONFIG / REGEX
# =========================
YT_RE = re.compile(r"\b\d{7}-\d\b|\b\d{8}\b")
EMAIL_RE = re.compile(r"[A-Za-z0-9_.+-]+@[A-Za-z0-9-]+\.[A-Za-z0-9-.]+")
EMAIL_A_RE = re.compile(r"[A-Za-z0-9_.+-]+\s*\(a\)\s*[A-Za-z0-9-]+\.[A-Za-z0-9-.]+", re.I)

KAUPPALEHTI_URL = "https://www.kauppalehti.fi/yritykset/protestilista"
KAUPPALEHTI_MATCH = "kauppalehti.fi/yritykset/protestilista"
YTJ_COMPANY_URL = "https://tietopalvelu.ytj.fi/yritys/{}"

# =========================
#   PATHS
# =========================
def get_exe_dir():
    if getattr(sys, "frozen", False):
        return os.path.dirname(sys.executable)
    return os.path.dirname(os.path.abspath(__file__))


def get_output_dir():
    base = get_exe_dir()
    # test write to exe dir
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


def extract_names_from_clipboard(text: str):
    """
    Heuristiikka: poimi “yritysnimiltä näyttävät” rivit copy/pastesta.
    Jos tekstissä on suoraan Y-tunnuksia, niitä käytetään ensin (tämä funktio vain nimille).
    """
    lines = split_lines(text)
    out = []
    seen = set()

    bad_contains = [
        "näytä lisää", "protestilista", "kauppalehti", "kirjaudu", "tilaa", "tilaajille",
        "€", "eur", "summa", "viiväst", "päivä", "päivää", "päivämäärä",
        "y-tunnus", "y tunnus", "ytunnus", "osoite", "postinumero",
    ]

    for ln in lines:
        low = ln.lower()

        # jos rivillä on suoraan ytunnus, ohitetaan nimilistasta
        if YT_RE.search(ln):
            continue

        # suodata roina
        if any(b in low for b in bad_contains):
            continue

        if len(ln) < 3:
            continue

        # liikaa numeroita -> tuskin nimi
        digits = sum(ch.isdigit() for ch in ln)
        if digits >= 3:
            continue

        if not any(ch.isalpha() for ch in ln):
            continue

        name = re.sub(r"\s{2,}", " ", ln).strip()
        if len(name) > 80:
            continue

        key = name.lower()
        if key in seen:
            continue

        seen.add(key)
        out.append(name)

    return out


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


def list_tabs(driver):
    tabs = []
    for h in driver.window_handles:
        try:
            driver.switch_to.window(h)
            tabs.append((driver.title or "", driver.current_url or ""))
        except Exception:
            tabs.append(("", ""))
    return tabs


def focus_kauppalehti_tab(driver, log_cb=None) -> bool:
    found = False
    for handle in driver.window_handles:
        try:
            driver.switch_to.window(handle)
            url = (driver.current_url or "")
            if KAUPPALEHTI_MATCH in url:
                found = True
                break
        except Exception:
            continue

    if log_cb:
        log_cb("Chrome TAB LISTA (title | url):")
        for title, url in list_tabs(driver):
            log_cb(f"  {title} | {url}")

    return found


# =========================
#   KAUPPALEHTI HELPERS
# =========================
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


def ensure_protestilista_open_and_ready(driver, status_cb, log_cb, max_wait_seconds=900) -> bool:
    # 1) löytyykö tab?
    if focus_kauppalehti_tab(driver, log_cb):
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
    """
    Kerää näkyvistä riveistä yrityslinkkien hrefit (robusti).
    """
    hrefs = []
    rows = driver.find_elements(By.XPATH, "//table//tbody//tr")
    for r in rows:
        try:
            if not r.is_displayed():
                continue
            links = r.find_elements(By.XPATH, ".//td[1]//a[contains(@href,'/yritykset/') and normalize-space(.)!='']")
            for a in links:
                try:
                    href = (a.get_attribute("href") or "").strip()
                    if href and "/yritykset/" in href:
                        hrefs.append(href)
                except Exception:
                    continue
        except Exception:
            continue

    # uniq preserve order
    out = []
    seen = set()
    for h in hrefs:
        if h not in seen:
            seen.add(h)
            out.append(h)
    return out


def extract_yt_from_text_anywhere(txt: str) -> str:
    if not txt:
        return ""
    for m in YT_RE.findall(txt):
        n = normalize_yt(m)
        if n:
            return n
    return ""


def extract_yt_from_company_page_in_new_tab(driver, href: str, stop_flag):
    """
    Avaa yrityssivu uuteen tabiin, poimi Y-tunnus sivun body-tekstistä, sulje tabi.
    """
    if stop_flag.is_set():
        return ""

    parent = driver.current_window_handle
    open_new_tab(driver, href)

    yt = ""
    try:
        WebDriverWait(driver, KL_COMPANY_PAGE_TIMEOUT).until(
            EC.presence_of_element_located((By.TAG_NAME, "body"))
        )
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
            try:
                driver.switch_to.window(driver.window_handles[0])
            except Exception:
                pass

    return yt


def collect_yts_from_kauppalehti(driver, status_cb, log_cb, stop_flag):
    WebDriverWait(driver, 25).until(EC.presence_of_element_located((By.TAG_NAME, "body")))
    try_accept_cookies(driver)

    if not ensure_protestilista_open_and_ready(driver, status_cb, log_cb, max_wait_seconds=900):
        return []

    collected = set()
    seen_hrefs = set()
    loops = 0

    while True:
        if stop_flag.is_set():
            status_cb("Pysäytetty.")
            break

        loops += 1
        hrefs = get_company_hrefs_from_visible_rows(driver)
        if not hrefs:
            status_cb("Kauppalehti: en löydä yrityslinkkejä (lista ei näy).")
            log_cb("ERROR: no company hrefs found")
            break

        new_hrefs = [h for h in hrefs if h not in seen_hrefs]
        status_cb(f"Kauppalehti: linkkejä {len(hrefs)} | uudet {len(new_hrefs)} | Y-tunnuksia {len(collected)}")

        got_this_pass = 0
        for href in new_hrefs:
            if stop_flag.is_set():
                status_cb("Pysäytetty.")
                break

            seen_hrefs.add(href)
            try:
                yt = extract_yt_from_company_page_in_new_tab(driver, href, stop_flag)
                if yt and yt not in collected:
                    collected.add(yt)
                    got_this_pass += 1
                    log_cb(f"+ {yt} (yht {len(collected)})")
                elif not yt:
                    log_cb("SKIP: Y-tunnusta ei löytynyt yrityssivulta")
            except StaleElementReferenceException:
                continue
            except Exception as e:
                log_cb(f"SKIP: yrityssivu error: {e}")

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

        if got_this_pass == 0 and len(new_hrefs) == 0:
            status_cb("Kauppalehti: ei uusia linkkejä + ei Näytä lisää -> valmis.")
            break

        if loops >= 3 and got_this_pass == 0 and len(new_hrefs) > 0:
            status_cb("Kauppalehti: uusia linkkejä mutta ei Y-tunnuksia -> lopetan (paywall/DOM-muutos?).")
            break

    return sorted(collected)


# =========================
#   YTJ EMAILS  (PDF->YTJ: ÄLÄ RIKO)
# =========================
def click_all_nayta_ytj(driver):
    # “Näytä” nappi voi olla useassa kohdassa (Puhelin, Email, yms)
    for _ in range(YTJ_NAYTA_PASSES):
        clicked = False
        # buttons
        for b in driver.find_elements(By.TAG_NAME, "button"):
            try:
                if (b.text or "").strip().lower() == "näytä" and b.is_displayed() and b.is_enabled():
                    safe_click(driver, b)
                    clicked = True
                    time.sleep(0.08)
            except Exception:
                continue
        # links
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
    # mailto first
    try:
        for a in driver.find_elements(By.TAG_NAME, "a"):
            href = (a.get_attribute("href") or "")
            if href.lower().startswith("mailto:"):
                return href.split(":", 1)[1].strip()
    except Exception:
        pass

    # row with “Sähköposti”
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

    # fallback whole body
    try:
        return pick_email_from_text(driver.find_element(By.TAG_NAME, "body").text or "")
    except Exception:
        return ""


def fetch_emails_from_ytj(driver, yt_list, status_cb, progress_cb, log_cb, stop_flag):
    emails = []
    seen = set()
    progress_cb(0, max(1, len(yt_list)))

    for i, yt in enumerate(yt_list, start=1):
        if stop_flag.is_set():
            status_cb("Pysäytetty.")
            break

        status_cb(f"YTJ: {i}/{len(yt_list)} {yt}")
        progress_cb(i - 1, len(yt_list))

        try:
            driver.get(YTJ_COMPANY_URL.format(yt))
        except TimeoutException:
            # try continue
            pass

        wait_ytj_loaded(driver)
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
                log_cb(email)

        time.sleep(YTJ_PER_COMPANY_SLEEP)

    progress_cb(len(yt_list), max(1, len(yt_list)))
    return emails


# =========================
#   YTJ NAME SEARCH (Clipboard moodi)
# =========================
def ytj_open_search_home(driver):
    driver.get("https://tietopalvelu.ytj.fi/")
    wait_ytj_loaded(driver)
    try_accept_cookies(driver)


def ytj_find_company_and_open_first(driver, name: str):
    """
    Hakee YTJ:stä nimellä ja avaa ensimmäisen osuman (automaattinen paras osuma).
    Palauttaa True jos päätyi yrityssivulle.
    """
    ytj_open_search_home(driver)

    candidates = []
    try:
        candidates += driver.find_elements(By.XPATH, "//input[@type='search']")
    except Exception:
        pass
    try:
        candidates += driver.find_elements(By.XPATH, "//input[contains(@placeholder,'Hae') or contains(@aria-label,'Hae')]")
    except Exception:
        pass
    try:
        candidates += driver.find_elements(By.XPATH, "//input")
    except Exception:
        pass

    search_box = None
    for inp in candidates:
        try:
            if not inp.is_displayed() or not inp.is_enabled():
                continue
            ph = (inp.get_attribute("placeholder") or "").lower()
            al = (inp.get_attribute("aria-label") or "").lower()
            if "hae" in ph or "hae" in al or (inp.get_attribute("type") == "search"):
                search_box = inp
                break
        except Exception:
            continue

    if not search_box:
        return False

    try:
        search_box.clear()
    except Exception:
        pass

    try:
        search_box.send_keys(name)
        search_box.send_keys(u"\ue007")  # ENTER
    except Exception:
        return False

    # Odota tuloksia ja avaa eka yrityslinkki
    for _ in range(40):
        try:
            try_accept_cookies(driver)
            links = driver.find_elements(By.XPATH, "//a[contains(@href,'/yritys/') or contains(@href,'/company/')]")
            if not links:
                links = driver.find_elements(By.XPATH, "//a[contains(@href,'yritys')]")
            for a in links:
                try:
                    href = (a.get_attribute("href") or "")
                    if href and ("tietopalvelu.ytj.fi" in href):
                        safe_click(driver, a)
                        wait_ytj_loaded(driver)
                        return True
                except Exception:
                    continue
        except Exception:
            pass
        time.sleep(0.15)

    return False


# =========================
#   CHROME BOT LAUNCHER (9222)
# =========================
def build_chrome_bot_command():
    # create a dedicated profile folder next to exe
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

    cmd = f"\"{chrome_exe}\" --remote-debugging-port=9222 --user-data-dir=\"{profile_dir}\""
    return cmd, profile_dir


def launch_chrome_bot():
    cmd, _ = build_chrome_bot_command()
    # open PowerShell and run chrome
    ps = f'Start-Process -FilePath powershell -ArgumentList \'-NoExit\', \'-Command\', \'{cmd}\''

    try:
        os.system(f'powershell -NoProfile -ExecutionPolicy Bypass -Command "{ps}"')
        return True
    except Exception:
        return False


# =========================
#   GUI (scroll + stop)
# =========================
class ScrollableFrame(ttk.Frame):
    def __init__(self, container, *args, **kwargs):
        super().__init__(container, *args, **kwargs)
        canvas = tk.Canvas(self, borderwidth=0, highlightthickness=0)
        scrollbar = ttk.Scrollbar(self, orient="vertical", command=canvas.yview)
        self.scrollable_frame = ttk.Frame(canvas)

        self.scrollable_frame.bind(
            "<Configure>",
            lambda e: canvas.configure(scrollregion=canvas.bbox("all"))
        )

        canvas.create_window((0, 0), window=self.scrollable_frame, anchor="nw")
        canvas.configure(yscrollcommand=scrollbar.set)

        canvas.pack(side="left", fill="both", expand=True)
        scrollbar.pack(side="right", fill="y")

        # mouse wheel
        def _on_mousewheel(event):
            canvas.yview_scroll(int(-1 * (event.delta / 120)), "units")
        canvas.bind_all("<MouseWheel>", _on_mousewheel)

        self.canvas = canvas


BaseTk = TkinterDnD.Tk if HAS_DND else tk.Tk


class App(BaseTk):
    def __init__(self):
        super().__init__()
        reset_log()

        self.stop_flag = threading.Event()

        self.title("ProtestiBotti (Kauppalehti + PDF + Clipboard -> YTJ)")
        self.geometry("980x780")

        # ESC STOP (oikeasti toimiva)
        self.bind_all("<Escape>", lambda e: self.request_stop())

        outer = ScrollableFrame(self)
        outer.pack(fill="both", expand=True)
        root = outer.scrollable_frame

        tk.Label(root, text="ProtestiBotti", font=("Arial", 18, "bold")).pack(pady=10)
        tk.Label(
            root,
            text="Moodit:\n"
                 "1) Kauppalehti (Chrome debug 9222) → Y-tunnukset → YTJ sähköpostit\n"
                 "2) PDF → Y-tunnukset → YTJ sähköpostit\n"
                 "3) Clipboard: Ctrl+C sivulta → Ctrl+V bottiin → YTJ sähköpostit\n\n"
                 "Hätäseis: Pysäytä-nappi tai ESC.",
            justify="center"
        ).pack(pady=4)

        btn_row = tk.Frame(root)
        btn_row.pack(pady=8)

        tk.Button(btn_row, text="Avaa Chrome-botti (9222)", font=("Arial", 12), command=self.open_chrome_bot).grid(row=0, column=0, padx=8)
        tk.Button(btn_row, text="Kauppalehti → YTJ", font=("Arial", 12), command=self.start_kauppalehti_mode).grid(row=0, column=1, padx=8)
        tk.Button(btn_row, text="PDF → YTJ", font=("Arial", 12), command=self.start_pdf_mode).grid(row=0, column=2, padx=8)
        tk.Button(btn_row, text="Pysäytä", font=("Arial", 12), command=self.request_stop).grid(row=0, column=3, padx=8)

        self.status = tk.Label(root, text="Valmiina.", font=("Arial", 11))
        self.status.pack(pady=6)

        self.progress = ttk.Progressbar(root, orient="horizontal", mode="determinate", length=920)
        self.progress.pack(pady=6)

        # PDF drop zone
        self.drop_var = tk.StringVar(value="PDF: Pudota tähän (tai paina PDF → YTJ ja valitse tiedosto)")
        drop = tk.Label(root, textvariable=self.drop_var, relief="groove", height=2)
        drop.pack(fill="x", padx=14, pady=6)
        if HAS_DND:
            drop.drop_target_register(DND_FILES)
            drop.dnd_bind("<<Drop>>", self._on_drop_pdf)

        # Clipboard input
        tk.Label(root, text="Clipboard (Ctrl+V tähän) → hae sähköpostit YTJ:stä:", font=("Arial", 11, "bold")).pack(pady=(10, 4))
        self.clip_text = tk.Text(root, height=7, wrap="word")
        self.clip_text.pack(fill="x", padx=14)

        clip_btn_row = tk.Frame(root)
        clip_btn_row.pack(pady=6)
        tk.Button(clip_btn_row, text="Clipboard → YTJ", font=("Arial", 12), command=self.start_clipboard_mode).pack()

        frame = tk.Frame(root)
        frame.pack(fill="both", expand=True, padx=14, pady=10)

        tk.Label(frame, text="Live-logi (uusimmat alimmaisena):").pack(anchor="w")
        self.listbox = tk.Listbox(frame, height=18)
        self.listbox.pack(side="left", fill="both", expand=True)

        sb = tk.Scrollbar(frame, orient="vertical", command=self.listbox.yview)
        sb.pack(side="right", fill="y")
        self.listbox.configure(yscrollcommand=sb.set)

        tk.Label(root, text=f"Tallennus: {OUT_DIR}", wraplength=940, justify="center").pack(pady=6)

    def request_stop(self):
        self.stop_flag.set()
        self.ui_log("STOP: käyttäjä pyysi pysäytystä.")
        self.status.config(text="Pysäytetään…")

    def clear_stop(self):
        self.stop_flag.clear()

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

            self.set_status("Kauppalehti: kerätään Y-tunnukset (yrityssivut linkeistä)…")
            yt_list = collect_yts_from_kauppalehti(driver, self.set_status, self.ui_log, self.stop_flag)

            if self.stop_flag.is_set():
                self.set_status("Pysäytetty.")
                return

            if not yt_list:
                self.set_status("Ei löytynyt Y-tunnuksia.")
                messagebox.showwarning("Ei löytynyt", "Y-tunnuksia ei saatu. Katso log.txt (kirjautuminen/paynwall/DOM).")
                return

            self.set_status("Avataan YTJ uuteen tabiin…")
            open_new_tab(driver, "about:blank")

            self.set_status("YTJ: haetaan sähköpostit…")
            emails = fetch_emails_from_ytj(driver, yt_list, self.set_status, self.set_progress, self.ui_log, self.stop_flag)

            if self.stop_flag.is_set():
                self.set_status("Pysäytetty.")
                return

            em_path = save_word_plain_lines(emails, "sahkopostit.docx")
            self.ui_log(f"Tallennettu: {em_path}")

            self.set_status("Valmis!")
            messagebox.showinfo(
                "Valmis",
                f"Valmis!\n\nKansio:\n{OUT_DIR}\n\nSähköposteja: {len(emails)}"
            )

        except WebDriverException as e:
            self.ui_log(f"SELENIUM VIRHE: {e}")
            self.set_status("Virhe. Katso log.txt")
            messagebox.showerror("Virhe", f"Selenium/Chrome virhe.\nKatso log.txt:\n{LOG_PATH}\n\n{e}")
        except Exception as e:
            self.ui_log(f"VIRHE: {e}")
            self.set_status("Virhe. Katso log.txt")
            messagebox.showerror("Virhe", f"Tuli virhe.\nKatso log.txt:\n{LOG_PATH}\n\n{e}")
        finally:
            pass  # ei suljeta käyttäjän Chromea

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

            em_path = save_word_plain_lines(emails, "sahkopostit.docx")
            self.ui_log(f"Tallennettu: {em_path}")

            self.set_status("Valmis!")
            messagebox.showinfo(
                "Valmis",
                f"Valmis!\n\nKansio:\n{OUT_DIR}\n\nSähköposteja: {len(emails)}"
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
            # 1) jos tekstissä on Y-tunnuksia, käytä niitä suoraan
            yt_list = extract_yts_from_text(text)

            if yt_list:
                self.set_status("Clipboard: löytyi Y-tunnukset tekstistä → haetaan sähköpostit…")
                driver = start_new_driver()
                emails = fetch_emails_from_ytj(driver, yt_list, self.set_status, self.set_progress, self.ui_log, self.stop_flag)

                if self.stop_flag.is_set():
                    self.set_status("Pysäytetty.")
                    return

                em_path = save_word_plain_lines(emails, "sahkopostit.docx")
                self.ui_log(f"Tallennettu: {em_path}")
                self.set_status("Valmis!")
                messagebox.showinfo("Valmis", f"Valmis!\n\nKansio:\n{OUT_DIR}\n\nSähköposteja: {len(emails)}")
                return

            # 2) muuten parsitaan nimet ja haetaan email nimellä (automaattinen paras osuma)
            names = extract_names_from_clipboard(text)
            if not names:
                self.set_status("Clipboard: en löytänyt Y-tunnuksia enkä yritysnimiä.")
                messagebox.showwarning("Ei löytynyt", "Tekstistä ei saatu irti Y-tunnuksia tai yritysnimiä.")
                return

            self.set_status("Käynnistetään Chrome ja haetaan sähköpostit YTJ:stä (nimihaku)…")
            driver = start_new_driver()

            emails = []
            seen = set()
            self.set_progress(0, max(1, len(names)))

            for i, name in enumerate(names, start=1):
                if self.stop_flag.is_set():
                    self.set_status("Pysäytetty.")
                    break

                self.set_status(f"YTJ nimihaku: {i}/{len(names)}  {name}")
                self.set_progress(i - 1, len(names))

                ok = ytj_find_company_and_open_first(driver, name)
                if not ok:
                    self.ui_log(f"NOT FOUND: {name}")
                    continue

                try_accept_cookies(driver)
                click_all_nayta_ytj(driver)

                email = ""
                for _ in range(YTJ_RETRY_READS):
                    if self.stop_flag.is_set():
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
                        self.ui_log(email)

                time.sleep(YTJ_PER_COMPANY_SLEEP)

            self.set_progress(len(names), max(1, len(names)))

            if self.stop_flag.is_set():
                self.set_status("Pysäytetty.")
                return

            em_path = save_word_plain_lines(emails, "sahkopostit.docx")
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


if __name__ == "__main__":
    App().mainloop()
