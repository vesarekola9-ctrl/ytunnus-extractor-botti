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
#   STOP / SLEEP
# =========================
def safe_sleep(stop_event: threading.Event, seconds: float, step: float = 0.05):
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
#   SELENIUM START (FAST)
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
        "profile.managed_default_content_settings.images": 2,
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
#   YTJ (FAST + Näytä)
# =========================
def wait_ytj_loaded_fast(driver):
    WebDriverWait(driver, 20).until(EC.presence_of_element_located((By.TAG_NAME, "body")))
    try:
        WebDriverWait(driver, 12).until(
            EC.presence_of_element_located(
                (By.XPATH, "//*[contains(normalize-space(.), 'Sähköposti') or contains(normalize-space(.), 'Y-tunnus') or contains(normalize-space(.), 'Toiminimi')]")
            )
        )
    except Exception:
        pass


def _find_label_blocks(driver, label_text: str):
    label = label_text.strip()
    blocks = []
    for xp in [
        f"//tr[.//*[contains(normalize-space(.), '{label}')]]",
        f"//*[self::div or self::section or self::li][.//*[contains(normalize-space(.), '{label}')]]",
    ]:
        try:
            blocks.extend(driver.find_elements(By.XPATH, xp))
        except Exception:
            pass

    out = []
    for b in blocks:
        try:
            if b.is_displayed():
                out.append(b)
        except Exception:
            pass
    return out


def click_show_for_labels(driver, labels=("Sähköposti",)):
    for _round in range(3):
        clicked_any = False
        for lab in labels:
            blocks = _find_label_blocks(driver, lab)
            for bl in blocks:
                try:
                    btns = bl.find_elements(
                        By.XPATH,
                        ".//button[normalize-space(.)='Näytä' or normalize-space(.)='näytä'] | "
                        ".//a[normalize-space(.)='Näytä' or normalize-space(.)='näytä'] | "
                        ".//*[@role='button' and (normalize-space(.)='Näytä' or normalize-space(.)='näytä')]"
                    )
                    for b in btns:
                        try:
                            if b.is_displayed() and b.is_enabled():
                                safe_click(driver, b)
                                clicked_any = True
                                time.sleep(0.12)
                        except Exception:
                            continue
                except Exception:
                    continue
        if not clicked_any:
            break


def extract_email_from_ytj_fast(driver) -> str:
    try:
        mail = driver.find_elements(By.XPATH, "//a[starts-with(translate(@href,'MAILTO','mailto'),'mailto:')]")
        if mail:
            href = mail[0].get_attribute("href") or ""
            e = href.split(":", 1)[1].strip()
            if e:
                return e
    except Exception:
        pass

    try:
        blocks = _find_label_blocks(driver, "Sähköposti")
        for b in blocks:
            txt = (b.text or "")
            if "@" in txt:
                e = pick_email_from_text(txt)
                if e:
                    return e
    except Exception:
        pass

    try:
        return pick_email_from_text(driver.find_element(By.TAG_NAME, "body").text or "")
    except Exception:
        return ""


def extract_yt_from_ytj_company_page(driver) -> str:
    # Use label "Y-tunnus" area or fallback regex from page text
    try:
        blocks = _find_label_blocks(driver, "Y-tunnus")
        for b in blocks:
            t = b.text or ""
            m = YT_RE.search(t)
            if m:
                n = normalize_yt(m.group(0))
                if n:
                    return n
    except Exception:
        pass

    try:
        body = driver.find_element(By.TAG_NAME, "body").text or ""
        m = YT_RE.search(body)
        if m:
            n = normalize_yt(m.group(0))
            if n:
                return n
    except Exception:
        pass

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

        status_cb(f"YTJ (Y-tunnus): {i}/{len(yt_list)} {yt}")
        progress_cb(i - 1, total)

        driver.get(YTJ_COMPANY_URL.format(yt))
        wait_ytj_loaded_fast(driver)
        try_accept_cookies(driver)

        email = ""
        for _ in range(8):
            if should_stop(stop_event):
                break
            click_show_for_labels(driver, labels=("Sähköposti",))
            email = extract_email_from_ytj_fast(driver)
            if email:
                break
            safe_sleep(stop_event, 0.15, step=0.05)

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
#   YTJ SEARCH BY NAME + LOCATION
# =========================
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


def ytj_search_company_url_by_name_and_location(driver, company_name: str, location_hint: str = "") -> str:
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

    results = []
    try:
        links = driver.find_elements(By.XPATH, "//a[contains(@href,'/yritys/')]")
    except Exception:
        links = []

    for a in links:
        try:
            if not a.is_displayed():
                continue
            href = (a.get_attribute("href") or "").strip()
            txt = (a.text or "").strip()
            if "/yritys/" not in href:
                continue
            around = ""
            try:
                parent = a.find_element(By.XPATH, "ancestor::*[self::li or self::div or self::tr][1]")
                around = (parent.text or "")
            except Exception:
                around = (txt or "")
            results.append((txt, href, around))
        except Exception:
            continue

    if not results:
        return ""

    low = name.lower().strip()
    loc = (location_hint or "").strip().lower()

    if loc:
        for txt, href, around in results:
            if txt and txt.strip().lower() == low and loc in (around or "").lower():
                return href

    for txt, href, around in results:
        if txt and txt.strip().lower() == low:
            return href

    if loc:
        for txt, href, around in results:
            if txt and low in txt.strip().lower() and loc in (around or "").lower():
                return href

    for txt, href, around in results:
        if txt and low in txt.strip().lower():
            return href

    if loc:
        for txt, href, around in results:
            if loc in (around or "").lower():
                return href

    return results[0][1]


# =========================
#   KAUPPALEHTI: LOAD ALL -> EXTRACT (NAME, LOCATION)
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
    # Rows in tbody; detail rows include "Y-TUNNUS"
    rows = []
    candidates = driver.find_elements(By.XPATH, "//table//tbody//tr")
    for r in candidates:
        try:
            if not r.is_displayed():
                continue
            txt = (r.text or "")
            if "Y-TUNNUS" in txt:
                continue
            # first td is company link; we only read it, never click it
            links = r.find_elements(By.XPATH, ".//td[1]//a[contains(@href,'/yritykset/') and normalize-space(.)!='']")
            if not links:
                continue
            rows.append(r)
        except Exception:
            continue
    return rows


def load_all_kauppalehti_entries(driver, stop_event, status_cb, log_cb, max_clicks=9999):
    # Scroll -> click "Näytä lisää" until gone
    clicks = 0
    last_count = -1
    stable_rounds = 0

    while True:
        if should_stop(stop_event):
            return

        rows = get_company_rows_table(driver)
        cnt = len(rows)

        status_cb(f"KL: ladattu rivejä {cnt} | Näytä lisää klikkejä {clicks}")
        if cnt == last_count:
            stable_rounds += 1
        else:
            stable_rounds = 0
        last_count = cnt

        # stop if stable and no button
        if stable_rounds >= 2:
            # try one final click if exists
            if not click_nayta_lisaa(driver):
                break

        if clicks >= max_clicks:
            break

        if click_nayta_lisaa(driver):
            clicks += 1
            # wait for new rows
            try:
                WebDriverWait(driver, 15).until(lambda d: len(get_company_rows_table(d)) > cnt)
            except Exception:
                safe_sleep(stop_event, 1.2)
            continue
        else:
            break

    status_cb(f"KL: lataus valmis. rivejä {len(get_company_rows_table(driver))}")


def extract_pairs_from_kauppalehti_table(driver, stop_event, status_cb, log_cb):
    rows = get_company_rows_table(driver)
    pairs = []
    seen = set()

    for i, r in enumerate(rows, start=1):
        if should_stop(stop_event):
            break
        try:
            name = r.find_element(By.XPATH, ".//td[1]//a").text.strip()
        except Exception:
            continue
        try:
            loc = r.find_element(By.XPATH, ".//td[2]").text.strip()
        except Exception:
            loc = ""

        key = (name.lower(), (loc or "").lower())
        if key in seen:
            # keep duplicates out of pair list; YTJ nimihaku on hitaampi.
            # jos haluat myös duplikaatit, poista tämä if.
            continue
        seen.add(key)
        pairs.append((name, loc))

        if i % 200 == 0:
            status_cb(f"KL: poimittu nimiä {len(pairs)} / rivejä {len(rows)}")

    status_cb(f"KL: poiminta valmis. uniikkeja nimi+paikka pareja: {len(pairs)}")
    return pairs


# =========================
#   KL -> (NAME,LOC) -> YTJ URL -> YT -> EMAIL
# =========================
def fetch_yts_from_pairs_via_ytj(driver, stop_event, pairs, status_cb, progress_cb, log_cb):
    yt_list = []
    seen = set()
    total = max(1, len(pairs))
    progress_cb(0, total)

    for i, (nm, loc) in enumerate(pairs, start=1):
        if should_stop(stop_event):
            status_cb("STOP: YTJ nimihaku (Y-tunnukset) keskeytetty.")
            break

        show_loc = f" ({loc})" if loc else ""
        status_cb(f"YTJ nimihaku (Y-tunnus): {i}/{len(pairs)} {nm}{show_loc}")
        progress_cb(i - 1, total)

        url = ytj_search_company_url_by_name_and_location(driver, nm, loc)
        if not url:
            log_cb(f"NO MATCH: {nm}{show_loc}")
            continue

        driver.get(url)
        wait_ytj_loaded_fast(driver)
        try_accept_cookies(driver)

        yt = extract_yt_from_ytj_company_page(driver)
        if yt:
            k = yt
            if k not in seen:
                seen.add(k)
                yt_list.append(k)
                log_cb(f"YT: {yt} <- {nm}{show_loc}")
        else:
            log_cb(f"NO YT: {nm}{show_loc}")

        safe_sleep(stop_event, 0.03, step=0.03)

    progress_cb(min(len(pairs), total), total)
    return yt_list


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
#   SCROLLABLE ROOT FRAME
# =========================
class ScrollableFrame(ttk.Frame):
    def __init__(self, container, *args, **kwargs):
        super().__init__(container, *args, **kwargs)

        self.canvas = tk.Canvas(self, highlightthickness=0)
        self.vsb = ttk.Scrollbar(self, orient="vertical", command=self.canvas.yview)
        self.inner = ttk.Frame(self.canvas)

        self.inner.bind("<Configure>", lambda e: self.canvas.configure(scrollregion=self.canvas.bbox("all")))
        self.window_id = self.canvas.create_window((0, 0), window=self.inner, anchor="nw")

        self.canvas.configure(yscrollcommand=self.vsb.set)
        self.canvas.pack(side="left", fill="both", expand=True)
        self.vsb.pack(side="right", fill="y")

        self.canvas.bind("<Configure>", self._on_canvas_configure)
        self.canvas.bind_all("<MouseWheel>", self._on_mousewheel)

    def _on_canvas_configure(self, event):
        try:
            self.canvas.itemconfigure(self.window_id, width=event.width)
        except Exception:
            pass

    def _on_mousewheel(self, event):
        try:
            self.canvas.yview_scroll(int(-1 * (event.delta / 120)), "units")
        except Exception:
            pass
        return "break"


# =========================
#   GUI
# =========================
class App(tk.Tk):
    def __init__(self):
        super().__init__()
        reset_log()

        self.title("ProtestiBotti (KL: lataa kaikki -> nimet -> YTJ -> YT -> email)")
        self.geometry("1120x860")

        self.stop_event = threading.Event()
        self.worker_thread = None
        self.running_driver = None
        self.locked_handle = None

        # Hotkey STOP
        self.bind_all("<Control-Shift-KeyPress-Q>", lambda e: self.emergency_stop())

        root = ScrollableFrame(self)
        root.pack(fill="both", expand=True)
        w = root.inner

        tk.Label(w, text="ProtestiBotti", font=("Arial", 18, "bold")).pack(pady=10)

        btn_row = tk.Frame(w)
        btn_row.pack(pady=6)

        tk.Button(btn_row, text="Avaa Chrome-botti (9222)", font=("Arial", 12), command=self.open_chrome_bot).grid(row=0, column=0, padx=6)
        tk.Button(btn_row, text="Kauppalehti → YTJ (uusi malli)", font=("Arial", 12), command=self.start_kauppalehti_mode).grid(row=0, column=1, padx=6)
        tk.Button(btn_row, text="PDF → YTJ", font=("Arial", 12), command=self.start_pdf_mode).grid(row=0, column=2, padx=6)

        tk.Button(
            btn_row,
            text="🛑 STOP (Ctrl+Shift+Q)",
            font=("Arial", 12, "bold"),
            fg="white",
            bg="#B00020",
            activebackground="#8C0019",
            command=self.emergency_stop
        ).grid(row=0, column=3, padx=10)

        self.status = tk.Label(w, text="Valmiina.", font=("Arial", 11))
        self.status.pack(pady=6)

        self.progress = ttk.Progressbar(w, orient="horizontal", mode="determinate", length=1060)
        self.progress.pack(pady=6)

        tk.Label(w, text=f"Tallennus: {OUT_DIR}", wraplength=1080, justify="center").pack(pady=6)

        logbox = tk.LabelFrame(w, text="Live-logi", padx=8, pady=8)
        logbox.pack(fill="both", expand=True, padx=12, pady=10)

        self.listbox = tk.Listbox(logbox, height=18)
        self.listbox.pack(side="left", fill="both", expand=True)

        sb = tk.Scrollbar(logbox, orient="vertical", command=self.listbox.yview)
        sb.pack(side="right", fill="y")
        self.listbox.configure(yscrollcommand=sb.set)

    # ---------- UI helpers ----------
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

    def _start_worker(self, target, args=()):
        if self.worker_thread and self.worker_thread.is_alive():
            messagebox.showwarning("Käynnissä", "Botti on jo käynnissä. Paina STOP jos haluat keskeyttää.")
            return
        self.stop_event.clear()
        self.worker_thread = threading.Thread(target=target, args=args, daemon=True)
        self.worker_thread.start()

    # ---------- STOP ----------
    def emergency_stop(self):
        self.stop_event.set()
        self.set_status("STOP pyydetty…")
        if self.running_driver is not None:
            try:
                self.running_driver.quit()
            except Exception:
                pass
            self.running_driver = None
        try:
            messagebox.showinfo("STOP", "Botti keskeytetty.\nVoit käynnistää uudestaan.")
        except Exception:
            pass

    # ---------- Chrome ----------
    def open_chrome_bot(self):
        ok, msg = launch_chrome_bot_9222()
        self.ui_log(msg)
        if ok:
            messagebox.showinfo("Chrome-botti", msg + "\nKirjaudu Kauppalehteen tässä ikkunassa.")
        else:
            messagebox.showerror("Chrome-botti", msg)

    # ---------- Mode 1: KL -> names -> YTJ -> YT -> email ----------
    def start_kauppalehti_mode(self):
        self._start_worker(self.run_kauppalehti_mode)

    def run_kauppalehti_mode(self):
        try:
            self.set_status("Liitytään Chrome-bottiin (9222)…")
            driver = attach_to_existing_chrome()

            self.set_status("Varmistetaan protestilista ja kirjautuminen…")
            if not ensure_protestilista_open_and_ready(driver, self.stop_event, self.set_status, self.ui_log):
                return

            if should_stop(self.stop_event):
                return

            self.set_status("KL: ladataan KAIKKI rivit (Näytä lisää loop)…")
            load_all_kauppalehti_entries(driver, self.stop_event, self.set_status, self.ui_log)

            if should_stop(self.stop_event):
                return

            self.set_status("KL: poimitaan yritysnimet + paikkakunnat…")
            pairs = extract_pairs_from_kauppalehti_table(driver, self.stop_event, self.set_status, self.ui_log)
            if not pairs:
                messagebox.showwarning("Ei löytynyt", "En saanut yritysnimiä/paikkakuntia KL:stä.")
                return

            # YTJ nimihaku tehdään uudessa tabissa (nopeampi / selkeämpi)
            self.set_status("Avataan YTJ uuteen tabiin (nimihaku)…")
            open_new_tab(driver, "about:blank")

            self.set_status("YTJ: haetaan Y-tunnukset nimillä + paikkakunnalla…")
            yts = fetch_yts_from_pairs_via_ytj(driver, self.stop_event, pairs, self.set_status, self.set_progress, self.ui_log)
            if should_stop(self.stop_event):
                return

            if not yts:
                messagebox.showwarning("Ei Y-tunnuksia", "Nimihaku ei palauttanut Y-tunnuksia.")
                return

            yt_path = save_word_plain_lines(yts, "ytunnukset.docx")
            self.ui_log(f"Tallennettu: {yt_path}")

            self.set_status("YTJ: haetaan sähköpostit Y-tunnuksilla…")
            emails = fetch_emails_from_ytj_by_yt_fast(driver, self.stop_event, yts, self.set_status, self.set_progress, self.ui_log)
            if should_stop(self.stop_event):
                return

            em_path = save_word_plain_lines(emails, "sahkopostit.docx")
            self.ui_log(f"Tallennettu: {em_path}")

            # Bonus: yhdistetty raportti
            report_lines = []
            # (tässä ei yritetä täydellistä mapitusta duplikaatteihin, mutta saat email-listan varmasti)
            report_lines.append(f"KL poimitut uniikit nimi+paikka: {len(pairs)}")
            report_lines.append(f"YTJ löydetyt uniikit Y-tunnukset: {len(yts)}")
            report_lines.append(f"YTJ löydetyt sähköpostit: {len(emails)}")
            rep_path = save_word_plain_lines(report_lines, "raportti.docx")
            self.ui_log(f"Tallennettu: {rep_path}")

            self.set_status("Valmis!")
            messagebox.showinfo(
                "Valmis",
                f"Valmis!\n\n"
                f"Nimi+paikka pareja: {len(pairs)}\n"
                f"Y-tunnuksia: {len(yts)}\n"
                f"Sähköposteja: {len(emails)}\n\n"
                f"Kansio:\n{OUT_DIR}"
            )

        except Exception as e:
            self.ui_log(f"VIRHE: {e}")
            messagebox.showerror("Virhe", f"Tuli virhe.\nKatso log.txt:\n{LOG_PATH}\n\n{e}")

    # ---------- Mode 2: PDF -> YTJ ----------
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
            messagebox.showinfo("Valmis", f"Valmis!\nSähköposteja: {len(emails)}\nKansio:\n{OUT_DIR}")

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
