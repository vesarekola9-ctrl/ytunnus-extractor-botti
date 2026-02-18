import os
import re
import sys
import time
import threading
import tkinter as tk
from tkinter import messagebox, filedialog
from tkinter import ttk
from urllib.parse import urlparse
import subprocess
import shutil

import PyPDF2
from docx import Document

from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import WebDriverException
from webdriver_manager.chrome import ChromeDriverManager


# =========================
#   CONFIG / REGEX
# =========================
YT_RE = re.compile(r"\b\d{7}-\d\b|\b\d{8}\b")
EMAIL_RE = re.compile(r"[A-Za-z0-9_.+-]+@[A-Za-z0-9-]+\.[A-Za-z0-9-]+\.[A-Za-z0-9-.]+")
EMAIL_A_RE = re.compile(r"[A-Za-z0-9_.+-]+\s*\(a\)\s*[A-Za-z0-9-]+\.[A-Za-z0-9-.]+", re.I)

KAUPPALEHTI_URL = "https://www.kauppalehti.fi/yritykset/protestilista"
YTJ_COMPANY_URL = "https://tietopalvelu.ytj.fi/yritys/{}"

ALLOWED_HOSTS_DURING_LOGIN = {
    "www.kauppalehti.fi",
    "kauppalehti.fi",
    "auth.kauppalehti.fi",
    "account.kauppalehti.fi",
    "alma.fi",
    "www.alma.fi",
    "login.alma.fi",
    "auth.alma.fi",
}

ALLOWED_KL_HOSTS_DURING_SCRAPE = {"www.kauppalehti.fi", "kauppalehti.fi"}


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
CHROMEDRIVER_LOG = os.path.join(OUT_DIR, "chromedriver.log")


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
    log_to_file(f"ChromeDriver log: {CHROMEDRIVER_LOG}")


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


def pick_email_from_text(text: str) -> str:
    if not text:
        return ""
    m = re.search(r"[A-Za-z0-9_.+-]+@[A-Za-z0-9-]+\.[A-Za-z0-9-.]+", text)
    if m:
        return m.group(0).strip().replace(" ", "")
    m2 = re.search(r"[A-Za-z0-9_.+-]+\s*\(a\)\s*[A-Za-z0-9-]+\.[A-Za-z0-9-.]+", text, re.I)
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


def current_host(driver) -> str:
    try:
        return urlparse(driver.current_url).netloc.lower()
    except Exception:
        return ""


def is_external_href(href: str, allowed_hosts: set) -> bool:
    if not href:
        return False
    try:
        host = urlparse(href).netloc.lower()
        if not host:
            return False
        return host not in allowed_hosts
    except Exception:
        return False


def safe_click_with_host_policy(driver, elem, allowed_hosts: set) -> bool:
    before = driver.current_url

    try:
        tag = (elem.tag_name or "").lower()
        if tag == "a":
            href = (elem.get_attribute("href") or "").strip()
            if is_external_href(href, allowed_hosts):
                return False
    except Exception:
        pass

    ok = safe_click(driver, elem)
    if not ok:
        return False

    time.sleep(0.15)
    host = current_host(driver)
    if host and host not in allowed_hosts:
        try:
            driver.get(before)
        except Exception:
            try:
                driver.back()
            except Exception:
                pass
        return False

    return True


def try_accept_cookies(driver, allowed_hosts: set):
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
                    safe_click_with_host_policy(driver, b, allowed_hosts)
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
#   CHROME: NORMAL LOGIN PROFILE (NO SELENIUM) + SELENIUM USES SAME DIR
# =========================
def get_bot_profile_dir():
    # oma profiili, ei Default -> ei kaadu jos Chrome on auki
    base = get_exe_dir()
    prof = os.path.join(base, "kl_profile")
    try:
        os.makedirs(prof, exist_ok=True)
        return prof
    except Exception:
        home = os.path.expanduser("~")
        docs = os.path.join(home, "Documents")
        prof = os.path.join(docs, "ProtestiBotti", "kl_profile")
        os.makedirs(prof, exist_ok=True)
        return prof


def find_chrome_exe():
    # 1) PATH
    p = shutil.which("chrome") or shutil.which("chrome.exe")
    if p:
        return p

    # 2) common Windows paths
    candidates = []
    pf = os.environ.get("PROGRAMFILES", r"C:\Program Files")
    pfx = os.environ.get("PROGRAMFILES(X86)", r"C:\Program Files (x86)")
    candidates += [
        os.path.join(pf, "Google", "Chrome", "Application", "chrome.exe"),
        os.path.join(pfx, "Google", "Chrome", "Application", "chrome.exe"),
    ]
    for c in candidates:
        if os.path.exists(c):
            return c
    return None


def open_normal_chrome_for_login(profile_dir: str):
    chrome = find_chrome_exe()
    if not chrome:
        raise RuntimeError("chrome.exe ei löytynyt. Asenna Google Chrome tai lisää se PATHiin.")
    # Normaali Chrome (EI selenium) -> kirjautuminen onnistuu
    args = [
        chrome,
        f"--user-data-dir={profile_dir}",
        "--new-window",
        KAUPPALEHTI_URL,
    ]
    subprocess.Popen(args, stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)


def start_selenium_with_profile(profile_dir: str, mobile: bool):
    options = webdriver.ChromeOptions()
    options.add_argument("--disable-gpu")
    options.add_argument("--disable-dev-shm-usage")
    options.add_argument("--no-first-run")
    options.add_argument("--no-default-browser-check")
    options.add_experimental_option("excludeSwitches", ["enable-automation"])
    options.add_experimental_option("useAutomationExtension", False)

    options.add_argument(f"--user-data-dir={profile_dir}")

    if mobile:
        mobile_emulation = {
            "deviceMetrics": {"width": 412, "height": 915, "pixelRatio": 2.625},
            "userAgent": "Mozilla/5.0 (Linux; Android 13; Pixel 7) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Mobile Safari/537.36"
        }
        options.add_experimental_option("mobileEmulation", mobile_emulation)
    else:
        options.add_argument("--start-maximized")

    driver_path = ChromeDriverManager().install()
    service = Service(driver_path, log_output=CHROMEDRIVER_LOG)
    return webdriver.Chrome(service=service, options=options)


# =========================
#   KL DETECT / GUARD
# =========================
def page_looks_like_protestilista(driver) -> bool:
    try:
        body = driver.find_element(By.TAG_NAME, "body").text or ""
        return "Protestilista" in body and ("Viimeisimmät protestit" in body or "protestia" in body)
    except Exception:
        return False


def enforce_kl_guard(driver, log_cb):
    host = current_host(driver)
    if host and host not in ALLOWED_KL_HOSTS_DURING_SCRAPE:
        log_cb(f"Guard: host '{host}' -> takaisin protestilistaan")
        driver.get(KAUPPALEHTI_URL)
        WebDriverWait(driver, 25).until(EC.presence_of_element_located((By.TAG_NAME, "body")))
        try_accept_cookies(driver, ALLOWED_KL_HOSTS_DURING_SCRAPE)
        time.sleep(0.6)

    if "protestilista" not in (driver.current_url or ""):
        log_cb("Guard: ei protestilista -> takaisin")
        driver.get(KAUPPALEHTI_URL)
        WebDriverWait(driver, 25).until(EC.presence_of_element_located((By.TAG_NAME, "body")))
        try_accept_cookies(driver, ALLOWED_KL_HOSTS_DURING_SCRAPE)
        time.sleep(0.6)


def click_nayta_lisaa(driver) -> bool:
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
                return safe_click_with_host_policy(driver, b, ALLOWED_KL_HOSTS_DURING_SCRAPE)
        except Exception:
            continue
    return False


# =========================
#   KL MOBILE SCRAPE
# =========================
def extract_yt_from_text_block(text: str) -> str:
    if not text:
        return ""
    for m in YT_RE.findall(text):
        n = normalize_yt(m)
        if n:
            return n
    return ""


def get_mobile_toggles(driver):
    toggles = []
    try:
        elems = driver.find_elements(By.XPATH, "//*[@aria-expanded='false' or @aria-expanded='true']")
        for e in elems:
            try:
                if e.is_displayed() and e.is_enabled():
                    toggles.append(e)
            except Exception:
                pass
    except Exception:
        pass

    if toggles:
        return toggles

    # fallback: buttonit joissa svg (nuoli)
    try:
        buttons = driver.find_elements(By.XPATH, "//button[.//*[name()='svg']]")
        for b in buttons:
            try:
                if b.is_displayed() and b.is_enabled():
                    toggles.append(b)
            except Exception:
                pass
    except Exception:
        pass

    return toggles


def mobile_toggle_fingerprint(toggle):
    for xp in ["ancestor::*[self::div or self::li or self::article][1]", "ancestor::tr[1]"]:
        try:
            c = toggle.find_element(By.XPATH, xp)
            txt = (c.text or "").strip().replace("\n", " ")
            if txt:
                return txt[:180]
        except Exception:
            continue
    try:
        return (toggle.get_attribute("class") or "")[:120]
    except Exception:
        return "unknown"


def extract_yt_after_expand(toggle):
    for xp in ["ancestor::*[self::div or self::li or self::article][1]", "ancestor::tr[1]", "ancestor::section[1]"]:
        try:
            c = toggle.find_element(By.XPATH, xp)
            txt = (c.text or "")
            if "Y-TUNNUS" in txt or "Y-tunnus" in txt:
                yt = extract_yt_from_text_block(txt)
                if yt:
                    return yt
        except Exception:
            continue

    try:
        c = toggle.find_element(By.XPATH, "ancestor::*[self::div or self::li or self::article][1]")
        txt = (c.get_attribute("innerText") or "")
        if "Y-TUNNUS" in txt or "Y-tunnus" in txt:
            yt = extract_yt_from_text_block(txt)
            if yt:
                return yt
    except Exception:
        pass

    return ""


def collect_yts_from_kauppalehti_mobile(driver, status_cb, log_cb, stop_evt, max_rounds=250):
    driver.get(KAUPPALEHTI_URL)
    WebDriverWait(driver, 25).until(EC.presence_of_element_located((By.TAG_NAME, "body")))
    try_accept_cookies(driver, ALLOWED_HOSTS_DURING_LOGIN)
    time.sleep(0.6)

    if not page_looks_like_protestilista(driver):
        status_cb("Protestilista ei näy. Kirjaudu ensin 'Avaa Login-Chrome' napilla ja sulje se, sitten aloita keruu.")
        html, png = dump_debug(driver, "kl_not_logged_in")
        log_cb(f"DEBUG dump: {html} | {png}")
        return []

    collected = set()
    processed = set()
    rounds_no_new = 0

    for _round in range(max_rounds):
        if stop_evt.is_set():
            break

        enforce_kl_guard(driver, log_cb)
        try_accept_cookies(driver, ALLOWED_KL_HOSTS_DURING_SCRAPE)

        toggles = get_mobile_toggles(driver)
        status_cb(f"KL(mobile): nuolia {len(toggles)} | kerätty {len(collected)}")

        new_this_round = 0
        for t in toggles:
            if stop_evt.is_set():
                break

            fp = mobile_toggle_fingerprint(t)
            if fp in processed:
                continue
            processed.add(fp)

            if not safe_click_with_host_policy(driver, t, ALLOWED_KL_HOSTS_DURING_SCRAPE):
                continue

            time.sleep(0.25)

            yt = ""
            for _ in range(25):
                yt = extract_yt_after_expand(t)
                if yt:
                    break
                time.sleep(0.06)

            if yt and yt not in collected:
                collected.add(yt)
                new_this_round += 1
                log_cb(f"+ {yt} (yht {len(collected)})")

        try:
            driver.execute_script("window.scrollBy(0, 950);")
        except Exception:
            pass
        time.sleep(0.35)

        if new_this_round == 0:
            rounds_no_new += 1
        else:
            rounds_no_new = 0

        if rounds_no_new >= 6:
            if click_nayta_lisaa(driver):
                rounds_no_new = 0
                time.sleep(1.0)
                continue
            break

    if not collected:
        html, png = dump_debug(driver, "kl_mobile_zero")
        log_cb(f"DEBUG dump: {html} | {png}")

    return sorted(collected)


# =========================
#   YTJ EMAILS (PDF-mode)
# =========================
def wait_ytj_loaded(driver):
    WebDriverWait(driver, 25).until(EC.presence_of_element_located((By.TAG_NAME, "body")))
    time.sleep(0.2)


def click_all_nayta_ytj(driver):
    for _ in range(3):
        clicked = False
        for b in driver.find_elements(By.XPATH, "//button|//*[@role='button']|//a"):
            try:
                if not b.is_displayed() or not b.is_enabled():
                    continue
                t = (b.text or "").strip().lower()
                if t == "näytä":
                    safe_click(driver, b)
                    clicked = True
                    time.sleep(0.15)
            except Exception:
                continue
        if not clicked:
            break


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


def fetch_emails_from_ytj(driver, yt_list, status_cb, progress_cb, log_cb, stop_evt):
    emails = []
    seen = set()
    progress_cb(0, max(1, len(yt_list)))

    for i, yt in enumerate(yt_list, start=1):
        if stop_evt.is_set():
            break

        status_cb(f"YTJ: {i}/{len(yt_list)} {yt}")
        progress_cb(i - 1, len(yt_list))

        driver.get(YTJ_COMPANY_URL.format(yt))
        wait_ytj_loaded(driver)
        click_all_nayta_ytj(driver)

        email = ""
        for _ in range(8):
            email = extract_email_from_ytj(driver)
            if email:
                break
            time.sleep(0.2)

        if email:
            k = email.lower()
            if k not in seen:
                seen.add(k)
                emails.append(email)
                log_cb(email)

        time.sleep(0.08)

    progress_cb(len(emails), max(1, len(yt_list)))
    return emails


# =========================
#   GUI
# =========================
class App(tk.Tk):
    def __init__(self):
        super().__init__()
        reset_log()

        self.stop_evt = threading.Event()
        self.worker = None

        self.driver_kl = None
        self.profile_dir = get_bot_profile_dir()

        self.title("ProtestiBotti (KL Login-Chrome -> Selenium scrape + PDF->YTJ)")
        self.geometry("1040x720")

        tk.Label(self, text="ProtestiBotti", font=("Arial", 18, "bold")).pack(pady=8)
        tk.Label(
            self,
            text="TÄRKEÄ: Kirjautuminen tehdään 'Avaa Login-Chrome' napilla (normaali Chrome, ei Selenium).\n"
                 "Kun olet kirjautunut, sulje Login-Chrome ja käynnistä keruu.\n\n"
                 "1) KL Mobile -> Y-tunnukset Wordiin\n"
                 "2) PDF -> Y-tunnukset -> YTJ sähköpostit Wordiin\n",
            justify="center"
        ).pack(pady=4)

        info = tk.Label(self, text=f"Botin KL-profiili: {self.profile_dir}", wraplength=980, justify="center")
        info.pack(pady=2)

        btn_row = tk.Frame(self)
        btn_row.pack(pady=10)

        self.btn_login = tk.Button(btn_row, text="Avaa Login-Chrome (normaali)", font=("Arial", 12, "bold"), command=self.open_login_chrome)
        self.btn_login.grid(row=0, column=0, padx=8)

        self.btn_kl_open = tk.Button(btn_row, text="1) Avaa KL (Selenium mobiili)", font=("Arial", 12, "bold"), command=self.start_kl_open)
        self.btn_kl_open.grid(row=0, column=1, padx=8)

        self.btn_kl_start = tk.Button(btn_row, text="ALOITA KERUU", font=("Arial", 12, "bold"), command=self.start_kl_scrape, state="disabled")
        self.btn_kl_start.grid(row=0, column=2, padx=8)

        self.btn_pdf = tk.Button(btn_row, text="2) PDF → YTJ sähköpostit", font=("Arial", 12, "bold"), command=self.start_pdf)
        self.btn_pdf.grid(row=0, column=3, padx=8)

        self.btn_stop = tk.Button(btn_row, text="STOP", font=("Arial", 12, "bold"), command=self.stop, state="disabled")
        self.btn_stop.grid(row=0, column=4, padx=8)

        self.status = tk.Label(self, text="Valmiina.", font=("Arial", 11))
        self.status.pack(pady=6)

        self.progress = ttk.Progressbar(self, orient="horizontal", mode="determinate", length=980)
        self.progress.pack(pady=6)

        frame = tk.Frame(self)
        frame.pack(fill="both", expand=True, padx=14, pady=10)

        tk.Label(frame, text="Live-logi (uusimmat alimmaisena):").pack(anchor="w")
        self.listbox = tk.Listbox(frame, height=22)
        self.listbox.pack(side="left", fill="both", expand=True)

        sb = tk.Scrollbar(frame, orient="vertical", command=self.listbox.yview)
        sb.pack(side="right", fill="y")
        self.listbox.configure(yscrollcommand=sb.set)

        tk.Label(self, text=f"Tallennus: {OUT_DIR}", wraplength=1020, justify="center").pack(pady=6)

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

    def lock_ui(self):
        self.stop_evt.clear()
        self.btn_stop.config(state="normal")
        self.btn_pdf.config(state="disabled")
        self.btn_kl_open.config(state="disabled")
        self.btn_login.config(state="disabled")

    def unlock_ui(self):
        self.btn_stop.config(state="disabled")
        self.btn_pdf.config(state="normal")
        self.btn_kl_open.config(state="normal")
        self.btn_login.config(state="normal")

    def stop(self):
        self.stop_evt.set()
        self.set_status("STOP pyydetty – lopetetaan hallitusti…")
        self.btn_stop.config(state="disabled")

    def open_login_chrome(self):
        try:
            self.set_status("Avaan normaali-Chromen botin profiililla. Kirjaudu KL:ään ja SULJE se Chrome sen jälkeen.")
            open_normal_chrome_for_login(self.profile_dir)
            messagebox.showinfo(
                "Kirjaudu Kauppalehteen",
                "Normaali Chrome avattiin botin profiililla.\n\n"
                "1) Kirjaudu Kauppalehteen\n"
                "2) Varmista että protestilista näkyy\n"
                "3) SULJE se Chrome kokonaan\n"
                "4) Palaa bottiin ja paina: '1) Avaa KL (Selenium mobiili)'"
            )
        except Exception as e:
            self.ui_log(f"VIRHE: {e}")
            messagebox.showerror("Virhe", str(e))

    # ---- MODE 1: KL Mobile ----
    def start_kl_open(self):
        if self.worker and self.worker.is_alive():
            return
        self.lock_ui()
        self.btn_kl_start.config(state="disabled")
        self.worker = threading.Thread(target=self.run_kl_open, daemon=True)
        self.worker.start()

    def run_kl_open(self):
        try:
            self.set_status("Käynnistän Selenium-Chromen mobiiliemulaatiolla (botin oma profiili)…")
            self.driver_kl = start_selenium_with_profile(self.profile_dir, mobile=True)
            self.set_status("Selenium Chrome auki. Paina nyt: ALOITA KERUU")
            self.btn_kl_start.config(state="normal")
        except Exception as e:
            self.ui_log(f"VIRHE: {e}")
            self.set_status("Virhe. Katso log.txt ja chromedriver.log")
            messagebox.showerror("Virhe", f"{e}\n\nKatso:\n{LOG_PATH}\n{CHROMEDRIVER_LOG}")
            try:
                if self.driver_kl:
                    self.driver_kl.quit()
            except Exception:
                pass
            self.driver_kl = None
            self.unlock_ui()

    def start_kl_scrape(self):
        if not self.driver_kl:
            messagebox.showwarning("Puuttuu", "Avaa ensin KL (Selenium).")
            return
        self.btn_kl_start.config(state="disabled")
        threading.Thread(target=self.run_kl_scrape, daemon=True).start()

    def run_kl_scrape(self):
        try:
            self.set_status("KL(mobile): kerätään Y-tunnukset…")
            yt_list = collect_yts_from_kauppalehti_mobile(self.driver_kl, self.set_status, self.ui_log, self.stop_evt)

            if self.stop_evt.is_set():
                self.set_status("Pysäytetty.")
                return

            if not yt_list:
                self.set_status("Ei löytynyt Y-tunnuksia. Katso output-kansion debug-kuvat.")
                messagebox.showwarning("Ei löytynyt", f"Y-tunnuksia ei saatu.\nKatso: {OUT_DIR}")
                return

            yt_path = save_word_plain_lines(yt_list, "ytunnukset_kl_mobile.docx")
            self.ui_log(f"Tallennettu: {yt_path}")

            self.set_status("Valmis!")
            messagebox.showinfo("Valmis", f"Valmis!\n\nY-tunnuksia: {len(yt_list)}\n\n{yt_path}")

        except Exception as e:
            self.ui_log(f"VIRHE: {e}")
            self.set_status("Virhe. Katso log.txt ja chromedriver.log")
            messagebox.showerror("Virhe", f"{e}\n\nKatso:\n{LOG_PATH}\n{CHROMEDRIVER_LOG}")
        finally:
            self.unlock_ui()

    # ---- MODE 2: PDF -> YTJ ----
    def start_pdf(self):
        path = filedialog.askopenfilename(filetypes=[("PDF files", "*.pdf")])
        if not path:
            return
        if self.worker and self.worker.is_alive():
            return
        self.lock_ui()
        self.worker = threading.Thread(target=self.run_pdf_mode, args=(path,), daemon=True)
        self.worker.start()

    def run_pdf_mode(self, pdf_path):
        driver = None
        try:
            self.set_status("Luetaan PDF ja kerätään Y-tunnukset…")
            yt_list = extract_ytunnukset_from_pdf(pdf_path)

            if not yt_list:
                self.set_status("Ei löytynyt Y-tunnuksia PDF:stä.")
                messagebox.showwarning("Ei löytynyt", "PDF:stä ei löytynyt yhtään Y-tunnusta.")
                return

            yt_path = save_word_plain_lines(yt_list, "ytunnukset_pdf.docx")
            self.ui_log(f"Tallennettu: {yt_path}")

            self.set_status("Käynnistetään Chrome ja haetaan sähköpostit YTJ:stä…")
            driver = start_selenium_with_profile(self.profile_dir, mobile=False)

            self.set_status("YTJ: haetaan sähköpostit…")
            emails = fetch_emails_from_ytj(driver, yt_list, self.set_status, self.set_progress, self.ui_log, self.stop_evt)

            em_path = save_word_plain_lines(emails, "sahkopostit_pdf_ytj.docx")
            self.ui_log(f"Tallennettu: {em_path}")

            self.set_status("Valmis!")
            messagebox.showinfo(
                "Valmis",
                f"Valmis!\n\nY-tunnuksia: {len(yt_list)}\nSähköposteja: {len(emails)}\n\n{yt_path}\n{em_path}"
            )

        except Exception as e:
            self.ui_log(f"VIRHE: {e}")
            self.set_status("Virhe. Katso log.txt ja chromedriver.log")
            messagebox.showerror("Virhe", f"{e}\n\nKatso:\n{LOG_PATH}\n{CHROMEDRIVER_LOG}")
        finally:
            try:
                if driver:
                    driver.quit()
            except Exception:
                pass
            self.unlock_ui()


if __name__ == "__main__":
    App().mainloop()
