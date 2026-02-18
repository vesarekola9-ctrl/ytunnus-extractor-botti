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
from selenium.common.exceptions import StaleElementReferenceException, WebDriverException
from webdriver_manager.chrome import ChromeDriverManager


# =========================
#   CONFIG / REGEX
# =========================
YT_RE = re.compile(r"\b\d{7}-\d\b|\b\d{8}\b")
EMAIL_RE = re.compile(r"[A-Za-z0-9_.+-]+@[A-Za-z0-9-]+\.[A-Za-z0-9-.]+")
EMAIL_A_RE = re.compile(r"[A-Za-z0-9_.+-]+\s*\(a\)\s*[A-Za-z0-9-]+\.[A-Za-z0-9-.]+", re.I)

KAUPPALEHTI_URL = "https://www.kauppalehti.fi/yritykset/protestilista"
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
    """Tallenna HTML + screenshot output-kansioon jotta nähdään mitä Selenium oikeasti näkee."""
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
    return webdriver.Chrome(service=Service(driver_path), options=options)


def attach_to_existing_chrome():
    options = webdriver.ChromeOptions()
    options.add_experimental_option("debuggerAddress", "127.0.0.1:9222")
    driver_path = ChromeDriverManager().install()
    return webdriver.Chrome(service=Service(driver_path), options=options)


def list_tabs(driver):
    tabs = []
    for h in driver.window_handles:
        try:
            driver.switch_to.window(h)
            tabs.append((driver.title or "", driver.current_url or ""))
        except Exception:
            tabs.append(("", ""))
    return tabs


def focus_protestilista_tab(driver, log_cb=None) -> bool:
    """Löydä tab missä URL sisältää kauppalehti.fi + protestilista."""
    target_handle = None
    for handle in driver.window_handles:
        try:
            driver.switch_to.window(handle)
            url = (driver.current_url or "")
            if "kauppalehti.fi" in url and "protestilista" in url:
                target_handle = handle
                break
        except Exception:
            continue

    if log_cb:
        log_cb("Chrome TAB LISTA (title | url):")
        for title, url in list_tabs(driver):
            log_cb(f"  {title} | {url}")

    if target_handle:
        driver.switch_to.window(target_handle)
        return True
    return False


def open_new_tab(driver, url="about:blank"):
    driver.execute_script("window.open(arguments[0], '_blank');", url)
    driver.switch_to.window(driver.window_handles[-1])


# =========================
#   KAUPPALEHTI PAGE CHECKS
# =========================
def page_looks_like_protestilista(driver) -> bool:
    """Heikko varmistus: sivulla näkyy 'Protestilista'."""
    try:
        body = (driver.find_element(By.TAG_NAME, "body").text or "")
        return "Protestilista" in body
    except Exception:
        return False


def page_looks_like_login_or_paywall(driver) -> bool:
    try:
        text = (driver.find_element(By.TAG_NAME, "body").text or "").lower()
        bad_words = [
            "kirjaudu", "tilaa", "tilaajille", "vahvista henkilöllisyytesi",
            "sign in", "subscribe", "login",
            "pääsy evätty", "access denied",
            "jokin meni pieleen", "something went wrong",
        ]
        return any(w in text for w in bad_words)
    except Exception:
        return False


def ensure_protestilista_open_and_ready(driver, status_cb, log_cb, max_wait_seconds=900) -> bool:
    # 1) löytyykö tab?
    if focus_protestilista_tab(driver, log_cb):
        status_cb("Löytyi protestilista-tab.")
    else:
        status_cb("Protestilista-tab ei löytynyt -> avaan protestilistan uuteen tabiin…")
        log_cb("AUTOFIX: opening protestilista in new tab")
        open_new_tab(driver, KAUPPALEHTI_URL)
        WebDriverWait(driver, 25).until(EC.presence_of_element_located((By.TAG_NAME, "body")))
        try_accept_cookies(driver)

    # 2) Odota että lista oikeasti näkyy (ja ei ole paywall)
    start = time.time()
    warned = False

    while True:
        try:
            try_accept_cookies(driver)
        except Exception:
            pass

        if page_looks_like_protestilista(driver) and not page_looks_like_login_or_paywall(driver):
            status_cb("Protestilista auki.")
            return True

        if page_looks_like_login_or_paywall(driver) and not warned:
            warned = True
            status_cb("Kauppalehti vaatii kirjautumisen/tilaajanäkymän. Kirjaudu nyt Chrome-bottiin (9222).")
            log_cb("AUTOFIX: waiting for user to login / unlock paywall…")
            try:
                messagebox.showinfo(
                    "Kirjaudu Kauppalehteen",
                    "Botti avasi protestilistan.\n\n"
                    "Kirjaudu nyt Kauppalehteen AUKI OLEVAAN Chrome-bottiin (9222).\n"
                    "Kun protestilista näkyy, botti jatkaa automaattisesti."
                )
            except Exception:
                pass

        if time.time() - start > max_wait_seconds:
            status_cb("Aikakatkaisu: protestilista ei tullut näkyviin. Tarkista kirjautuminen.")
            log_cb("ERROR: timeout waiting protestilista")
            html, png = dump_debug(driver, "kl_timeout")
            log_cb(f"DEBUG: {html} | {png}")
            return False

        time.sleep(2)


# =========================
#   KAUPPALEHTI SCRAPE (ROBUST)
# =========================
def click_nayta_lisaa(driver) -> bool:
    """Robusti 'Näytä lisää' / 'Lataa lisää'."""
    try:
        driver.execute_script("window.scrollTo(0, document.body.scrollHeight);")
    except Exception:
        pass
    time.sleep(0.3)

    for b in driver.find_elements(By.XPATH, "//button|//*[@role='button']|//a"):
        try:
            if not b.is_displayed() or not b.is_enabled():
                continue
            txt = (b.text or "").strip().lower()
            if ("näytä" in txt and "lisää" in txt) or ("lataa" in txt and "lisää" in txt):
                return safe_click(driver, b)
        except Exception:
            continue
    return False


def get_protestilista_row_toggles(driver):
    """
    Etsi toggle-nuolia vain listan riveistä.
    Ensisijaisesti table/tbody -> aria-expanded.
    """
    toggles = []

    # 1) table/tbody
    xpath = (
        "//table//tbody"
        "//*[(@aria-expanded='false' or @aria-expanded='true') and "
        "(self::button or self::a or @role='button')]"
    )
    elems = driver.find_elements(By.XPATH, xpath)
    for e in elems:
        try:
            if e.is_displayed() and e.is_enabled():
                toggles.append(e)
        except Exception:
            continue

    # 2) fallback (div-lista tms)
    if not toggles:
        elems = driver.find_elements(By.XPATH, "//*[@aria-expanded='false' or @aria-expanded='true']")
        for e in elems:
            try:
                if not (e.is_displayed() and e.is_enabled()):
                    continue
                # pitää olla rivikontekstissa
                e.find_element(By.XPATH, "ancestor::*[self::tr or @role='row'][1]")
                toggles.append(e)
            except Exception:
                continue

    return toggles


def toggle_fingerprint(driver, toggle):
    """Fingerprint togglelle sen rivitekstistä, jotta ei klikata samaa uudelleen."""
    try:
        container = toggle.find_element(By.XPATH, "ancestor::tr[1]")
        txt = (container.text or "").strip().replace("\n", " ")
        return txt[:200]
    except Exception:
        pass

    try:
        container = toggle.find_element(By.XPATH, "ancestor::*[@role='row'][1]")
        txt = (container.text or "").strip().replace("\n", " ")
        return txt[:200]
    except Exception:
        pass

    try:
        return (toggle.get_attribute("aria-expanded") or "") + "|" + (toggle.get_attribute("class") or "")[:120]
    except Exception:
        return "unknown"


def wait_toggle_state(driver, toggle, want_true: bool, timeout=3.0):
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


def extract_yt_from_opened_detail(toggle):
    """
    Toggle on rivissä. Detail aukeaa yleensä seuraavaksi <tr>:ksi.
    Etsitään 'Y-tunnus' ja poimitaan regexillä.
    """
    # 1) seuraavat sibling-tr:t
    try:
        tr = toggle.find_element(By.XPATH, "ancestor::tr[1]")
        for k in range(1, 7):
            try:
                sib = tr.find_element(By.XPATH, f"following-sibling::tr[{k}]")
                txt = (sib.text or "")
                if "Y-tunnus" in txt or "Y-TUNNUS" in txt:
                    for m in YT_RE.findall(txt):
                        n = normalize_yt(m)
                        if n:
                            return n
            except Exception:
                continue
    except Exception:
        pass

    # 2) fallback: lähikontaineri
    try:
        container = toggle.find_element(By.XPATH, "ancestor::*[self::tr or @role='row' or self::div][1]")
        txt = container.text or ""
        if "Y-tunnus" in txt or "Y-TUNNUS" in txt:
            for m in YT_RE.findall(txt):
                n = normalize_yt(m)
                if n:
                    return n
    except Exception:
        pass

    return ""


def collect_yts_from_kauppalehti(driver, status_cb, log_cb):
    WebDriverWait(driver, 25).until(EC.presence_of_element_located((By.TAG_NAME, "body")))
    try_accept_cookies(driver)

    if not ensure_protestilista_open_and_ready(driver, status_cb, log_cb, max_wait_seconds=900):
        return []

    collected = set()
    processed = set()
    rounds_without_new = 0

    time.sleep(1.0)

    while True:
        try:
            try_accept_cookies(driver)
        except Exception:
            pass

        # pysy oikealla sivulla
        cur = driver.current_url or ""
        if "kauppalehti.fi" not in cur or "protestilista" not in cur:
            status_cb("KL: ei olla protestilistassa -> palaan/avaan uudelleen…")
            focus_protestilista_tab(driver, log_cb)
            cur = driver.current_url or ""
            if "protestilista" not in cur:
                driver.get(KAUPPALEHTI_URL)
                WebDriverWait(driver, 25).until(EC.presence_of_element_located((By.TAG_NAME, "body")))
                try_accept_cookies(driver)

        if page_looks_like_login_or_paywall(driver):
            status_cb("KL: näyttää paywallilta/virheeltä -> teen debug-dumpin ja odotan.")
            html, png = dump_debug(driver, "kl_blocked")
            log_cb(f"DEBUG: {html} | {png}")
            time.sleep(3)
            continue

        toggles = get_protestilista_row_toggles(driver)
        if not toggles:
            status_cb("KL: toggleja ei löydy -> teen debug-dumpin.")
            html, png = dump_debug(driver, "kl_no_toggles")
            log_cb(f"DEBUG: {html} | {png}")
            break

        status_cb(f"KL: toggleja {len(toggles)} näkyvissä | kerätty {len(collected)}")
        new_in_pass = 0

        for idx in range(len(toggles)):
            try:
                toggles = get_protestilista_row_toggles(driver)
                if idx >= len(toggles):
                    break
                t = toggles[idx]

                fp = toggle_fingerprint(driver, t)
                if fp in processed:
                    continue
                processed.add(fp)

                if not safe_click(driver, t):
                    continue

                wait_toggle_state(driver, t, want_true=True, timeout=3.0)

                yt = ""
                for _ in range(40):
                    yt = extract_yt_from_opened_detail(t)
                    if yt:
                        break
                    time.sleep(0.08)

                if yt and yt not in collected:
                    collected.add(yt)
                    new_in_pass += 1
                    log_cb(f"+ {yt} (yht {len(collected)})")
                elif not yt:
                    log_cb("SKIP: Y-tunnusta ei löytynyt avatusta detailistä")

                # sulje takaisin
                try:
                    safe_click(driver, t)
                    wait_toggle_state(driver, t, want_true=False, timeout=1.5)
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

        # Näytä lisää / Lataa lisää
        if click_nayta_lisaa(driver):
            status_cb("KL: Näytä lisää…")
            time.sleep(1.4)
            continue

        if rounds_without_new >= 2:
            status_cb("KL: ei uusia + ei Näytä lisää -> valmis.")
            break

    if not collected:
        status_cb("KL: Keräys tuotti 0 YT -> teen debug-dumpin.")
        html, png = dump_debug(driver, "kl_zero_result")
        log_cb(f"DEBUG: {html} | {png}")

    return sorted(collected)


# =========================
#   YTJ EMAILS
# =========================
def click_all_nayta_ytj(driver):
    for _ in range(3):
        clicked = False
        for b in driver.find_elements(By.TAG_NAME, "button"):
            try:
                if (b.text or "").strip().lower() == "näytä" and b.is_displayed() and b.is_enabled():
                    safe_click(driver, b)
                    clicked = True
                    time.sleep(0.15)
            except Exception:
                continue
        for a in driver.find_elements(By.TAG_NAME, "a"):
            try:
                if (a.text or "").strip().lower() == "näytä" and a.is_displayed():
                    safe_click(driver, a)
                    clicked = True
                    time.sleep(0.15)
            except Exception:
                continue
        if not clicked:
            break


def wait_ytj_loaded(driver):
    wait = WebDriverWait(driver, 25)
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


def fetch_emails_from_ytj(driver, yt_list, status_cb, progress_cb, log_cb):
    emails = []
    seen = set()
    progress_cb(0, max(1, len(yt_list)))

    for i, yt in enumerate(yt_list, start=1):
        status_cb(f"YTJ: {i}/{len(yt_list)} {yt}")
        progress_cb(i - 1, len(yt_list))

        driver.get(YTJ_COMPANY_URL.format(yt))
        wait_ytj_loaded(driver)
        try_accept_cookies(driver)

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

    progress_cb(len(yt_list), max(1, len(yt_list)))
    return emails


# =========================
#   GUI
# =========================
class App(tk.Tk):
    def __init__(self):
        super().__init__()
        reset_log()

        self.title("ProtestiBotti (KL Protestilista FIX)")
        self.geometry("940x600")

        tk.Label(self, text="ProtestiBotti", font=("Arial", 18, "bold")).pack(pady=10)
        tk.Label(
            self,
            text="Moodit:\n1) Kauppalehti (Chrome debug 9222) → Y-tunnukset → YTJ sähköpostit\n2) PDF → Y-tunnukset → YTJ sähköpostit",
            justify="center"
        ).pack(pady=4)

        btn_row = tk.Frame(self)
        btn_row.pack(pady=8)

        tk.Button(btn_row, text="Kauppalehti → YTJ", font=("Arial", 12), command=self.start_kauppalehti_mode).grid(row=0, column=0, padx=8)
        tk.Button(btn_row, text="PDF → YTJ", font=("Arial", 12), command=self.start_pdf_mode).grid(row=0, column=1, padx=8)

        self.status = tk.Label(self, text="Valmiina.", font=("Arial", 11))
        self.status.pack(pady=6)

        self.progress = ttk.Progressbar(self, orient="horizontal", mode="determinate", length=880)
        self.progress.pack(pady=6)

        frame = tk.Frame(self)
        frame.pack(fill="both", expand=True, padx=14, pady=10)

        tk.Label(frame, text="Live-logi (uusimmat alimmaisena):").pack(anchor="w")
        self.listbox = tk.Listbox(frame, height=18)
        self.listbox.pack(side="left", fill="both", expand=True)

        sb = tk.Scrollbar(frame, orient="vertical", command=self.listbox.yview)
        sb.pack(side="right", fill="y")
        self.listbox.configure(yscrollcommand=sb.set)

        tk.Label(self, text=f"Tallennus: {OUT_DIR}", wraplength=920, justify="center").pack(pady=6)

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

    def start_kauppalehti_mode(self):
        threading.Thread(target=self.run_kauppalehti_mode, daemon=True).start()

    def run_kauppalehti_mode(self):
        driver = None
        try:
            self.set_status("Liitytään Chrome-bottiin (9222)…")
            driver = attach_to_existing_chrome()

            self.set_status("Kauppalehti: kerätään Y-tunnukset (protestilista fix)…")
            yt_list = collect_yts_from_kauppalehti(driver, self.set_status, self.ui_log)

            if not yt_list:
                self.set_status("Ei löytynyt Y-tunnuksia.")
                messagebox.showwarning(
                    "Ei löytynyt",
                    "Y-tunnuksia ei saatu. Katso output-kansio: kl_*.html ja kl_*.png."
                )
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
            messagebox.showinfo(
                "Valmis",
                f"Valmis!\n\nKansio:\n{OUT_DIR}\n\nY-tunnuksia: {len(yt_list)}\nSähköposteja: {len(emails)}"
            )

        except WebDriverException as e:
            self.ui_log(f"SELENIUM VIRHE: {e}")
            self.set_status("Virhe. Katso log.txt / debug-dumpit.")
            messagebox.showerror("Virhe", f"Selenium/Chrome virhe.\nKatso log.txt:\n{LOG_PATH}\n\n{e}")
        except Exception as e:
            self.ui_log(f"VIRHE: {e}")
            self.set_status("Virhe. Katso log.txt / debug-dumpit.")
            messagebox.showerror("Virhe", f"Tuli virhe.\nKatso log.txt:\n{LOG_PATH}\n\n{e}")

    def start_pdf_mode(self):
        path = filedialog.askopenfilename(filetypes=[("PDF files", "*.pdf")])
        if path:
            threading.Thread(target=self.run_pdf_mode, args=(path,), daemon=True).start()

    def run_pdf_mode(self, pdf_path):
        driver = None
        try:
            self.set_status("Luetaan PDF ja kerätään Y-tunnukset…")
            yt_list = extract_ytunnukset_from_pdf(pdf_path)

            if not yt_list:
                self.set_status("Ei löytynyt Y-tunnuksia PDF:stä.")
                messagebox.showwarning("Ei löytynyt", "PDF:stä ei löytynyt yhtään Y-tunnusta.")
                return

            yt_path = save_word_plain_lines(yt_list, "ytunnukset.docx")
            self.ui_log(f"Tallennettu: {yt_path}")

            self.set_status("Käynnistetään Chrome ja haetaan sähköpostit YTJ:stä…")
            driver = start_new_driver()

            emails = fetch_emails_from_ytj(driver, yt_list, self.set_status, self.set_progress, self.ui_log)
            em_path = save_word_plain_lines(emails, "sahkopostit.docx")
            self.ui_log(f"Tallennettu: {em_path}")

            self.set_status("Valmis!")
            messagebox.showinfo(
                "Valmis",
                f"Valmis!\n\nKansio:\n{OUT_DIR}\n\nY-tunnuksia: {len(yt_list)}\nSähköposteja: {len(emails)}"
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
