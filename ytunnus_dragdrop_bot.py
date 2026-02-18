import os
import re
import sys
import time
import threading
import tkinter as tk
from tkinter import messagebox, filedialog
from tkinter import ttk
from dataclasses import dataclass

import PyPDF2
from docx import Document

from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
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
YTJ_HOME = "https://tietopalvelu.ytj.fi/"
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
    s = s.replace("asunto oy", "as oy")
    s = s.replace("asunto-osakeyhtiö", "as oy")
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
    overlap = len(qtoks & ctoks)
    score += overlap * 8

    loc = normalize_name(query_location)
    if loc and loc in cl:
        score += 35

    for suf in [" oy", " ab", " ky", " oyj", " tmi", " ry"]:
        if suf.strip() in qn and suf.strip() in cl:
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
    hdr_cells = table.rows[0].cells
    for i, h in enumerate(headers):
        hdr_cells[i].text = str(h)

    for r in rows:
        row_cells = table.add_row().cells
        for i, v in enumerate(r):
            row_cells[i].text = "" if v is None else str(v)

    doc.save(path)
    return path


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
#   SELENIUM START (NO 9222)
# =========================
def get_profile_dir():
    """
    Pysyvä Chrome-profiili botille.
    Tänne tallentuu login Kauppalehteen, evästeet jne.
    """
    base = get_exe_dir()
    prof = os.path.join(base, "chrome_profile")
    # jos ei kirjoitusoikeutta exe-kansioon, käytä Documents
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
    """
    Käynnistää Chromen omalla profiililla (ei 9222).
    Kirjaudut kerran -> jatkossa sessio pysyy.
    """
    options = webdriver.ChromeOptions()
    options.add_argument("--start-maximized")
    options.add_argument("--disable-gpu")
    options.add_argument("--disable-dev-shm-usage")

    profile_dir = get_profile_dir()
    options.add_argument(f"--user-data-dir={profile_dir}")

    driver_path = ChromeDriverManager().install()
    return webdriver.Chrome(service=Service(driver_path), options=options)


# =========================
#   KL checks
# =========================
def page_looks_like_login_or_paywall(driver) -> bool:
    try:
        text = (driver.find_element(By.TAG_NAME, "body").text or "").lower()
        bad_words = [
            "kirjaudu", "tilaa", "tilaajille", "vahvista henkilöllb", "vahvista henkilöllisyytesi",
            "sign in", "subscribe", "login",
            "pääsy evätty", "access denied",
            "jokin meni pieleen", "something went wrong",
        ]
        return any(w.lower() in text for w in bad_words)
    except Exception:
        return False


def ensure_kl_ready(driver, status_cb, log_cb, max_wait_seconds=900):
    driver.get(KAUPPALEHTI_URL)
    WebDriverWait(driver, 25).until(EC.presence_of_element_located((By.TAG_NAME, "body")))
    try_accept_cookies(driver)

    start = time.time()
    warned = False

    while True:
        try_accept_cookies(driver)
        body = ""
        try:
            body = driver.find_element(By.TAG_NAME, "body").text or ""
        except Exception:
            pass

        if "Protestilista" in body and not page_looks_like_login_or_paywall(driver):
            status_cb("KL: protestilista näkyy.")
            time.sleep(0.8)
            return True

        if not warned:
            warned = True
            status_cb("KL: kirjaudu Kauppalehteen tässä botti-Chromessa, sitten botti jatkaa.")
            log_cb("Waiting for user to login on Kauppalehti (persistent profile)…")
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


def click_nayta_lisaa(driver) -> bool:
    try:
        driver.execute_script("window.scrollTo(0, document.body.scrollHeight);")
    except Exception:
        pass
    time.sleep(0.25)

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


# =========================
#   HYBRID: KL -> YTJ by name
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


def collect_kl_rows_all(driver, status_cb, log_cb, max_pages=999):
    if not ensure_kl_ready(driver, status_cb, log_cb, max_wait_seconds=900):
        return []

    collected = []
    seen = set()
    rounds_no_new = 0
    page_clicks = 0

    while True:
        try_accept_cookies(driver)

        if page_looks_like_login_or_paywall(driver):
            status_cb("KL: näyttää blokilta/paywallilta -> debug-dump ja odotan…")
            html, png = dump_debug(driver, "kl_blocked")
            log_cb(f"DEBUG dump: {html} | {png}")
            time.sleep(3)
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
            time.sleep(1.2)
            continue

        if rounds_no_new >= 2:
            break

    if not collected:
        html, png = dump_debug(driver, "kl_zero_rows")
        log_cb(f"DEBUG dump: {html} | {png}")

    return collected


# ---- YTJ helpers ----
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

    for a in links[:50]:
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


def run_hybrid(driver, status_cb, progress_cb, log_cb):
    status_cb("KL: kerätään yritysrivit (nimi+sijainti)…")
    kl_rows = collect_kl_rows_all(driver, status_cb, log_cb)
    if not kl_rows:
        return [], []

    # tallenna KL rivit
    kl_table = [[r.company, r.location, r.date, r.amount, r.ptype, r.source] for r in kl_rows]
    kl_path = save_word_table(
        kl_table,
        headers=["Yritys", "Sijainti", "Häiriöpäivä", "Summa", "Tyyppi", "Lähde"],
        filename="kl_rivit.docx",
    )
    log_cb(f"Tallennettu: {kl_path}")

    # YTJ-haku (sama chrome, mutta navigoidaan sivusta toiseen)
    results = []
    progress_cb(0, len(kl_rows))

    for i, r in enumerate(kl_rows, start=1):
        status_cb(f"YTJ-haku: {i}/{len(kl_rows)} — {r.company} ({r.location})")
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

        time.sleep(0.15)

    progress_cb(len(kl_rows), len(kl_rows))

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

        self.title("ProtestiBotti (HYBRID, ilman 9222)")
        self.geometry("980x620")

        tk.Label(self, text="ProtestiBotti (Hybrid, NO 9222)", font=("Arial", 18, "bold")).pack(pady=10)
        tk.Label(
            self,
            text="Tämä versio avaa oman Chromen pysyvällä profiililla.\n"
                 "Kirjaudu Kauppalehteen kerran -> jatkossa kirjautuminen pysyy.\n\n"
                 "Nappi: Kauppalehti (HYBRID) → Tulokset",
            justify="center"
        ).pack(pady=4)

        btn_row = tk.Frame(self)
        btn_row.pack(pady=8)

        tk.Button(btn_row, text="Kauppalehti (HYBRID) → Tulokset", font=("Arial", 12), command=self.start_hybrid_mode).grid(row=0, column=0, padx=8)

        self.status = tk.Label(self, text="Valmiina.", font=("Arial", 11))
        self.status.pack(pady=6)

        self.progress = ttk.Progressbar(self, orient="horizontal", mode="determinate", length=920)
        self.progress.pack(pady=6)

        frame = tk.Frame(self)
        frame.pack(fill="both", expand=True, padx=14, pady=10)

        tk.Label(frame, text="Live-logi (uusimmat alimmaisena):").pack(anchor="w")
        self.listbox = tk.Listbox(frame, height=20)
        self.listbox.pack(side="left", fill="both", expand=True)

        sb = tk.Scrollbar(frame, orient="vertical", command=self.listbox.yview)
        sb.pack(side="right", fill="y")
        self.listbox.configure(yscrollcommand=sb.set)

        tk.Label(self, text=f"Tallennus: {OUT_DIR}", wraplength=960, justify="center").pack(pady=6)

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

    def start_hybrid_mode(self):
        threading.Thread(target=self.run_hybrid_mode, daemon=True).start()

    def run_hybrid_mode(self):
        driver = None
        try:
            self.set_status("Käynnistetään Chrome (pysyvä profiili, ei 9222)…")
            driver = start_persistent_driver()

            self.set_status("Aloitetaan HYBRID: KL rivit → YTJ nimihaku…")
            kl_rows, results = run_hybrid(driver, self.set_status, self.set_progress, self.ui_log)

            if not kl_rows:
                self.set_status("KL: Ei saatu rivejä.")
                messagebox.showwarning("Ei rivejä", "Kauppalehdestä ei saatu rivejä. Katso log.txt / debug-dumpit.")
                return

            emails = [r[3] for r in results if r and r[3]]
            uniq_emails = len(set([e.lower() for e in emails]))

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
            messagebox.showerror("Virhe", f"Selenium/Chrome virhe.\nKatso log.txt:\n{LOG_PATH}\n\n{e}")
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
