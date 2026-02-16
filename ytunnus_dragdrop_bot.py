import os
import re
import sys
import time
import threading
import tkinter as tk
from tkinter import messagebox, filedialog

from concurrent.futures import ThreadPoolExecutor, as_completed

import PyPDF2
from docx import Document

from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC

from webdriver_manager.chrome import ChromeDriverManager


EMAIL_RE = re.compile(r"[A-Za-z0-9_.+-]+@[A-Za-z0-9-]+\.[A-Za-z0-9-.]+")
EMAIL_A_RE = re.compile(r"[A-Za-z0-9_.+-]+\s*\(a\)\s*[A-Za-z0-9-]+\.[A-Za-z0-9-.]+")


# -----------------------------
# Output: exe-kansio ensisijaisesti, fallback Documents\ProtestiBotti
# + päivämääräkansio YYYY-MM-DD
# -----------------------------
def get_exe_dir():
    if getattr(sys, "frozen", False):
        return os.path.dirname(sys.executable)
    return os.path.dirname(os.path.abspath(__file__))


def get_output_dir():
    base = get_exe_dir()
    try:
        test_path = os.path.join(base, "_write_test.tmp")
        with open(test_path, "w", encoding="utf-8") as f:
            f.write("ok")
        os.remove(test_path)
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

_log_lock = threading.Lock()


def log(msg: str):
    ts = time.strftime("%H:%M:%S")
    line = f"[{ts}] {msg}"
    with _log_lock:
        try:
            with open(LOG_PATH, "a", encoding="utf-8") as f:
                f.write(line + "\n")
        except Exception:
            pass
    print(line)


def reset_log():
    try:
        with open(LOG_PATH, "w", encoding="utf-8") as f:
            f.write("=== BOTTI KÄYNNISTETTY ===\n")
    except Exception:
        pass
    log(f"Output: {OUT_DIR}")
    log(f"Logi: {LOG_PATH}")


# -----------------------------
# PDF -> Y-tunnukset
# -----------------------------
def normalize_yt(yt: str):
    yt = yt.strip().replace(" ", "")
    if re.fullmatch(r"\d{7}-\d", yt):
        return yt
    if re.fullmatch(r"\d{8}", yt):
        return yt[:7] + "-" + yt[7]
    return None


def extract_ytunnukset_from_pdf(pdf_path: str):
    yt_set = set()
    with open(pdf_path, "rb") as f:
        reader = PyPDF2.PdfReader(f)
        for page in reader.pages:
            text = page.extract_text() or ""
            matches = re.findall(r"\b\d{7}-\d\b|\b\d{8}\b", text)
            for m in matches:
                n = normalize_yt(m)
                if n:
                    yt_set.add(n)
    return sorted(yt_set)


# -----------------------------
# Word-tallennus: puhdas lista
# - ei otsikkoa, ei mitään extraa
# -----------------------------
def save_word_plain_lines(lines, filename):
    path = os.path.join(OUT_DIR, filename)
    doc = Document()
    for line in lines:
        doc.add_paragraph(line)
    doc.save(path)
    log(f"Tallennettu: {path}")


# -----------------------------
# Email normalisointi
# -----------------------------
def normalize_email_candidate(raw: str) -> str:
    raw = (raw or "").strip()
    raw = raw.replace(" ", "")
    raw = raw.replace("(a)", "@").replace("[at]", "@")
    return raw


def pick_email_from_text(text: str) -> str:
    """
    Palauttaa yhden emailin tekstistä (ensisijaisesti @-muoto, sitten (a)-muoto).
    """
    if not text:
        return ""

    m = EMAIL_RE.search(text)
    if m:
        return normalize_email_candidate(m.group(0))

    m2 = EMAIL_A_RE.search(text)
    if m2:
        return normalize_email_candidate(m2.group(0))

    return ""


# -----------------------------
# VIRRE: sähköposti vain "Sähköposti"-kohdan yhteydestä
# (ja poimitaan regexillä vain email)
# -----------------------------
def extract_company_email_from_dom(driver):
    # Etsi label "Sähköposti"
    labels = driver.find_elements(By.XPATH, "//*[normalize-space(.)='Sähköposti']")
    for lab in labels:
        try:
            # 1) katso sama "rivi": parentin tekstistä poimi email
            parent = lab.find_element(By.XPATH, "..")
            email = pick_email_from_text(parent.text)
            if email:
                return email

            # 2) katso seuraava sisarus (arvo voi olla siinä)
            sibs = lab.find_elements(By.XPATH, "following-sibling::*[1]")
            for s in sibs:
                email = pick_email_from_text(s.text)
                if email:
                    return email

            # 3) katso linkit parentissa
            links = parent.find_elements(By.XPATH, ".//a")
            for a in links:
                email = pick_email_from_text(a.text)
                if email:
                    return email

        except Exception:
            continue

    # varmistus: jos labelia ei löydy mutta yrityssivulla silti email, etsitään vain sisältöalueesta
    # (ei koko sivu - vältetään footteri)
    try:
        main = driver.find_element(By.TAG_NAME, "main")
        email = pick_email_from_text(main.text)
        if email:
            return email
    except Exception:
        pass

    return ""


def click_hae_button(driver):
    buttons = driver.find_elements(By.TAG_NAME, "button")
    for b in buttons:
        try:
            if b.is_displayed() and "hae" in (b.text or "").strip().lower():
                b.click()
                return True
        except Exception:
            continue

    elems = driver.find_elements(By.XPATH, "//*[contains(translate(normalize-space(.),'hae','HAE'),'HAE')]")
    for e in elems:
        try:
            if e.is_displayed() and e.is_enabled():
                e.click()
                return True
        except Exception:
            continue

    return False


def start_driver(headless=True):
    options = webdriver.ChromeOptions()

    if headless:
        options.add_argument("--headless=new")
        options.add_argument("--window-size=1920,1080")

    options.add_argument("--disable-gpu")
    options.add_argument("--disable-dev-shm-usage")
    options.add_argument("--no-sandbox")

    driver_path = ChromeDriverManager().install()
    return webdriver.Chrome(service=Service(driver_path), options=options)


def virre_fetch_email(driver, yt):
    wait = WebDriverWait(driver, 12)

    driver.get("https://virre.prh.fi/novus/home")
    wait.until(EC.presence_of_element_located((By.TAG_NAME, "body")))

    search = wait.until(EC.presence_of_element_located((By.CSS_SELECTOR, "input[type='text'], input[type='search']")))
    search.clear()
    search.send_keys(yt)

    if not click_hae_button(driver):
        return ""

    # odota että ainakin body on olemassa, ja yritä myös sähköposti-labelia (jos löytyy)
    try:
        wait.until(EC.presence_of_element_located((By.XPATH, "//*[normalize-space(.)='Sähköposti']")))
    except Exception:
        pass

    return extract_company_email_from_dom(driver)


# -----------------------------
# PRO/ULTIMATE: worker pool
# -----------------------------
def worker_run(worker_id, yt_list, headless, per_request_delay):
    emails = []
    failed = []
    driver = None

    try:
        log(f"[W{worker_id}] Chrome start (headless={headless})...")
        driver = start_driver(headless=headless)
        log(f"[W{worker_id}] Ready. Items={len(yt_list)}")

        for yt in yt_list:
            try:
                email = virre_fetch_email(driver, yt)
                if email:
                    emails.append(email)
                else:
                    failed.append(yt)

                if per_request_delay > 0:
                    time.sleep(per_request_delay)

            except Exception as e:
                log(f"[W{worker_id}] Error {yt}: {e}")
                failed.append(yt)

    except Exception as e:
        log(f"[W{worker_id}] Worker crashed: {e}")
        failed.extend(yt_list)

    finally:
        if driver:
            try:
                driver.quit()
            except Exception:
                pass
        log(f"[W{worker_id}] Closed.")
    return emails, failed


def split_round_robin(items, n):
    buckets = [[] for _ in range(n)]
    for i, it in enumerate(items):
        buckets[i % n].append(it)
    return buckets


# -----------------------------
# GUI
# -----------------------------
class App(tk.Tk):
    def __init__(self):
        super().__init__()
        reset_log()

        self.title("ProtestiBotti ULTIMATE (nopea + fallback)")
        self.geometry("720x420")

        tk.Label(self, text="ProtestiBotti ULTIMATE", font=("Arial", 18, "bold")).pack(pady=12)
        tk.Label(self, text="Rinnakkaishaku + headless + fallback näkyvään Chromeen\nTallentaa vain Wordit", justify="center").pack(pady=6)

        row = tk.Frame(self)
        row.pack(pady=8)

        tk.Label(row, text="Workers:", font=("Arial", 11)).pack(side="left", padx=6)
        self.workers_var = tk.IntVar(value=3)
        tk.OptionMenu(row, self.workers_var, 1, 2, 3, 4, 5, 6).pack(side="left")

        self.headless_var = tk.BooleanVar(value=True)
        tk.Checkbutton(row, text="Headless (nopeampi)", variable=self.headless_var).pack(side="left", padx=12)

        self.fallback_var = tk.BooleanVar(value=True)
        tk.Checkbutton(row, text="Fallback näkyvään Chromeen (vain epäonnistuneet)", variable=self.fallback_var).pack(side="left", padx=12)

        row2 = tk.Frame(self)
        row2.pack(pady=6)
        tk.Label(row2, text="Viive / haku (s):", font=("Arial", 11)).pack(side="left", padx=6)
        self.delay_var = tk.DoubleVar(value=0.25)
        tk.Entry(row2, textvariable=self.delay_var, width=6).pack(side="left")
        tk.Label(row2, text="(0.2–0.5 suositus, nosta jos blokkaa)", font=("Arial", 10)).pack(side="left", padx=8)

        tk.Button(self, text="Valitse PDF", font=("Arial", 12), command=self.pick_pdf).pack(pady=12)

        self.status = tk.Label(self, text="Valmiina.", font=("Arial", 11))
        self.status.pack(pady=10)

        tk.Label(self, text=f"Tallennus: {OUT_DIR}", wraplength=680, justify="center").pack(pady=6)

    def set_status(self, s):
        self.status.config(text=s)
        self.update_idletasks()
        log(s)

    def pick_pdf(self):
        path = filedialog.askopenfilename(filetypes=[("PDF files", "*.pdf")])
        if path:
            threading.Thread(target=self.run_job, args=(path,), daemon=True).start()

    def run_job(self, pdf_path):
        try:
            self.set_status("Luetaan PDF ja kerätään Y-tunnukset…")
            yt_list = extract_ytunnukset_from_pdf(pdf_path)

            if not yt_list:
                self.set_status("Ei löytynyt Y-tunnuksia.")
                messagebox.showwarning("Ei löytynyt", "PDF:stä ei löytynyt yhtään Y-tunnusta.")
                return

            # ytunnukset.docx (otsikolla ok)
            save_word_plain_lines(yt_list, "ytunnukset.docx")  # pelkkä lista

            workers = int(self.workers_var.get())
            headless = bool(self.headless_var.get())
            fallback = bool(self.fallback_var.get())
            per_delay = float(self.delay_var.get())

            workers = max(1, min(workers, len(yt_list), 6))

            self.set_status(f"Headless-haku… workers={workers}")
            buckets = split_round_robin(yt_list, workers)

            all_emails = []
            failed_all = []

            with ThreadPoolExecutor(max_workers=workers) as ex:
                futures = []
                for wid, chunk in enumerate(buckets, start=1):
                    futures.append(ex.submit(worker_run, wid, chunk, headless, per_delay))

                done = 0
                for fut in as_completed(futures):
                    done += 1
                    emails, failed = fut.result()
                    all_emails.extend(emails)
                    failed_all.extend(failed)
                    self.set_status(f"Workers valmiina {done}/{workers}…")

            # ULTIMATE fallback: näkyvä chrome vain epäonnistuneille
            if fallback and failed_all:
                self.set_status(f"Fallback (näkyvä Chrome) epäonnistuneille: {len(failed_all)}")
                log(f"Fallback start. Failed={len(failed_all)}")

                # 1 näkyvä chrome (ei rinnakkain -> turvallisempi)
                try:
                    driver = start_driver(headless=False)
                    for i, yt in enumerate(failed_all, start=1):
                        self.set_status(f"Fallback {i}/{len(failed_all)}: {yt}")
                        try:
                            email = virre_fetch_email(driver, yt)
                            if email:
                                all_emails.append(email)
                        except Exception as e:
                            log(f"Fallback error {yt}: {e}")
                        time.sleep(max(per_delay, 0.25))
                finally:
                    try:
                        driver.quit()
                    except Exception:
                        pass

            # dedupe + puhdistus
            emails_only = []
            seen = set()
            for e in all_emails:
                e = normalize_email_candidate(e)
                # viimeinen varmistus: poimi pelkkä email jos jotain roskaa mukana
                e2 = pick_email_from_text(e) or e
                e2 = normalize_email_candidate(e2)
                if e2 and ("@" in e2) and ("." in e2):
                    k = e2.lower()
                    if k not in seen:
                        seen.add(k)
                        emails_only.append(e2)

            # Word: vain sähköpostit allekkain, EI OTSIKKOA
            self.set_status("Tallennetaan sahkopostit.docx…")
            save_word_plain_lines(emails_only, "sahkopostit.docx")

            self.set_status("Valmis!")
            messagebox.showinfo(
                "Valmis",
                f"Valmis!\n\nKansio:\n{OUT_DIR}\n\nTiedostot:\n- ytunnukset.docx\n- sahkopostit.docx\n\nLogi:\n{LOG_PATH}"
            )

        except Exception as e:
            log(f"VIRHE: {e}")
            self.set_status("Virhe. Katso log.txt")
            messagebox.showerror("Virhe", f"Tuli virhe.\nKatso log.txt:\n{LOG_PATH}\n\n{e}")


if __name__ == "__main__":
    App().mainloop()
