import os
import re
import sys
import time
import random
import threading
import tkinter as tk
from tkinter import messagebox, filedialog

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

BLOCK_HINTS = [
    "captcha", "robot", "access denied", "liian monta", "too many requests",
    "rate limit", "kielletty", "forbidden", "blocked", "estetty", "varmistus"
]


# -----------------------------
# Output + päivämääräkansio
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


def log(msg: str):
    ts = time.strftime("%H:%M:%S")
    line = f"[{ts}] {msg}"
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
# Word: pelkkä lista riveinä (ei otsikkoa)
# -----------------------------
def save_word_plain_lines(lines, filename):
    path = os.path.join(OUT_DIR, filename)
    doc = Document()
    for line in lines:
        doc.add_paragraph(line)
    doc.save(path)
    log(f"Tallennettu: {path}")


# -----------------------------
# Email poiminta (vain email)
# -----------------------------
def normalize_email_candidate(raw: str) -> str:
    raw = (raw or "").strip().replace(" ", "")
    raw = raw.replace("(a)", "@").replace("[at]", "@")
    return raw


def pick_email_from_text(text: str) -> str:
    if not text:
        return ""
    m = EMAIL_RE.search(text)
    if m:
        return normalize_email_candidate(m.group(0))
    m2 = EMAIL_A_RE.search(text)
    if m2:
        return normalize_email_candidate(m2.group(0))
    return ""


def extract_company_email_from_dom(driver):
    labels = driver.find_elements(By.XPATH, "//*[normalize-space(.)='Sähköposti']")
    for lab in labels:
        try:
            parent = lab.find_element(By.XPATH, "..")
            email = pick_email_from_text(parent.text)
            if email:
                return email

            sib = lab.find_elements(By.XPATH, "following-sibling::*[1]")
            for s in sib:
                email = pick_email_from_text(s.text)
                if email:
                    return email

            links = parent.find_elements(By.XPATH, ".//a")
            for a in links:
                email = pick_email_from_text(a.text)
                if email:
                    return email
        except Exception:
            continue

    # fallback: etsi mail main-alueelta (ei koko sivulta -> vähentää footteri-osumia)
    try:
        main = driver.find_element(By.TAG_NAME, "main")
        email = pick_email_from_text(main.text)
        if email:
            return email
    except Exception:
        pass

    return ""


# -----------------------------
# Virre: 1 selain, NOPEA, blokin tunnistus
# -----------------------------
def start_driver():
    options = webdriver.ChromeOptions()
    options.add_argument("--start-maximized")
    options.add_argument("--disable-gpu")
    options.add_argument("--disable-dev-shm-usage")

    driver_path = ChromeDriverManager().install()
    log("ChromeDriver OK")
    return webdriver.Chrome(service=Service(driver_path), options=options)


def click_hae_button(driver):
    buttons = driver.find_elements(By.TAG_NAME, "button")
    for b in buttons:
        try:
            if b.is_displayed() and "hae" in (b.text or "").strip().lower():
                b.click()
                return True
        except Exception:
            continue
    return False


def looks_blocked(page_source_lower: str) -> bool:
    return any(h in page_source_lower for h in BLOCK_HINTS)


def virre_fetch_email(driver, yt):
    """
    Returns: (email, blocked_flag)
    blocked_flag=True jos havaitaan blokki/captcha/robot.
    """
    wait = WebDriverWait(driver, 20)

    driver.get("https://virre.prh.fi/novus/home")
    wait.until(EC.presence_of_element_located((By.TAG_NAME, "body")))

    src = (driver.page_source or "").lower()
    if looks_blocked(src):
        return "", True

    search = wait.until(EC.presence_of_element_located((By.CSS_SELECTOR, "input[type='text'], input[type='search']")))
    search.clear()
    search.send_keys(yt)

    if not click_hae_button(driver):
        # ei välttämättä blokki, voi olla UI muutos
        return "", False

    # pieni odotus että tulos ehtii päivittyä (ei pitkä sleep)
    try:
        wait.until(EC.presence_of_element_located((By.XPATH, "//*[normalize-space(.)='Sähköposti']")))
    except Exception:
        pass

    src2 = (driver.page_source or "").lower()
    if looks_blocked(src2):
        return "", True

    email = extract_company_email_from_dom(driver)
    return email, False


# -----------------------------
# GUI
# -----------------------------
class App(tk.Tk):
    def __init__(self):
        super().__init__()
        reset_log()

        self.title("ProtestiBotti (nopea + blokki-hälytys)")
        self.geometry("720x420")

        tk.Label(self, text="ProtestiBotti", font=("Arial", 18, "bold")).pack(pady=12)
        tk.Label(self, text="Nopea yksinajo + pysäyttää jos Virre blokkaa/captcha", justify="center").pack(pady=6)

        row = tk.Frame(self)
        row.pack(pady=6)

        tk.Label(row, text="Perusviive (s):", font=("Arial", 11)).pack(side="left", padx=6)
        self.base_delay_var = tk.DoubleVar(value=0.12)  # NOPEA
        tk.Entry(row, textvariable=self.base_delay_var, width=6).pack(side="left")

        row2 = tk.Frame(self)
        row2.pack(pady=6)
        tk.Label(row2, text="Max viive (s):", font=("Arial", 11)).pack(side="left", padx=6)
        self.max_delay_var = tk.DoubleVar(value=1.5)  # ei pitkiä viiveitä
        tk.Entry(row2, textvariable=self.max_delay_var, width=6).pack(side="left")

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

            save_word_plain_lines(yt_list, "ytunnukset.docx")

            self.set_status("Käynnistetään Chrome…")
            driver = start_driver()

            emails = []
            seen = set()

            base_delay = max(0.05, float(self.base_delay_var.get()))
            max_delay = max(base_delay, float(self.max_delay_var.get()))
            cur_delay = base_delay

            try:
                for i, yt in enumerate(yt_list, start=1):
                    self.set_status(f"Haku {i}/{len(yt_list)} (viive {cur_delay:.2f}s)")

                    email, blocked = virre_fetch_email(driver, yt)

                    if blocked:
                        # Tallenna tähän asti saadut ja pysäytä
                        save_word_plain_lines(emails, "sahkopostit_PARTIAL.docx")
                        log("BLOKKI HAVAITTU -> pysäytetään.")
                        messagebox.showerror(
                            "Blokki havaittu",
                            "Virre näyttää blokanneen / CAPTCHA / robot-tarkistus.\n\n"
                            "Tallensin tähän asti saadut sähköpostit:\n"
                            f"{os.path.join(OUT_DIR, 'sahkopostit_PARTIAL.docx')}\n\n"
                            "Pidä tauko ja kokeile myöhemmin."
                        )
                        self.set_status("Pysäytetty: blokki havaittu.")
                        return

                    if email:
                        k = email.lower()
                        if k not in seen:
                            seen.add(k)
                            emails.append(email)

                    # NOPEA, mutta pieni jitter ettei ole täysin tasainen
                    jitter = random.uniform(0.0, 0.08)

                    # kevyt “mikro-backoff” jos ei löydy mitään usein (mutta ei pitkiä viiveitä)
                    if not email:
                        cur_delay = min(max_delay, max(cur_delay * 1.08, base_delay))
                    else:
                        cur_delay = max(base_delay, cur_delay * 0.95)

                    time.sleep(cur_delay + jitter)

                    # kevyt mini-tauko joka 100 haku (EI pitkä)
                    if i % 100 == 0:
                        time.sleep(1.0)

            finally:
                try:
                    driver.quit()
                except Exception:
                    pass

            save_word_plain_lines(emails, "sahkopostit.docx")
            self.set_status("Valmis!")
            messagebox.showinfo("Valmis", f"Valmis!\n\nKansio:\n{OUT_DIR}")

        except Exception as e:
            log(f"VIRHE: {e}")
            self.set_status("Virhe. Katso log.txt")
            messagebox.showerror("Virhe", f"Tuli virhe.\nKatso log.txt:\n{LOG_PATH}\n\n{e}")


if __name__ == "__main__":
    App().mainloop()
