import os
import re
import sys
import time
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
EMAIL_A_RE = re.compile(r"[A-Za-z0-9_.+-]+\s*\(a\)\s*[A-Za-z0-9-]+\.[A-Za-z0-9-.]+", re.I)


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
# Email helpers
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


# -----------------------------
# Selenium / YTJ
# -----------------------------
def start_driver():
    options = webdriver.ChromeOptions()
    options.add_argument("--start-maximized")
    options.add_argument("--disable-gpu")
    options.add_argument("--disable-dev-shm-usage")

    driver_path = ChromeDriverManager().install()
    log("ChromeDriver OK")
    return webdriver.Chrome(service=Service(driver_path), options=options)


def click_all_nayta(driver):
    """
    Klikkaa kaikki näkyvät 'Näytä' -napit (button tai linkki).
    """
    for _ in range(3):
        clicked = False

        # button
        for b in driver.find_elements(By.TAG_NAME, "button"):
            try:
                if (b.text or "").strip().lower() == "näytä" and b.is_displayed() and b.is_enabled():
                    driver.execute_script("arguments[0].scrollIntoView({block:'center'});", b)
                    b.click()
                    clicked = True
                    time.sleep(0.15)
            except Exception:
                continue

        # linkki
        for a in driver.find_elements(By.TAG_NAME, "a"):
            try:
                if (a.text or "").strip().lower() == "näytä" and a.is_displayed():
                    driver.execute_script("arguments[0].scrollIntoView({block:'center'});", a)
                    a.click()
                    clicked = True
                    time.sleep(0.15)
            except Exception:
                continue

        if not clicked:
            break


def extract_email_from_ytj(driver):
    """
    Poimi sähköposti taulukkoriviltä:
    - löydä elementti "Sähköposti"
    - ota lähin <tr> ja poimi email sen tekstistä
    """
    nodes = driver.find_elements(By.XPATH, "//*[normalize-space(.)='Sähköposti']")
    for n in nodes:
        try:
            tr = n.find_element(By.XPATH, "ancestor::tr[1]")
            email = pick_email_from_text(tr.text or "")
            if email:
                return email
        except Exception:
            continue

    # fallback: main-alueesta
    try:
        main = driver.find_element(By.TAG_NAME, "main")
        return pick_email_from_text(main.text or "")
    except Exception:
        return ""


def ytj_fetch_email_direct(driver, yt):
    """
    UUSI LOGIIKKA: avaa suoraan yrityssivu:
    https://tietopalvelu.ytj.fi/yritys/<Y-TUNNUS>
    """
    wait = WebDriverWait(driver, 20)
    url = f"https://tietopalvelu.ytj.fi/yritys/{yt}"
    driver.get(url)

    # odota että sivu latautuu
    wait.until(EC.presence_of_element_located((By.TAG_NAME, "body")))

    # varmistus: jos tuli jokin virhesivu, palauta tyhjä
    src = (driver.page_source or "").lower()
    if "404" in src or "not found" in src:
        return ""

    # klikkaa "Näytä" jos tarvitaan
    click_all_nayta(driver)

    # poimi email
    return extract_email_from_ytj(driver)


# -----------------------------
# GUI
# -----------------------------
class App(tk.Tk):
    def __init__(self):
        super().__init__()
        reset_log()

        self.title("ProtestiBotti (YTJ suora linkki)")
        self.geometry("620x360")

        tk.Label(self, text="ProtestiBotti", font=("Arial", 18, "bold")).pack(pady=12)
        tk.Label(self, text="Valitse PDF → kerää Y-tunnukset → avaa YTJ yrityssivu suoraan (/yritys/..)\n→ klikkaa Näytä → poimi sähköposti",
                 justify="center").pack(pady=6)

        tk.Button(self, text="Valitse PDF", font=("Arial", 12), command=self.pick_pdf).pack(pady=12)

        self.status = tk.Label(self, text="Valmiina.", font=("Arial", 11))
        self.status.pack(pady=10)

        tk.Label(self, text=f"Tallennus: {OUT_DIR}", wraplength=580, justify="center").pack(pady=6)

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

            self.set_status("Käynnistetään Chrome ja haetaan sähköpostit (YTJ)…")
            driver = start_driver()

            emails = []
            seen = set()

            try:
                for i, yt in enumerate(yt_list, start=1):
                    self.set_status(f"Haku {i}/{len(yt_list)}: {yt}")
                    email = ytj_fetch_email_direct(driver, yt)

                    if email:
                        k = email.lower()
                        if k not in seen:
                            seen.add(k)
                            emails.append(email)

                    time.sleep(0.1)  # nopea
            finally:
                try:
                    driver.quit()
                except Exception:
                    pass

            save_word_plain_lines(emails, "sahkopostit.docx")

            self.set_status("Valmis!")
            messagebox.showinfo("Valmis", f"Valmis!\n\nKansio:\n{OUT_DIR}\n\nLöydetyt emailit: {len(emails)}")

        except Exception as e:
            log(f"VIRHE: {e}")
            self.set_status("Virhe. Katso log.txt")
            messagebox.showerror("Virhe", f"Tuli virhe.\nKatso log.txt:\n{LOG_PATH}\n\n{e}")


if __name__ == "__main__":
    App().mainloop()
