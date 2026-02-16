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


# -----------------------------
# Output-kansio + Päivämääräkansio
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
    except:
        pass

    print(line)


def reset_log():
    try:
        with open(LOG_PATH, "w", encoding="utf-8") as f:
            f.write("=== BOTTI KÄYNNISTETTY ===\n")
    except:
        pass

    log(f"Output: {OUT_DIR}")


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
# Word tallennus
# -----------------------------
def save_word_list(lines, filename, title=None):
    path = os.path.join(OUT_DIR, filename)

    doc = Document()

    if title:
        doc.add_heading(title, level=1)

    for line in lines:
        doc.add_paragraph(line)

    doc.save(path)
    log(f"Tallennettu: {filename}")


# -----------------------------
# VIRRE sähköposti haku
# -----------------------------
def normalize_email(text: str):
    return (text or "").strip().replace(" ", "").replace("(a)", "@").replace("[at]", "@")


def extract_company_email_from_dom(driver):
    labels = driver.find_elements(By.XPATH, "//*[normalize-space(.)='Sähköposti']")

    for lab in labels:
        try:
            parent = lab.find_element(By.XPATH, "..")
            txt = normalize_email(parent.text)

            if "@" in txt and "." in txt:
                return txt

        except:
            pass

    return ""


def start_chrome_driver():
    log("Käynnistetään headless Chrome...")

    options = webdriver.ChromeOptions()

    # 🔥 HEADLESS MODE
    options.add_argument("--headless=new")
    options.add_argument("--window-size=1920,1080")

    options.add_argument("--disable-gpu")
    options.add_argument("--disable-dev-shm-usage")

    driver_path = ChromeDriverManager().install()
    log(f"ChromeDriver OK")

    return webdriver.Chrome(service=Service(driver_path), options=options)


def click_hae_button(driver):
    buttons = driver.find_elements(By.TAG_NAME, "button")

    for b in buttons:
        try:
            if "hae" in b.text.lower():
                b.click()
                return True
        except:
            pass

    return False


def virre_fetch_email_for_yt(driver, yt):
    wait = WebDriverWait(driver, 10)

    driver.get("https://virre.prh.fi/novus/home")
    wait.until(EC.presence_of_element_located((By.TAG_NAME, "body")))

    try:
        search = driver.find_element(By.CSS_SELECTOR, "input[type='text']")
        search.clear()
        search.send_keys(yt)

        if not click_hae_button(driver):
            return "", "EI HAE-NAPPIA"

        try:
            wait.until(EC.presence_of_element_located((By.XPATH, "//*[normalize-space(.)='Sähköposti']")))
        except:
            pass

        email = extract_company_email_from_dom(driver)

        if email:
            return email, "OK"

        return "", "EI SÄHKÖPOSTIA"

    except Exception as e:
        return "", f"VIRHE: {str(e)}"


# -----------------------------
# GUI
# -----------------------------
class App(tk.Tk):
    def __init__(self):
        super().__init__()

        reset_log()

        self.title("ProtestiBotti (HEADLESS)")
        self.geometry("600x340")

        tk.Label(self, text="ProtestiBotti", font=("Arial", 18, "bold")).pack(pady=12)

        tk.Label(self, text="Valitse PDF → sähköpostit haetaan taustalla", justify="center").pack(pady=6)

        tk.Button(self, text="Valitse PDF", font=("Arial", 12), command=self.pick_pdf).pack(pady=10)

        self.status = tk.Label(self, text="Valmiina.")
        self.status.pack(pady=10)

    def set_status(self, s):
        self.status.config(text=s)
        self.update_idletasks()
        log(s)

    def pick_pdf(self):
        path = filedialog.askopenfilename(filetypes=[("PDF files", "*.pdf")])
        if path:
            threading.Thread(target=self.run_job, args=(path,), daemon=True).start()

    def run_job(self, pdf):
        ytunnukset = extract_ytunnukset_from_pdf(pdf)

        if not ytunnukset:
            messagebox.showwarning("Ei löytynyt", "Ei Y-tunnuksia")
            return

        save_word_list(ytunnukset, "ytunnukset.docx", "Y-tunnukset")

        driver = start_chrome_driver()

        emails = []
        seen = set()

        for i, yt in enumerate(ytunnukset, start=1):
            self.set_status(f"Haku {i}/{len(ytunnukset)}")

            email, _ = virre_fetch_email_for_yt(driver, yt)

            if email:
                key = email.lower()
                if key not in seen:
                    seen.add(key)
                    emails.append(email)

        driver.quit()

        save_word_list(emails, "sahkopostit.docx", "Sähköpostit")

        self.set_status("VALMIS")
        log("VALMIS")


if __name__ == "__main__":
    App().mainloop()
