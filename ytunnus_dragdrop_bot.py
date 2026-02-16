import os
import re
import time
import threading
import tkinter as tk
from tkinter import messagebox, filedialog

import PyPDF2
from docx import Document
import openpyxl

from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from webdriver_manager.chrome import ChromeDriverManager


# -----------------------------
# Output kansio (aina toimii)
# -----------------------------
def get_output_dir():
    home = os.path.expanduser("~")
    docs = os.path.join(home, "Documents")
    out = os.path.join(docs, "ProtestiBotti")
    os.makedirs(out, exist_ok=True)
    return out

OUT_DIR = get_output_dir()
LOG_PATH = os.path.join(OUT_DIR, "log.txt")


def log(msg):
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


# -----------------------------
# Y-tunnus logiikka
# -----------------------------
def normalize_yt(yt):
    yt = yt.strip().replace(" ", "")

    if re.fullmatch(r"\d{7}-\d", yt):
        return yt

    if re.fullmatch(r"\d{8}", yt):
        return yt[:7] + "-" + yt[7]

    return None


def extract_ytunnukset_from_pdf(pdf_path):
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
# Tallennukset
# -----------------------------
def save_word(lines, filename, title):
    path = os.path.join(OUT_DIR, filename)

    doc = Document()
    doc.add_heading(title, level=1)

    for line in lines:
        doc.add_paragraph(line)

    doc.save(path)
    log(f"Word tallennettu: {filename}")


def save_excel(rows, filename, headers):
    path = os.path.join(OUT_DIR, filename)

    wb = openpyxl.Workbook()
    ws = wb.active

    ws.append(headers)

    for r in rows:
        ws.append(list(r))

    wb.save(path)
    log(f"Excel tallennettu: {filename}")


# -----------------------------
# Virre sähköposti haku
# -----------------------------
EMAIL_REGEX = re.compile(r"[a-zA-Z0-9_.+-]+@[a-zA-Z0-9-]+\.[a-zA-Z0-9-.]+")


def start_driver():
    log("Käynnistetään Chrome...")

    options = webdriver.ChromeOptions()
    options.add_argument("--start-maximized")

    driver_path = ChromeDriverManager().install()
    log(f"ChromeDriver: {driver_path}")

    return webdriver.Chrome(service=Service(driver_path), options=options)


def fetch_email(driver, yt):
    wait = WebDriverWait(driver, 20)

    driver.get("https://virre.prh.fi/novus/home")
    wait.until(EC.presence_of_element_located((By.TAG_NAME, "body")))
    time.sleep(1)

    try:
        search = driver.find_element(By.CSS_SELECTOR, "input[type='text']")
        search.clear()
        search.send_keys(yt)

        buttons = driver.find_elements(By.TAG_NAME, "button")
        for b in buttons:
            if "hae" in b.text.lower():
                b.click()
                break

        time.sleep(3)

        page = driver.page_source
        match = EMAIL_REGEX.search(page)

        if match:
            return match.group(0)

        return "EI SÄHKÖPOSTIA"

    except Exception as e:
        return f"VIRHE: {str(e)}"


# -----------------------------
# GUI
# -----------------------------
class App(tk.Tk):
    def __init__(self):
        super().__init__()

        reset_log()
        log("GUI käynnistyi")

        self.title("ProtestiBotti")
        self.geometry("500x300")

        tk.Label(self, text="Valitse PDF", font=("Arial", 16)).pack(pady=20)

        tk.Button(self, text="Valitse PDF", command=self.pick_pdf).pack()

        self.status = tk.Label(self, text="")
        self.status.pack(pady=10)

    def pick_pdf(self):
        path = filedialog.askopenfilename(filetypes=[("PDF", "*.pdf")])
        if path:
            threading.Thread(target=self.run_bot, args=(path,), daemon=True).start()

    def run_bot(self, pdf):
        self.status.config(text="Luetaan PDF...")
        log("Luetaan PDF")

        ytunnukset = extract_ytunnukset_from_pdf(pdf)

        if not ytunnukset:
            log("Ei Y-tunnuksia")
            messagebox.showwarning("Ei löytynyt", "Ei Y-tunnuksia")
            return

        save_word(ytunnukset, "ytunnukset.docx", "Y-tunnukset")
        save_excel([(y,) for y in ytunnukset], "ytunnukset.xlsx", ["Y-tunnus"])

        try:
            driver = start_driver()
        except Exception as e:
            log(f"Chrome ei käynnisty: {e}")
            return

        results = []

        for yt in ytunnukset:
            self.status.config(text=f"Haku: {yt}")
            log(f"Haku Virrestä: {yt}")

            email = fetch_email(driver, yt)
            results.append((yt, email))

        driver.quit()

        save_word([f"{y} -> {e}" for y, e in results],
                  "virre_sahkopostit.docx",
                  "Virre sähköpostit")

        save_excel(results, "virre_sahkopostit.xlsx",
                   ["Y-tunnus", "Sähköposti"])

        self.status.config(text="VALMIS")
        log("VALMIS")


if __name__ == "__main__":
    App().mainloop()
