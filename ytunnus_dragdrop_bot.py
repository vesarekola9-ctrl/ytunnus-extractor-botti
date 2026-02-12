import os
import re
import threading
import tkinter as tk
from tkinter import messagebox

from tkinterdnd2 import DND_FILES, TkinterDnD

import PyPDF2
from docx import Document
import openpyxl

# Selenium + driver manager
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from webdriver_manager.chrome import ChromeDriverManager


# -----------------------------
# Y-tunnus logiikka
# -----------------------------
def fix_ytunnus(yt):
    yt = yt.strip()
    if re.match(r"^\d{7}-\d$", yt):
        return yt
    if re.match(r"^\d{8}$", yt):
        return yt[:7] + "-" + yt[7]
    return None


def extract_ytunnukset_from_pdf(pdf_path):
    ytunnukset = set()

    with open(pdf_path, "rb") as f:
        reader = PyPDF2.PdfReader(f)
        for page in reader.pages:
            text = page.extract_text()
            if not text:
                continue

            # Etsitään kaikki mahdolliset Y-tunnukset
            matches = re.findall(r"\b\d{7}-\d\b|\b\d{8}\b", text)

            for m in matches:
                fixed = fix_ytunnus(m)
                if fixed:
                    ytunnukset.add(fixed)

    return sorted(list(ytunnukset))


# -----------------------------
# Tallennus Word + Excel
# -----------------------------
def save_ytunnukset_word(ytunnukset, filename="ytunnukset.docx"):
    doc = Document()
    doc.add_heading("Y-tunnukset", level=1)

    for yt in ytunnukset:
        doc.add_paragraph(yt)

    doc.save(filename)


def save_ytunnukset_excel(ytunnukset, filename="ytunnukset.xlsx"):
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Y-tunnukset"

    ws["A1"] = "Y-tunnus"

    for i, yt in enumerate(ytunnukset, start=2):
        ws[f"A{i}"] = yt

    wb.save(filename)


def save_emails_word(results, filename="virre_sahkopostit.docx"):
    doc = Document()
    doc.add_heading("Virre - Sähköpostit", level=1)

    for yt, email in results:
        doc.add_paragraph(f"{yt}  ->  {email}")

    doc.save(filename)


def save_emails_excel(results, filename="virre_sahkopostit.xlsx"):
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Sähköpostit"

    ws["A1"] = "Y-tunnus"
    ws["B1"] = "Sähköposti"

    for i, (yt, email) in enumerate(results, start=2):
        ws[f"A{i}"] = yt
        ws[f"B{i}"] = email

    wb.save(filename)


# -----------------------------
# Virre sähköpostien haku
# -----------------------------
def fetch_email_from_virre(driver, ytunnus):
    """
    Avaa Virre, hakee Y-tunnuksen ja palauttaa sähköpostin jos löytyy.
    """
    driver.get("https://virre.prh.fi/novus/home")

    wait = WebDriverWait(driver, 20)

    try:
        # Etsitään hakukenttä (Virre voi käyttää eri nimityksiä)
        search_input = wait.until(
            EC.presence_of_element_located((By.CSS_SELECTOR, "input[type='text']"))
        )
        search_input.clear()
        search_input.send_keys(ytunnus)

        # Etsitään Hae/Search nappi
        buttons = driver.find_elements(By.TAG_NAME, "button")

        found = False
        for b in buttons:
            txt = b.text.strip().lower()
            if "hae" in txt or "search" in txt:
                b.click()
                found = True
                break

        if not found:
            return "EI HAKUNAPPIA"

        # Odotetaan että sivu latautuu ja sähköposti kohta löytyy
        wait.until(EC.presence_of_element_located((By.TAG_NAME, "body")))

        # Haetaan koko sivun teksti
        page_text = driver.page_source

        # Etsitään sähköposti
        email_match = re.search(r"[a-zA-Z0-9_.+-]+@[a-zA-Z0-9-]+\.[a-zA-Z0-9-.]+", page_text)
        if email_match:
            return email_match.group(0)

        # Jos ei löytynyt sähköpostia
        return "EI SÄHKÖPOSTIA"

    except Exception as e:
        return f"VIRHE: {str(e)}"


def fetch_all_emails(ytunnukset, log_func):
    """
    Hakee sähköpostit kaikille Y-tunnuksille Seleniumilla.
    """
    log_func("Käynnistetään Chrome...")

    options = webdriver.ChromeOptions()
    options.add_argument("--disable-gpu")
    options.add_argument("--no-sandbox")
    options.add_argument("--start-maximized")

    driver = webdriver.Chrome(service=Service(ChromeDriverManager().install()), options=options)

    results = []

    try:
        for yt in ytunnukset:
            log_func(f"Haetaan Virrestä: {yt}")
            email = fetch_email_from_virre(driver, yt)
            results.append((yt, email))

    finally:
        driver.quit()

    return results


# -----------------------------
# GUI
# -----------------------------
class App(TkinterDnD.Tk):
    def __init__(self):
        super().__init__()

        self.title("Y-tunnus + Virre sähköposti botti")
        self.geometry("600x420")

        self.label = tk.Label(
            self,
            text="Vedä ja pudota PDF tähän ikkunaan",
            font=("Arial", 14),
            relief="ridge",
            width=50,
            height=4
        )
        self.label.pack(pady=20)

        self.log_box = tk.Text(self, height=12, width=70)
        self.log_box.pack(pady=10)

        self.label.drop_target_register(DND_FILES)
        self.label.dnd_bind("<<Drop>>", self.drop_file)

    def log(self, msg):
        self.log_box.insert(tk.END, msg + "\n")
        self.log_box.see(tk.END)
        self.update_idletasks()

    def drop_file(self, event):
        file_path = event.data.strip()

        # Jos Windows antaa { } ympärille
        if file_path.startswith("{") and file_path.endswith("}"):
            file_path = file_path[1:-1]

        if not file_path.lower().endswith(".pdf"):
            messagebox.showerror("Virhe", "Vain PDF-tiedostot sallittu!")
            return

        self.log(f"PDF vastaanotettu: {file_path}")
        threading.Thread(target=self.process_pdf, args=(file_path,), daemon=True).start()

    def process_pdf(self, pdf_path):
        try:
            self.log("Luetaan PDF...")
            ytunnukset = extract_ytunnukset_from_pdf(pdf_path)

            if not ytunnukset:
                self.log("Ei löytynyt yhtään Y-tunnusta.")
                return

            self.log(f"Löytyi {len(ytunnukset)} Y-tunnusta.")

            # Tallennetaan Y-tunnukset Word + Excel
            save_ytunnukset_word(ytunnukset, "ytunnukset.docx")
            save_ytunnukset_excel(ytunnukset, "ytunnukset.xlsx")

            self.log("Tallennettu: ytunnukset.docx")
            self.log("Tallennettu: ytunnukset.xlsx")

            # Virre sähköpostien haku
            self.log("Aloitetaan Virre sähköpostien haku... (Chrome avautuu)")
            results = fetch_all_emails(ytunnukset, self.log)

            # Tallennetaan sähköpostit Word + Excel
            save_emails_word(results, "virre_sahkopostit.docx")
            save_emails_excel(results, "virre_sahkopostit.xlsx")

            self.log("Tallennettu: virre_sahkopostit.docx")
            self.log("Tallennettu: virre_sahkopostit.xlsx")

            self.log("VALMIS!")

            messagebox.showinfo("Valmis", "Kaikki tiedostot luotu!\nKatso sama kansio missä exe on.")

        except Exception as e:
            messagebox.showerror("Virhe", str(e))
            self.log(f"VIRHE: {str(e)}")


if __name__ == "__main__":
    app = App()
    app.mainloop()
