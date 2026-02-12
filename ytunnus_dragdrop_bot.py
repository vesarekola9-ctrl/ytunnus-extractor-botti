import os
import re
import threading
import tkinter as tk
from tkinter import messagebox

from tkinterdnd2 import DND_FILES, TkinterDnD

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
# LOG TIEDOSTO
# -----------------------------
def write_log_to_file(msg):
    try:
        with open("log.txt", "a", encoding="utf-8") as f:
            f.write(msg + "\n")
    except:
        pass


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

            matches = re.findall(r"\b\d{7}-\d\b|\b\d{8}\b", text)

            for m in matches:
                fixed = fix_ytunnus(m)
                if fixed:
                    ytunnukset.add(fixed)

    return sorted(list(ytunnukset))


# -----------------------------
# Tallennus Word + Excel
# -----------------------------
def save_word_list(lines, filename, title):
    doc = Document()
    doc.add_heading(title, level=1)

    for line in lines:
        doc.add_paragraph(line)

    doc.save(filename)


def save_excel_table(rows, filename, headers):
    wb = openpyxl.Workbook()
    ws = wb.active

    for col, h in enumerate(headers, start=1):
        ws.cell(row=1, column=col).value = h

    for r, row_data in enumerate(rows, start=2):
        for c, value in enumerate(row_data, start=1):
            ws.cell(row=r, column=c).value = value

    wb.save(filename)


# -----------------------------
# VIRRE HAKU
# -----------------------------
def virre_get_email(driver, ytunnus, log_func):
    wait = WebDriverWait(driver, 25)

    driver.get("https://virre.prh.fi/novus/home")

    try:
        log_func("Etsitään hakukenttä...")

        search_input = wait.until(
            EC.presence_of_element_located((By.CSS_SELECTOR, "input[type='text']"))
        )

        search_input.clear()
        search_input.send_keys(ytunnus)

        log_func("Klikataan HAE...")

        hae_btn = wait.until(
            EC.element_to_be_clickable((By.XPATH, "//button[contains(., 'HAE')]"))
        )
        hae_btn.click()

        log_func("Odotetaan tulossivua...")

        wait.until(
            EC.presence_of_element_located((By.TAG_NAME, "body"))
        )

        page_source = driver.page_source

        email_match = re.search(
            r"[a-zA-Z0-9_.+-]+@[a-zA-Z0-9-]+\.[a-zA-Z0-9-.]+",
            page_source
        )

        if email_match:
            log_func(f"Sähköposti löytyi: {email_match.group(0)}")
            return email_match.group(0)

        log_func("Ei sähköpostia löytynyt.")
        return "EI SÄHKÖPOSTIA"

    except Exception as e:
        log_func(f"Virre virhe: {str(e)}")
        return f"VIRHE: {str(e)}"


def fetch_all_emails_from_virre(ytunnukset, log_func):
    log_func("Käynnistetään Chrome...")

    try:
        options = webdriver.ChromeOptions()
        options.add_argument("--start-maximized")
        options.add_argument("--disable-gpu")
        options.add_argument("--disable-dev-shm-usage")

        driver_path = ChromeDriverManager().install()
        log_func(f"ChromeDriver polku: {driver_path}")

        driver = webdriver.Chrome(service=Service(driver_path), options=options)

    except Exception as e:
        log_func("CHROME EI KÄYNNISTY!")
        log_func(str(e))
        return [(yt, "CHROME EI KÄYNNISTY") for yt in ytunnukset]

    results = []

    try:
        for yt in ytunnukset:
            log_func(f"Haku Virrestä: {yt}")
            email = virre_get_email(driver, yt, log_func)
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

        self.title("Protestilista botti (Y-tunnukset + Virre sähköpostit)")
        self.geometry("720x520")

        self.drop_label = tk.Label(
            self,
            text="VEDÄ JA PUDOTA PDF TÄHÄN",
            font=("Arial", 16, "bold"),
            bg="lightgray",
            relief="ridge",
            width=60,
            height=5
        )
        self.drop_label.pack(pady=20)

        self.log_box = tk.Text(self, height=18, width=85)
        self.log_box.pack(pady=10)

        self.drop_target_register(DND_FILES)
        self.dnd_bind("<<Drop>>", self.drop_file)

        self.drop_label.drop_target_register(DND_FILES)
        self.drop_label.dnd_bind("<<Drop>>", self.drop_file)

        # Tyhjennetään vanha logi aina kun botti avataan
        try:
            with open("log.txt", "w", encoding="utf-8") as f:
                f.write("=== BOTTI KÄYNNISTETTY ===\n")
        except:
            pass

    def log(self, msg):
        try:
            self.log_box.insert(tk.END, msg + "\n")
            self.log_box.see(tk.END)
            self.update_idletasks()
        except:
            pass

        write_log_to_file(msg)

    def drop_file(self, event):
        file_path = event.data.strip()

        if file_path.startswith("{") and file_path.endswith("}"):
            file_path = file_path[1:-1]

        self.log(f"PDF vastaanotettu: {file_path}")

        if not file_path.lower().endswith(".pdf"):
            self.log("VIRHE: Ei PDF tiedosto.")
            messagebox.showerror("Virhe", "Vain PDF-tiedostot sallittu!")
            return

        threading.Thread(target=self.process_pdf, args=(file_path,), daemon=True).start()

    def process_pdf(self, pdf_path):
        try:
            self.log("Luetaan PDF ja etsitään Y-tunnukset...")

            ytunnukset = extract_ytunnukset_from_pdf(pdf_path)

            if not ytunnukset:
                self.log("Ei löytynyt yhtään Y-tunnusta.")
                messagebox.showwarning("Ei löytynyt", "PDF:stä ei löytynyt yhtään Y-tunnusta.")
                return

            self.log(f"Löytyi {len(ytunnukset)} Y-tunnusta.")

            save_word_list(ytunnukset, "ytunnukset.docx", "Y-tunnukset")
            save_excel_table([(yt,) for yt in ytunnukset], "ytunnukset.xlsx", ["Y-tunnus"])

            self.log("Tallennettu: ytunnukset.docx")
            self.log("Tallennettu: ytunnukset.xlsx")

            self.log("Aloitetaan Virre haku... Chrome avautuu nyt.")
            results = fetch_all_emails_from_virre(ytunnukset, self.log)

            word_lines = [f"{yt} -> {email}" for yt, email in results]

            save_word_list(word_lines, "virre_sahkopostit.docx", "Virre sähköpostit")
            save_excel_table(results, "virre_sahkopostit.xlsx", ["Y-tunnus", "Sähköposti"])

            self.log("Tallennettu: virre_sahkopostit.docx")
            self.log("Tallennettu: virre_sahkopostit.xlsx")

            self.log("VALMIS!")
            messagebox.showinfo("Valmis", "Valmis!\nTiedostot löytyvät exe:n kansiosta.")

        except Exception as e:
            self.log(f"VIRHE: {str(e)}")
            messagebox.showerror("Virhe", str(e))


if __name__ == "__main__":
    app = App()
    app.mainloop()
