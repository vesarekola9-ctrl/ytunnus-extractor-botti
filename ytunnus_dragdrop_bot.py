import re
import time
import tkinter as tk
from tkinter import messagebox
from PyPDF2 import PdfReader
from docx import Document
import openpyxl

# Selenium
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.options import Options


# --- PDF -> teksti ---
def extract_text_from_pdf(pdf_path):
    try:
        reader = PdfReader(pdf_path)
        full_text = ""
        for page in reader.pages:
            page_text = page.extract_text()
            if page_text:
                full_text += page_text + "\n"
        return full_text
    except Exception as e:
        print(f"Virhe PDF:n lukemisessa ({pdf_path}): {e}")
        return ""


# --- Etsi Y-tunnukset ---
def find_ytunnukset(text):
    yt1 = re.findall(r"\b\d{7}-\d\b", text)
    yt2 = re.findall(r"\b\d{8}\b", text)
    yt2_fixed = [y[:7] + "-" + y[7:] for y in yt2]
    all_yt = yt1 + yt2_fixed
    return sorted(list(dict.fromkeys(all_yt)))


# --- Tallenna Word ---
def save_to_word(lines, output_file, title="Tulokset"):
    doc = Document()
    doc.add_heading(title, level=1)
    for line in lines:
        doc.add_paragraph(line)
    doc.save(output_file)


# --- Tallenna Excel ---
def save_to_excel(rows, output_file, headers):
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Tulokset"
    ws.append(headers)
    for r in rows:
        ws.append(r)
    wb.save(output_file)


# --- Muunna virren sähköposti (a) -> @ ---
def normalize_email(email):
    email = email.strip()
    email = email.replace("(a)", "@")
    email = email.replace("[at]", "@")
    return email


# --- Hae sähköposti virrestä ---
def fetch_email_from_virre(driver, ytunnus):
    # tämä execution muuttuu aina, mutta virre ohjaa yleensä automaattisesti
    url = "https://virre.prh.fi/novus/home"
    driver.get(url)

    time.sleep(2)

    # Yritetään etsiä hakukenttä
    # (tämä voi muuttua, mutta toimii usein)
    try:
        search_input = driver.find_element(By.TAG_NAME, "input")
        search_input.clear()
        search_input.send_keys(ytunnus)
    except:
        return ""

    # Yritetään painaa Enter tai klikata hakunappia
    try:
        search_input.submit()
    except:
        pass

    time.sleep(3)

    page_text = driver.page_source

    # etsitään sähköposti sivulta
    emails = re.findall(r"[a-zA-Z0-9_.+-]+\s*\(a\)\s*[a-zA-Z0-9-]+\.[a-zA-Z0-9-.]+", page_text)
    if emails:
        return normalize_email(emails[0])

    # varalla suora @-muotoinen
    emails2 = re.findall(r"[a-zA-Z0-9_.+-]+@[a-zA-Z0-9-]+\.[a-zA-Z0-9-.]+", page_text)
    if emails2:
        return emails2[0]

    return ""


# --- Käsittele PDF:t ---
def process_pdfs(pdf_files):
    all_ytunnukset = []
    for pdf_file in pdf_files:
        text = extract_text_from_pdf(pdf_file)
        yt = find_ytunnukset(text)
        all_ytunnukset.extend(yt)

    if not all_ytunnukset:
        messagebox.showinfo("Valmis", "Ei löytynyt yhtään Y-tunnusta.")
        return

    all_ytunnukset = sorted(list(dict.fromkeys(all_ytunnukset)))

    # tallenna ytunnukset wordiin/exceliin
    save_to_word(all_ytunnukset, "ytunnukset.docx", title="Y-tunnukset")
    save_to_excel([[y] for y in all_ytunnukset], "ytunnukset.xlsx", ["Y-tunnus"])

    # --- Selenium käyntiin ---
    chrome_options = Options()
    chrome_options.add_argument("--headless=new")
    chrome_options.add_argument("--disable-gpu")
    chrome_options.add_argument("--window-size=1920,1080")

    driver = webdriver.Chrome(options=chrome_options)

    results = []
    word_lines = []

    try:
        for i, yt in enumerate(all_ytunnukset, start=1):
            print(f"[{i}/{len(all_ytunnukset)}] Haetaan sähköposti: {yt}")

            email = ""
            try:
                email = fetch_email_from_virre(driver, yt)
            except Exception as e:
                print("Virhe:", e)

            results.append([yt, email])

            if email:
                word_lines.append(f"{yt}  ->  {email}")
            else:
                word_lines.append(f"{yt}  ->  EI SÄHKÖPOSTIA")

            time.sleep(2)  # viive ettei hakata virreä

    finally:
        driver.quit()

    # tallenna sähköpostitulokset
    save_to_word(word_lines, "virre_sahkopostit.docx", title="Virre sähköpostit")
    save_to_excel(results, "virre_sahkopostit.xlsx", ["Y-tunnus", "Sähköposti"])

    messagebox.showinfo(
        "Valmis",
        f"Löytyi {len(all_ytunnukset)} Y-tunnusta.\n\n"
        "Tallennettu tiedostot:\n"
        "- ytunnukset.docx\n"
        "- ytunnukset.xlsx\n"
        "- virre_sahkopostit.docx\n"
        "- virre_sahkopostit.xlsx"
    )


# --- GUI ---
def on_drop(event):
    files = root.tk.splitlist(event.data)
    pdf_files = [f.strip("{}") for f in files if f.lower().endswith(".pdf")]
    if pdf_files:
        process_pdfs(pdf_files)


root = tk.Tk()
root.title("Y-tunnus + Virre sähköposti botti")
root.geometry("450x200")

label = tk.Label(
    root,
    text="Vedä ja pudota PDF-tiedostot tähän ikkunaan\nBotti kerää Y-tunnukset ja hakee sähköpostit Virrestä",
    wraplength=420,
    justify="center"
)
label.pack(expand=True)

try:
    import tkinterdnd2
    root = tkinterdnd2.TkinterDnD.Tk()
    root.title("Y-tunnus + Virre sähköposti botti")
    root.geometry("450x200")
    label = tk.Label(
        root,
        text="Vedä ja pudota PDF-tiedostot tähän ikkunaan\nBotti kerää Y-tunnukset ja hakee sähköpostit Virrestä",
        wraplength=420,
        justify="center"
    )
    label.pack(expand=True)

    root.drop_target_register(tkinterdnd2.DND_FILES)
    root.dnd_bind("<<Drop>>", on_drop)

except ImportError:
    pass

root.mainloop()
