import re
import time
import tkinter as tk
from tkinter import messagebox
from PyPDF2 import PdfReader
from docx import Document
import openpyxl

from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys


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


# --- Word tallennus ---
def save_to_word(lines, output_file, title="Tulokset"):
    doc = Document()
    doc.add_heading(title, level=1)
    for line in lines:
        doc.add_paragraph(line)
    doc.save(output_file)


# --- Excel tallennus ---
def save_to_excel(rows, output_file, headers):
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Tulokset"
    ws.append(headers)
    for r in rows:
        ws.append(r)
    wb.save(output_file)


# --- sähköpostin muoto (a) -> @ ---
def normalize_email(email):
    return email.replace("(a)", "@").replace(" ", "").strip()


# --- Etsi sähköposti HTML:stä ---
def find_email_from_html(html):
    # Virre näyttää usein info(a)domain.fi
    match = re.search(r"[a-zA-Z0-9_.+-]+\s*\(a\)\s*[a-zA-Z0-9-]+\.[a-zA-Z0-9-.]+", html)
    if match:
        return normalize_email(match.group(0))

    match2 = re.search(r"[a-zA-Z0-9_.+-]+@[a-zA-Z0-9-]+\.[a-zA-Z0-9-.]+", html)
    if match2:
        return match2.group(0)

    return ""


# --- Virre haku ---
def fetch_email_from_virre(driver, ytunnus):
    driver.get("https://virre.prh.fi/novus/home")

    time.sleep(3)

    # etsitään hakukenttä (usein ensimmäinen text input)
    inputs = driver.find_elements(By.TAG_NAME, "input")

    search_box = None
    for inp in inputs:
        try:
            t = inp.get_attribute("type")
            if t and t.lower() in ["text", "search"]:
                search_box = inp
                break
        except:
            continue

    if not search_box:
        return ""

    search_box.clear()
    search_box.send_keys(ytunnus)
    search_box.send_keys(Keys.ENTER)

    time.sleep(4)

    # nyt pitäisi olla yrityssivu tai hakutulos
    html = driver.page_source

    email = find_email_from_html(html)
    return email


# --- pääprosessi ---
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

    save_to_word(all_ytunnukset, "ytunnukset.docx", title="Y-tunnukset")
    save_to_excel([[y] for y in all_ytunnukset], "ytunnukset.xlsx", ["Y-tunnus"])

    # --- Selenium käynnistys näkyvänä ---
    try:
        driver = webdriver.Chrome()
    except Exception as e:
        messagebox.showerror("Virhe", f"Selenium ei käynnistynyt.\n\nVirhe:\n{e}")
        return

    results_excel = []
    results_word = []

    try:
        for i, yt in enumerate(all_ytunnukset, start=1):
            print(f"[{i}/{len(all_ytunnukset)}] Haetaan sähköposti: {yt}")

            try:
                email = fetch_email_from_virre(driver, yt)
            except Exception as e:
                email = ""
                print("Virhe:", e)

            results_excel.append([yt, email])

            if email:
                results_word.append(f"{yt} -> {email}")
            else:
                results_word.append(f"{yt} -> EI SÄHKÖPOSTIA")

            time.sleep(2)

    finally:
        driver.quit()

    save_to_word(results_word, "virre_sahkopostit.docx", title="Virre sähköpostit")
    save_to_excel(results_excel, "virre_sahkopostit.xlsx", ["Y-tunnus", "Sähköposti"])

    messagebox.showinfo(
        "Valmis",
        f"Valmis!\n\nTallennettu:\n"
        f"- ytunnukset.docx\n"
        f"- ytunnukset.xlsx\n"
        f"- virre_sahkopostit.docx\n"
        f"- virre_sahkopostit.xlsx"
    )


# --- GUI Drop ---
def on_drop(event):
    files = root.tk.splitlist(event.data)
    pdf_files = [f.strip("{}") for f in files if f.lower().endswith(".pdf")]
    if pdf_files:
        process_pdfs(pdf_files)


root = tk.Tk()
root.title("Y-tunnus + Virre sähköposti botti")
root.geometry("500x220")

label = tk.Label(
    root,
    text="Vedä ja pudota PDF tähän.\n\nBotti kerää Y-tunnukset ja hakee sähköpostit Virre-palvelusta.",
    wraplength=480,
    justify="center"
)
label.pack(expand=True)

try:
    import tkinterdnd2
    root = tkinterdnd2.TkinterDnD.Tk()
    root.title("Y-tunnus + Virre sähköposti botti")
    root.geometry("500x220")

    label = tk.Label(
        root,
        text="Vedä ja pudota PDF tähän.\n\nBotti kerää Y-tunnukset ja hakee sähköpostit Virre-palvelusta.",
        wraplength=480,
        justify="center"
    )
    label.pack(expand=True)

    root.drop_target_register(tkinterdnd2.DND_FILES)
    root.dnd_bind("<<Drop>>", on_drop)

except ImportError:
    pass

root.mainloop()

root.mainloop()
