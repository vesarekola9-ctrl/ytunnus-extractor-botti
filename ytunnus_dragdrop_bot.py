import re
import tkinter as tk
from tkinter import messagebox
from PyPDF2 import PdfReader
from docx import Document
import openpyxl

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
    yt1 = re.findall(r"\b\d{7}-\d\b", text)  # jo viivalliset
    yt2 = re.findall(r"\b\d{8}\b", text)     # ilman viivaa
    yt2_fixed = [y[:7] + "-" + y[7:] for y in yt2]
    all_yt = yt1 + yt2_fixed
    return sorted(list(dict.fromkeys(all_yt)))  # poista duplikaatit

# --- Lisää yksi viiva tarvittaessa ---
def add_extra_dash(ytunnus):
    if "-" in ytunnus:
        return ytunnus  # kopioidaan sellaisenaan
    else:
        return ytunnus[:7] + "-" + ytunnus[7:]

# --- Tallenna Word ---
def save_to_word(ytunnukset, output_file):
    doc = Document()
    doc.add_heading("Y-tunnukset", level=1)
    for y in ytunnukset:
        doc.add_paragraph(y)
    doc.save(output_file)

# --- Tallenna Excel ---
def save_to_excel(ytunnukset, output_file):
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Y-tunnukset"
    ws.append(["Y-tunnukset"])
    for y in ytunnukset:
        ws.append([y])
    wb.save(output_file)

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
    all_ytunnukset_modified = [add_extra_dash(y) for y in all_ytunnukset]

    word_file = "ytunnukset.docx"
    excel_file = "ytunnukset.xlsx"
    save_to_word(all_ytunnukset_modified, word_file)
    save_to_excel(all_ytunnukset_modified, excel_file)

    messagebox.showinfo(
        "Valmis",
        f"Löytyi {len(all_ytunnukset_modified)} Y-tunnusta.\n"
        f"Tallennettu:\n - {word_file}\n - {excel_file}"
    )

# --- GUI ---
def on_drop(event):
    files = root.tk.splitlist(event.data)
    pdf_files = [f.strip("{}") for f in files if f.lower().endswith(".pdf")]
    if pdf_files:
        process_pdfs(pdf_files)

root = tk.Tk()
root.title("Y-tunnus Drag & Drop Botti")
root.geometry("450x200")

label = tk.Label(
    root,
    text="Vedä ja pudota PDF‑tiedostot tähän ikkunaan",
    wraplength=400,
    justify="center"
)
label.pack(expand=True)

# Drag & drop tuki
try:
    import tkinterdnd2
    root = tkinterdnd2.TkinterDnD.Tk()
    label.pack(expand=True)
    root.drop_target_register(tkinterdnd2.DND_FILES)
    root.dnd_bind('<<Drop>>', on_drop)
except ImportError:
    from tkinter import filedialog
    def fallback():
        pdf_files = filedialog.askopenfilenames(filetypes=[("PDF Files", "*.pdf")])
        if pdf_files:
            process_pdfs(pdf_files)
    button = tk.Button(root, text="Valitse PDF‑tiedostot", command=fallback)
    button.pack(pady=20)

root.mainloop()
