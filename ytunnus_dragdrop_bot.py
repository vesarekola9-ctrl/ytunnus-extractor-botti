import os
import re
import sys
import time
import threading
import tkinter as tk
from tkinter import messagebox, filedialog
from tkinter import ttk

import PyPDF2
from docx import Document
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from webdriver_manager.chrome import ChromeDriverManager


# =========================
#   CONFIG / REGEX
# =========================
YT_RE = re.compile(r"\b\d{7}-\d\b|\b\d{8}\b")

# Yritysnimi: protestilistassa nimet ovat usein omalla rivillä.
# Tämä regex poimii “nimi-rivejä” ja suodattaa selviä taulukko-otsikoita.
BAD_LINES = {
    "yritys", "sijainti", "summa", "häiriöpäivä", "tyyppi", "lähde",
    "viimeisimmät protestit", "protestilista",
    "y-tunnus", "julkaisupäivä", "alue", "velkoja",
}

# YTJ
YTJ_COMPANY_URL = "https://tietopalvelu.ytj.fi/yritys/{}"
EMAIL_RE = re.compile(r"[A-Za-z0-9_.+-]+@[A-Za-z0-9-]+\.[A-Za-z0-9-.]+")
EMAIL_A_RE = re.compile(r"[A-Za-z0-9_.+-]+\s*\(a\)\s*[A-Za-z0-9-]+\.[A-Za-z0-9-.]+", re.I)


# =========================
#   PATHS / LOG
# =========================
def get_exe_dir():
    if getattr(sys, "frozen", False):
        return os.path.dirname(sys.executable)
    return os.path.dirname(os.path.abspath(__file__))


def get_output_dir():
    base = get_exe_dir()
    try:
        p = os.path.join(base, "_write_test.tmp")
        with open(p, "w", encoding="utf-8") as f:
            f.write("ok")
        os.remove(p)
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


def log_to_file(msg: str):
    ts = time.strftime("%H:%M:%S")
    line = f"[{ts}] {msg}"
    try:
        with open(LOG_PATH, "a", encoding="utf-8") as f:
            f.write(line + "\n")
    except Exception:
        pass
    return line


def reset_log():
    try:
        with open(LOG_PATH, "w", encoding="utf-8") as f:
            f.write("=== BOTTI KÄYNNISTETTY ===\n")
    except Exception:
        pass
    log_to_file(f"Output: {OUT_DIR}")
    log_to_file(f"Logi: {LOG_PATH}")


# =========================
#   UTIL
# =========================
def normalize_yt(yt: str):
    yt = (yt or "").strip().replace(" ", "")
    if re.fullmatch(r"\d{7}-\d", yt):
        return yt
    if re.fullmatch(r"\d{8}", yt):
        return yt[:7] + "-" + yt[7]
    return None


def save_word_plain_lines(lines, filename):
    path = os.path.join(OUT_DIR, filename)
    doc = Document()
    for line in lines:
        if line and str(line).strip():
            doc.add_paragraph(str(line).strip())
    doc.save(path)
    return path


def pick_email_from_text(text: str) -> str:
    if not text:
        return ""
    m = EMAIL_RE.search(text)
    if m:
        return m.group(0).strip().replace(" ", "")
    m2 = EMAIL_A_RE.search(text)
    if m2:
        return m2.group(0).replace(" ", "").replace("(a)", "@")
    return ""


# =========================
#   MODE 1: KL COPY-PASTE PARSER
# =========================
def parse_company_names_from_copied_text(raw: str):
    """
    Idea: käyttäjä kopioi protestilistan sisällön (Ctrl+A/Ctrl+C) ja liittää tänne.
    Poimitaan yritysnimet: otetaan rivit, jotka:
      - eivät ole otsikkoja
      - eivät näytä euro-summilta / päivämääriltä / tyypeiltä / lähteiltä
      - ovat “nimi-tyyppisiä” (sisältää kirjaimia, eikä ole liian lyhyt)
    """
    if not raw:
        return []

    lines = [ln.strip() for ln in raw.splitlines() if ln.strip()]
    names = []
    seen = set()

    for ln in lines:
        low = ln.lower()

        # suodata otsikoita / selviä kenttiä
        if low in BAD_LINES:
            continue
        if low.startswith("viimeisimmät protestit"):
            continue

        # suodata rahasummat, päivät, tyypit, lähteet
        if re.fullmatch(r"\d{1,3}(\s?\d{3})*\s?€", ln):  # 1 234 €
            continue
        if re.fullmatch(r"\d{1,2}\.\d{1,2}\.\d{4}", ln):  # 18.02.2026
            continue
        if "dun" in low and "bradstreet" in low:
            continue
        if "velkomustuomiot" in low or "ulosotto" in low or "konkurssi" in low:
            continue

        # jos rivi on y-tunnus tai sisältää y-tunnus, ohita nimi-listasta
        if "y-tunnus" in low:
            continue
        if YT_RE.search(ln):
            continue

        # nimiheuristiikka
        if len(ln) < 3:
            continue
        if not re.search(r"[A-Za-zÅÄÖåäö]", ln):
            continue

        # poista ihan selviä “alue/sijainti” -rivejä (pelkät kaupungit)
        # tämä on karkea: jos rivillä on vain yksi sana ja se alkaa isolla, skip
        if len(ln.split()) == 1 and ln[:1].isupper():
            # esim "Helsinki", "Turku"
            # (tämä voi joskus poistaa yhden sanan yrityksiä, mutta protestilistassa harvinaista)
            continue

        key = ln.strip().lower()
        if key not in seen:
            seen.add(key)
            names.append(ln.strip())

    return names


# =========================
#   MODE 2: PDF -> YTs -> YTJ emails
# =========================
def extract_ytunnukset_from_pdf(pdf_path: str):
    yt_set = set()
    with open(pdf_path, "rb") as f:
        reader = PyPDF2.PdfReader(f)
        for page in reader.pages:
            text = page.extract_text() or ""
            for m in YT_RE.findall(text):
                n = normalize_yt(m)
                if n:
                    yt_set.add(n)
    return sorted(yt_set)


def start_new_driver():
    options = webdriver.ChromeOptions()
    options.add_argument("--start-maximized")
    options.add_argument("--disable-gpu")
    options.add_argument("--disable-dev-shm-usage")
    driver_path = ChromeDriverManager().install()
    return webdriver.Chrome(service=Service(driver_path), options=options)


def extract_email_from_ytj_driver(driver):
    try:
        for a in driver.find_elements("tag name", "a"):
            href = (a.get_attribute("href") or "")
            if href.lower().startswith("mailto:"):
                return href.split(":", 1)[1].strip()
    except Exception:
        pass

    try:
        body = driver.find_element("tag name", "body").text or ""
        return pick_email_from_text(body)
    except Exception:
        return ""


def fetch_emails_from_ytj(driver, yt_list, status_cb, progress_cb, log_cb, stop_evt):
    emails = []
    seen = set()
    progress_cb(0, max(1, len(yt_list)))

    for i, yt in enumerate(yt_list, start=1):
        if stop_evt.is_set():
            break

        status_cb(f"YTJ: {i}/{len(yt_list)} {yt}")
        progress_cb(i - 1, len(yt_list))

        driver.get(YTJ_COMPANY_URL.format(yt))
        time.sleep(1.0)

        email = ""
        for _ in range(10):
            email = extract_email_from_ytj_driver(driver)
            if email:
                break
            time.sleep(0.2)

        if email:
            k = email.lower()
            if k not in seen:
                seen.add(k)
                emails.append(email)
                log_cb(email)

    progress_cb(len(yt_list), max(1, len(yt_list)))
    return emails


# =========================
#   GUI
# =========================
class App(tk.Tk):
    def __init__(self):
        super().__init__()
        reset_log()

        self.stop_evt = threading.Event()
        self.worker = None

        self.title("ProtestiBotti (YKSINKERTAINEN: KL copy/paste -> yritysnimet + PDF->YTJ)")
        self.geometry("1040x760")

        tk.Label(self, text="ProtestiBotti", font=("Arial", 18, "bold")).pack(pady=8)

        tk.Label(
            self,
            text="Moodit:\n"
                 "1) Kauppalehti: SINÄ avaat Chromen & protestilistan → Ctrl+A / Ctrl+C → liitä tähän → bot tallentaa yritysnimet Wordiin\n"
                 "2) PDF → Y-tunnukset → YTJ sähköpostit → Word",
            justify="center"
        ).pack(pady=4)

        # ----- MODE 1 UI -----
        box = tk.LabelFrame(self, text="1) Kauppalehti (copy/paste)", padx=10, pady=10)
        box.pack(fill="x", padx=12, pady=10)

        tk.Label(
            box,
            text="Avaa protestilista Chromessa (kirjaudu itse). Valitse sivu → Ctrl+A → Ctrl+C.\n"
                 "Liitä sisältö tähän ja paina 'Poimi yritysnimet'.",
            justify="left"
        ).pack(anchor="w")

        self.text = tk.Text(box, height=10)
        self.text.pack(fill="x", pady=6)

        row = tk.Frame(box)
        row.pack(fill="x")

        tk.Button(row, text="Poimi yritysnimet → Word", font=("Arial", 12, "bold"), command=self.start_parse_names).pack(side="left")
        tk.Button(row, text="Tyhjennä", command=lambda: self.text.delete("1.0", tk.END)).pack(side="left", padx=8)

        # ----- MODE 2 UI -----
        box2 = tk.LabelFrame(self, text="2) PDF → YTJ", padx=10, pady=10)
        box2.pack(fill="x", padx=12, pady=10)

        tk.Button(box2, text="Valitse PDF → YTJ sähköpostit", font=("Arial", 12, "bold"), command=self.start_pdf_mode).pack(anchor="w")

        self.status = tk.Label(self, text="Valmiina.", font=("Arial", 11))
        self.status.pack(pady=6)

        self.progress = ttk.Progressbar(self, orient="horizontal", mode="determinate", length=980)
        self.progress.pack(pady=6)

        frame = tk.Frame(self)
        frame.pack(fill="both", expand=True, padx=14, pady=10)

        tk.Label(frame, text="Live-logi (uusimmat alimmaisena):").pack(anchor="w")
        self.listbox = tk.Listbox(frame, height=16)
        self.listbox.pack(side="left", fill="both", expand=True)

        sb = tk.Scrollbar(frame, orient="vertical", command=self.listbox.yview)
        sb.pack(side="right", fill="y")
        self.listbox.configure(yscrollcommand=sb.set)

        tk.Label(self, text=f"Tallennus: {OUT_DIR}", wraplength=1020, justify="center").pack(pady=6)

    def ui_log(self, msg):
        line = log_to_file(msg)
        self.listbox.insert(tk.END, line)
        self.listbox.yview_moveto(1.0)
        self.update_idletasks()

    def set_status(self, s):
        self.status.config(text=s)
        self.update_idletasks()
        self.ui_log(s)

    def set_progress(self, value, maximum):
        self.progress["maximum"] = max(1, maximum)
        self.progress["value"] = value
        self.update_idletasks()

    # ---- MODE 1 ----
    def start_parse_names(self):
        raw = self.text.get("1.0", tk.END)
        names = parse_company_names_from_copied_text(raw)

        if not names:
            self.set_status("Ei löytynyt yritysnimiä liitetystä tekstistä.")
            messagebox.showwarning("Ei löytynyt", "Liitetystä tekstistä ei löytynyt yritysnimiä. Kopioithan koko listan?")
            return

        path = save_word_plain_lines(names, "yritysnimet_kauppalehti.docx")
        self.set_status(f"Valmis: {len(names)} yritystä")
        self.ui_log(f"Tallennettu: {path}")
        messagebox.showinfo("Valmis", f"Poimittu {len(names)} yritysnimeä.\n\n{path}")

    # ---- MODE 2 ----
    def start_pdf_mode(self):
        path = filedialog.askopenfilename(filetypes=[("PDF files", "*.pdf")])
        if not path:
            return
        threading.Thread(target=self.run_pdf_mode, args=(path,), daemon=True).start()

    def run_pdf_mode(self, pdf_path):
        driver = None
        try:
            self.stop_evt.clear()
            self.set_status("Luetaan PDF ja kerätään Y-tunnukset…")
            yt_list = extract_ytunnukset_from_pdf(pdf_path)

            if not yt_list:
                self.set_status("Ei löytynyt Y-tunnuksia PDF:stä.")
                messagebox.showwarning("Ei löytynyt", "PDF:stä ei löytynyt yhtään Y-tunnusta.")
                return

            yt_path = save_word_plain_lines(yt_list, "ytunnukset_pdf.docx")
            self.ui_log(f"Tallennettu: {yt_path}")

            self.set_status("Käynnistetään Chrome ja haetaan sähköpostit YTJ:stä…")
            driver = start_new_driver()

            emails = fetch_emails_from_ytj(driver, yt_list, self.set_status, self.set_progress, self.ui_log, self.stop_evt)
            em_path = save_word_plain_lines(emails, "sahkopostit_pdf_ytj.docx")
            self.ui_log(f"Tallennettu: {em_path}")

            self.set_status("Valmis!")
            messagebox.showinfo(
                "Valmis",
                f"Valmis!\n\nKansio:\n{OUT_DIR}\n\nY-tunnuksia: {len(yt_list)}\nSähköposteja: {len(emails)}"
            )

        except Exception as e:
            self.ui_log(f"VIRHE: {e}")
            self.set_status("Virhe. Katso log.txt")
            messagebox.showerror("Virhe", f"Tuli virhe.\nKatso log.txt:\n{LOG_PATH}\n\n{e}")
        finally:
            if driver:
                try:
                    driver.quit()
                except Exception:
                    pass


if __name__ == "__main__":
    App().mainloop()
