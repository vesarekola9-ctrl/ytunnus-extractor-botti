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

from reportlab.lib.pagesizes import A4
from reportlab.pdfgen import canvas

from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from webdriver_manager.chrome import ChromeDriverManager


# =========================
#   REGEX
# =========================
YT_RE = re.compile(r"\b\d{7}-\d\b|\b\d{8}\b")
EMAIL_RE = re.compile(r"[A-Za-z0-9_.+-]+@[A-Za-z0-9-]+\.[A-Za-z0-9-.]+")
EMAIL_A_RE = re.compile(r"[A-Za-z0-9_.+-]+\s*\(a\)\s*[A-Za-z0-9-]+\.[A-Za-z0-9-.]+", re.I)

# YTJ
YTJ_HOME = "https://tietopalvelu.ytj.fi/"
YTJ_COMPANY_URL = "https://tietopalvelu.ytj.fi/yritys/{}"

BAD_LINES = {
    "yritys", "sijainti", "summa", "häiriöpäivä", "tyyppi", "lähde",
    "viimeisimmät protestit", "protestilista",
    "y-tunnus", "julkaisupäivä", "alue", "velkoja",
}
BAD_CONTAINS = [
    "velkomustuomiot", "ulosotto", "konkurssi", "dun & bradstreet", "bisnode",
    "protestit", "protestia",
]


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


def save_word_plain_lines(lines, filename):
    path = os.path.join(OUT_DIR, filename)
    doc = Document()
    for line in lines:
        if line and str(line).strip():
            doc.add_paragraph(str(line).strip())
    doc.save(path)
    return path


def save_pdf_lines(lines, filename, title=None):
    path = os.path.join(OUT_DIR, filename)
    c = canvas.Canvas(path, pagesize=A4)
    w, h = A4
    y = h - 60

    if title:
        c.setFont("Helvetica-Bold", 14)
        c.drawString(50, y, title)
        y -= 28

    c.setFont("Helvetica", 11)

    for line in lines:
        if not line:
            continue
        # sivunvaihto
        if y < 60:
            c.showPage()
            c.setFont("Helvetica", 11)
            y = h - 60
        c.drawString(50, y, str(line)[:140])
        y -= 16

    c.save()
    return path


# =========================
#   MODE 1: KL COPY/PASTE -> NAMES + YTS
# =========================
def parse_names_and_yts_from_copied_text(raw: str):
    if not raw:
        return [], []

    lines = [ln.strip() for ln in raw.splitlines() if ln.strip()]
    names = []
    yts = []
    seen_names = set()
    seen_yts = set()

    # poimi y-tunnukset kaikkialta
    for m in YT_RE.findall(raw):
        n = normalize_yt(m)
        if n and n not in seen_yts:
            seen_yts.add(n)
            yts.append(n)

    for ln in lines:
        low = ln.lower()

        if low in BAD_LINES:
            continue
        if any(b in low for b in BAD_CONTAINS):
            continue
        if low.startswith("viimeisimmät protestit"):
            continue

        # rahasummat / päivämäärät pois
        if re.fullmatch(r"\d{1,3}(\s?\d{3})*\s?€", ln):
            continue
        if re.fullmatch(r"\d{1,2}\.\d{1,2}\.\d{4}", ln):
            continue

        # jos rivi sisältää y-tunnuksen tai on y-tunnuslabel, ohita nimistä
        if "y-tunnus" in low:
            continue
        if YT_RE.search(ln):
            continue

        # nimiheuristiikka
        if len(ln) < 3:
            continue
        if not re.search(r"[A-Za-zÅÄÖåäö]", ln):
            continue

        # pudota selviä sijainteja (yksisanaiset kaupungit jne.)
        if len(ln.split()) == 1 and ln[:1].isupper():
            continue

        key = ln.strip().lower()
        if key not in seen_names:
            seen_names.add(key)
            names.append(ln.strip())

    return names, yts


# =========================
#   MODE 2: PDF -> (NAMES + YTS) -> YTJ EMAILS
# =========================
def extract_names_and_yts_from_pdf(pdf_path: str):
    text_all = ""
    with open(pdf_path, "rb") as f:
        reader = PyPDF2.PdfReader(f)
        for page in reader.pages:
            text_all += (page.extract_text() or "") + "\n"

    names, yts = parse_names_and_yts_from_copied_text(text_all)
    return names, yts


def start_new_driver():
    options = webdriver.ChromeOptions()
    options.add_argument("--start-maximized")
    options.add_argument("--disable-gpu")
    options.add_argument("--disable-dev-shm-usage")
    driver_path = ChromeDriverManager().install()
    return webdriver.Chrome(service=Service(driver_path), options=options)


def wait_body(driver, timeout=20):
    WebDriverWait(driver, timeout).until(EC.presence_of_element_located((By.TAG_NAME, "body")))


def ytj_open_company_by_name(driver, name: str, status_cb):
    """
    Avaa YTJ ja hakee nimellä.
    Valitsee ensimmäisen osuman listasta (best effort).
    """
    status_cb(f"YTJ: haku nimellä: {name}")

    driver.get(YTJ_HOME)
    wait_body(driver, 25)
    time.sleep(0.4)

    # etsi hakukenttä (YTJ:n etusivulla on yleensä yksi iso input)
    candidates = []
    for css in ["input[type='search']", "input[type='text']"]:
        try:
            candidates += driver.find_elements(By.CSS_SELECTOR, css)
        except Exception:
            pass

    search_input = None
    for inp in candidates:
        try:
            if inp.is_displayed() and inp.is_enabled():
                ph = (inp.get_attribute("placeholder") or "").lower()
                # hyväksy vaikka ei placeholderia, jos näkyy
                search_input = inp
                # jos placeholder viittaa hakuun, priorisoi
                if "hae" in ph or "lei" in ph or "y-tunnus" in ph or "nimellä" in ph:
                    break
        except Exception:
            continue

    if not search_input:
        return False

    try:
        search_input.clear()
    except Exception:
        pass
    search_input.send_keys(name)
    search_input.send_keys(Keys.ENTER)

    # odota että tuloksia tulee tai suoraan yrityssivulle
    time.sleep(1.2)

    # jos mennään suoraan yrityssivulle, ok
    if "/yritys/" in (driver.current_url or ""):
        return True

    # muuten klikkaa ensimmäinen tuloslinkki (best effort)
    links = []
    try:
        links = driver.find_elements(By.XPATH, "//a[contains(@href,'/yritys/')]")
    except Exception:
        links = []

    for a in links:
        try:
            if a.is_displayed():
                a.click()
                time.sleep(0.9)
                if "/yritys/" in (driver.current_url or ""):
                    return True
        except Exception:
            continue

    return False


def extract_email_from_ytj_page(driver):
    # mailto
    try:
        for a in driver.find_elements(By.TAG_NAME, "a"):
            href = (a.get_attribute("href") or "")
            if href.lower().startswith("mailto:"):
                return href.split(":", 1)[1].strip()
    except Exception:
        pass

    # body regex
    try:
        body = driver.find_element(By.TAG_NAME, "body").text or ""
        return pick_email_from_text(body)
    except Exception:
        return ""


def fetch_emails_from_ytj_by_yts_and_names(driver, yts, names, status_cb, progress_cb, log_cb, stop_evt):
    """
    1) Y-tunnukset (varma)
    2) jos nimiä jäi, haetaan nimellä YTJ:stä (best effort)
    """
    emails = []
    seen = set()

    total = len(yts) + len(names)
    progress_cb(0, max(1, total))
    done = 0

    # 1) Y-tunnuksilla
    for yt in yts:
        if stop_evt.is_set():
            break
        done += 1
        status_cb(f"YTJ (Y-tunnus): {yt} ({done}/{total})")
        progress_cb(done, total)

        driver.get(YTJ_COMPANY_URL.format(yt))
        wait_body(driver, 25)
        time.sleep(0.6)

        email = ""
        for _ in range(10):
            email = extract_email_from_ytj_page(driver)
            if email:
                break
            time.sleep(0.2)

        if email:
            k = email.lower()
            if k not in seen:
                seen.add(k)
                emails.append(email)
                log_cb(email)

    # 2) Nimillä (best effort)
    for nm in names:
        if stop_evt.is_set():
            break
        done += 1
        progress_cb(done, total)

        ok = ytj_open_company_by_name(driver, nm, status_cb)
        if not ok:
            log_cb(f"YTJ: ei löytynyt osumaa nimelle: {nm}")
            continue

        email = ""
        for _ in range(10):
            email = extract_email_from_ytj_page(driver)
            if email:
                break
            time.sleep(0.2)

        if email:
            k = email.lower()
            if k not in seen:
                seen.add(k)
                emails.append(email)
                log_cb(email)

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

        self.title("ProtestiBotti (KL copy/paste -> PDF + PDF->YTJ)")
        self.geometry("1080x820")

        tk.Label(self, text="ProtestiBotti", font=("Arial", 18, "bold")).pack(pady=8)
        tk.Label(
            self,
            text="1) Kauppalehti: SINÄ avaat Chromen ja protestilistan → Ctrl+A/Ctrl+C → liitä tähän.\n"
                 "   Botti tekee: yritysnimet.pdf + yritysnimet_ja_ytunnukset.pdf (+ Wordit)\n"
                 "2) PDF → poimii yritysnimet + Y-tunnukset → hakee sähköpostit YTJ:stä (Y-tunnus ensin, muuten nimi)\n",
            justify="center"
        ).pack(pady=4)

        # MODE 1
        box = tk.LabelFrame(self, text="1) Kauppalehti (copy/paste)", padx=10, pady=10)
        box.pack(fill="x", padx=12, pady=10)

        tk.Label(
            box,
            text="Vinkki: Jos avaat siniset nuolet ennen kopiointia, myös Y-tunnukset tulevat mukaan ja haku on varmempi.",
            justify="left"
        ).pack(anchor="w")

        self.text = tk.Text(box, height=10)
        self.text.pack(fill="x", pady=6)

        row = tk.Frame(box)
        row.pack(fill="x")
        tk.Button(row, text="Tee PDF + Word (nimet + yt)", font=("Arial", 12, "bold"), command=self.run_make_pdfs).pack(side="left")
        tk.Button(row, text="Tyhjennä", command=lambda: self.text.delete("1.0", tk.END)).pack(side="left", padx=8)

        # MODE 2
        box2 = tk.LabelFrame(self, text="2) PDF → YTJ", padx=10, pady=10)
        box2.pack(fill="x", padx=12, pady=10)

        tk.Button(box2, text="Valitse PDF ja hae sähköpostit YTJ:stä", font=("Arial", 12, "bold"), command=self.start_pdf_to_ytj).pack(anchor="w")

        self.status = tk.Label(self, text="Valmiina.", font=("Arial", 11))
        self.status.pack(pady=6)

        self.progress = ttk.Progressbar(self, orient="horizontal", mode="determinate", length=1020)
        self.progress.pack(pady=6)

        frame = tk.Frame(self)
        frame.pack(fill="both", expand=True, padx=14, pady=10)

        tk.Label(frame, text="Live-logi (uusimmat alimmaisena):").pack(anchor="w")
        self.listbox = tk.Listbox(frame, height=18)
        self.listbox.pack(side="left", fill="both", expand=True)

        sb = tk.Scrollbar(frame, orient="vertical", command=self.listbox.yview)
        sb.pack(side="right", fill="y")
        self.listbox.configure(yscrollcommand=sb.set)

        tk.Label(self, text=f"Tallennus: {OUT_DIR}", wraplength=1040, justify="center").pack(pady=6)

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

    # ---------- MODE 1 ----------
    def run_make_pdfs(self):
        raw = self.text.get("1.0", tk.END)
        names, yts = parse_names_and_yts_from_copied_text(raw)

        if not names and not yts:
            self.set_status("Ei löytynyt yritysnimiä tai Y-tunnuksia liitetystä tekstistä.")
            messagebox.showwarning("Ei löytynyt", "Liitä protestilistan sisältö (Ctrl+A/Ctrl+C) ja yritä uudestaan.")
            return

        self.set_status(f"Poimittu: nimet={len(names)} | y-tunnukset={len(yts)}")

        # Wordit
        if names:
            p1 = save_word_plain_lines(names, "yritysnimet_kauppalehti.docx")
            self.ui_log(f"Tallennettu: {p1}")
        if yts:
            p2 = save_word_plain_lines(yts, "ytunnukset_kauppalehti.docx")
            self.ui_log(f"Tallennettu: {p2}")

        # PDF: pelkät nimet
        if names:
            pdf1 = save_pdf_lines(names, "yritysnimet_kauppalehti.pdf", title="Yritysnimet (Kauppalehti protestilista)")
            self.ui_log(f"Tallennettu: {pdf1}")

        # PDF: nimet + yts (jos löytyy)
        combined = []
        if names:
            combined.append("YRITYSNIMET")
            combined += names
            combined.append("")
        if yts:
            combined.append("Y-TUNNUKSET")
            combined += yts

        pdf2 = save_pdf_lines(combined, "yritysnimet_ja_ytunnukset.pdf", title="Kauppalehti -> poimitut tiedot")
        self.ui_log(f"Tallennettu: {pdf2}")

        self.set_status("Valmis (PDF + Word).")
        messagebox.showinfo(
            "Valmis",
            f"Valmis!\n\nYritysnimiä: {len(names)}\nY-tunnuksia: {len(yts)}\n\nKansio:\n{OUT_DIR}"
        )

    # ---------- MODE 2 ----------
    def start_pdf_to_ytj(self):
        pdf_path = filedialog.askopenfilename(filetypes=[("PDF files", "*.pdf")])
        if not pdf_path:
            return
        threading.Thread(target=self.run_pdf_to_ytj, args=(pdf_path,), daemon=True).start()

    def run_pdf_to_ytj(self, pdf_path):
        driver = None
        try:
            self.set_status("Luetaan PDF: yritysnimet + Y-tunnukset…")
            names, yts = extract_names_and_yts_from_pdf(pdf_path)

            if not names and not yts:
                self.set_status("PDF:stä ei löytynyt nimiä tai Y-tunnuksia.")
                messagebox.showwarning("Ei löytynyt", "PDF:stä ei löytynyt nimiä tai Y-tunnuksia.")
                return

            # tallenna poimitut myös Wordiin
            if names:
                p1 = save_word_plain_lines(names, "pdf_poimitut_yritysnimet.docx")
                self.ui_log(f"Tallennettu: {p1}")
            if yts:
                p2 = save_word_plain_lines(yts, "pdf_poimitut_ytunnukset.docx")
                self.ui_log(f"Tallennettu: {p2}")

            self.set_status("Käynnistetään Chrome ja haetaan sähköpostit YTJ:stä…")
            driver = start_new_driver()

            emails = fetch_emails_from_ytj_by_yts_and_names(
                driver,
                yts=yts,
                names=names,
                status_cb=self.set_status,
                progress_cb=self.set_progress,
                log_cb=self.ui_log,
                stop_evt=self.stop_evt
            )

            em_path = save_word_plain_lines(emails, "sahkopostit_ytj.docx")
            self.ui_log(f"Tallennettu: {em_path}")

            self.set_status("Valmis!")
            messagebox.showinfo(
                "Valmis",
                f"Valmis!\n\nNimiä PDF:stä: {len(names)}\nY-tunnuksia PDF:stä: {len(yts)}\nSähköposteja löytyi: {len(emails)}\n\nKansio:\n{OUT_DIR}"
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
