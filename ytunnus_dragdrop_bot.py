import os
import re
import sys
import time
import threading
import tkinter as tk
from tkinter import messagebox, filedialog

# Drag & drop
try:
    from tkinterdnd2 import DND_FILES, TkinterDnD
    HAS_DND = True
except Exception:
    HAS_DND = False

# PDF + Word + Excel
import PyPDF2
from docx import Document
import openpyxl

# Selenium
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC

# Driver manager
from webdriver_manager.chrome import ChromeDriverManager


# -----------------------------
# Polut (exe-kansioon)
# -----------------------------
def base_dir():
    # exe: sys.executable -> .../ytunnus_dragdrop_bot.exe
    if getattr(sys, "frozen", False):
        return os.path.dirname(sys.executable)
    return os.path.dirname(os.path.abspath(__file__))


OUT_DIR = base_dir()
LOG_PATH = os.path.join(OUT_DIR, "log.txt")


def log(msg: str):
    ts = time.strftime("%Y-%m-%d %H:%M:%S")
    line = f"[{ts}] {msg}"
    try:
        with open(LOG_PATH, "a", encoding="utf-8") as f:
            f.write(line + "\n")
    except Exception:
        pass
    # myös konsoliin (GitHub Actions / debug)
    print(line)


# -----------------------------
# Y-tunnukset PDF:stä
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
            # hae 1234567-8 tai 12345678
            matches = re.findall(r"\b\d{7}-\d\b|\b\d{8}\b", text)
            for m in matches:
                n = normalize_yt(m)
                if n:
                    yt_set.add(n)
    return sorted(yt_set)


# -----------------------------
# Tallennus Word / Excel
# -----------------------------
def save_word_list(lines, filename, title):
    path = os.path.join(OUT_DIR, filename)
    doc = Document()
    doc.add_heading(title, level=1)
    for line in lines:
        doc.add_paragraph(line)
    doc.save(path)
    log(f"Tallennettu Word: {path}")


def save_excel_table(rows, filename, headers):
    path = os.path.join(OUT_DIR, filename)
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Tulokset"
    ws.append(headers)
    for r in rows:
        ws.append(list(r))
    wb.save(path)
    log(f"Tallennettu Excel: {path}")


# -----------------------------
# VIRRE: sähköpostin poiminta
# -----------------------------
EMAIL_AT = re.compile(r"[a-zA-Z0-9_.+-]+@[a-zA-Z0-9-]+\.[a-zA-Z0-9-.]+")
EMAIL_A = re.compile(r"[a-zA-Z0-9_.+-]+\s*\(a\)\s*[a-zA-Z0-9-]+\.[a-zA-Z0-9-.]+")


def normalize_email(raw: str):
    raw = raw.strip().replace(" ", "")
    raw = raw.replace("(a)", "@").replace("[at]", "@")
    return raw


def find_email_anywhere(page_source: str):
    m = EMAIL_AT.search(page_source)
    if m:
        return m.group(0)
    m2 = EMAIL_A.search(page_source)
    if m2:
        return normalize_email(m2.group(0))
    return ""


def find_search_input(driver):
    """
    Etsitään Virren etusivun hakukenttä robustisti:
    - input[type=text] jossa placeholder/aria sisältää 'Y-tunnus'
    - fallback: ensimmäinen näkyvä input[type=text]
    """
    inputs = driver.find_elements(By.CSS_SELECTOR, "input[type='text'], input[type='search']")
    candidates = []
    for inp in inputs:
        try:
            if not inp.is_displayed():
                continue
            ph = (inp.get_attribute("placeholder") or "").lower()
            aria = (inp.get_attribute("aria-label") or "").lower()
            name = (inp.get_attribute("name") or "").lower()
            if ("y-tunnus" in ph) or ("y-tunnus" in aria) or ("y-tunnus" in name):
                return inp
            candidates.append(inp)
        except Exception:
            continue
    return candidates[0] if candidates else None


def click_hae_button(driver):
    """
    Klikkaa HAE-nappi robustisti:
    - button jossa teksti sisältää HAE
    - fallback: elementti jossa teksti HAE
    """
    buttons = driver.find_elements(By.TAG_NAME, "button")
    for b in buttons:
        try:
            if not b.is_displayed():
                continue
            if "hae" in (b.text or "").strip().lower():
                b.click()
                return True
        except Exception:
            continue

    # fallback: mikä tahansa klikattava elementti jossa HAE
    elems = driver.find_elements(By.XPATH, "//*[contains(translate(normalize-space(.),'hae','HAE'),'HAE')]")
    for e in elems:
        try:
            if e.is_displayed() and e.is_enabled():
                e.click()
                return True
        except Exception:
            continue
    return False


def try_open_first_result_if_list(driver):
    """
    Jos Virre näyttää listan (useampi osuma), klikataan ensimmäinen osuma.
    Tämä on heuristiikka: klikataan ensimmäinen näkyvä linkki/row, jossa on Y-tunnus-muotoinen teksti.
    """
    page = driver.page_source
    if "Y-tunnus" not in page and "y-tunnus" not in page.lower():
        return False

    # etsitään linkkejä, joissa näkyy y-tunnus
    links = driver.find_elements(By.TAG_NAME, "a")
    for a in links:
        try:
            txt = (a.text or "").strip()
            if re.search(r"\b\d{7}-\d\b", txt) and a.is_displayed():
                a.click()
                return True
        except Exception:
            continue
    return False


def virre_fetch_email_for_yt(driver, yt):
    """
    Avaa Virre, hakee ytunnuksella, ja palauttaa email tai tyhjä/teksti.
    """
    wait = WebDriverWait(driver, 25)

    driver.get("https://virre.prh.fi/novus/home")
    wait.until(EC.presence_of_element_located((By.TAG_NAME, "body")))
    time.sleep(1)

    inp = find_search_input(driver)
    if not inp:
        return "", "EI LÖYTYNYT HAKUKENTTÄÄ"

    try:
        inp.clear()
        inp.send_keys(yt)
    except Exception as e:
        return "", f"HAKUKENTTÄ VIRHE: {e}"

    ok = click_hae_button(driver)
    if not ok:
        return "", "EI LÖYTYNYT HAE-NAPPIA"

    # odotetaan, että sivu muuttuu / sisältö päivittyy
    time.sleep(3)
    wait.until(EC.presence_of_element_located((By.TAG_NAME, "body")))

    # jos listanäkymä, yritä avata eka osuma
    try:
        opened = try_open_first_result_if_list(driver)
        if opened:
            time.sleep(3)
    except Exception:
        pass

    src = driver.page_source
    email = find_email_anywhere(src)
    if email:
        return email, "OK"

    # jos ei löydy mitään, merkitään "EI SÄHKÖPOSTIA"
    return "", "EI SÄHKÖPOSTIA"


def start_chrome_driver():
    """
    Käynnistää Chromen. Jos tämä epäonnistuu, log.txt kertoo miksi.
    """
    options = webdriver.ChromeOptions()
    # näkyvä Chrome helpottaa (ja ei jää “salaisesti” piiloon)
    options.add_argument("--start-maximized")
    options.add_argument("--disable-gpu")
    options.add_argument("--disable-dev-shm-usage")

    driver_path = ChromeDriverManager().install()
    log(f"ChromeDriver: {driver_path}")

    driver = webdriver.Chrome(service=Service(driver_path), options=options)
    return driver


# -----------------------------
# GUI
# -----------------------------
class App(TkinterDnD.Tk if HAS_DND else tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("Protestilista botti (PDF → Y-tunnukset → Virre sähköpostit)")
        self.geometry("760x520")

        # reset log
        try:
            with open(LOG_PATH, "w", encoding="utf-8") as f:
                f.write("=== BOTTI KÄYNNISTETTY ===\n")
        except Exception:
            pass

        self.info = tk.Label(
            self,
            text="VEDÄ JA PUDOTA PDF TÄHÄN\n(Tai paina 'Valitse PDF')\n\nTulokset ja log.txt tallentuu samaan kansioon kuin exe",
            font=("Arial", 14, "bold"),
            bg="lightgray",
            relief="ridge",
            width=70,
            height=6
        )
        self.info.pack(pady=18)

        self.btn = tk.Button(self, text="Valitse PDF", font=("Arial", 12), command=self.pick_pdf)
        self.btn.pack(pady=6)

        self.status = tk.Label(self, text="Valmiina.", font=("Arial", 11))
        self.status.pack(pady=6)

        # Drag&Drop: rekisteröi koko ikkuna + label (ettei tipu “väärään paikkaan”)
        if HAS_DND:
            self.drop_target_register(DND_FILES)
            self.dnd_bind("<<Drop>>", self.on_drop)
            self.info.drop_target_register(DND_FILES)
            self.info.dnd_bind("<<Drop>>", self.on_drop)

        log(f"OUT_DIR = {OUT_DIR}")

    def set_status(self, s):
        self.status.config(text=s)
        self.update_idletasks()

    def pick_pdf(self):
        path = filedialog.askopenfilename(filetypes=[("PDF files", "*.pdf")])
        if path:
            self.run_job(path)

    def on_drop(self, event):
        path = event.data.strip()
        if path.startswith("{") and path.endswith("}"):
            path = path[1:-1]
        self.run_job(path)

    def run_job(self, pdf_path):
        if not pdf_path.lower().endswith(".pdf"):
            messagebox.showerror("Virhe", "Valitse PDF.")
            return

        self.set_status(f"PDF vastaanotettu: {pdf_path}")
        log(f"PDF vastaanotettu: {pdf_path}")

        t = threading.Thread(target=self.job, args=(pdf_path,), daemon=True)
        t.start()

    def job(self, pdf_path):
        try:
            self.set_status("Luetaan PDF ja kerätään Y-tunnukset…")
            ytunnukset = extract_ytunnukset_from_pdf(pdf_path)
            log(f"Löytyi Y-tunnuksia: {len(ytunnukset)}")

            if not ytunnukset:
                self.set_status("Ei löytynyt Y-tunnuksia.")
                messagebox.showwarning("Ei löytynyt", "PDF:stä ei löytynyt yhtään Y-tunnusta.")
                return

            # tallenna ytunnukset
            save_word_list(ytunnukset, "ytunnukset.docx", "Y-tunnukset")
            save_excel_table([(y,) for y in ytunnukset], "ytunnukset.xlsx", ["Y-tunnus"])

            # virre
            self.set_status("Käynnistetään Chrome ja haetaan Virrestä sähköpostit…")
            try:
                driver = start_chrome_driver()
            except Exception as e:
                log(f"CHROME EI KÄYNNISTY: {e}")
                messagebox.showerror("Chrome ei käynnisty",
                                     f"Chrome/Selenium ei käynnistynyt.\nAvaa log.txt ja lähetä se minulle.\n\nVirhe:\n{e}")
                self.set_status("Chrome ei käynnistynyt. Katso log.txt")
                return

            results = []
            try:
                for idx, yt in enumerate(ytunnukset, start=1):
                    self.set_status(f"Virre-haku {idx}/{len(ytunnukset)}: {yt}")
                    log(f"Virre-haku {idx}/{len(ytunnukset)}: {yt}")

                    email, status = virre_fetch_email_for_yt(driver, yt)
                    if not email:
                        email_out = status
                    else:
                        email_out = email

                    results.append((yt, email_out))

                    # pieni viive ettei hakata palvelua
                    time.sleep(2)

            finally:
                try:
                    driver.quit()
                except Exception:
                    pass

            # tallenna sähköpostit
            word_lines = [f"{yt} -> {email}" for yt, email in results]
            save_word_list(word_lines, "virre_sahkopostit.docx", "Virre sähköpostit")
            save_excel_table(results, "virre_sahkopostit.xlsx", ["Y-tunnus", "Sähköposti / status"])

            self.set_status("Valmis! Katso exe-kansiosta tulokset + log.txt")
            messagebox.showinfo(
                "Valmis",
                "Valmis!\n\nTiedostot luotu exe:n kansioon:\n"
                "- ytunnukset.docx / ytunnukset.xlsx\n"
                "- virre_sahkopostit.docx / virre_sahkopostit.xlsx\n"
                "- log.txt"
            )

        except Exception as e:
            log(f"VIRHE: {e}")
            self.set_status("Virhe. Katso log.txt")
            messagebox.showerror("Virhe", f"Tuli virhe. Katso log.txt.\n\n{e}")


if __name__ == "__main__":
    app = App()
    app.mainloop()
