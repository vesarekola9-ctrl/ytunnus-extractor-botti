import os
import re
import sys
import time
import threading
import tkinter as tk
from tkinter import messagebox, filedialog

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
# Output-kansio: EXE-kansio ensisijaisesti, fallback Documentsiin
# -----------------------------
def get_exe_dir():
    if getattr(sys, "frozen", False):
        return os.path.dirname(sys.executable)
    return os.path.dirname(os.path.abspath(__file__))


def get_output_dir():
    exe_dir = get_exe_dir()
    try:
        test_path = os.path.join(exe_dir, "_write_test.tmp")
        with open(test_path, "w", encoding="utf-8") as f:
            f.write("ok")
        os.remove(test_path)
        return exe_dir
    except Exception:
        home = os.path.expanduser("~")
        docs = os.path.join(home, "Documents")
        out = os.path.join(docs, "ProtestiBotti")
        os.makedirs(out, exist_ok=True)
        return out


OUT_DIR = get_output_dir()
LOG_PATH = os.path.join(OUT_DIR, "log.txt")


def log(msg: str):
    ts = time.strftime("%Y-%m-%d %H:%M:%S")
    line = f"[{ts}] {msg}"
    try:
        with open(LOG_PATH, "a", encoding="utf-8") as f:
            f.write(line + "\n")
    except Exception:
        pass
    print(line)


def reset_log():
    try:
        with open(LOG_PATH, "w", encoding="utf-8") as f:
            f.write("=== BOTTI KÄYNNISTETTY ===\n")
    except Exception:
        pass
    log(f"Output-kansio: {OUT_DIR}")
    log(f"Logi: {LOG_PATH}")


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
# Tallennus Word/Excel
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
# VIRRE: sähköposti nimenomaan "Sähköposti"-kentästä
# -----------------------------
def normalize_email(text: str):
    # Virre näyttää usein info(a)firma.fi
    return (text or "").strip().replace(" ", "").replace("(a)", "@").replace("[at]", "@")


def extract_company_email_from_dom(driver):
    """
    Poimii yrityksen sähköpostin DOMista kohdasta 'Sähköposti' (ei footterin virre-mailia).
    Palauttaa "" jos ei löydy.
    """
    # Etsi label-elementti jonka tekstinä tarkalleen Sähköposti
    labels = driver.find_elements(By.XPATH, "//*[normalize-space(.)='Sähköposti']")
    for lab in labels:
        try:
            # Kokeillaan yleiset: sisarus / seuraava elementti / parentin seuraava
            candidates = []
            candidates += lab.find_elements(By.XPATH, "following-sibling::*[1]")
            candidates += lab.find_elements(By.XPATH, "following::*[1]")
            parent = lab.find_element(By.XPATH, "..")
            candidates += parent.find_elements(By.XPATH, "following-sibling::*[1]")

            # Lisäksi: joskus arvo on linkissä <a>
            candidates += parent.find_elements(By.XPATH, ".//a")

            for c in candidates:
                txt = normalize_email(getattr(c, "text", "") or "")
                if "@" in txt and "." in txt and len(txt) >= 6:
                    return txt

                # varalla: jos mukana muuta tekstiä, poimi email-regexillä
                m = re.search(r"[a-zA-Z0-9_.+-]+@[a-zA-Z0-9-]+\.[a-zA-Z0-9-.]+", txt)
                if m:
                    return m.group(0)
        except Exception:
            continue

    return ""


def find_search_input(driver):
    """
    Etsitään Virre etusivun hakukenttä luotettavasti.
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
    # Ensisijaisesti <button> jossa teksti "HAE"
    buttons = driver.find_elements(By.TAG_NAME, "button")
    for b in buttons:
        try:
            if b.is_displayed() and "hae" in (b.text or "").strip().lower():
                b.click()
                return True
        except Exception:
            continue

    # Fallback: mikä tahansa klikattava elementti, jossa HAE
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
    Jos tulee tuloslista, klikataan ensimmäinen linkki jossa näkyy Y-tunnus.
    """
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


def start_chrome_driver():
    options = webdriver.ChromeOptions()
    # näkyvä Chrome (luotettavin)
    options.add_argument("--start-maximized")
    options.add_argument("--disable-gpu")
    options.add_argument("--disable-dev-shm-usage")

    driver_path = ChromeDriverManager().install()
    log(f"ChromeDriver: {driver_path}")
    return webdriver.Chrome(service=Service(driver_path), options=options)


def virre_fetch_email_for_yt(driver, yt):
    """
    Palauttaa (email tai "", status-teksti)
    """
    wait = WebDriverWait(driver, 25)

    driver.get("https://virre.prh.fi/novus/home")
    wait.until(EC.presence_of_element_located((By.TAG_NAME, "body")))

    inp = find_search_input(driver)
    if not inp:
        return "", "EI HAKUKENTTÄÄ"

    try:
        inp.clear()
        inp.send_keys(yt)
    except Exception as e:
        return "", f"HAKUKENTTÄ VIRHE: {e}"

    if not click_hae_button(driver):
        return "", "EI HAE-NAPPIA"

    # Odota että tulos/yrityssivu latautuu (body löytyy aina)
    wait.until(EC.presence_of_element_located((By.TAG_NAME, "body")))

    # Jos listaus, avaa eka osuma
    try:
        if try_open_first_result_if_list(driver):
            wait.until(EC.presence_of_element_located((By.TAG_NAME, "body")))
    except Exception:
        pass

    # TÄRKEIN: poimi sähköposti vain "Sähköposti"-kentästä
    email = extract_company_email_from_dom(driver)
    if email:
        return email, "OK"

    # Jos sähköposti-label ei löydy tai ei ole arvoa
    return "", "EI SÄHKÖPOSTIA"


# -----------------------------
# GUI (valitse PDF)
# -----------------------------
class App(tk.Tk):
    def __init__(self):
        super().__init__()
        reset_log()

        self.title("Y-tunnus extractor + Virre sähköpostit")
        self.geometry("560x320")

        title = tk.Label(self, text="ProtestiBotti", font=("Arial", 18, "bold"))
        title.pack(pady=12)

        info = tk.Label(
            self,
            text="Valitse PDF → botti kerää Y-tunnukset → hakee Virrestä sähköpostit\n"
                 "Tallennus: EXE-kansio (tai fallback Documents\\ProtestiBotti)",
            justify="center"
        )
        info.pack(pady=8)

        btn = tk.Button(self, text="Valitse PDF", font=("Arial", 12), command=self.pick_pdf)
        btn.pack(pady=10)

        self.status = tk.Label(self, text="Valmiina.", font=("Arial", 11))
        self.status.pack(pady=8)

        out = tk.Label(self, text=f"Output: {OUT_DIR}", wraplength=520, justify="center")
        out.pack(pady=6)

    def pick_pdf(self):
        path = filedialog.askopenfilename(filetypes=[("PDF files", "*.pdf")])
        if path:
            threading.Thread(target=self.run_job, args=(path,), daemon=True).start()

    def set_status(self, s):
        self.status.config(text=s)
        self.update_idletasks()
        log(s)

    def run_job(self, pdf_path):
        try:
            self.set_status("Luetaan PDF ja kerätään Y-tunnukset…")
            yt = extract_ytunnukset_from_pdf(pdf_path)
            if not yt:
                self.set_status("Ei löytynyt Y-tunnuksia.")
                messagebox.showwarning("Ei löytynyt", "PDF:stä ei löytynyt yhtään Y-tunnusta.")
                return

            self.set_status(f"Löytyi {len(yt)} Y-tunnusta. Tallennetaan…")
            save_word_list(yt, "ytunnukset.docx", "Y-tunnukset")
            save_excel_table([(x,) for x in yt], "ytunnukset.xlsx", ["Y-tunnus"])

            self.set_status("Käynnistetään Chrome ja haetaan Virrestä sähköpostit…")
            try:
                driver = start_chrome_driver()
            except Exception as e:
                log(f"CHROME EI KÄYNNISTY: {e}")
                messagebox.showerror("Chrome ei käynnisty", f"Katso log.txt:\n{LOG_PATH}\n\nVirhe:\n{e}")
                self.set_status("Chrome ei käynnistynyt.")
                return

            results = []
            try:
                for i, y in enumerate(yt, start=1):
                    self.set_status(f"Virre {i}/{len(yt)}: {y}")
                    email, status = virre_fetch_email_for_yt(driver, y)
                    results.append((y, email if email else status))
                    # pieni viive ettei hakata palvelua
                    time.sleep(1)
            finally:
                try:
                    driver.quit()
                except Exception:
                    pass

            self.set_status("Tallennetaan sähköpostit…")
            word_lines = [f"{y} -> {e}" for y, e in results]
            save_word_list(word_lines, "virre_sahkopostit.docx", "Virre sähköpostit")
            save_excel_table(results, "virre_sahkopostit.xlsx", ["Y-tunnus", "Sähköposti / status"])

            self.set_status("Valmis!")
            messagebox.showinfo("Valmis", f"Valmis!\n\nTiedostot:\n{OUT_DIR}\n\nLogi:\n{LOG_PATH}")

        except Exception as e:
            log(f"VIRHE: {e}")
            messagebox.showerror("Virhe", f"Tuli virhe.\nKatso log.txt:\n{LOG_PATH}\n\n{e}")


if __name__ == "__main__":
    App().mainloop()
