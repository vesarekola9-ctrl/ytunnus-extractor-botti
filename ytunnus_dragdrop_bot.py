import os
import re
import sys
import time
import threading
import tkinter as tk
from tkinter import messagebox, filedialog

import PyPDF2
from docx import Document

from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from webdriver_manager.chrome import ChromeDriverManager


# -----------------------------
# Output-kansio: EXE-kansio ensisijaisesti, fallback Documentsiin
# + Päivämääräkansio YYYY-MM-DD
# -----------------------------
def get_exe_dir():
    if getattr(sys, "frozen", False):
        return os.path.dirname(sys.executable)
    return os.path.dirname(os.path.abspath(__file__))


def get_output_dir():
    base = get_exe_dir()

    # testaa kirjoitusoikeus exe-kansioon
    try:
        test_path = os.path.join(base, "_write_test.tmp")
        with open(test_path, "w", encoding="utf-8") as f:
            f.write("ok")
        os.remove(test_path)
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
# Word-tallennus
# -----------------------------
def save_word_list(lines, filename, title=None):
    path = os.path.join(OUT_DIR, filename)
    doc = Document()
    if title:
        doc.add_heading(title, level=1)
    for line in lines:
        doc.add_paragraph(line)
    doc.save(path)
    log(f"Tallennettu Word: {path}")


# -----------------------------
# VIRRE: hae sähköposti "Sähköposti"-kentästä
# -----------------------------
def normalize_email(text: str):
    return (text or "").strip().replace(" ", "").replace("(a)", "@").replace("[at]", "@")


def extract_company_email_from_dom(driver):
    labels = driver.find_elements(By.XPATH, "//*[normalize-space(.)='Sähköposti']")
    for lab in labels:
        try:
            candidates = []
            candidates += lab.find_elements(By.XPATH, "following-sibling::*[1]")
            candidates += lab.find_elements(By.XPATH, "following::*[1]")
            parent = lab.find_element(By.XPATH, "..")
            candidates += parent.find_elements(By.XPATH, "following-sibling::*[1]")
            candidates += parent.find_elements(By.XPATH, ".//a")

            for c in candidates:
                txt = normalize_email(getattr(c, "text", "") or "")
                if "@" in txt and "." in txt and len(txt) >= 6:
                    return txt
                m = re.search(r"[a-zA-Z0-9_.+-]+@[a-zA-Z0-9-]+\.[a-zA-Z0-9-.]+", txt)
                if m:
                    return m.group(0)
        except Exception:
            continue
    return ""


def find_search_input(driver):
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
    buttons = driver.find_elements(By.TAG_NAME, "button")
    for b in buttons:
        try:
            if b.is_displayed() and "hae" in (b.text or "").strip().lower():
                b.click()
                return True
        except Exception:
            continue

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
    options.add_argument("--start-maximized")
    options.add_argument("--disable-gpu")
    options.add_argument("--disable-dev-shm-usage")

    driver_path = ChromeDriverManager().install()
    log(f"ChromeDriver: {driver_path}")
    return webdriver.Chrome(service=Service(driver_path), options=options)


def virre_fetch_email_for_yt(driver, yt):
    wait = WebDriverWait(driver, 20)

    # Ladataan etusivu joka haulla (varmin, vaikka vähän hitaampi)
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

    # Odota että yritystiedot latautuvat – nopea yritys: "Sähköposti"-label
    try:
        wait.until(EC.presence_of_element_located((By.XPATH, "//*[normalize-space(.)='Sähköposti']")))
    except Exception:
        # joskus ei ole sähköposti-kenttää, mutta sivu voi silti olla yrityssivu
        pass

    # Jos listaus, avaa eka osuma
    try:
        if try_open_first_result_if_list(driver):
            # odota että sivu latautuu
            wait.until(EC.presence_of_element_located((By.TAG_NAME, "body")))
    except Exception:
        pass

    email = extract_company_email_from_dom(driver)
    if email:
        return email, "OK"
    return "", "EI SÄHKÖPOSTIA"


# -----------------------------
# GUI
# -----------------------------
class App(tk.Tk):
    def __init__(self):
        super().__init__()
        reset_log()

        self.title("ProtestiBotti (valmis)")
        self.geometry("600x340")

        title = tk.Label(self, text="ProtestiBotti", font=("Arial", 18, "bold"))
        title.pack(pady=12)

        info = tk.Label(
            self,
            text="Valitse PDF → kerää Y-tunnukset → hakee Virrestä sähköpostit\n"
                 "Tuloksena vain Wordit + päivämääräkansio",
            justify="center"
        )
        info.pack(pady=6)

        btn = tk.Button(self, text="Valitse PDF", font=("Arial", 12), command=self.pick_pdf)
        btn.pack(pady=10)

        self.status = tk.Label(self, text="Valmiina.", font=("Arial", 11))
        self.status.pack(pady=8)

        out = tk.Label(self, text=f"Tallennus: {OUT_DIR}", wraplength=560, justify="center")
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

            self.set_status(f"Löytyi {len(yt)} Y-tunnusta. Tallennetaan ytunnukset.docx…")
            save_word_list(yt, "ytunnukset.docx", title="Y-tunnukset")

            self.set_status("Käynnistetään Chrome ja haetaan sähköpostit…")
            try:
                driver = start_chrome_driver()
            except Exception as e:
                log(f"CHROME EI KÄYNNISTY: {e}")
                messagebox.showerror("Chrome ei käynnisty", f"Katso log.txt:\n{LOG_PATH}\n\nVirhe:\n{e}")
                self.set_status("Chrome ei käynnistynyt.")
                return

            emails_only = []
            seen = set()

            try:
                for i, y in enumerate(yt, start=1):
                    self.set_status(f"Virre {i}/{len(yt)}: {y}")
                    email, _status = virre_fetch_email_for_yt(driver, y)
                    if email:
                        key = email.lower()
                        if key not in seen:
                            seen.add(key)
                            emails_only.append(email)

                    # NOPEUTUS: pienempi viive
                    time.sleep(0.3)

            finally:
                try:
                    driver.quit()
                except Exception:
                    pass

            self.set_status("Tallennetaan sahkopostit.docx…")
            save_word_list(emails_only, "sahkopostit.docx", title="Sähköpostit")

            self.set_status("Valmis!")
            messagebox.showinfo(
                "Valmis",
                f"Valmis!\n\nKansio:\n{OUT_DIR}\n\nTiedostot:\n- ytunnukset.docx\n- sahkopostit.docx\n\nLogi:\n{LOG_PATH}"
            )

        except Exception as e:
            log(f"VIRHE: {e}")
            messagebox.showerror("Virhe", f"Tuli virhe.\nKatso log.txt:\n{LOG_PATH}\n\n{e}")


if __name__ == "__main__":
    App().mainloop()
