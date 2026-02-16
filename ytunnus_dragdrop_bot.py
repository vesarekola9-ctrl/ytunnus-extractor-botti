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


EMAIL_RE = re.compile(r"[A-Za-z0-9_.+-]+@[A-Za-z0-9-]+\.[A-Za-z0-9-.]+")
EMAIL_A_RE = re.compile(r"[A-Za-z0-9_.+-]+\s*\(a\)\s*[A-Za-z0-9-]+\.[A-Za-z0-9-.]+", re.I)

# Debug: jos ei löydy emailia, tallenna sivun HTML ensimmäisistä N tapauksesta
DEBUG_DUMP_NO_EMAIL_FIRST_N = 5


# -----------------------------
# Output + päivämääräkansio
# -----------------------------
def get_exe_dir():
    if getattr(sys, "frozen", False):
        return os.path.dirname(sys.executable)
    return os.path.dirname(os.path.abspath(__file__))


def get_output_dir():
    base = get_exe_dir()
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
    ts = time.strftime("%H:%M:%S")
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
    log(f"Output: {OUT_DIR}")
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
# Word: pelkkä lista riveinä (ei otsikkoa)
# -----------------------------
def save_word_plain_lines(lines, filename):
    path = os.path.join(OUT_DIR, filename)
    doc = Document()
    for line in lines:
        doc.add_paragraph(line)
    doc.save(path)
    log(f"Tallennettu: {path}")


# -----------------------------
# Email helpers
# -----------------------------
def normalize_email_candidate(raw: str) -> str:
    raw = (raw or "").strip().replace(" ", "")
    raw = raw.replace("(a)", "@").replace("[at]", "@")
    return raw


def pick_email_from_text(text: str) -> str:
    if not text:
        return ""
    m = EMAIL_RE.search(text)
    if m:
        return normalize_email_candidate(m.group(0))
    m2 = EMAIL_A_RE.search(text)
    if m2:
        return normalize_email_candidate(m2.group(0))
    return ""


# -----------------------------
# Selenium / YTJ
# -----------------------------
def start_driver():
    options = webdriver.ChromeOptions()
    options.add_argument("--start-maximized")
    options.add_argument("--disable-gpu")
    options.add_argument("--disable-dev-shm-usage")
    driver_path = ChromeDriverManager().install()
    log("ChromeDriver OK")
    return webdriver.Chrome(service=Service(driver_path), options=options)


def click_all_nayta(driver):
    """
    Klikkaa kaikki näkyvät 'Näytä' -napit (button tai linkki).
    """
    for _ in range(3):
        clicked = False

        for b in driver.find_elements(By.TAG_NAME, "button"):
            try:
                if (b.text or "").strip().lower() == "näytä" and b.is_displayed() and b.is_enabled():
                    driver.execute_script("arguments[0].scrollIntoView({block:'center'});", b)
                    b.click()
                    clicked = True
                    time.sleep(0.15)
            except Exception:
                continue

        for a in driver.find_elements(By.TAG_NAME, "a"):
            try:
                if (a.text or "").strip().lower() == "näytä" and a.is_displayed():
                    driver.execute_script("arguments[0].scrollIntoView({block:'center'});", a)
                    a.click()
                    clicked = True
                    time.sleep(0.15)
            except Exception:
                continue

        if not clicked:
            break


def wait_company_page_loaded(driver):
    """
    Odota että yrityssivun sisältö on oikeasti ladattu.
    Odotetaan tekstiä 'Y-tunnus' tai 'Sähköposti' (kumpi tulee ensin).
    """
    wait = WebDriverWait(driver, 25)
    wait.until(EC.presence_of_element_located((By.TAG_NAME, "body")))

    # YTJ on JS-heavy: odotetaan vielä että yritystiedot-tekstit löytyy DOM:sta
    try:
        wait.until(EC.presence_of_element_located(
            (By.XPATH, "//*[contains(normalize-space(.), 'Y-tunnus') or contains(normalize-space(.), 'Sähköposti')]")
        ))
    except Exception:
        # jos ei löydy, jatketaan silti (fallback regex)
        pass


def extract_email_from_row_cells(driver):
    """
    Etsi 'Sähköposti' sisältävä solu ja ota seuraava solu samalta riviltä.
    Toimii kun sivu on taulukko/tr/td (kuten screenshot).
    """
    # Solu voi olla td tai th
    label_cells = driver.find_elements(
        By.XPATH,
        "//tr//*[self::td or self::th][contains(normalize-space(.), 'Sähköposti')]"
    )

    for cell in label_cells:
        try:
            tr = cell.find_element(By.XPATH, "ancestor::tr[1]")

            # 1) kokeile mailto-linkki ensin
            for a in tr.find_elements(By.TAG_NAME, "a"):
                href = (a.get_attribute("href") or "")
                if href.lower().startswith("mailto:"):
                    return normalize_email_candidate(href.split(":", 1)[1])
                e = pick_email_from_text(a.text or "")
                if e:
                    return e

            # 2) ota seuraava solu samalla rivillä
            # jos label cell on 1. solu, seuraava on 2.
            next_cell = None
            try:
                next_cell = cell.find_element(By.XPATH, "following-sibling::*[1]")
            except Exception:
                # vaihtoehto: ota rivin kaikki solut ja etsi labelin jälkeen
                tds = tr.find_elements(By.XPATH, ".//*[self::td or self::th]")
                for idx, c in enumerate(tds):
                    if c == cell and idx + 1 < len(tds):
                        next_cell = tds[idx + 1]
                        break

            if next_cell:
                e = pick_email_from_text(next_cell.text or "")
                if e:
                    return e
                # myös mailto mahdollinen tässä solussa
                for a in next_cell.find_elements(By.TAG_NAME, "a"):
                    href = (a.get_attribute("href") or "")
                    if href.lower().startswith("mailto:"):
                        return normalize_email_candidate(href.split(":", 1)[1])
                    e = pick_email_from_text(a.text or "")
                    if e:
                        return e

            # 3) fallback: koko riviteksti
            e = pick_email_from_text(tr.text or "")
            if e:
                return e

        except Exception:
            continue

    return ""


def extract_email_fallback_main(driver):
    try:
        main = driver.find_element(By.TAG_NAME, "main")
        return pick_email_from_text(main.text or "")
    except Exception:
        return ""


def dump_debug_html(driver, yt):
    try:
        path = os.path.join(OUT_DIR, f"debug_no_email_{yt}.html")
        with open(path, "w", encoding="utf-8") as f:
            f.write(driver.page_source or "")
        log(f"DEBUG: tallennettu HTML: {path}")
    except Exception:
        pass


def ytj_fetch_email_direct(driver, yt, debug_counter):
    """
    Avaa suoraan yrityssivu /yritys/<yt>, klikkaa 'Näytä', poimi email.
    """
    url = f"https://tietopalvelu.ytj.fi/yritys/{yt}"
    driver.get(url)

    wait_company_page_loaded(driver)

    # Avaa piilotetut tiedot
    click_all_nayta(driver)

    # Poimi sähköposti
    email = extract_email_from_row_cells(driver)
    if not email:
        email = extract_email_fallback_main(driver)

    # Debug dump jos ei löydy
    if not email and debug_counter < DEBUG_DUMP_NO_EMAIL_FIRST_N:
        dump_debug_html(driver, yt)

    return email


# -----------------------------
# GUI
# -----------------------------
class App(tk.Tk):
    def __init__(self):
        super().__init__()
        reset_log()

        self.title("ProtestiBotti (YTJ suora linkki)")
        self.geometry("620x360")

        tk.Label(self, text="ProtestiBotti", font=("Arial", 18, "bold")).pack(pady=12)
        tk.Label(
            self,
            text="Valitse PDF → kerää Y-tunnukset → avaa YTJ yrityssivu suoraan (/yritys/..)\n→ klikkaa Näytä → poimi sähköposti",
            justify="center"
        ).pack(pady=6)

        tk.Button(self, text="Valitse PDF", font=("Arial", 12), command=self.pick_pdf).pack(pady=12)

        self.status = tk.Label(self, text="Valmiina.", font=("Arial", 11))
        self.status.pack(pady=10)

        tk.Label(self, text=f"Tallennus: {OUT_DIR}", wraplength=580, justify="center").pack(pady=6)

    def set_status(self, s):
        self.status.config(text=s)
        self.update_idletasks()
        log(s)

    def pick_pdf(self):
        path = filedialog.askopenfilename(filetypes=[("PDF files", "*.pdf")])
        if path:
            threading.Thread(target=self.run_job, args=(path,), daemon=True).start()

    def run_job(self, pdf_path):
        try:
            self.set_status("Luetaan PDF ja kerätään Y-tunnukset…")
            yt_list = extract_ytunnukset_from_pdf(pdf_path)

            if not yt_list:
                self.set_status("Ei löytynyt Y-tunnuksia.")
                messagebox.showwarning("Ei löytynyt", "PDF:stä ei löytynyt yhtään Y-tunnusta.")
                return

            save_word_plain_lines(yt_list, "ytunnukset.docx")

            self.set_status("Käynnistetään Chrome ja haetaan sähköpostit (YTJ)…")
            driver = start_driver()

            emails = []
            seen = set()
            debug_no_email = 0

            try:
                for i, yt in enumerate(yt_list, start=1):
                    self.set_status(f"Haku {i}/{len(yt_list)}: {yt}")
                    email = ytj_fetch_email_direct(driver, yt, debug_no_email)

                    if email:
                        key = email.lower()
                        if key not in seen:
                            seen.add(key)
                            emails.append(email)
                    else:
                        debug_no_email += 1

                    time.sleep(0.1)  # nopea, kuten halusit

            finally:
                try:
                    driver.quit()
                except Exception:
                    pass

            save_word_plain_lines(emails, "sahkopostit.docx")

            self.set_status("Valmis!")
            messagebox.showinfo(
                "Valmis",
                f"Valmis!\n\nKansio:\n{OUT_DIR}\n\nLöydetyt emailit: {len(emails)}\n\n"
                f"Jos nolla, katso debug_no_email_*.html tiedostot kansiosta."
            )

        except Exception as e:
            log(f"VIRHE: {e}")
            self.set_status("Virhe. Katso log.txt")
            messagebox.showerror("Virhe", f"Tuli virhe.\nKatso log.txt:\n{LOG_PATH}\n\n{e}")


if __name__ == "__main__":
    App().mainloop()
