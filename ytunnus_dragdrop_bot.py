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
# YTJ Selenium
# -----------------------------
def start_driver():
    options = webdriver.ChromeOptions()
    options.add_argument("--start-maximized")
    options.add_argument("--disable-gpu")
    options.add_argument("--disable-dev-shm-usage")

    driver_path = ChromeDriverManager().install()
    log("ChromeDriver OK")
    return webdriver.Chrome(service=Service(driver_path), options=options)


def find_ytj_search_input(driver):
    # Etsi label jossa "Y-tunnus tai LEI-tunnus" ja ota seuraava input
    candidates = driver.find_elements(By.XPATH, "//*[contains(normalize-space(.), 'Y-tunnus') and contains(normalize-space(.), 'LEI')]")
    for lab in candidates:
        try:
            inp = lab.find_element(By.XPATH, "following::input[1]")
            if inp.is_displayed():
                return inp
        except Exception:
            continue

    # fallback: eka näkyvä input
    inputs = driver.find_elements(By.CSS_SELECTOR, "input[type='text'], input[type='search']")
    for i in inputs:
        try:
            if i.is_displayed():
                return i
        except Exception:
            pass
    return None


def click_ytj_hae(driver):
    btns = driver.find_elements(By.TAG_NAME, "button")
    for b in btns:
        try:
            if b.is_displayed() and (b.text or "").strip().lower() == "hae":
                b.click()
                return True
        except Exception:
            continue

    # fallback: elementti jossa teksti HAE
    elems = driver.find_elements(By.XPATH, "//*[normalize-space(.)='HAE']")
    for e in elems:
        try:
            if e.is_displayed() and e.is_enabled():
                e.click()
                return True
        except Exception:
            continue
    return False


def open_first_result_if_list(driver, yt):
    """
    Jos hakutulos on listana, avaa linkki jossa näkyy Y-tunnus.
    """
    links = driver.find_elements(By.TAG_NAME, "a")
    for a in links:
        try:
            t = (a.text or "").strip()
            if yt in t and a.is_displayed():
                a.click()
                return True
        except Exception:
            continue
    return False


def click_all_nayta_buttons(driver):
    """
    Klikkaa kaikki näkyvät 'Näytä' -napit yrityssivulla.
    Joissain tapauksissa ne voivat olla buttonit, joskus myös linkkejä.
    """
    clicked_any = False

    for _round in range(3):
        clicked_this_round = False

        # 1) button "Näytä"
        for b in driver.find_elements(By.TAG_NAME, "button"):
            try:
                if (b.text or "").strip().lower() == "näytä" and b.is_displayed() and b.is_enabled():
                    driver.execute_script("arguments[0].scrollIntoView({block:'center'});", b)
                    b.click()
                    clicked_this_round = True
                    clicked_any = True
                    time.sleep(0.12)
            except Exception:
                continue

        # 2) linkki <a> jossa "Näytä"
        for a in driver.find_elements(By.TAG_NAME, "a"):
            try:
                if (a.text or "").strip().lower() == "näytä" and a.is_displayed():
                    driver.execute_script("arguments[0].scrollIntoView({block:'center'});", a)
                    a.click()
                    clicked_this_round = True
                    clicked_any = True
                    time.sleep(0.12)
            except Exception:
                continue

        if not clicked_this_round:
            break

    return clicked_any


def extract_email_from_ytj_table(driver):
    """
    Oikea tapa YTJ:ssä:
    - etsi rivi jossa ensimmäisessä solussa lukee 'Sähköposti'
    - lue saman rivin muut solut ja poimi email
    """
    # Etsi elementti jonka tekstinä 'Sähköposti'
    nodes = driver.find_elements(By.XPATH, "//*[normalize-space(.)='Sähköposti']")
    for node in nodes:
        try:
            # mene ylös rivitasolle (tr) jos taulukko, tai lähimpään "row" konttiin
            row = None
            try:
                row = node.find_element(By.XPATH, "ancestor::tr[1]")
            except Exception:
                row = node.find_element(By.XPATH, "ancestor::*[self::div or self::section][1]")

            txt = row.text or ""
            email = pick_email_from_text(txt)
            if email:
                return email

            # jos ei löytynyt row.textistä, etsi linkeistä/soluista
            for a in row.find_elements(By.TAG_NAME, "a"):
                email = pick_email_from_text(a.text or "")
                if email:
                    return email

        except Exception:
            continue

    return ""


def extract_email_fallback_main(driver):
    # fallback: main-alueesta
    try:
        main = driver.find_element(By.TAG_NAME, "main")
        return pick_email_from_text(main.text or "")
    except Exception:
        return ""


def ytj_fetch_email(driver, yt):
    wait = WebDriverWait(driver, 20)

    driver.get("https://tietopalvelu.ytj.fi/")
    wait.until(EC.presence_of_element_located((By.TAG_NAME, "body")))

    inp = find_ytj_search_input(driver)
    if not inp:
        log("Ei löydy hakukenttää.")
        return ""

    inp.clear()
    inp.send_keys(yt)

    if not click_ytj_hae(driver):
        log("Ei löydy HAE-nappia.")
        return ""

    wait.until(EC.presence_of_element_located((By.TAG_NAME, "body")))

    # jos tuli listaus -> avaa
    try:
        if open_first_result_if_list(driver, yt):
            wait.until(EC.presence_of_element_located((By.TAG_NAME, "body")))
    except Exception:
        pass

    # Nyt pitäisi olla yrityssivu (URL usein /yritys/xxxxxxx-x)
    # Klikkaa "Näytä" napit, jotta piilotetut tiedot paljastuu
    try:
        click_all_nayta_buttons(driver)
        # anna DOM:n päivittyä
        time.sleep(0.15)
    except Exception:
        pass

    # 1) poimi sähköposti oikealta riviltä
    email = extract_email_from_ytj_table(driver)
    if email:
        return email

    # 2) fallback: main-alueesta
    return extract_email_fallback_main(driver)


# -----------------------------
# GUI
# -----------------------------
class App(tk.Tk):
    def __init__(self):
        super().__init__()
        reset_log()

        self.title("ProtestiBotti (YTJ)")
        self.geometry("600x340")

        tk.Label(self, text="ProtestiBotti", font=("Arial", 18, "bold")).pack(pady=12)
        tk.Label(self, text="Valitse PDF → kerää Y-tunnukset → hae sähköpostit YTJ:stä", justify="center").pack(pady=6)

        tk.Button(self, text="Valitse PDF", font=("Arial", 12), command=self.pick_pdf).pack(pady=12)

        self.status = tk.Label(self, text="Valmiina.", font=("Arial", 11))
        self.status.pack(pady=10)

        tk.Label(self, text=f"Tallennus: {OUT_DIR}", wraplength=560, justify="center").pack(pady=6)

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

            try:
                for i, yt in enumerate(yt_list, start=1):
                    self.set_status(f"Haku {i}/{len(yt_list)}: {yt}")
                    email = ytj_fetch_email(driver, yt)

                    if email:
                        k = email.lower()
                        if k not in seen:
                            seen.add(k)
                            emails.append(email)

                    time.sleep(0.1)  # nopea kuten pyysit
            finally:
                try:
                    driver.quit()
                except Exception:
                    pass

            save_word_plain_lines(emails, "sahkopostit.docx")

            self.set_status("Valmis!")
            messagebox.showinfo("Valmis", f"Valmis!\n\nKansio:\n{OUT_DIR}\n\nLöydetyt emailit: {len(emails)}")

        except Exception as e:
            log(f"VIRHE: {e}")
            self.set_status("Virhe. Katso log.txt")
            messagebox.showerror("Virhe", f"Tuli virhe.\nKatso log.txt:\n{LOG_PATH}\n\n{e}")


if __name__ == "__main__":
    App().mainloop()
