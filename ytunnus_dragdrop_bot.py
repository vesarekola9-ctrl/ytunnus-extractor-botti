import os
import re
import sys
import time
import random
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
EMAIL_A_RE = re.compile(r"[A-Za-z0-9_.+-]+\s*\(a\)\s*[A-Za-z0-9-]+\.[A-Za-z0-9-.]+")


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


def save_word_plain_lines(lines, filename):
    path = os.path.join(OUT_DIR, filename)
    doc = Document()
    for line in lines:
        doc.add_paragraph(line)
    doc.save(path)
    log(f"Tallennettu: {path}")


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


def extract_company_email_from_dom(driver):
    labels = driver.find_elements(By.XPATH, "//*[normalize-space(.)='Sähköposti']")
    for lab in labels:
        try:
            parent = lab.find_element(By.XPATH, "..")
            email = pick_email_from_text(parent.text)
            if email:
                return email

            sibs = lab.find_elements(By.XPATH, "following-sibling::*[1]")
            for s in sibs:
                email = pick_email_from_text(s.text)
                if email:
                    return email

            links = parent.find_elements(By.XPATH, ".//a")
            for a in links:
                email = pick_email_from_text(a.text)
                if email:
                    return email
        except Exception:
            continue

    try:
        main = driver.find_element(By.TAG_NAME, "main")
        email = pick_email_from_text(main.text)
        if email:
            return email
    except Exception:
        pass

    return ""


def click_hae_button(driver):
    buttons = driver.find_elements(By.TAG_NAME, "button")
    for b in buttons:
        try:
            if b.is_displayed() and "hae" in (b.text or "").strip().lower():
                b.click()
                return True
        except Exception:
            continue
    return False


def start_driver():
    options = webdriver.ChromeOptions()
    options.add_argument("--start-maximized")
    options.add_argument("--disable-gpu")
    options.add_argument("--disable-dev-shm-usage")
    driver_path = ChromeDriverManager().install()
    log("ChromeDriver OK")
    return webdriver.Chrome(service=Service(driver_path), options=options)


def virre_fetch_email(driver, yt):
    """
    Palauttaa (email, ok_flag).
    ok_flag False jos epäillään blokkia / captcha / virhettä.
    """
    wait = WebDriverWait(driver, 20)
    try:
        driver.get("https://virre.prh.fi/novus/home")
        wait.until(EC.presence_of_element_located((By.TAG_NAME, "body")))

        # heuristiikka blokkisivulle
        src = driver.page_source.lower()
        if "captcha" in src or "robot" in src or "liian monta" in src or "access denied" in src:
            return "", False

        search = wait.until(EC.presence_of_element_located((By.CSS_SELECTOR, "input[type='text'], input[type='search']")))
        search.clear()
        search.send_keys(yt)

        if not click_hae_button(driver):
            return "", True  # ei nappia, mutta ei välttämättä blokki

        # odota hetki että tulos ehtii vaihtua (nopeasti)
        try:
            wait.until(EC.presence_of_element_located((By.XPATH, "//*[normalize-space(.)='Sähköposti']")))
        except Exception:
            pass

        # tarkista taas blokkisanoja tulossivulta
        src2 = driver.page_source.lower()
        if "captcha" in src2 or "robot" in src2 or "access denied" in src2:
            return "", False

        email = extract_company_email_from_dom(driver)
        return email, True

    except Exception:
        # jos tulee paljon poikkeuksia putkeen, se voi olla blokki/hidastus
        return "", False


class App(tk.Tk):
    def __init__(self):
        super().__init__()
        reset_log()

        self.title("ProtestiBotti (nopeampi + automaattinen backoff)")
        self.geometry("680x390")

        tk.Label(self, text="ProtestiBotti", font=("Arial", 18, "bold")).pack(pady=12)
        tk.Label(self, text="Nopea yksinajo + hidastaa itse jos Virre alkaa blokata", justify="center").pack(pady=6)

        row = tk.Frame(self)
        row.pack(pady=6)

        tk.Label(row, text="Alkuviive (s):", font=("Arial", 11)).pack(side="left", padx=6)
        self.base_delay_var = tk.DoubleVar(value=0.25)
        tk.Entry(row, textvariable=self.base_delay_var, width=6).pack(side="left")

        row2 = tk.Frame(self)
        row2.pack(pady=6)
        tk.Label(row2, text="Max viive (s):", font=("Arial", 11)).pack(side="left", padx=6)
        self.max_delay_var = tk.DoubleVar(value=3.0)
        tk.Entry(row2, textvariable=self.max_delay_var, width=6).pack(side="left")

        tk.Button(self, text="Valitse PDF", font=("Arial", 12), command=self.pick_pdf).pack(pady=12)

        self.status = tk.Label(self, text="Valmiina.", font=("Arial", 11))
        self.status.pack(pady=10)

        tk.Label(self, text=f"Tallennus: {OUT_DIR}", wraplength=640, justify="center").pack(pady=6)

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

            self.set_status("Käynnistetään Chrome…")
            driver = start_driver()

            emails = []
            seen = set()

            base_delay = float(self.base_delay_var.get())
            max_delay = float(self.max_delay_var.get())
            cur_delay = max(0.1, base_delay)

            failures_in_row = 0

            try:
                for i, yt in enumerate(yt_list, start=1):
                    self.set_status(f"Haku {i}/{len(yt_list)} (viive {cur_delay:.2f}s)")

                    email, ok = virre_fetch_email(driver, yt)

                    if email:
                        k = email.lower()
                        if k not in seen:
                            seen.add(k)
                            emails.append(email)

                    # backoff logiikka:
                    if ok:
                        failures_in_row = 0
                        # jos toimii, tiputa viivettä hitaasti kohti base_delay
                        cur_delay = max(base_delay, cur_delay * 0.9)
                    else:
                        failures_in_row += 1
                        # jos epäily blokkia -> nosta viivettä
                        cur_delay = min(max_delay, max(cur_delay * 1.7, base_delay))

                    # mikrotauko joka 25 haku (vähentää blokkia)
                    if i % 25 == 0:
                        cool = min(10.0, 2.0 + cur_delay * 2)
                        log(f"Cooldown {cool:.1f}s (25 haun välein)")
                        time.sleep(cool)

                    # satunnainen jitter (ettei näytä robottimaiselta)
                    jitter = random.uniform(0.0, 0.15)
                    time.sleep(cur_delay + jitter)

                    # jos epäonnistumisia putkeen paljon -> pidä pidempi tauko
                    if failures_in_row >= 5:
                        cool = min(30.0, 8.0 + cur_delay * 5)
                        log(f"Useita epäonnistumisia peräkkäin -> tauko {cool:.1f}s")
                        time.sleep(cool)
                        failures_in_row = 0

            finally:
                try:
                    driver.quit()
                except Exception:
                    pass

            save_word_plain_lines(emails, "sahkopostit.docx")

            self.set_status("Valmis!")
            messagebox.showinfo(
                "Valmis",
                f"Valmis!\n\nKansio:\n{OUT_DIR}\n\nTiedostot:\n- ytunnukset.docx\n- sahkopostit.docx\n\nLogi:\n{LOG_PATH}"
            )

        except Exception as e:
            log(f"VIRHE: {e}")
            self.set_status("Virhe. Katso log.txt")
            messagebox.showerror("Virhe", f"Tuli virhe.\nKatso log.txt:\n{LOG_PATH}\n\n{e}")


if __name__ == "__main__":
    App().mainloop()
