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
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from webdriver_manager.chrome import ChromeDriverManager


# ---------- Regex ----------
YT_RE = re.compile(r"\b\d{7}-\d\b|\b\d{8}\b")
EMAIL_RE = re.compile(r"[A-Za-z0-9_.+-]+@[A-Za-z0-9-]+\.[A-Za-z0-9-.]+")
EMAIL_A_RE = re.compile(r"[A-Za-z0-9_.+-]+\s*\(a\)\s*[A-Za-z0-9-]+\.[A-Za-z0-9-.]+", re.I)

KAUPPALEHTI_URL = "https://www.kauppalehti.fi/yritykset/protestilista"
YTJ_COMPANY_URL = "https://tietopalvelu.ytj.fi/yritys/{}"


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
# Helpers
# -----------------------------
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
        doc.add_paragraph(line)
    doc.save(path)
    log(f"Tallennettu: {path}")
    return path


def try_accept_cookies(driver):
    texts = [
        "Hyväksy", "Hyväksy kaikki", "Salli kaikki", "Accept", "Accept all",
        "I agree", "OK", "Selvä"
    ]
    for _ in range(2):
        found = False
        for e in driver.find_elements(By.XPATH, "//button|//a|//*[@role='button']"):
            try:
                t = (e.text or "").strip()
                if not t:
                    continue
                low = t.lower()
                if any(x.lower() in low for x in texts):
                    if e.is_displayed() and e.is_enabled():
                        driver.execute_script("arguments[0].scrollIntoView({block:'center'});", e)
                        e.click()
                        time.sleep(0.2)
                        found = True
                        break
            except Exception:
                continue
        if not found:
            break


# -----------------------------
# PDF -> Y-tunnukset
# -----------------------------
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


# -----------------------------
# Selenium start modes
# -----------------------------
def start_new_driver():
    options = webdriver.ChromeOptions()
    options.add_argument("--start-maximized")
    options.add_argument("--disable-gpu")
    options.add_argument("--disable-dev-shm-usage")
    driver_path = ChromeDriverManager().install()
    log("ChromeDriver OK (new session)")
    return webdriver.Chrome(service=Service(driver_path), options=options)


def attach_to_existing_chrome():
    """
    Liity jo auki olevaan Chromeen, joka on käynnistetty:
    --remote-debugging-port=9222 --user-data-dir=...
    """
    options = webdriver.ChromeOptions()
    options.add_experimental_option("debuggerAddress", "127.0.0.1:9222")
    driver_path = ChromeDriverManager().install()
    log("Yritetään liittyä olemassa olevaan Chromeen portissa 9222…")
    return webdriver.Chrome(service=Service(driver_path), options=options)


def focus_kauppalehti_tab(driver):
    """
    Etsii avoimista välilehdistä Kauppalehden protestilistan ja vaihtaa siihen.
    """
    for handle in driver.window_handles:
        driver.switch_to.window(handle)
        url = (driver.current_url or "")
        if "kauppalehti.fi/yritykset/protestilista" in url:
            return True
    return False


# -----------------------------
# Kauppalehti: kerää Y-tunnukset
# -----------------------------
def expand_and_collect_yts(driver, collected_set):
    before = len(collected_set)

    # Klikkaa mahdollisia rivin avaajia (nuoli/chevron). Heuristiikka: tyhjät napit tai aria/title.
    candidates = driver.find_elements(By.XPATH, "//button|//*[@role='button']")
    for c in candidates:
        try:
            if not c.is_displayed() or not c.is_enabled():
                continue
            txt = (c.text or "").strip().lower()
            aria = (c.get_attribute("aria-label") or "").strip().lower()
            title = (c.get_attribute("title") or "").strip().lower()

            if "näytä lisää" in txt:
                continue

            if (txt == "" and (aria or title)) or ("avaa" in aria) or ("laajenna" in aria) or ("expand" in aria) or ("details" in aria):
                driver.execute_script("arguments[0].scrollIntoView({block:'center'});", c)
                c.click()
                time.sleep(0.03)
        except Exception:
            continue

    # Poimi Y-tunnukset sivun tekstistä
    try:
        body_text = driver.find_element(By.TAG_NAME, "body").text
        for m in YT_RE.findall(body_text):
            n = normalize_yt(m)
            if n:
                collected_set.add(n)
    except Exception:
        pass

    return len(collected_set) - before


def click_nayta_lisaa(driver):
    # Klikkaa "Näytä lisää" (button tai role=button)
    for b in driver.find_elements(By.XPATH, "//button|//*[@role='button']"):
        try:
            if not b.is_displayed() or not b.is_enabled():
                continue
            if (b.text or "").strip().lower() == "näytä lisää":
                driver.execute_script("arguments[0].scrollIntoView({block:'center'});", b)
                b.click()
                return True
        except Exception:
            continue
    return False


def collect_yts_from_kauppalehti_already_open(driver, status_cb):
    wait = WebDriverWait(driver, 25)
    wait.until(EC.presence_of_element_located((By.TAG_NAME, "body")))
    try_accept_cookies(driver)

    collected = set()
    rounds = 0
    stable_rounds = 0
    last_count = 0

    while True:
        rounds += 1
        status_cb(f"Kauppalehti: kerätään Y-tunnuksia… (kierros {rounds})")

        added = expand_and_collect_yts(driver, collected)

        if len(collected) == last_count and added == 0:
            stable_rounds += 1
        else:
            stable_rounds = 0
            last_count = len(collected)

        status_cb(f"Kauppalehti: kasassa {len(collected)} (lisätty {added})")

        # Näytä lisää
        if click_nayta_lisaa(driver):
            time.sleep(1.0)
            try:
                driver.find_element(By.TAG_NAME, "body").send_keys(Keys.END)
            except Exception:
                pass
            time.sleep(0.8)
            continue

        if stable_rounds >= 2:
            break

        time.sleep(0.7)

    return sorted(collected)


# -----------------------------
# YTJ: hae sähköpostit
# -----------------------------
def click_all_nayta_ytj(driver):
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


def wait_ytj_loaded(driver):
    wait = WebDriverWait(driver, 25)
    wait.until(EC.presence_of_element_located((By.TAG_NAME, "body")))
    try:
        wait.until(EC.presence_of_element_located(
            (By.XPATH, "//*[contains(normalize-space(.), 'Y-tunnus') or contains(normalize-space(.), 'Toiminimi') or contains(normalize-space(.), 'Sähköposti')]")
        ))
    except Exception:
        pass


def extract_email_from_ytj(driver):
    # mailto
    try:
        for a in driver.find_elements(By.TAG_NAME, "a"):
            href = (a.get_attribute("href") or "")
            if href.lower().startswith("mailto:"):
                return href.split(":", 1)[1].strip()
    except Exception:
        pass

    # rivi "Sähköposti"
    try:
        cells = driver.find_elements(
            By.XPATH,
            "//tr//*[self::td or self::th][contains(normalize-space(.), 'Sähköposti')]"
        )
        for c in cells:
            try:
                tr = c.find_element(By.XPATH, "ancestor::tr[1]")
                email = pick_email_from_text(tr.text or "")
                if email:
                    return email
            except Exception:
                continue
    except Exception:
        pass

    # fallback main/body
    try:
        main = driver.find_element(By.TAG_NAME, "main")
        email = pick_email_from_text(main.text or "")
        if email:
            return email
    except Exception:
        pass

    try:
        return pick_email_from_text(driver.find_element(By.TAG_NAME, "body").text or "")
    except Exception:
        return ""


def fetch_emails_from_ytj(driver, yt_list, status_cb):
    emails = []
    seen = set()

    for i, yt in enumerate(yt_list, start=1):
        status_cb(f"YTJ: {i}/{len(yt_list)} {yt}")
        driver.get(YTJ_COMPANY_URL.format(yt))
        wait_ytj_loaded(driver)
        try_accept_cookies(driver)

        click_all_nayta_ytj(driver)

        email = ""
        for _ in range(6):  # ~1.2s max
            email = extract_email_from_ytj(driver)
            if email:
                break
            time.sleep(0.2)

        if email:
            k = email.lower()
            if k not in seen:
                seen.add(k)
                emails.append(email)

        time.sleep(0.1)

    return emails


# -----------------------------
# GUI
# -----------------------------
class App(tk.Tk):
    def __init__(self):
        super().__init__()
        reset_log()

        self.title("ProtestiBotti (Kauppalehti / PDF → YTJ)")
        self.geometry("760x420")

        tk.Label(self, text="ProtestiBotti", font=("Arial", 18, "bold")).pack(pady=12)

        tk.Label(
            self,
            text="Moodit:\n1) Kauppalehti (kirjautunut Chrome auki) → Y-tunnukset → YTJ sähköpostit\n2) PDF → Y-tunnukset → YTJ sähköpostit",
            justify="center"
        ).pack(pady=6)

        btn_row = tk.Frame(self)
        btn_row.pack(pady=10)

        tk.Button(
            btn_row,
            text="Kauppalehti (avaa kirjautuneesta Chromesta)",
            font=("Arial", 12),
            command=self.start_kauppalehti_mode
        ).grid(row=0, column=0, padx=8, pady=6)

        tk.Button(
            btn_row,
            text="PDF → YTJ sähköpostit",
            font=("Arial", 12),
            command=self.start_pdf_mode
        ).grid(row=0, column=1, padx=8, pady=6)

        self.status = tk.Label(self, text="Valmiina.", font=("Arial", 11))
        self.status.pack(pady=10)

        tk.Label(self, text=f"Tallennus: {OUT_DIR}", wraplength=720, justify="center").pack(pady=6)

    def set_status(self, s):
        self.status.config(text=s)
        self.update_idletasks()
        log(s)

    # ---- Mode 1: Kauppalehti existing logged-in chrome ----
    def start_kauppalehti_mode(self):
        threading.Thread(target=self.run_kauppalehti_mode, daemon=True).start()

    def run_kauppalehti_mode(self):
        driver = None
        try:
            self.set_status("Liitytään kirjautuneeseen Chromeen (portti 9222)…")
            driver = attach_to_existing_chrome()

            # Etsi protestilista-tab
            if not focus_kauppalehti_tab(driver):
                messagebox.showerror(
                    "Ei löytynyt Kauppalehteä",
                    "En löytänyt välilehteä jossa on kauppalehti.fi/yritykset/protestilista.\n\n"
                    "Varmista että:\n"
                    "1) Chrome avattiin komennolla remote-debugging-port=9222\n"
                    "2) Olet kirjautunut Kauppalehteen siinä Chromessa\n"
                    "3) Protestilista on auki välilehdessä"
                )
                self.set_status("Keskeytetty.")
                return

            self.set_status("Kauppalehti: kerätään Y-tunnukset…")
            yt_list = collect_yts_from_kauppalehti_already_open(driver, self.set_status)

            if not yt_list:
                self.set_status("Ei löytynyt Y-tunnuksia.")
                messagebox.showwarning("Ei löytynyt", "Kauppalehden listalta ei löytynyt Y-tunnuksia.")
                return

            save_word_plain_lines(yt_list, "ytunnukset.docx")

            self.set_status("YTJ: haetaan sähköpostit…")
            emails = fetch_emails_from_ytj(driver, yt_list, self.set_status)
            save_word_plain_lines(emails, "sahkopostit.docx")

            self.set_status("Valmis!")
            messagebox.showinfo(
                "Valmis",
                f"Valmis!\n\nKansio:\n{OUT_DIR}\n\nY-tunnuksia: {len(yt_list)}\nSähköposteja: {len(emails)}"
            )

        except Exception as e:
            log(f"VIRHE: {e}")
            self.set_status("Virhe. Katso log.txt")
            messagebox.showerror("Virhe", f"Tuli virhe.\nKatso log.txt:\n{LOG_PATH}\n\n{e}")
        finally:
            # ÄLÄ sulje käyttäjän chromea (koska se on “existing”)
            # driver.quit() sulkee joskus koko sessiota; jätetään pois.
            pass

    # ---- Mode 2: PDF -> YTJ ----
    def start_pdf_mode(self):
        path = filedialog.askopenfilename(filetypes=[("PDF files", "*.pdf")])
        if path:
            threading.Thread(target=self.run_pdf_mode, args=(path,), daemon=True).start()

    def run_pdf_mode(self, pdf_path):
        driver = None
        try:
            self.set_status("Luetaan PDF ja kerätään Y-tunnukset…")
            yt_list = extract_ytunnukset_from_pdf(pdf_path)

            if not yt_list:
                self.set_status("Ei löytynyt Y-tunnuksia PDF:stä.")
                messagebox.showwarning("Ei löytynyt", "PDF:stä ei löytynyt yhtään Y-tunnusta.")
                return

            save_word_plain_lines(yt_list, "ytunnukset.docx")

            self.set_status("Käynnistetään Chrome ja haetaan sähköpostit YTJ:stä…")
            driver = start_new_driver()

            emails = fetch_emails_from_ytj(driver, yt_list, self.set_status)
            save_word_plain_lines(emails, "sahkopostit.docx")

            self.set_status("Valmis!")
            messagebox.showinfo(
                "Valmis",
                f"Valmis!\n\nKansio:\n{OUT_DIR}\n\nY-tunnuksia: {len(yt_list)}\nSähköposteja: {len(emails)}"
            )

        except Exception as e:
            log(f"VIRHE: {e}")
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
