import os
import re
import sys
import time
import threading
import tkinter as tk
from tkinter import messagebox

from docx import Document

from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from webdriver_manager.chrome import ChromeDriverManager


# --------- Regex ----------
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


def start_driver():
    options = webdriver.ChromeOptions()
    options.add_argument("--start-maximized")
    options.add_argument("--disable-gpu")
    options.add_argument("--disable-dev-shm-usage")
    driver_path = ChromeDriverManager().install()
    log("ChromeDriver OK")
    return webdriver.Chrome(service=Service(driver_path), options=options)


def try_accept_cookies(driver):
    """
    Yritä klikata yleisimmät cookie-accept napit (fi/en).
    Ei kaadu jos ei löydy.
    """
    texts = [
        "Hyväksy", "Hyväksy kaikki", "Salli kaikki", "Accept", "Accept all",
        "I agree", "OK", "Selvä"
    ]
    # kokeillaan button + a + div role=button
    for _ in range(3):
        found = False
        elems = driver.find_elements(By.XPATH, "//button|//a|//*[@role='button']")
        for e in elems:
            try:
                t = (e.text or "").strip()
                if not t:
                    continue
                low = t.lower()
                if any(x.lower() in low for x in texts):
                    if e.is_displayed() and e.is_enabled():
                        driver.execute_script("arguments[0].scrollIntoView({block:'center'});", e)
                        e.click()
                        time.sleep(0.3)
                        found = True
                        break
            except Exception:
                continue
        if not found:
            break


# -----------------------------
# 1) Kauppalehti: kerää Y-tunnukset
# -----------------------------
def expand_all_visible_rows_and_collect_yts(driver, collected_set):
    """
    Klikkaa näkyvien rivien "avaa"/chevronit ja poimii Y-tunnukset aukeavista lisäriveistä.
    Tehdään turvallisesti: ensin klikkaus, sitten luetaan sivun teksti-alueelta Y-tunnuksia.
    """
    before = len(collected_set)

    # KL:n listassa on riveissä oikealla "nuoli" (button / role=button / svg parent).
    # Etsitään kaikki klikattavat elementit joissa aria-label tai title voi viitata avaamiseen,
    # mutta jos ei löydy, klikataan myös kaikki pienet buttonit joiden tekstinä ei ole mitään.
    candidates = driver.find_elements(By.XPATH, "//button|//*[@role='button']")

    for c in candidates:
        try:
            if not c.is_displayed() or not c.is_enabled():
                continue
            txt = (c.text or "").strip().lower()
            aria = (c.get_attribute("aria-label") or "").strip().lower()
            title = (c.get_attribute("title") or "").strip().lower()

            # Skip: Näytä lisää nappi käsitellään erikseen
            if "näytä lisää" in txt:
                continue

            # Heuristiikka: avausnapit usein tyhjiä tai aria/title kertoo
            if (txt == "" and (aria != "" or title != "")) or ("avaa" in aria) or ("laajenna" in aria) or ("expand" in aria) or ("details" in aria):
                driver.execute_script("arguments[0].scrollIntoView({block:'center'});", c)
                c.click()
                time.sleep(0.05)
        except Exception:
            continue

    # Nyt etsitään Y-tunnuksia sivun näkyvästä tekstistä
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
    """
    Klikkaa "Näytä lisää" jos löytyy. Palauttaa True jos klikattiin.
    """
    # Etsi selkeä nappi
    btns = driver.find_elements(By.XPATH, "//button|//*[@role='button']")
    for b in btns:
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


def collect_yts_from_kauppalehti(driver, status_cb):
    wait = WebDriverWait(driver, 25)
    driver.get(KAUPPALEHTI_URL)
    wait.until(EC.presence_of_element_located((By.TAG_NAME, "body")))
    try_accept_cookies(driver)

    collected = set()
    rounds = 0
    stable_rounds = 0
    last_count = 0

    while True:
        rounds += 1
        status_cb(f"Kauppalehti: käydään listaa läpi… (kierros {rounds})")

        added = expand_all_visible_rows_and_collect_yts(driver, collected)

        # jos mikään ei lisäänny moneen kierrokseen, ollaan ehkä valmiita
        if len(collected) == last_count and added == 0:
            stable_rounds += 1
        else:
            stable_rounds = 0
            last_count = len(collected)

        status_cb(f"Kauppalehti: Y-tunnuksia kasassa {len(collected)} (lisätty {added})")

        # kokeile "Näytä lisää"
        clicked = click_nayta_lisaa(driver)
        if clicked:
            time.sleep(1.0)
            try:
                # scroll down vähän että uusi sisältö latautuu
                driver.find_element(By.TAG_NAME, "body").send_keys(Keys.END)
            except Exception:
                pass
            time.sleep(0.8)
            continue

        # jos "Näytä lisää" ei löydy ja sisältö ei kasva -> lopeta
        if stable_rounds >= 2:
            break

        # varmistus: jos ei näytä lisää mutta ehkä lataus viive -> pieni odotus ja uusi kierros
        time.sleep(0.8)

    return sorted(collected)


# -----------------------------
# 2) YTJ: hae sähköposti suoraan yrityssivulta
# -----------------------------
def click_all_nayta_ytj(driver):
    """
    Klikkaa kaikki näkyvät 'Näytä' napit (button tai linkki) YTJ:n yrityssivulla.
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


def wait_ytj_loaded(driver):
    wait = WebDriverWait(driver, 25)
    wait.until(EC.presence_of_element_located((By.TAG_NAME, "body")))
    # odota, että yritystiedot näkyy (Y-tunnus/Toiminimi tms)
    try:
        wait.until(EC.presence_of_element_located(
            (By.XPATH, "//*[contains(normalize-space(.), 'Y-tunnus') or contains(normalize-space(.), 'Toiminimi') or contains(normalize-space(.), 'Sähköposti')]")
        ))
    except Exception:
        pass


def extract_email_from_ytj(driver):
    # 1) mailto-linkit
    try:
        for a in driver.find_elements(By.TAG_NAME, "a"):
            href = (a.get_attribute("href") or "")
            if href.lower().startswith("mailto:"):
                return href.split(":", 1)[1].strip()
    except Exception:
        pass

    # 2) etsi rivi jossa 'Sähköposti' ja poimi saman rivin teksti
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

    # 3) fallback: main-alueen tekstistä
    try:
        main = driver.find_element(By.TAG_NAME, "main")
        email = pick_email_from_text(main.text or "")
        if email:
            return email
    except Exception:
        pass

    # 4) fallback: koko sivu
    try:
        email = pick_email_from_text(driver.find_element(By.TAG_NAME, "body").text or "")
        return email
    except Exception:
        return ""


def fetch_emails_from_ytj(driver, yt_list, status_cb):
    emails = []
    seen = set()

    for i, yt in enumerate(yt_list, start=1):
        status_cb(f"YTJ: {i}/{len(yt_list)} {yt}")
        url = YTJ_COMPANY_URL.format(yt)
        driver.get(url)
        wait_ytj_loaded(driver)
        try_accept_cookies(driver)

        # klikkaa "Näytä" jos tietoja piilossa
        click_all_nayta_ytj(driver)

        # odota lyhyesti että email ehtii ilmestyä JS:n jälkeen
        email = ""
        for _ in range(6):  # max ~1.2s
            email = extract_email_from_ytj(driver)
            if email:
                break
            time.sleep(0.2)

        if email:
            key = email.lower()
            if key not in seen:
                seen.add(key)
                emails.append(email)

        # nopea loop
        time.sleep(0.1)

    return emails


# -----------------------------
# GUI
# -----------------------------
class App(tk.Tk):
    def __init__(self):
        super().__init__()
        reset_log()

        self.title("ProtestiBotti (Kauppalehti → YTJ)")
        self.geometry("680x360")

        tk.Label(self, text="ProtestiBotti", font=("Arial", 18, "bold")).pack(pady=12)
        tk.Label(self, text="Kerää Y-tunnukset Kauppalehden protestilistalta → hae sähköpostit YTJ:stä → Word",
                 justify="center").pack(pady=6)

        tk.Button(self, text="Aloita (Kauppalehti → YTJ)", font=("Arial", 12), command=self.start_job).pack(pady=12)

        self.status = tk.Label(self, text="Valmiina.", font=("Arial", 11))
        self.status.pack(pady=10)

        tk.Label(self, text=f"Tallennus: {OUT_DIR}", wraplength=640, justify="center").pack(pady=6)

    def set_status(self, s):
        self.status.config(text=s)
        self.update_idletasks()
        log(s)

    def start_job(self):
        threading.Thread(target=self.run_job, daemon=True).start()

    def run_job(self):
        driver = None
        try:
            self.set_status("Käynnistetään Chrome…")
            driver = start_driver()

            self.set_status("Kerätään Y-tunnukset Kauppalehdestä…")
            yt_list = collect_yts_from_kauppalehti(driver, self.set_status)

            if not yt_list:
                self.set_status("Ei löytynyt Y-tunnuksia Kauppalehdestä.")
                messagebox.showwarning("Ei löytynyt", "Kauppalehden listalta ei löytynyt Y-tunnuksia.")
                return

            save_word_plain_lines(yt_list, "ytunnukset.docx")

            self.set_status("Haetaan sähköpostit YTJ:stä…")
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
