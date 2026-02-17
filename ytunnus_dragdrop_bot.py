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
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import StaleElementReferenceException
from webdriver_manager.chrome import ChromeDriverManager


# ---------- Regex ----------
YT_RE = re.compile(r"\b\d{7}-\d\b|\b\d{8}\b")
EMAIL_RE = re.compile(r"[A-Za-z0-9_.+-]+@[A-Za-z0-9-]+\.[A-Za-z0-9-.]+")
EMAIL_A_RE = re.compile(r"[A-Za-z0-9_.+-]+\s*\(a\)\s*[A-Za-z0-9-]+\.[A-Za-z0-9-.]+", re.I)

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
    return path


def try_accept_cookies(driver):
    texts = ["Hyväksy", "Hyväksy kaikki", "Salli kaikki", "Accept", "Accept all", "I agree", "OK", "Selvä"]
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
    return webdriver.Chrome(service=Service(driver_path), options=options)


def attach_to_existing_chrome():
    options = webdriver.ChromeOptions()
    options.add_experimental_option("debuggerAddress", "127.0.0.1:9222")
    driver_path = ChromeDriverManager().install()
    return webdriver.Chrome(service=Service(driver_path), options=options)


def focus_kauppalehti_tab(driver):
    for handle in driver.window_handles:
        driver.switch_to.window(handle)
        url = (driver.current_url or "")
        if "kauppalehti.fi/yritykset/protestilista" in url:
            return True
    return False


# -----------------------------
# Kauppalehti: click only company row cell (not header)
# -----------------------------
def click_nayta_lisaa(driver):
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


def get_company_rows_tbody(driver):
    """
    Palauttaa taulukon TBODY-rivit, joissa on yrityslinkki.
    Tämä EI ikinä ota headeria.
    """
    rows = []
    # Usein taulukko on thead + tbody
    candidates = driver.find_elements(By.XPATH, "//table//tbody//tr")
    for r in candidates:
        try:
            if not r.is_displayed():
                continue
            # skip expanded detail rows that contain Y-TUNNUS text
            if "Y-TUNNUS" in (r.text or ""):
                continue
            # must contain a company link
            links = r.find_elements(By.XPATH, ".//a[contains(@href,'/yritykset/') and normalize-space(.)!='']")
            if not links:
                continue
            rows.append(r)
        except Exception:
            continue
    return rows


def open_row_by_hairiopaiva_cell(row):
    """
    Klikkaa Häiriöpäivä-sarakkeen solua (4. sarake).
    Sarakkeet: Yritys(1) Sijainti(2) Summa(3) Häiriöpäivä(4) Tyyppi(5) Lähde(6) Nuoli(7)
    """
    try:
        cell = row.find_element(By.XPATH, ".//td[4]")
        return cell
    except Exception:
        return None


def open_row_by_toggle_button(row):
    """
    Fallback: oikean laidan nuoli/button (aria-expanded)
    """
    try:
        btn = row.find_element(By.XPATH, ".//button[@aria-expanded='true' or @aria-expanded='false']")
        return btn
    except Exception:
        return None


def extract_yt_after_open(anchor_elem):
    """
    Avauksen jälkeen rivin alle tulee detail-alue jossa 'Y-TUNNUS' + numero.
    Haetaan anchorin jälkeen lähin Y-TUNNUS ja otetaan siitä löytyvä y-tunnus.
    """
    try:
        y_label = anchor_elem.find_element(By.XPATH, "following:://*[contains(normalize-space(.), 'Y-TUNNUS')][1]")
        block = y_label.find_element(By.XPATH, "ancestor::*[self::div or self::tr][1]")
        found = YT_RE.findall(block.text or "")
        for m in found:
            n = normalize_yt(m)
            if n:
                return n
    except Exception:
        pass
    return ""


def collect_yts_from_kauppalehti(driver, status_cb, log_cb):
    wait = WebDriverWait(driver, 25)
    wait.until(EC.presence_of_element_located((By.TAG_NAME, "body")))
    try_accept_cookies(driver)

    collected = set()
    seen_rows = set()

    def row_signature(row):
        # sormenjälki: yritysnimi-linkin teksti + häiriöpäivä + summa
        try:
            name = row.find_element(By.XPATH, ".//a[contains(@href,'/yritykset/') and normalize-space(.)!='']").text.strip()
        except Exception:
            name = ""
        try:
            date = row.find_element(By.XPATH, ".//td[4]").text.strip()
        except Exception:
            date = ""
        try:
            amount = row.find_element(By.XPATH, ".//td[3]").text.strip()
        except Exception:
            amount = ""
        return f"{name}|{date}|{amount}"

    while True:
        rows = get_company_rows_tbody(driver)
        if not rows:
            status_cb("Kauppalehti: en löydä yritysrivejä (tbody). Onko lista näkyvissä + kirjautuneena?")
            break

        status_cb(f"Kauppalehti: näkyviä yritysrivejä {len(rows)} | kerätty {len(collected)}")

        new_in_pass = 0

        for row in rows:
            try:
                sig = row_signature(row)
                if sig in seen_rows:
                    continue
                seen_rows.add(sig)

                # 1) yritä klikata häiriöpäivä-solu
                anchor = open_row_by_hairiopaiva_cell(row)

                # jos ei onnistu, käytä nuolta
                if anchor is None:
                    anchor = open_row_by_toggle_button(row)
                if anchor is None:
                    continue

                driver.execute_script("arguments[0].scrollIntoView({block:'center'});", anchor)
                time.sleep(0.03)

                try:
                    driver.execute_script("arguments[0].click();", anchor)
                except StaleElementReferenceException:
                    continue

                yt = ""
                for _ in range(25):  # max ~2.5s
                    yt = extract_yt_after_open(anchor)
                    if yt:
                        break
                    time.sleep(0.1)

                if yt and yt not in collected:
                    collected.add(yt)
                    new_in_pass += 1
                    log_cb(f"+ {yt} (yht {len(collected)})")

                # yritä sulkea: klikkaa samaa anchor-elementtiä uudestaan
                try:
                    driver.execute_script("arguments[0].click();", anchor)
                except Exception:
                    pass

                time.sleep(0.02)

            except Exception:
                continue

        # Näytä lisää
        if click_nayta_lisaa(driver):
            status_cb("Kauppalehti: Näytä lisää…")
            time.sleep(1.2)
            try:
                driver.execute_script("window.scrollTo(0, document.body.scrollHeight);")
            except Exception:
                pass
            time.sleep(0.8)
            continue

        # jos ei enää uusia ja ei lisää-nappia
        if new_in_pass == 0:
            status_cb("Kauppalehti: ei uusia + ei Näytä lisää -> valmis.")
            break

    return sorted(collected)


# -----------------------------
# YTJ: sähköpostit
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
    try:
        for a in driver.find_elements(By.TAG_NAME, "a"):
            href = (a.get_attribute("href") or "")
            if href.lower().startswith("mailto:"):
                return href.split(":", 1)[1].strip()
    except Exception:
        pass

    try:
        cells = driver.find_elements(By.XPATH, "//tr//*[self::td or self::th][contains(normalize-space(.), 'Sähköposti')]")
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


def fetch_emails_from_ytj(driver, yt_list, status_cb, progress_cb, log_cb):
    emails = []
    seen = set()

    progress_cb(0, max(1, len(yt_list)))

    for i, yt in enumerate(yt_list, start=1):
        status_cb(f"YTJ: {i}/{len(yt_list)} {yt}")
        progress_cb(i - 1, len(yt_list))

        driver.get(YTJ_COMPANY_URL.format(yt))
        wait_ytj_loaded(driver)
        try_accept_cookies(driver)

        click_all_nayta_ytj(driver)

        email = ""
        for _ in range(8):
            email = extract_email_from_ytj(driver)
            if email:
                break
            time.sleep(0.2)

        if email:
            k = email.lower()
            if k not in seen:
                seen.add(k)
                emails.append(email)
                log_cb(email)

        time.sleep(0.1)

    progress_cb(len(yt_list), max(1, len(yt_list)))
    return emails


# -----------------------------
# GUI
# -----------------------------
class App(tk.Tk):
    def __init__(self):
        super().__init__()
        reset_log()

        self.title("ProtestiBotti (tbody-row click)")
        self.geometry("920x580")

        tk.Label(self, text="ProtestiBotti", font=("Arial", 18, "bold")).pack(pady=10)

        tk.Label(
            self,
            text="Moodit:\n1) Kauppalehti (kirjautunut Chrome auki, portti 9222) → Y-tunnukset → YTJ sähköpostit\n2) PDF → Y-tunnukset → YTJ sähköpostit",
            justify="center"
        ).pack(pady=4)

        btn_row = tk.Frame(self)
        btn_row.pack(pady=8)

        tk.Button(
            btn_row,
            text="Kauppalehti → YTJ",
            font=("Arial", 12),
            command=self.start_kauppalehti_mode
        ).grid(row=0, column=0, padx=8)

        tk.Button(
            btn_row,
            text="PDF → YTJ",
            font=("Arial", 12),
            command=self.start_pdf_mode
        ).grid(row=0, column=1, padx=8)

        self.status = tk.Label(self, text="Valmiina.", font=("Arial", 11))
        self.status.pack(pady=6)

        self.progress = ttk.Progressbar(self, orient="horizontal", mode="determinate", length=840)
        self.progress.pack(pady=6)

        frame = tk.Frame(self)
        frame.pack(fill="both", expand=True, padx=14, pady=10)

        tk.Label(frame, text="Live-logi (uusimmat alimmaisena):").pack(anchor="w")
        self.listbox = tk.Listbox(frame, height=16)
        self.listbox.pack(side="left", fill="both", expand=True)

        sb = tk.Scrollbar(frame, orient="vertical", command=self.listbox.yview)
        sb.pack(side="right", fill="y")
        self.listbox.configure(yscrollcommand=sb.set)

        tk.Label(self, text=f"Tallennus: {OUT_DIR}", wraplength=900, justify="center").pack(pady=6)

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
        self.progress["maximum"] = maximum
        self.progress["value"] = value
        self.update_idletasks()

    def start_kauppalehti_mode(self):
        threading.Thread(target=self.run_kauppalehti_mode, daemon=True).start()

    def run_kauppalehti_mode(self):
        driver = None
        try:
            self.set_status("Liitytään kirjautuneeseen Chromeen (9222)…")
            driver = attach_to_existing_chrome()

            if not focus_kauppalehti_tab(driver):
                messagebox.showerror(
                    "Ei löytynyt Kauppalehteä",
                    "En löytänyt välilehteä jossa on kauppalehti.fi/yritykset/protestilista.\n\n"
                    "Varmista että avasit Chromen debug-tilassa (portti 9222) ja protestilista on auki."
                )
                self.set_status("Keskeytetty.")
                return

            self.set_status("Kauppalehti: kerätään Y-tunnukset (tbody rivit + td[4] klikkaus)…")
            yt_list = collect_yts_from_kauppalehti(driver, self.set_status, self.ui_log)

            if not yt_list:
                self.set_status("Ei löytynyt Y-tunnuksia.")
                messagebox.showwarning("Ei löytynyt", "Kauppalehden listalta ei löytynyt Y-tunnuksia.")
                return

            yt_path = save_word_plain_lines(yt_list, "ytunnukset.docx")
            self.ui_log(f"Tallennettu: {yt_path}")

            self.set_status("YTJ: haetaan sähköpostit…")
            emails = fetch_emails_from_ytj(driver, yt_list, self.set_status, self.set_progress, self.ui_log)

            em_path = save_word_plain_lines(emails, "sahkopostit.docx")
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
            pass

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

            yt_path = save_word_plain_lines(yt_list, "ytunnukset.docx")
            self.ui_log(f"Tallennettu: {yt_path}")

            self.set_status("Käynnistetään Chrome ja haetaan sähköpostit YTJ:stä…")
            driver = start_new_driver()

            emails = fetch_emails_from_ytj(driver, yt_list, self.set_status, self.set_progress, self.ui_log)
            em_path = save_word_plain_lines(emails, "sahkopostit.docx")
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
