# Finnish Business Email Finder (fix: avoid selenium.webdriver.chrome.options import)
# Modes: PDF->YTJ, Clipboard->(emails/yt/names)->YTJ
import os
import re
import sys
import time
import threading
from difflib import SequenceMatcher

import tkinter as tk
from tkinter import ttk, messagebox, filedialog

import PyPDF2
from docx import Document
from openpyxl import Workbook
from openpyxl.utils import get_column_letter
from openpyxl.styles import Font, Alignment

from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import TimeoutException
from webdriver_manager.chrome import ChromeDriverManager

try:
    from tkinterdnd2 import DND_FILES, TkinterDnD  # type: ignore
    HAS_DND = True
except Exception:
    HAS_DND = False


APP_BUILD = "2026-02-27_fix_selenium_options"

YTJ_PAGE_LOAD_TIMEOUT = 18
YTJ_RETRY_READS = 5
YTJ_RETRY_SLEEP = 0.12
YTJ_NAYTA_PASSES = 2
YTJ_PER_COMPANY_SLEEP = 0.02

NAME_SEARCH_TIMEOUT = 10.0
NAME_SEARCH_SLEEP = 0.18

YT_RE = re.compile(r"\b\d{7}-\d\b|\b\d{8}\b")
EMAIL_RE = re.compile(r"[A-Za-z0-9_.+-]+@[A-Za-z0-9-]+\.[A-Za-z0-9-.]+")
EMAIL_A_RE = re.compile(r"[A-Za-z0-9_.+-]+\s*\(a\)\s*[A-Za-z0-9-]+\.[A-Za-z0-9-.]+", re.I)

YTJ_COMPANY_URL = "https://tietopalvelu.ytj.fi/yritys/{}"
STRICT_FORMS_RE = re.compile(r"\b(oy|ab|ky|tmi|oyj|osakeyhtiö|kommandiittiyhtiö|toiminimi|as\.|ltd|llc|inc|gmbh)\b", re.I)


def exe_dir() -> str:
    if getattr(sys, "frozen", False):
        return os.path.dirname(sys.executable)
    return os.path.dirname(os.path.abspath(__file__))


def base_output_dir() -> str:
    base = exe_dir()
    try:
        p = os.path.join(base, "_write_test.tmp")
        with open(p, "w", encoding="utf-8") as f:
            f.write("ok")
        os.remove(p)
        return os.path.join(base, "FinnishBusinessEmailFinder")
    except Exception:
        home = os.path.expanduser("~")
        docs = os.path.join(home, "Documents")
        return os.path.join(docs, "FinnishBusinessEmailFinder")


def create_run_dir() -> str:
    root = base_output_dir()
    date_folder = time.strftime("%Y-%m-%d")
    run_folder = "run_" + time.strftime("%H-%M-%S")
    out = os.path.join(root, date_folder, run_folder)
    os.makedirs(out, exist_ok=True)
    return out


class RunContext:
    def __init__(self):
        self.run_dir = create_run_dir()
        self.log_path = os.path.join(self.run_dir, "log.txt")
        self._lock = threading.Lock()
        self.log("=== RUN START ===")
        self.log(f"Build: {APP_BUILD}")
        self.log(f"RunDir: {self.run_dir}")

    def log(self, msg: str):
        ts = time.strftime("%H:%M:%S")
        line = f"[{ts}] {msg}"
        with self._lock:
            try:
                with open(self.log_path, "a", encoding="utf-8") as f:
                    f.write(line + "\n")
            except Exception:
                pass
        return line


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


def extract_yts_from_text(text: str):
    yts = set()
    for m in YT_RE.findall(text or ""):
        n = normalize_yt(m)
        if n:
            yts.add(n)
    return sorted(yts)


def split_lines(text: str):
    if not text:
        return []
    text = text.replace("\r\n", "\n").replace("\r", "\n")
    return [ln.strip() for ln in text.split("\n") if ln.strip()]


def extract_names_from_text(text: str, strict: bool, max_names: int):
    lines = split_lines(text)
    out = []
    seen = set()
    bad_contains = ["näytä lisää", "kirjaudu", "tilaa", "tilaajille", "€", "y-tunnus", "ytunnus", "sähköposti", "puhelin", "www.", "http"]

    for ln in lines:
        if len(out) >= max_names:
            break
        if YT_RE.search(ln):
            continue
        low = ln.lower()
        if any(b in low for b in bad_contains):
            continue
        if len(ln) < 3:
            continue
        if sum(ch.isdigit() for ch in ln) >= 3:
            continue
        if not any(ch.isalpha() for ch in ln):
            continue
        name = re.sub(r"\s{2,}", " ", ln).strip()
        if len(name) > 90:
            continue
        if strict and not STRICT_FORMS_RE.search(name):
            continue
        key = name.lower()
        if key in seen:
            continue
        seen.add(key)
        out.append(name)
    return out


def save_emails_docx(run: RunContext, emails: list[str]):
    path = os.path.join(run.run_dir, "emails.docx")
    doc = Document()
    for e in emails:
        if e:
            doc.add_paragraph(e)
    doc.save(path)
    return path


def save_results_xlsx(run: RunContext, rows: list[dict], filename="results.xlsx"):
    path = os.path.join(run.run_dir, filename)
    wb = Workbook()
    ws = wb.active
    ws.title = "Results"
    headers = ["Name", "Y-tunnus", "Email", "Source", "Notes"]
    ws.append(headers)
    header_font = Font(bold=True)
    for col in range(1, len(headers) + 1):
        c = ws.cell(row=1, column=col)
        c.font = header_font
        c.alignment = Alignment(horizontal="left")

    for r in rows:
        ws.append([r.get("name", ""), r.get("yt", ""), r.get("email", ""), r.get("source", ""), r.get("notes", "")])

    for col in range(1, len(headers) + 1):
        max_len = 0
        for row in range(1, min(ws.max_row, 500) + 1):
            v = ws.cell(row=row, column=col).value
            if v is None:
                continue
            max_len = max(max_len, len(str(v)))
        ws.column_dimensions[get_column_letter(col)].width = min(60, max(12, max_len + 2))

    wb.save(path)
    return path


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


def start_new_driver():
    # IMPORTANT: do NOT import selenium.webdriver.chrome.options.Options
    options = webdriver.ChromeOptions()
    options.add_argument("--start-maximized")
    options.add_argument("--disable-gpu")
    options.add_argument("--disable-dev-shm-usage")

    driver_path = ChromeDriverManager().install()
    drv = webdriver.Chrome(service=Service(driver_path), options=options)
    drv.set_page_load_timeout(YTJ_PAGE_LOAD_TIMEOUT)
    return drv


def safe_click(driver, elem) -> bool:
    try:
        driver.execute_script("arguments[0].scrollIntoView({block:'center'});", elem)
        time.sleep(0.01)
        try:
            elem.click()
        except Exception:
            driver.execute_script("arguments[0].click();", elem)
        return True
    except Exception:
        return False


def try_accept_cookies(driver):
    texts = ["Hyväksy", "Hyväksy kaikki", "Salli kaikki", "Accept", "Accept all", "I agree", "OK", "Selvä"]
    for _ in range(2):
        for e in driver.find_elements(By.XPATH, "//button|//a|//*[@role='button']"):
            try:
                t = (e.text or "").strip()
                if not t:
                    continue
                if any(x.lower() in t.lower() for x in texts):
                    if e.is_displayed() and e.is_enabled():
                        safe_click(driver, e)
                        time.sleep(0.15)
                        break
            except Exception:
                continue


def wait_ytj_loaded(driver):
    WebDriverWait(driver, YTJ_PAGE_LOAD_TIMEOUT).until(EC.presence_of_element_located((By.TAG_NAME, "body")))


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
        return pick_email_from_text(driver.find_element(By.TAG_NAME, "body").text or "")
    except Exception:
        return ""


def click_all_nayta_ytj(driver):
    for _ in range(YTJ_NAYTA_PASSES):
        clicked = False
        for b in driver.find_elements(By.TAG_NAME, "button"):
            try:
                if (b.text or "").strip().lower() == "näytä" and b.is_displayed() and b.is_enabled():
                    safe_click(driver, b)
                    clicked = True
                    time.sleep(0.06)
            except Exception:
                continue
        for a in driver.find_elements(By.TAG_NAME, "a"):
            try:
                if (a.text or "").strip().lower() == "näytä" and a.is_displayed():
                    safe_click(driver, a)
                    clicked = True
                    time.sleep(0.06)
            except Exception:
                continue
        if not clicked:
            break


def fetch_email_by_yt(driver, yt: str, stop_flag: threading.Event):
    if stop_flag.is_set():
        return ""
    try:
        driver.get(YTJ_COMPANY_URL.format(yt))
    except TimeoutException:
        pass
    try:
        wait_ytj_loaded(driver)
    except Exception:
        pass
    try_accept_cookies(driver)

    email = ""
    for _ in range(2):
        if stop_flag.is_set():
            return ""
        email = extract_email_from_ytj(driver)
        if email:
            return email
        time.sleep(YTJ_RETRY_SLEEP)

    click_all_nayta_ytj(driver)
    for _ in range(YTJ_RETRY_READS):
        if stop_flag.is_set():
            return ""
        email = extract_email_from_ytj(driver)
        if email:
            return email
        time.sleep(YTJ_RETRY_SLEEP)
    return ""


def ytj_open_search_home(driver):
    driver.get("https://tietopalvelu.ytj.fi/")
    try:
        wait_ytj_loaded(driver)
    except Exception:
        pass
    try_accept_cookies(driver)


def find_ytj_search_input(driver):
    xpaths = [
        "//input[@type='search']",
        "//input[contains(translate(@placeholder,'HAE','hae'),'hae')]",
        "//input[contains(translate(@aria-label,'HAE','hae'),'hae')]",
        "//form//input",
        "//input",
    ]
    for xp in xpaths:
        try:
            cands = driver.find_elements(By.XPATH, xp)
        except Exception:
            cands = []
        for inp in cands:
            try:
                if not inp.is_displayed() or not inp.is_enabled():
                    continue
                t = (inp.get_attribute("type") or "").lower()
                if t in ("hidden", "password", "checkbox", "radio", "submit", "button"):
                    continue
                return inp
            except Exception:
                continue
    return None


def extract_yt_from_text_anywhere(txt: str) -> str:
    if not txt:
        return ""
    for m in YT_RE.findall(txt):
        n = normalize_yt(m)
        if n:
            return n
    return ""


def score_result(name_query: str, card_text: str) -> float:
    txt = (card_text or "").strip()
    ratio = SequenceMatcher(None, (name_query or "").lower(), txt.lower()).ratio()
    score = ratio * 100.0
    if extract_yt_from_text_anywhere(txt):
        score += 20.0
    return score


def ytj_name_to_yt(driver, name: str, stop_flag: threading.Event):
    ytj_open_search_home(driver)
    if stop_flag.is_set():
        return ""
    inp = find_ytj_search_input(driver)
    if not inp:
        return ""

    try:
        try:
            inp.clear()
        except Exception:
            pass
        inp.send_keys(name)
        inp.send_keys(u"\ue007")
    except Exception:
        return ""

    best_link = None
    best_score = -1.0
    t0 = time.time()
    while time.time() - t0 < NAME_SEARCH_TIMEOUT:
        if stop_flag.is_set():
            return ""
        candidate_links = []
        for xp in ("//a[contains(@href,'/yritys/')]", "//a[contains(@href,'yritys')]"):
            try:
                candidate_links.extend(driver.find_elements(By.XPATH, xp))
            except Exception:
                pass

        checked = 0
        for a in candidate_links:
            if checked >= 30:
                break
            try:
                if not a.is_displayed():
                    continue
                href = (a.get_attribute("href") or "")
                if not href or "tietopalvelu.ytj.fi" not in href:
                    continue
                try:
                    card = a.find_element(By.XPATH, "ancestor::*[self::li or self::div or self::article][1]")
                    card_text = (card.text or "")
                except Exception:
                    card_text = (a.text or "")
                s = score_result(name, card_text)
                checked += 1
                if s > best_score:
                    best_score = s
                    best_link = a
            except Exception:
                continue

        if best_link:
            break
        time.sleep(NAME_SEARCH_SLEEP)

    if not best_link:
        return ""

    try:
        safe_click(driver, best_link)
        try:
            wait_ytj_loaded(driver)
        except Exception:
            pass
        try_accept_cookies(driver)
        try:
            body = driver.find_element(By.TAG_NAME, "body").text or ""
        except Exception:
            body = ""
        return extract_yt_from_text_anywhere(body)
    except Exception:
        return ""


def pipeline_pdf(run: RunContext, pdf_path: str, status_cb, progress_cb, stop_flag: threading.Event):
    status_cb("Luetaan PDF ja kerätään Y-tunnukset…")
    yts = extract_ytunnukset_from_pdf(pdf_path)
    if not yts:
        status_cb("PDF: ei löytynyt Y-tunnuksia.")
        return [], []

    status_cb(f"PDF: löytyi {len(yts)} Y-tunnusta. Haetaan emailit YTJ:stä…")
    driver = start_new_driver()
    rows = []
    try:
        progress_cb(0, max(1, len(yts)))
        cache = {}
        for i, yt in enumerate(yts, start=1):
            if stop_flag.is_set():
                status_cb("Pysäytetty.")
                break
            progress_cb(i - 1, len(yts))
            status_cb(f"YTJ: {i}/{len(yts)} {yt}")
            email = cache.get(yt)
            if email is None:
                email = fetch_email_by_yt(driver, yt, stop_flag)
                cache[yt] = email
            rows.append({"name": "", "yt": yt, "email": email, "source": "pdf->ytj", "notes": ""})
            time.sleep(YTJ_PER_COMPANY_SLEEP)
        progress_cb(len(yts), max(1, len(yts)))
    finally:
        try:
            driver.quit()
        except Exception:
            pass

    emails = sorted({r["email"].lower(): r["email"] for r in rows if r.get("email")}.values())
    return rows, emails


def pipeline_clipboard(run: RunContext, text: str, strict: bool, max_names: int, status_cb, progress_cb, stop_flag: threading.Event):
    status_cb("Clipboard: poimitaan sähköpostit, Y-tunnukset ja yritysnimet…")
    direct_emails = set(e.strip().lower() for e in EMAIL_RE.findall(text or "") if e.strip())
    yts = extract_yts_from_text(text)
    names = extract_names_from_text(text, strict=strict, max_names=max_names)

    rows = []
    for em in sorted(direct_emails):
        rows.append({"name": "", "yt": "", "email": em, "source": "clipboard", "notes": "email found in pasted text"})
    for yt in yts:
        rows.append({"name": "", "yt": yt, "email": "", "source": "clipboard->ytj", "notes": "yt found in pasted text"})
    for nm in names:
        rows.append({"name": nm, "yt": "", "email": "", "source": "clipboard->ytj", "notes": "name found in pasted text"})

    seen = set()
    dedup = []
    for r in rows:
        key = (r.get("name", "").lower(), r.get("yt", ""), r.get("email", "").lower())
        if key in seen:
            continue
        seen.add(key)
        dedup.append(r)
    rows = dedup

    if not direct_emails and not yts and not names:
        status_cb("Clipboard: en löytänyt mitään (email / Y-tunnus / nimi).")
        return [], []

    status_cb("Käynnistetään YTJ-haku…")
    driver = start_new_driver()
    yt_cache: dict[str, str] = {}
    name_cache: dict[str, str] = {}
    try:
        todo = [r for r in rows if not r.get("email")]
        progress_cb(0, max(1, len(todo)))
        for idx, r in enumerate(todo, start=1):
            if stop_flag.is_set():
                status_cb("Pysäytetty.")
                break
            progress_cb(idx - 1, len(todo))

            name = (r.get("name") or "").strip()
            yt = (r.get("yt") or "").strip()

            if yt:
                status_cb(f"YTJ email: {idx}/{len(todo)} {yt}")
                email = yt_cache.get(yt)
                if email is None:
                    email = fetch_email_by_yt(driver, yt, stop_flag)
                    yt_cache[yt] = email
                if email:
                    r["email"] = email
            else:
                if not name:
                    continue
                status_cb(f"YTJ haku: {idx}/{len(todo)} {name}")
                yt2 = name_cache.get(name)
                if yt2 is None:
                    yt2 = ytj_name_to_yt(driver, name, stop_flag)
                    name_cache[name] = yt2
                if yt2:
                    r["yt"] = yt2
                    email2 = yt_cache.get(yt2)
                    if email2 is None:
                        email2 = fetch_email_by_yt(driver, yt2, stop_flag)
                        yt_cache[yt2] = email2
                    if email2:
                        r["email"] = email2
            time.sleep(YTJ_PER_COMPANY_SLEEP)
        progress_cb(len(todo), max(1, len(todo)))
    finally:
        try:
            driver.quit()
        except Exception:
            pass

    emails = sorted({(r.get("email") or "").lower(): (r.get("email") or "") for r in rows if r.get("email")}.values())
    return rows, emails


BaseTk = TkinterDnD.Tk if HAS_DND else tk.Tk


class App(BaseTk):
    def __init__(self):
        super().__init__()
        self.stop_flag = threading.Event()

        self.BG = "#ffffff"
        self.CARD = "#f6f8ff"
        self.BORDER = "#d7def7"
        self.TEXT = "#0b1b3a"
        self.MUTED = "#465a82"
        self.BLUE = "#1d4ed8"
        self.BLUE_H = "#1e40af"
        self.DANGER = "#b91c1c"
        self.DANGER_H = "#991b1b"
        self.OK = "#16a34a"

        self.title("Finnish Business Email Finder")
        self.geometry("980x860")
        self.configure(bg=self.BG)

        self.run = None
        self._build_ui()

    def _card(self, parent):
        return tk.Frame(parent, bg=self.CARD, highlightthickness=1, highlightbackground=self.BORDER)

    def _btn(self, parent, text, cmd, danger=False):
        bg = self.DANGER if danger else self.BLUE
        hover = self.DANGER_H if danger else self.BLUE_H
        b = tk.Button(
            parent, text=text, command=cmd,
            bg=bg, fg="#ffffff",
            activebackground=hover, activeforeground="#ffffff",
            relief="flat", padx=14, pady=10,
            font=("Segoe UI", 10, "bold")
        )
        b.bind("<Enter>", lambda e: b.configure(bg=hover))
        b.bind("<Leave>", lambda e: b.configure(bg=bg))
        return b

    def _ui_log(self, msg):
        ts = time.strftime("%H:%M:%S")
        line = f"[{ts}] {msg}"
        try:
            self.listbox.insert(tk.END, line)
            self.listbox.yview_moveto(1.0)
        except Exception:
            pass
        if self.run:
            self.run.log(msg)

    def _set_status(self, s):
        self.status.config(text=s)
        self.update_idletasks()
        self._ui_log(s)

    def _set_progress(self, v, mx):
        self.progress["maximum"] = mx
        self.progress["value"] = v
        self.update_idletasks()

    def request_stop(self):
        self.stop_flag.set()
        self._set_status("Pysäytetään…")

    def _clear_stop(self):
        self.stop_flag.clear()

    def _build_ui(self):
        header = tk.Frame(self, bg=self.BG)
        header.pack(fill="x", padx=16, pady=(16, 10))

        tk.Label(header, text="Finnish Business Email Finder", bg=self.BG, fg=self.TEXT,
                 font=("Segoe UI", 20, "bold")).pack(anchor="w")

        tk.Label(header, text=f"Build: {APP_BUILD}", bg=self.BG, fg=self.MUTED, font=("Segoe UI", 9)).pack(anchor="w")

        tk.Label(
            header,
            text="Liitä teksti (Ctrl+V) mistä tahansa sivusta tai valitse PDF — botti poimii emailit ja täydentää YTJ:stä.",
            bg=self.BG, fg=self.MUTED, font=("Segoe UI", 10)
        ).pack(anchor="w", pady=(4, 0))

        container = tk.Frame(self, bg=self.BG)
        container.pack(fill="both", expand=True, padx=16, pady=(0, 16))

        actions = self._card(container)
        actions.pack(fill="x", pady=(0, 12))
        row = tk.Frame(actions, bg=self.CARD)
        row.pack(fill="x", padx=12, pady=12)

        self._btn(row, "PDF → YTJ", self.start_pdf_mode).pack(side="left", padx=6)
        self._btn(row, "Clipboard → Finder", self.start_clipboard_mode).pack(side="left", padx=6)
        self._btn(row, "Pysäytä", self.request_stop, danger=True).pack(side="right", padx=6)

        status_card = self._card(container)
        status_card.pack(fill="x", pady=(0, 12))

        self.status = tk.Label(status_card, text="Valmiina.", bg=self.CARD, fg=self.TEXT, font=("Segoe UI", 11))
        self.status.pack(anchor="w", padx=12, pady=(12, 6))

        self.progress = ttk.Progressbar(status_card, orient="horizontal", mode="determinate", length=920)
        self.progress.pack(fill="x", padx=12, pady=(0, 12))

        pdf_card = self._card(container)
        pdf_card.pack(fill="x", pady=(0, 12))

        self.drop_var = tk.StringVar(value="PDF: Pudota tähän (tai paina PDF → YTJ ja valitse tiedosto)")
        drop = tk.Label(pdf_card, textvariable=self.drop_var,
                        bg="#ffffff", fg=self.MUTED, font=("Segoe UI", 10),
                        padx=12, pady=10,
                        highlightthickness=1, highlightbackground=self.BORDER)
        drop.pack(fill="x", padx=12, pady=12)

        if HAS_DND:
            drop.drop_target_register(DND_FILES)
            drop.dnd_bind("<<Drop>>", self._on_drop_pdf)

        clip_card = self._card(container)
        clip_card.pack(fill="x", pady=(0, 12))

        top = tk.Frame(clip_card, bg=self.CARD)
        top.pack(fill="x", padx=12, pady=(12, 6))

        self.strict_var = tk.BooleanVar(value=True)
        tk.Checkbutton(
            top, text="Tiukka parsinta (vain Oy/Ab/Ky/Tmi/...)",
            variable=self.strict_var,
            bg=self.CARD, fg=self.TEXT,
            selectcolor="#ffffff",
            activebackground=self.CARD, activeforeground=self.TEXT
        ).pack(side="left")

        tk.Label(top, text="Max nimeä:", bg=self.CARD, fg=self.TEXT, font=("Segoe UI", 10)).pack(side="left", padx=(16, 6))
        self.max_names_var = tk.IntVar(value=400)
        tk.Spinbox(top, from_=10, to=5000, textvariable=self.max_names_var, width=7,
                   bg="#ffffff", fg=self.TEXT, insertbackground=self.TEXT,
                   highlightthickness=1, highlightbackground=self.BORDER).pack(side="left")

        tk.Label(clip_card, text="Liitä tähän (Ctrl+V):", bg=self.CARD, fg=self.MUTED,
                 font=("Segoe UI", 10)).pack(anchor="w", padx=12, pady=(6, 6))

        self.clip_text = tk.Text(clip_card, height=9, wrap="word",
                                 bg="#ffffff", fg=self.TEXT, insertbackground=self.TEXT,
                                 highlightthickness=1, highlightbackground=self.BORDER)
        self.clip_text.pack(fill="x", padx=12, pady=(0, 12))

        log_card = self._card(container)
        log_card.pack(fill="both", expand=True)

        tk.Label(log_card, text="Logi:", bg=self.CARD, fg=self.MUTED, font=("Segoe UI", 10)).pack(anchor="w", padx=12, pady=(12, 6))

        body = tk.Frame(log_card, bg=self.CARD)
        body.pack(fill="both", expand=True, padx=12, pady=(0, 12))

        self.listbox = tk.Listbox(body, height=14, bg="#ffffff", fg=self.TEXT,
                                  highlightthickness=1, highlightbackground=self.BORDER,
                                  selectbackground="#dbeafe")
        self.listbox.pack(side="left", fill="both", expand=True)

        sb = ttk.Scrollbar(body, orient="vertical", command=self.listbox.yview)
        sb.pack(side="right", fill="y")
        self.listbox.configure(yscrollcommand=sb.set)

        self._ui_log("Ready.")

    def _on_drop_pdf(self, event):
        path = (event.data or "").strip()
        if path.startswith("{") and path.endswith("}"):
            path = path[1:-1]
        if path.lower().endswith(".pdf") and os.path.exists(path):
            self.drop_var.set(f"PDF valittu: {path}")
            self._clear_stop()
            threading.Thread(target=self._run_pdf, args=(path,), daemon=True).start()
        else:
            messagebox.showwarning("Ei PDF", "Pudotettu tiedosto ei ollut .pdf")

    def start_pdf_mode(self):
        self._clear_stop()
        path = filedialog.askopenfilename(filetypes=[("PDF files", "*.pdf")])
        if path:
            self.drop_var.set(f"PDF valittu: {path}")
            threading.Thread(target=self._run_pdf, args=(path,), daemon=True).start()

    def _run_pdf(self, pdf_path):
        self.run = RunContext()
        try:
            self._set_status("Aloitetaan PDF ajo…")
            rows, emails = pipeline_pdf(self.run, pdf_path, self._set_status, self._set_progress, self.stop_flag)
            if self.stop_flag.is_set():
                self._set_status("Pysäytetty.")
                return
            if not rows:
                self._set_status("Ei löytynyt tuloksia.")
                return
            save_results_xlsx(self.run, rows)
            save_emails_docx(self.run, emails)
            messagebox.showinfo("Valmis", f"Valmis!\n\nKansio:\n{self.run.run_dir}\n\nRivejä: {len(rows)}\nSähköposteja: {len(emails)}")
            self._set_status("Valmis!")
        except Exception as e:
            self._ui_log(f"VIRHE: {e}")
            messagebox.showerror("Virhe", f"Tuli virhe:\n\n{e}\n\nKatso log:\n{self.run.log_path}")

    def start_clipboard_mode(self):
        self._clear_stop()
        text = self.clip_text.get("1.0", tk.END).strip()
        if not text:
            messagebox.showwarning("Tyhjä", "Liitä ensin teksti (Ctrl+V) kenttään.")
            return
        threading.Thread(target=self._run_clipboard, args=(text,), daemon=True).start()

    def _run_clipboard(self, text: str):
        self.run = RunContext()
        try:
            self._set_status("Aloitetaan Clipboard ajo…")
            strict = bool(self.strict_var.get())
            max_names = int(self.max_names_var.get() or 400)
            rows, emails = pipeline_clipboard(self.run, text, strict, max_names, self._set_status, self._set_progress, self.stop_flag)
            if self.stop_flag.is_set():
                self._set_status("Pysäytetty.")
                return
            if not rows:
                self._set_status("Ei löytynyt tuloksia.")
                return
            save_results_xlsx(self.run, rows)
            save_emails_docx(self.run, emails)
            messagebox.showinfo("Valmis", f"Valmis!\n\nKansio:\n{self.run.run_dir}\n\nRivejä: {len(rows)}\nSähköposteja: {len(emails)}")
            self._set_status("Valmis!")
        except Exception as e:
            self._ui_log(f"VIRHE: {e}")
            messagebox.showerror("Virhe", f"Tuli virhe:\n\n{e}\n\nKatso log:\n{self.run.log_path}")


if __name__ == "__main__":
    App().mainloop()
