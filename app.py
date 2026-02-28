# app.py
# Finnish Business Email Finder (MVP Sellable EXE)
# Modes:
# 1) KL Protesti (URL) -> Expand rows -> Extract Y-tunnus -> YTJ -> Email
# 2) PDF -> extract Y-tunnus -> YTJ -> Email
# 3) Clipboard -> extract emails + Y-tunnus + (strict names) -> YTJ -> Email
#
# IMPORTANT:
# - User logs in manually for KL (paywall safe). Bot starts only after "Jatka".
# - Output folder is created ONLY at the end, and only if there are results.
# - Output files: results.xlsx, results.csv, emails.docx

import os
import re
import sys
import time
import csv
import threading
from dataclasses import dataclass
from difflib import SequenceMatcher
from typing import Optional

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


APP_BUILD = "2026-02-28_MVP_KL_URL_login_output_end_only"

# --- tuning ---
PAGE_LOAD_TIMEOUT = 18
YTJ_PAGE_LOAD_TIMEOUT = 18
YTJ_RETRY_READS = 6
YTJ_RETRY_SLEEP = 0.15
YTJ_NAYTA_PASSES = 3
YTJ_PER_COMPANY_SLEEP = 0.05

NAME_SEARCH_TIMEOUT = 10.0
NAME_SEARCH_SLEEP = 0.18

KL_SHOW_MORE_MAX_PASSES = 300
KL_AFTER_CLICK_SLEEP = 0.18
KL_SCROLL_SLEEP = 0.12

# --- regex ---
YT_RE = re.compile(r"\b\d{7}-\d\b|\b\d{8}\b")
EMAIL_RE = re.compile(r"[A-Za-z0-9_.+-]+@[A-Za-z0-9-]+\.[A-Za-z0-9-.]+")
EMAIL_A_RE = re.compile(r"[A-Za-z0-9_.+-]+\s*\(a\)\s*[A-Za-z0-9-]+\.[A-Za-z0-9-.]+", re.I)

YTJ_COMPANY_URL = "https://tietopalvelu.ytj.fi/yritys/{}"
YTJ_SEARCH_URLS = ["https://tietopalvelu.ytj.fi/haku", "https://tietopalvelu.ytj.fi/"]

STRICT_FORMS_RE = re.compile(
    r"\b(oy|ab|ky|tmi|oyj|osakeyhtiö|kommandiittiyhtiö|toiminimi|as\.|ltd|llc|inc|gmbh)\b",
    re.I,
)

KL_PROTEST_DEFAULT_URL = "https://www.kauppalehti.fi/yritykset/protestilista"


# =========================
#   SCROLLABLE UI WRAPPER
# =========================
class ScrollableFrame(ttk.Frame):
    """Canvas-based scrollable frame. Mouse wheel scrolls the whole UI."""
    def __init__(self, container, *args, **kwargs):
        super().__init__(container, *args, **kwargs)

        self.canvas = tk.Canvas(self, borderwidth=0, highlightthickness=0)
        self.vsb = ttk.Scrollbar(self, orient="vertical", command=self.canvas.yview)
        self.canvas.configure(yscrollcommand=self.vsb.set)

        self.inner = ttk.Frame(self.canvas)
        self._win = self.canvas.create_window((0, 0), window=self.inner, anchor="nw")

        self.vsb.pack(side="right", fill="y")
        self.canvas.pack(side="left", fill="both", expand=True)

        self.inner.bind("<Configure>", self._on_inner_configure)
        self.canvas.bind("<Configure>", self._on_canvas_configure)

        self.canvas.bind_all("<MouseWheel>", self._on_mousewheel)  # Windows

    def _on_inner_configure(self, _event):
        self.canvas.configure(scrollregion=self.canvas.bbox("all"))

    def _on_canvas_configure(self, event):
        self.canvas.itemconfigure(self._win, width=event.width)

    def _on_mousewheel(self, event):
        self.canvas.yview_scroll(int(-1 * (event.delta / 120)), "units")


BaseTk = TkinterDnD.Tk if HAS_DND else tk.Tk


# =========================
#   PATHS (OUTPUT END ONLY)
# =========================
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


def create_final_run_dir() -> str:
    root = base_output_dir()
    date_folder = time.strftime("%Y-%m-%d")
    run_folder = "run_" + time.strftime("%H-%M-%S")
    out = os.path.join(root, date_folder, run_folder)
    os.makedirs(out, exist_ok=True)
    return out


# =========================
#   DATA MODEL
# =========================
@dataclass
class Row:
    name: str = ""
    location: str = ""
    amount: str = ""
    date: str = ""
    type: str = ""
    source_name: str = ""  # e.g. "KL"
    source_url: str = ""
    yt: str = ""
    email: str = ""
    notes: str = ""


# =========================
#   UTIL
# =========================
def normalize_yt(yt: str):
    yt = (yt or "").strip().replace(" ", "")
    if re.fullmatch(r"\d{7}-\d", yt):
        return yt
    if re.fullmatch(r"\d{8}", yt):
        return yt[:7] + "-" + yt[7]
    return None


def extract_yts_from_text(text: str):
    yts = set()
    for m in YT_RE.findall(text or ""):
        n = normalize_yt(m)
        if n:
            yts.add(n)
    return sorted(yts)


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


def split_lines(text: str):
    if not text:
        return []
    text = text.replace("\r\n", "\n").replace("\r", "\n")
    return [ln.strip() for ln in text.split("\n") if ln.strip()]


def extract_names_from_text(text: str, strict: bool, max_names: int):
    lines = split_lines(text)
    out = []
    seen = set()
    bad_contains = [
        "näytä lisää", "kirjaudu", "tilaa", "tilaajille",
        "€", "y-tunnus", "ytunnus", "sähköposti", "puhelin",
        "www.", "http"
    ]

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


# =========================
#   OUTPUT (END ONLY)
# =========================
def save_emails_docx(out_dir: str, emails: list[str]):
    path = os.path.join(out_dir, "emails.docx")
    doc = Document()
    for e in emails:
        if e:
            doc.add_paragraph(e)
    doc.save(path)
    return path


def _autosize_columns(ws, max_rows=800):
    for col in range(1, ws.max_column + 1):
        max_len = 0
        for row in range(1, min(ws.max_row, max_rows) + 1):
            v = ws.cell(row=row, column=col).value
            if v is None:
                continue
            max_len = max(max_len, len(str(v)))
        ws.column_dimensions[get_column_letter(col)].width = min(70, max(12, max_len + 2))


def save_results_xlsx(out_dir: str, rows: list[Row], source_label: str, source_url: str):
    path = os.path.join(out_dir, "results.xlsx")
    wb = Workbook()

    # Results
    ws = wb.active
    ws.title = "Results"

    headers = ["Name", "Location", "Amount", "Date", "Type", "Y-tunnus", "Email", "Source", "Notes"]
    ws.append(headers)
    header_font = Font(bold=True)
    for col in range(1, len(headers) + 1):
        c = ws.cell(row=1, column=col)
        c.font = header_font
        c.alignment = Alignment(horizontal="left")

    ok = [r for r in rows if (r.email or r.yt)]
    missing = [r for r in rows if not (r.email or r.yt)]

    for r in ok:
        ws.append([r.name, r.location, r.amount, r.date, r.type, r.yt, r.email, r.source_name, r.notes])
    ws.freeze_panes = "A2"
    ws.auto_filter.ref = f"A1:I{ws.max_row}"
    _autosize_columns(ws)

    # Missing sheet
    ws2 = wb.create_sheet("Missing")
    ws2.append(headers)
    for col in range(1, len(headers) + 1):
        c = ws2.cell(row=1, column=col)
        c.font = header_font
        c.alignment = Alignment(horizontal="left")
    for r in missing:
        ws2.append([r.name, r.location, r.amount, r.date, r.type, r.yt, r.email, r.source_name, r.notes])
    ws2.freeze_panes = "A2"
    ws2.auto_filter.ref = f"A1:I{ws2.max_row}"
    _autosize_columns(ws2)

    # Summary
    ws3 = wb.create_sheet("Summary")
    ws3.append(["Build", APP_BUILD])
    ws3.append(["Source", source_label])
    ws3.append(["Source URL", source_url])
    ws3.append(["Total rows", len(rows)])
    ws3.append(["Rows with Y-tunnus", sum(1 for r in rows if r.yt)])
    ws3.append(["Rows with Email", sum(1 for r in rows if r.email)])
    ws3.append(["Unique Y-tunnus", len({r.yt for r in rows if r.yt})])
    ws3.append(["Unique Emails", len({r.email.lower() for r in rows if r.email})])
    _autosize_columns(ws3)

    wb.save(path)
    return path


def save_results_csv(out_dir: str, rows: list[Row]):
    path = os.path.join(out_dir, "results.csv")
    with open(path, "w", newline="", encoding="utf-8") as f:
        w = csv.writer(f, delimiter=";")
        w.writerow(["Name", "Location", "Amount", "Date", "Type", "Y-tunnus", "Email", "Source", "Notes"])
        for r in rows:
            w.writerow([r.name, r.location, r.amount, r.date, r.type, r.yt, r.email, r.source_name, r.notes])
    return path


def open_folder(path: str):
    try:
        if sys.platform.startswith("win"):
            os.startfile(path)  # type: ignore
        elif sys.platform == "darwin":
            os.system(f'open "{path}"')
        else:
            os.system(f'xdg-open "{path}"')
    except Exception:
        pass


# =========================
#   PDF -> YTs
# =========================
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


# =========================
#   SELENIUM COMMON
# =========================
def start_new_driver():
    options = webdriver.ChromeOptions()
    options.add_argument("--start-maximized")
    options.add_argument("--disable-gpu")
    options.add_argument("--disable-dev-shm-usage")
    # keep visible browser (paywall/login friendly)
    driver_path = ChromeDriverManager().install()
    drv = webdriver.Chrome(service=Service(driver_path), options=options)
    drv.set_page_load_timeout(PAGE_LOAD_TIMEOUT)
    return drv


def safe_click(driver, elem) -> bool:
    try:
        driver.execute_script("arguments[0].scrollIntoView({block:'center'});", elem)
        time.sleep(0.02)
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
                        time.sleep(0.20)
                        break
            except Exception:
                continue


def wait_loaded(driver, timeout=PAGE_LOAD_TIMEOUT):
    WebDriverWait(driver, timeout).until(
        EC.presence_of_element_located((By.TAG_NAME, "body"))
    )


# =========================
#   YTJ EMAIL FETCH
# =========================
def extract_email_from_ytj(driver):
    # mailto:
    try:
        for a in driver.find_elements(By.TAG_NAME, "a"):
            href = (a.get_attribute("href") or "")
            if href.lower().startswith("mailto:"):
                return href.split(":", 1)[1].strip()
    except Exception:
        pass

    # labeled row
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

    # fallback
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
                    time.sleep(0.08)
            except Exception:
                continue
        for a in driver.find_elements(By.TAG_NAME, "a"):
            try:
                if (a.text or "").strip().lower() == "näytä" and a.is_displayed():
                    safe_click(driver, a)
                    clicked = True
                    time.sleep(0.08)
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
        wait_loaded(driver, timeout=YTJ_PAGE_LOAD_TIMEOUT)
    except Exception:
        pass
    try_accept_cookies(driver)

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


# =========================
#   YTJ NAME SEARCH (fallback)
# =========================
def _attr(el, name: str) -> str:
    try:
        return (el.get_attribute(name) or "").strip()
    except Exception:
        return ""


def find_ytj_company_search_input(driver):
    candidates = []
    try:
        inputs = driver.find_elements(By.XPATH, "//input")
    except Exception:
        inputs = []

    for inp in inputs:
        try:
            if not inp.is_displayed() or not inp.is_enabled():
                continue

            itype = (_attr(inp, "type") or "").lower()
            if itype in ("hidden", "password", "checkbox", "radio", "submit", "button", "file"):
                continue

            ph = _attr(inp, "placeholder").lower()
            al = _attr(inp, "aria-label").lower()
            nm = _attr(inp, "name").lower()
            iid = _attr(inp, "id").lower()
            text = " ".join([ph, al, nm, iid]).strip()

            if "y-tunnus" in text and ("yritys" not in text and "toiminimi" not in text and "nimi" not in text):
                continue

            score = 0
            if "yritys" in text:
                score += 50
            if "toiminimi" in text:
                score += 35
            if "nimi" in text:
                score += 25
            if itype == "search":
                score += 15
            if "hae" in text:
                score += 10
            if "y-tunnus" in text:
                score -= 5

            if score <= 0:
                continue

            candidates.append((score, inp))
        except Exception:
            continue

    if not candidates:
        try:
            for inp in driver.find_elements(By.XPATH, "//input[@type='search']"):
                if inp.is_displayed() and inp.is_enabled():
                    return inp
        except Exception:
            pass
        return None

    candidates.sort(key=lambda x: x[0], reverse=True)
    return candidates[0][1]


def ytj_open_search_home(driver):
    for url in YTJ_SEARCH_URLS:
        try:
            driver.get(url)
            wait_loaded(driver, timeout=YTJ_PAGE_LOAD_TIMEOUT)
            try_accept_cookies(driver)
            if find_ytj_company_search_input(driver):
                return
        except Exception:
            continue


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

    inp = find_ytj_company_search_input(driver)
    if not inp:
        return ""

    try:
        try:
            inp.clear()
        except Exception:
            pass
        inp.send_keys(name)
        inp.send_keys(u"\ue007")  # ENTER
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
            wait_loaded(driver, timeout=YTJ_PAGE_LOAD_TIMEOUT)
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


# =========================
#   KL PROTESTI SCRAPE
# =========================
def _text(el) -> str:
    try:
        return (el.text or "").strip()
    except Exception:
        return ""


def kl_click_show_more(driver, status_cb, stop_flag: threading.Event):
    """Click 'Näytä lisää' until it's gone or max passes reached."""
    passes = 0
    last_height = 0
    while passes < KL_SHOW_MORE_MAX_PASSES and not stop_flag.is_set():
        passes += 1

        # scroll a bit to make button appear
        try:
            driver.execute_script("window.scrollTo(0, document.body.scrollHeight);")
        except Exception:
            pass
        time.sleep(KL_SCROLL_SLEEP)

        try_accept_cookies(driver)

        btn = None
        # Find by text
        for xp in (
            "//button[contains(translate(normalize-space(.),'ÄÖÅ','äöå'),'näytä lisää')]",
            "//a[contains(translate(normalize-space(.),'ÄÖÅ','äöå'),'näytä lisää')]",
        ):
            try:
                candidates = driver.find_elements(By.XPATH, xp)
                for c in candidates:
                    if c.is_displayed() and c.is_enabled():
                        btn = c
                        break
                if btn:
                    break
            except Exception:
                continue

        if not btn:
            # no more button
            status_cb(f"KL: 'Näytä lisää' loppui ({passes-1} klikkausta).")
            return

        status_cb(f"KL: Klikataan 'Näytä lisää'... ({passes})")
        safe_click(driver, btn)
        time.sleep(KL_AFTER_CLICK_SLEEP)

        # detect height change (optional)
        try:
            h = int(driver.execute_script("return document.body.scrollHeight;") or 0)
            if h == last_height and passes > 3:
                # still continue a bit; some pages don't change height visibly
                pass
            last_height = h
        except Exception:
            pass

    status_cb("KL: Lopetettiin 'Näytä lisää' (max passes / stop).")


def kl_extract_rows_from_visible_page(driver, status_cb, stop_flag: threading.Event, max_rows: int):
    """
    Extract company rows from the page after expanding.
    Approach:
    - Find table-like rows by scanning link texts + nearby container text.
    - Expand details by clicking chevrons / aria-expanded.
    - After expansions, parse all visible texts containing 'Y-TUNNUS'.
    """
    out: list[Row] = []
    seen_yt = set()

    # Step 1: try expanding anything that looks expandable
    status_cb("KL: Avataan rivien lisätiedot (nuolet)...")
    for _pass in range(6):
        if stop_flag.is_set():
            break

        expanded_any = False
        # Best guess: buttons with aria-expanded false
        try:
            exp_buttons = driver.find_elements(By.XPATH, "//button[@aria-expanded='false']")
        except Exception:
            exp_buttons = []

        # Also try chevrons: buttons that contain svg and are in row end
        try:
            exp_buttons2 = driver.find_elements(By.XPATH, "//button[.//svg]")
        except Exception:
            exp_buttons2 = []

        candidates = []
        for b in exp_buttons + exp_buttons2:
            try:
                if not b.is_displayed() or not b.is_enabled():
                    continue
                candidates.append(b)
            except Exception:
                continue

        # Click a limited number per pass
        clicked = 0
        for b in candidates:
            if stop_flag.is_set():
                break
            if clicked >= 120:
                break
            try:
                # avoid clicking irrelevant (header filters etc.)
                aria = (b.get_attribute("aria-label") or "").lower()
                # if aria-label hints expand/collapse or empty -> ok
                # we still allow empties to catch icon-only buttons
                if aria and not any(k in aria for k in ["avaa", "näytä", "lisätied", "expand", "open", "details", "toggle"]):
                    # could still be ok, but skip to reduce risk
                    pass

                ok = safe_click(driver, b)
                if ok:
                    expanded_any = True
                    clicked += 1
                    time.sleep(0.04)
            except Exception:
                continue

        if not expanded_any:
            break
        time.sleep(0.20)

    # Step 2: parse all blocks that contain Y-TUNNUS
    status_cb("KL: Poimitaan Y-tunnukset sivulta...")
    blocks = []
    for xp in (
        "//*[contains(translate(normalize-space(.),'YÄÖÅ','yäöå'),'y-tunnus')]",
        "//*[contains(translate(normalize-space(.),'YÄÖÅ','yäöå'),'ytunnus')]",
    ):
        try:
            blocks.extend(driver.find_elements(By.XPATH, xp))
        except Exception:
            pass

    # Prefer ancestor blocks that look like detail sections
    candidate_containers = []
    for el in blocks:
        if stop_flag.is_set():
            break
        try:
            cont = el.find_element(By.XPATH, "ancestor::*[self::div or self::li or self::tr][1]")
            candidate_containers.append(cont)
        except Exception:
            candidate_containers.append(el)

    # Dedup containers by id / object
    uniq = []
    seen_ids = set()
    for c in candidate_containers:
        try:
            hid = c.id  # selenium internal id
        except Exception:
            hid = None
        key = hid or str(c)
        if key in seen_ids:
            continue
        seen_ids.add(key)
        uniq.append(c)

    # Extract row fields heuristically
    for cont in uniq:
        if stop_flag.is_set():
            break
        txt = _text(cont)
        yt = extract_yts_from_text(txt)
        if not yt:
            continue
        yt = yt[0]
        if yt in seen_yt:
            continue
        seen_yt.add(yt)

        # find name from nearest clickable link or previous row container
        name = ""
        location = ""
        amount = ""
        date = ""
        typ = ""

        try:
            # Look for an <a> in nearby previous sibling blocks
            try:
                link = cont.find_element(By.XPATH, ".//a[normalize-space(string())!='']")
                name = _text(link)
            except Exception:
                # fallback: try nearest preceding link
                try:
                    link2 = cont.find_element(By.XPATH, "preceding::a[normalize-space(string())!=''][1]")
                    name = _text(link2)
                except Exception:
                    name = ""

            # Try pulling summary row text from preceding container
            try:
                prev = cont.find_element(By.XPATH, "preceding::*[self::div or self::tr or self::li][1]")
                ptxt = _text(prev)
                # heuristics: first line = name, second = maybe location
                lines = split_lines(ptxt)
                if not name and lines:
                    name = lines[0]
                # Try guess location/amount/date/type from previous lines
                if lines:
                    for ln in lines:
                        if not location and any(ch.isalpha() for ch in ln) and len(ln) <= 30 and "€" not in ln and "." not in ln and "-" not in ln:
                            # location often short (e.g. Tampere)
                            location = ln
                        if not amount and "€" in ln:
                            amount = ln.strip()
                        if not date and re.search(r"\b\d{2}\.\d{2}\.\d{4}\b", ln):
                            date = re.search(r"\b\d{2}\.\d{2}\.\d{4}\b", ln).group(0)  # type: ignore
                        if not typ and any(k in ln.lower() for k in ["velkomus", "protest", "tuomio"]):
                            typ = ln.strip()
            except Exception:
                pass

            # Parse from detail block too (sometimes has area/velkoja etc.)
            # We keep only key items.
        except Exception:
            pass

        out.append(Row(
            name=name.strip(),
            location=location.strip(),
            amount=amount.strip(),
            date=date.strip(),
            type=typ.strip(),
            source_name="KL",
            source_url=driver.current_url,
            yt=yt,
            email="",
            notes="from KL (expanded row)"
        ))

        if max_rows and len(out) >= max_rows:
            break

    status_cb(f"KL: Löytyi {len(out)} uniikkia Y-tunnusta.")
    return out


# =========================
#   PIPELINES
# =========================
def pipeline_pdf(pdf_path: str, status_cb, progress_cb, stop_flag: threading.Event):
    status_cb("PDF: Luetaan ja kerätään Y-tunnukset…")
    yts = extract_ytunnukset_from_pdf(pdf_path)
    if not yts:
        status_cb("PDF: ei löytynyt Y-tunnuksia.")
        return [], []

    status_cb(f"PDF: löytyi {len(yts)} Y-tunnusta. Haetaan emailit YTJ:stä…")
    driver = start_new_driver()
    try:
        progress_cb(0, max(1, len(yts)))
        rows: list[Row] = []
        for i, yt in enumerate(yts, start=1):
            if stop_flag.is_set():
                break
            progress_cb(i - 1, len(yts))
            status_cb(f"YTJ: {i}/{len(yts)} {yt}")
            email = fetch_email_by_yt(driver, yt, stop_flag)
            rows.append(Row(source_name="PDF", source_url=pdf_path, yt=yt, email=email))
            time.sleep(YTJ_PER_COMPANY_SLEEP)
        progress_cb(len(yts), max(1, len(yts)))
    finally:
        try:
            driver.quit()
        except Exception:
            pass

    emails = sorted({r.email.lower(): r.email for r in rows if r.email}.values())
    return rows, emails


def pipeline_clipboard(text: str, strict: bool, max_names: int, status_cb, progress_cb, stop_flag: threading.Event):
    status_cb("Clipboard: poimitaan sähköpostit, Y-tunnukset ja yritysnimet…")

    direct_emails = set(e.strip().lower() for e in EMAIL_RE.findall(text or "") if e.strip())
    yts = extract_yts_from_text(text)
    names = extract_names_from_text(text, strict=strict, max_names=max_names)

    rows: list[Row] = []
    for em in sorted(direct_emails):
        rows.append(Row(email=em, source_name="Clipboard", source_url="clipboard", notes="email found in pasted text"))
    for yt in yts:
        rows.append(Row(yt=yt, source_name="Clipboard->YTJ", source_url="clipboard", notes="yt found in pasted text"))
    for nm in names:
        rows.append(Row(name=nm, source_name="Clipboard->YTJ", source_url="clipboard", notes="name found in pasted text"))

    # dedup (name, yt, email)
    deduped = []
    seen = set()
    for r in rows:
        key = (r.name.lower(), r.yt, r.email.lower())
        if key in seen:
            continue
        seen.add(key)
        deduped.append(r)
    rows = deduped

    if not direct_emails and not yts and not names:
        status_cb("Clipboard: en löytänyt mitään (email / Y-tunnus / nimi).")
        return [], []

    status_cb("Clipboard: Käynnistetään YTJ-haku…")
    driver = start_new_driver()

    yt_cache: dict[str, str] = {}
    name_cache: dict[str, str] = {}

    try:
        todo = [r for r in rows if not r.email]
        progress_cb(0, max(1, len(todo)))

        for idx, r in enumerate(todo, start=1):
            if stop_flag.is_set():
                break
            progress_cb(idx - 1, len(todo))

            if r.yt:
                status_cb(f"YTJ email: {idx}/{len(todo)} {r.yt}")
                if r.yt not in yt_cache:
                    yt_cache[r.yt] = fetch_email_by_yt(driver, r.yt, stop_flag)
                r.email = yt_cache.get(r.yt, "") or ""
            else:
                if not r.name:
                    continue
                status_cb(f"YTJ yrityshaku: {idx}/{len(todo)} {r.name}")
                if r.name not in name_cache:
                    name_cache[r.name] = ytj_name_to_yt(driver, r.name, stop_flag) or ""
                yt2 = name_cache.get(r.name, "") or ""
                r.yt = yt2
                if yt2:
                    if yt2 not in yt_cache:
                        yt_cache[yt2] = fetch_email_by_yt(driver, yt2, stop_flag)
                    r.email = yt_cache.get(yt2, "") or ""

            time.sleep(YTJ_PER_COMPANY_SLEEP)

        progress_cb(len(todo), max(1, len(todo)))
    finally:
        try:
            driver.quit()
        except Exception:
            pass

    emails = sorted({r.email.lower(): r.email for r in rows if r.email}.values())
    return rows, emails


def pipeline_kl_protest(url: str, max_companies: int, status_cb, progress_cb, stop_flag: threading.Event, driver):
    """
    Uses an existing driver (user logged in).
    1) Click show-more until all loaded
    2) Expand rows and extract YTs (+ some metadata)
    3) For each YT, fetch email from YTJ
    """
    if stop_flag.is_set():
        return [], []

    status_cb("KL: Ladataan sivu…")
    try:
        driver.get(url)
    except TimeoutException:
        pass
    try:
        wait_loaded(driver)
    except Exception:
        pass
    try_accept_cookies(driver)

    status_cb("KL: Ladataan kaikki rivit (Näytä lisää)…")
    kl_click_show_more(driver, status_cb, stop_flag)

    # Extract YTs from KL
    status_cb("KL: Poimitaan yritykset ja Y-tunnukset…")
    rows = kl_extract_rows_from_visible_page(driver, status_cb, stop_flag, max_rows=max_companies)

    if not rows:
        status_cb("KL: Ei löytynyt yhtään Y-tunnusta. (Oletko varmasti avannut protestilistan ja detailit?)")
        return [], []

    # Fetch emails from YTJ using a NEW driver to avoid messing KL session
    status_cb("YTJ: Käynnistetään email-haku…")
    ytj_driver = start_new_driver()
    try:
        todo = [r for r in rows if r.yt]
        progress_cb(0, max(1, len(todo)))
        yt_cache: dict[str, str] = {}

        for i, r in enumerate(todo, start=1):
            if stop_flag.is_set():
                break
            progress_cb(i - 1, len(todo))
            status_cb(f"YTJ: {i}/{len(todo)} {r.yt}")

            if r.yt not in yt_cache:
                yt_cache[r.yt] = fetch_email_by_yt(ytj_driver, r.yt, stop_flag)
            r.email = yt_cache.get(r.yt, "") or ""

            time.sleep(YTJ_PER_COMPANY_SLEEP)

        progress_cb(len(todo), max(1, len(todo)))
    finally:
        try:
            ytj_driver.quit()
        except Exception:
            pass

    emails = sorted({r.email.lower(): r.email for r in rows if r.email}.values())
    return rows, emails


# =========================
#   APP UI
# =========================
class App(BaseTk):
    def __init__(self):
        super().__init__()
        self.stop_flag = threading.Event()

        # Drivers
        self.kl_driver = None  # kept alive between "Open" and "Continue"
        self.last_output_dir: Optional[str] = None

        # theme
        self.BG = "#ffffff"
        self.CARD = "#f6f8ff"
        self.BORDER = "#d7def7"
        self.TEXT = "#0b1b3a"
        self.MUTED = "#465a82"
        self.BLUE = "#1d4ed8"
        self.BLUE_H = "#1e40af"
        self.DANGER = "#b91c1c"
        self.DANGER_H = "#991b1b"
        self.GREY_BTN = "#64748b"
        self.GREY_BTN_H = "#475569"

        self.title("Finnish Business Email Finder (MVP)")
        self.geometry("1020x940")
        self.configure(bg=self.BG)

        self._build_ui()

    def _card(self, parent):
        return tk.Frame(parent, bg=self.CARD, highlightthickness=1, highlightbackground=self.BORDER)

    def _btn(self, parent, text, cmd, kind="blue"):
        if kind == "danger":
            bg, hover = self.DANGER, self.DANGER_H
        elif kind == "grey":
            bg, hover = self.GREY_BTN, self.GREY_BTN_H
        else:
            bg, hover = self.BLUE, self.BLUE_H

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

    def clear_clipboard_text(self):
        self.clip_text.delete("1.0", tk.END)
        self._ui_log("Clipboard tyhjennetty.")

    def open_last_output(self):
        if self.last_output_dir and os.path.isdir(self.last_output_dir):
            open_folder(self.last_output_dir)
        else:
            messagebox.showinfo("Ei kansiota", "Ei vielä tuloskansiota (ajo ei valmis / ei tuloksia).")

    def _build_ui(self):
        outer = ScrollableFrame(self)
        outer.pack(fill="both", expand=True)
        root = outer.inner

        header = tk.Frame(root, bg=self.BG)
        header.pack(fill="x", padx=16, pady=(16, 10))

        tk.Label(header, text="Finnish Business Email Finder (MVP)", bg=self.BG, fg=self.TEXT,
                 font=("Segoe UI", 20, "bold")).pack(anchor="w")
        tk.Label(header, text=f"Build: {APP_BUILD}", bg=self.BG, fg=self.MUTED, font=("Segoe UI", 9)).pack(anchor="w")

        actions = self._card(root)
        actions.pack(fill="x", padx=16, pady=(0, 12))
        row = tk.Frame(actions, bg=self.CARD)
        row.pack(fill="x", padx=12, pady=12)

        self._btn(row, "KL Protesti (URL) → Avaa selain", self.kl_open_browser).pack(side="left", padx=6)
        self._btn(row, "KL → Jatka (aloita keruu)", self.kl_continue_scrape).pack(side="left", padx=6)

        self._btn(row, "PDF → YTJ", self.start_pdf_mode).pack(side="left", padx=6)
        self._btn(row, "Clipboard → Finder", self.start_clipboard_mode).pack(side="left", padx=6)

        self._btn(row, "Avaa tuloskansio", self.open_last_output, kind="grey").pack(side="right", padx=6)
        self._btn(row, "Pysäytä", self.request_stop, kind="danger").pack(side="right", padx=6)

        status_card = self._card(root)
        status_card.pack(fill="x", padx=16, pady=(0, 12))

        self.status = tk.Label(status_card, text="Valmiina.", bg=self.CARD, fg=self.TEXT, font=("Segoe UI", 11))
        self.status.pack(anchor="w", padx=12, pady=(12, 6))

        self.progress = ttk.Progressbar(status_card, orient="horizontal", mode="determinate", length=920)
        self.progress.pack(fill="x", padx=12, pady=(0, 12))

        # --- KL CARD ---
        kl_card = self._card(root)
        kl_card.pack(fill="x", padx=16, pady=(0, 12))

        top = tk.Frame(kl_card, bg=self.CARD)
        top.pack(fill="x", padx=12, pady=(12, 6))

        tk.Label(top, text="Kauppalehti Protestilista URL:", bg=self.CARD, fg=self.TEXT, font=("Segoe UI", 10, "bold")).pack(side="left")
        self.kl_url_var = tk.StringVar(value=KL_PROTEST_DEFAULT_URL)
        tk.Entry(top, textvariable=self.kl_url_var, width=70, bg="#ffffff", fg=self.TEXT,
                 highlightthickness=1, highlightbackground=self.BORDER).pack(side="left", padx=10)

        tk.Label(top, text="Max yritystä:", bg=self.CARD, fg=self.TEXT, font=("Segoe UI", 10)).pack(side="left", padx=(12, 6))
        self.kl_max_var = tk.IntVar(value=2000)
        tk.Spinbox(top, from_=50, to=50000, textvariable=self.kl_max_var, width=8,
                   bg="#ffffff", fg=self.TEXT, insertbackground=self.TEXT,
                   highlightthickness=1, highlightbackground=self.BORDER).pack(side="left")

        tk.Label(kl_card,
                 text="KL käyttö: Paina 'Avaa selain' → kirjaudu itse → varmista että protestilista näkyy → paina 'Jatka'.",
                 bg=self.CARD, fg=self.MUTED, font=("Segoe UI", 9)).pack(anchor="w", padx=12, pady=(0, 12))

        # --- PDF CARD ---
        pdf_card = self._card(root)
        pdf_card.pack(fill="x", padx=16, pady=(0, 12))
        self.drop_var = tk.StringVar(value="PDF: Pudota tähän (tai paina PDF → YTJ ja valitse tiedosto)")
        drop = tk.Label(pdf_card, textvariable=self.drop_var,
                        bg="#ffffff", fg=self.MUTED, font=("Segoe UI", 10),
                        padx=12, pady=10,
                        highlightthickness=1, highlightbackground=self.BORDER)
        drop.pack(fill="x", padx=12, pady=12)
        if HAS_DND:
            drop.drop_target_register(DND_FILES)
            drop.dnd_bind("<<Drop>>", self._on_drop_pdf)

        # --- CLIPBOARD CARD ---
        clip_card = self._card(root)
        clip_card.pack(fill="x", padx=16, pady=(0, 12))

        top2 = tk.Frame(clip_card, bg=self.CARD)
        top2.pack(fill="x", padx=12, pady=(12, 6))

        self.strict_var = tk.BooleanVar(value=True)
        tk.Checkbutton(
            top2, text="Tiukka parsinta (vain Oy/Ab/Ky/Tmi/...)",
            variable=self.strict_var,
            bg=self.CARD, fg=self.TEXT,
            selectcolor="#ffffff",
            activebackground=self.CARD, activeforeground=self.TEXT
        ).pack(side="left")

        tk.Label(top2, text="Max nimeä:", bg=self.CARD, fg=self.TEXT, font=("Segoe UI", 10)).pack(side="left", padx=(16, 6))
        self.max_names_var = tk.IntVar(value=400)
        tk.Spinbox(top2, from_=10, to=5000, textvariable=self.max_names_var, width=7,
                   bg="#ffffff", fg=self.TEXT, insertbackground=self.TEXT,
                   highlightthickness=1, highlightbackground=self.BORDER).pack(side="left")

        self._btn(top2, "Tyhjennä", self.clear_clipboard_text, kind="grey").pack(side="right", padx=6)

        tk.Label(clip_card, text="Liitä tähän (Ctrl+V):", bg=self.CARD, fg=self.MUTED,
                 font=("Segoe UI", 10)).pack(anchor="w", padx=12, pady=(6, 6))

        self.clip_text = tk.Text(clip_card, height=9, wrap="word",
                                 bg="#ffffff", fg=self.TEXT, insertbackground=self.TEXT,
                                 highlightthickness=1, highlightbackground=self.BORDER)
        self.clip_text.pack(fill="x", padx=12, pady=(0, 12))

        # --- LOG CARD ---
        log_card = self._card(root)
        log_card.pack(fill="both", expand=True, padx=16, pady=(0, 16))

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

    # ========= KL =========
    def kl_open_browser(self):
        self._clear_stop()
        url = (self.kl_url_var.get() or "").strip() or KL_PROTEST_DEFAULT_URL
        self._set_status("KL: Avataan selain…")
        try:
            if self.kl_driver:
                try:
                    self.kl_driver.quit()
                except Exception:
                    pass
                self.kl_driver = None

            self.kl_driver = start_new_driver()
            try:
                self.kl_driver.get(url)
            except TimeoutException:
                pass
            try:
                wait_loaded(self.kl_driver)
            except Exception:
                pass
            try_accept_cookies(self.kl_driver)

            self._set_status("KL: Selain auki. Kirjaudu itse ja varmista protestilista näkyy. Paina sitten 'Jatka'.")
        except Exception as e:
            self._set_status(f"KL: Virhe selaimen avauksessa: {e}")
            messagebox.showerror("Virhe", str(e))

    def kl_continue_scrape(self):
        self._clear_stop()
        if not self.kl_driver:
            messagebox.showwarning("Ei selainta", "Avaa ensin KL-selain (Kirjaudu).")
            return
        url = (self.kl_url_var.get() or "").strip() or (self.kl_driver.current_url or KL_PROTEST_DEFAULT_URL)
        max_companies = int(self.kl_max_var.get() or 2000)
        threading.Thread(target=self._run_kl, args=(url, max_companies), daemon=True).start()

    def _run_kl(self, url: str, max_companies: int):
        try:
            self._set_status("KL: Aloitetaan keruu…")
            rows, emails = pipeline_kl_protest(url, max_companies, self._set_status, self._set_progress, self.stop_flag, self.kl_driver)
            if self.stop_flag.is_set():
                self._set_status("Pysäytetty.")
                return
            if not rows:
                self._set_status("Ei löytynyt tuloksia.")
                return

            # create output ONLY now
            out_dir = create_final_run_dir()
            self.last_output_dir = out_dir
            save_results_xlsx(out_dir, rows, source_label="KL Protestilista", source_url=url)
            save_results_csv(out_dir, rows)
            save_emails_docx(out_dir, emails)

            messagebox.showinfo("Valmis", f"Valmis!\n\nKansio:\n{out_dir}\n\nRivejä: {len(rows)}\nSähköposteja: {len(emails)}")
            self._set_status("Valmis!")
        except Exception as e:
            self._ui_log(f"VIRHE: {e}")
            messagebox.showerror("Virhe", f"Tuli virhe:\n\n{e}")

    # ========= PDF =========
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
        try:
            self._set_status("PDF: Aloitetaan ajo…")
            rows, emails = pipeline_pdf(pdf_path, self._set_status, self._set_progress, self.stop_flag)
            if self.stop_flag.is_set():
                self._set_status("Pysäytetty.")
                return
            if not rows:
                self._set_status("Ei löytynyt tuloksia.")
                return

            out_dir = create_final_run_dir()
            self.last_output_dir = out_dir
            save_results_xlsx(out_dir, rows, source_label="PDF", source_url=pdf_path)
            save_results_csv(out_dir, rows)
            save_emails_docx(out_dir, emails)

            messagebox.showinfo("Valmis", f"Valmis!\n\nKansio:\n{out_dir}\n\nRivejä: {len(rows)}\nSähköposteja: {len(emails)}")
            self._set_status("Valmis!")
        except Exception as e:
            self._ui_log(f"VIRHE: {e}")
            messagebox.showerror("Virhe", f"Tuli virhe:\n\n{e}")

    # ========= Clipboard =========
    def start_clipboard_mode(self):
        self._clear_stop()
        text = self.clip_text.get("1.0", tk.END).strip()
        if not text:
            messagebox.showwarning("Tyhjä", "Liitä ensin teksti (Ctrl+V) kenttään.")
            return
        threading.Thread(target=self._run_clipboard, args=(text,), daemon=True).start()

    def _run_clipboard(self, text: str):
        try:
            self._set_status("Clipboard: Aloitetaan ajo…")
            strict = bool(self.strict_var.get())
            max_names = int(self.max_names_var.get() or 400)

            rows, emails = pipeline_clipboard(
                text, strict, max_names,
                self._set_status, self._set_progress, self.stop_flag
            )

            if self.stop_flag.is_set():
                self._set_status("Pysäytetty.")
                return
            if not rows:
                self._set_status("Ei löytynyt tuloksia.")
                return

            out_dir = create_final_run_dir()
            self.last_output_dir = out_dir
            save_results_xlsx(out_dir, rows, source_label="Clipboard", source_url="clipboard")
            save_results_csv(out_dir, rows)
            save_emails_docx(out_dir, emails)

            messagebox.showinfo("Valmis", f"Valmis!\n\nKansio:\n{out_dir}\n\nRivejä: {len(rows)}\nSähköposteja: {len(emails)}")
            self._set_status("Valmis!")
        except Exception as e:
            self._ui_log(f"VIRHE: {e}")
            messagebox.showerror("Virhe", f"Tuli virhe:\n\n{e}")


if __name__ == "__main__":
    App().mainloop()
