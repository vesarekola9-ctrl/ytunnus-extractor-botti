import os
import re
import sys
import time
import json
import csv
import random
import threading
import difflib
import traceback
import tkinter as tk
from tkinter import messagebox, filedialog
from tkinter import ttk

import PyPDF2
from docx import Document

from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import TimeoutException, WebDriverException, StaleElementReferenceException
from webdriver_manager.chrome import ChromeDriverManager


# =========================
#   CONFIG / REGEX
# =========================
YT_RE = re.compile(r"\b\d{7}-\d\b|\b\d{8}\b")
EMAIL_RE = re.compile(r"[A-Za-z0-9_.+-]+@[A-Za-z0-9-]+\.[A-Za-z0-9-.]+")
EMAIL_A_RE = re.compile(r"[A-Za-z0-9_.+-]+\s*\(a\)\s*[A-Za-z0-9-]+\.[A-Za-z0-9-.]+", re.I)

YTJ_HOME = "https://tietopalvelu.ytj.fi/"
YTJ_COMPANY_URL = "https://tietopalvelu.ytj.fi/yritys/{}"

BAD_LINES = {
    "yritys", "sijainti", "summa", "häiriöpäivä", "tyyppi", "lähde",
    "viimeisimmät protestit", "protestilista",
    "y-tunnus", "julkaisupäivä", "alue", "velkoja",
}
BAD_CONTAINS = [
    "velkomustuomiot", "ulosotto", "konkurssi", "dun & bradstreet", "bisnode",
    "protestit", "protestia",
]

GENERIC_EMAIL_PREFIXES = (
    "info@", "office@", "support@", "asiakaspalvelu@", "sales@", "myynti@",
    "noreply@", "no-reply@", "donotreply@", "admin@", "contact@", "hello@"
)


# =========================
#   PATHS / LOG
# =========================
def get_exe_dir():
    if getattr(sys, "frozen", False):
        return os.path.dirname(sys.executable)
    return os.path.dirname(os.path.abspath(__file__))


def get_output_dir():
    base = get_exe_dir()
    try:
        p = os.path.join(base, "_write_test.tmp")
        with open(p, "w", encoding="utf-8") as f:
            f.write("ok")
        os.remove(p)
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

CACHE_DIR = os.path.join(os.path.expanduser("~"), "Documents", "ProtestiBotti")
os.makedirs(CACHE_DIR, exist_ok=True)
CACHE_PATH = os.path.join(CACHE_DIR, "ytj_cache.json")


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
    log_to_file(f"Cache: {CACHE_PATH}")
    log_to_file("PDF: sisäinen minipdf-writer (ei reportlabia)")
    log_to_file("YTJ: Näytä-avaukseen shadowDOM + aria-label/title + contains('Näytä')")


# =========================
#   CACHE
# =========================
def _load_cache():
    try:
        if os.path.exists(CACHE_PATH):
            with open(CACHE_PATH, "r", encoding="utf-8") as f:
                data = json.load(f)
                if isinstance(data, dict):
                    data.setdefault("yt", {})
                    data.setdefault("name", {})
                    return data
    except Exception:
        pass
    return {"yt": {}, "name": {}}


def _save_cache(cache):
    try:
        tmp = CACHE_PATH + ".tmp"
        with open(tmp, "w", encoding="utf-8") as f:
            json.dump(cache, f, ensure_ascii=False, indent=2)
        os.replace(tmp, CACHE_PATH)
    except Exception:
        pass


def cache_clear():
    try:
        if os.path.exists(CACHE_PATH):
            os.remove(CACHE_PATH)
    except Exception:
        pass


def cache_get_by_yt(cache, yt: str):
    return cache.get("yt", {}).get((yt or "").strip())


def cache_put_yt(cache, yt: str, email: str, name: str = "", status: str = "DONE"):
    yt = (yt or "").strip()
    if not yt:
        return
    cache.setdefault("yt", {})
    cache["yt"][yt] = {
        "email": (email or "").strip(),
        "name": (name or "").strip(),
        "status": status,
        "ts": int(time.time()),
    }


def normalize_company_name(name: str) -> str:
    s = (name or "").strip().casefold()
    s = re.sub(r"\s+", " ", s)
    s = re.sub(r"[.,;:()\[\]{}\"'´`]", "", s)
    s = s.replace("&", " and ")
    s = re.sub(r"\s+", " ", s).strip()
    s = re.sub(r"\b(osakeyhtiö|osakeyhtio)\b", "oy", s)
    return s


def cache_get_by_name(cache, name: str):
    return cache.get("name", {}).get(normalize_company_name(name))


def cache_put_name(cache, name: str, email: str, yt: str = "", status: str = "DONE"):
    key = normalize_company_name(name)
    if not key:
        return
    cache.setdefault("name", {})
    cache["name"][key] = {
        "email": (email or "").strip(),
        "yt": (yt or "").strip(),
        "status": status,
        "ts": int(time.time()),
        "raw": (name or "").strip(),
    }


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


def classify_email(email: str) -> str:
    e = (email or "").strip().lower()
    if not e:
        return ""
    if e.startswith(GENERIC_EMAIL_PREFIXES):
        return "GENEERINEN"
    return "OK"


def similarity(a: str, b: str) -> float:
    a = normalize_company_name(a)
    b = normalize_company_name(b)
    if not a or not b:
        return 0.0
    return difflib.SequenceMatcher(None, a, b).ratio()


def save_word_plain_lines(lines, filename):
    path = os.path.join(OUT_DIR, filename)
    doc = Document()
    for line in lines:
        if line and str(line).strip():
            doc.add_paragraph(str(line).strip())
    doc.save(path)
    return path


def save_txt_lines(lines, filename):
    path = os.path.join(OUT_DIR, filename)
    with open(path, "w", encoding="utf-8") as f:
        for line in lines:
            if line and str(line).strip():
                f.write(str(line).strip() + "\n")
    return path


def save_csv_rows(rows, headers, filename):
    path = os.path.join(OUT_DIR, filename)
    with open(path, "w", encoding="utf-8", newline="") as f:
        w = csv.writer(f, delimiter=";")
        w.writerow(headers)
        for r in rows:
            w.writerow(r)
    return path


# =========================
#   PDF WITHOUT REPORTLAB (minipdf)
# =========================
def _pdf_escape_text(s: str) -> str:
    s = s.replace("\\", "\\\\").replace("(", "\\(").replace(")", "\\)")
    s = s.replace("\t", "    ")
    return s


def save_pdf_lines(lines, filename, title=None):
    path = os.path.join(OUT_DIR, filename)
    PAGE_W, PAGE_H = 595.28, 841.89
    LEFT = 50
    TOP = PAGE_H - 60
    BOTTOM = 60
    FONT_SIZE = 11
    LEADING = 14

    out_lines = []
    if title:
        out_lines.append(title)
        out_lines.append("")
    for ln in lines:
        if ln is None:
            continue
        t = str(ln).strip()
        if t:
            out_lines.append(t)

    max_lines_per_page = int((TOP - BOTTOM) // LEADING)
    pages, cur = [], []
    for ln in out_lines:
        cur.append(ln)
        if len(cur) >= max_lines_per_page:
            pages.append(cur)
            cur = []
    if cur:
        pages.append(cur)
    if not pages:
        pages = [[""]]

    objects = []

    def add_obj(data_bytes: bytes) -> int:
        objects.append(data_bytes)
        return len(objects)

    font_obj = add_obj(b"<< /Type /Font /Subtype /Type1 /BaseFont /Helvetica /Encoding /WinAnsiEncoding >>")
    page_objs = []
    pages_obj_id = add_obj(b"<< /Type /Pages /Kids [] /Count 0 >>")

    for page_lines in pages:
        y = TOP
        chunks = ["BT\n", f"/F1 {FONT_SIZE} Tf\n", f"{LEFT} {y:.2f} Td\n"]
        first = True
        for ln in page_lines:
            txt = _pdf_escape_text(ln)
            if not first:
                chunks.append(f"0 -{LEADING} Td\n")
            first = False
            chunks.append(f"({txt}) Tj\n")
        chunks.append("ET\n")
        content_str = "".join(chunks)
        content_bytes = content_str.encode("latin-1", errors="replace")
        stream = b"<< /Length %d >>\nstream\n%s\nendstream" % (len(content_bytes), content_bytes)
        content_obj = add_obj(stream)

        page_dict = (
            b"<< /Type /Page /Parent %d 0 R "
            b"/MediaBox [0 0 %.2f %.2f] "
            b"/Resources << /Font << /F1 %d 0 R >> >> "
            b"/Contents %d 0 R >>"
            % (pages_obj_id, PAGE_W, PAGE_H, font_obj, content_obj)
        )
        page_obj_id = add_obj(page_dict)
        page_objs.append(page_obj_id)

    kids = " ".join([f"{pid} 0 R" for pid in page_objs]).encode("ascii")
    objects[pages_obj_id - 1] = b"<< /Type /Pages /Kids [%s] /Count %d >>" % (kids, len(page_objs))
    catalog_obj_id = add_obj(b"<< /Type /Catalog /Pages %d 0 R >>" % pages_obj_id)

    header = b"%PDF-1.4\n%\xe2\xe3\xcf\xd3\n"
    offsets = [0]
    body = b""
    for i, obj in enumerate(objects, start=1):
        offsets.append(len(header) + len(body))
        body += f"{i} 0 obj\n".encode("ascii") + obj + b"\nendobj\n"

    xref_start = len(header) + len(body)
    xref = [b"xref\n", f"0 {len(objects)+1}\n".encode("ascii")]
    xref.append
