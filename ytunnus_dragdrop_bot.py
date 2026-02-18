import os
import re
import sys
import time
import subprocess
import threading
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
from webdriver_manager.chrome import ChromeDriverManager


# =========================
#   CONFIG / REGEX
# =========================
KL_URL = "https://www.kauppalehti.fi/yritykset/protestilista"

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
    log_to_file("PDF: sisäinen minipdf-writer (ei reportlabia)")


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


# =========================
#   PDF WITHOUT REPORTLAB (pure python)
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
    pages = []
    cur = []
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

    font_obj = add_obj(
        b"<< /Type /Font /Subtype /Type1 /BaseFont /Helvetica /Encoding /WinAnsiEncoding >>"
    )

    page_objs = []
    pages_obj_id = add_obj(b"<< /Type /Pages /Kids [] /Count 0 >>")

    for page_lines in pages:
        y = TOP
        chunks = []
        chunks.append("BT\n")
        chunks.append(f"/F1 {FONT_SIZE} Tf\n")
        chunks.append(f"{LEFT} {y:.2f} Td\n")

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
    pages_dict = b"<< /Type /Pages /Kids [%s] /Count %d >>" % (kids, len(page_objs))
    objects[pages_obj_id - 1] = pages_dict

    catalog_obj_id = add_obj(b"<< /Type /Catalog /Pages %d 0 R >>" % pages_obj_id)

    header = b"%PDF-1.4\n%\xe2\xe3\xcf\xd3\n"
    offsets = [0]
    body = b""
    for i, obj in enumerate(objects, start=1):
        offsets.append(len(header) + len(body))
        body += f"{i} 0 obj\n".encode("ascii") + obj + b"\nendobj\n"

    xref_start = len(header) + len(body)
    xref = [b"xref\n", f"0 {len(objects)+1}\n".encode("ascii")]
    xref.append(b"0000000000 65535 f \n")
    for off in offsets[1:]:
        xref.append(f"{off:010d} 00000 n \n".encode("ascii"))
    xref_bytes = b"".join(xref)

    trailer = (
        b"trailer\n"
        b"<< /Size %d /Root %d 0 R >>\n"
        b"startxref\n%d\n%%%%EOF\n"
        % (len(objects) + 1, catalog_obj_id, xref_start)
    )

    pdf_bytes = header + body + xref_bytes + trailer
    with open(path, "wb") as f:
        f.write(pdf_bytes)
    return path


# =========================
#   SMART PARSER (names + yts)
# =========================
def _is_money_line(s: str) -> bool:
    return bool(re.fullmatch(r"\d{1,3}(\s?\d{3})*\s?€", s.strip()))


def _is_date_line(s: str) -> bool:
    return bool(re.fullmatch(r"\d{1,2}\.\d{1,2}\.\d{4}", s.strip()))


def _looks_like_location(s: str) -> bool:
    s = s.strip()
    if not s or any(ch.isdigit() for ch in s):
        return False
    parts = s.split()
    if len(parts) != 1:
        return False
    if not parts[0][:1].isupper():
        return False
    low = s.lower()
    if any(x in low for x in [" oy", " ab", " ky", " ay", " ry", " tmi", " oyj", " lkv", " osakeyhtiö"]):
        return False
    return True


def _looks_like_company_name(s: str) -> bool:
    s = s.strip()
    if len(s) < 3:
        return False
    if not re.search(r"[A-Za-zÅÄÖåäö]", s):
        return False

    low = s.lower()
    if low in BAD_LINES:
        return False
    if any(b in low for b in BAD_CONTAINS):
        return False
    if _is_money_line(s) or _is_date_line(s):
        return False
    if "y-tunnus" in low or YT_RE.search(s):
        return False

    suffix_ok = any(x in low for x in [" oy", " ab", " ky", " ay", " ry", " tmi", " oyj", " ltd", " gmbh", " inc", " lkv"])
    many_words = len(s.split()) >= 2

    if len(s.split()) == 1 and _looks_like_location(s):
        return False

    return suffix_ok or many_words


def _normalize_name(s: str) -> str:
    s = re.sub(r"\s+", " ", s.strip())
    s = s.strip(" -•\u2022")
    return s


def parse_names_and_yts_from_copied_text(raw: str):
    if not raw:
        return [], []

    yts = []
    seen_yts = set()
    for m in YT_RE.findall(raw):
        n = normalize_yt(m)
        if n and n not in seen_yts:
            seen_yts.add(n)
            yts.append(n)

    lines = [ln.strip() for ln in raw.splitlines() if ln and ln.strip()]
    lines = [ln for ln in lines if len(ln) >= 2]

    names = []
    seen = set()

    money_idxs = [i for i, ln in enumerate(lines) if _is_money_line(ln)]
    for i in money_idxs:
        candidates = []
        for back in (1, 2, 3):
            j = i - back
            if j >= 0:
                candidates.append(lines[j])

        same_line = lines[i]
        m = re.search(r"^(.*)\s+\d{1,3}(\s?\d{3})*\s?€$", same_line)
        if m:
            candidates.insert(0, m.group(1).strip())

        best = ""
        for c in candidates:
            c = _normalize_name(c)
            if _looks_like_company_name(c):
                best = c
                break

        if best:
            key = best.lower()
            if key not in seen:
                seen.add(key)
                names.append(best)

    if not names:
        for ln in lines:
            ln = _normalize_name(ln)
            if _looks_like_company_name(ln):
                key = ln.lower()
                if key not in seen:
                    seen.add(key)
                    names.append(ln)

    return names, yts


# =========================
#   PDF -> (names + yts)
# =========================
def extract_names_and_yts_from_pdf(pdf_path: str):
    text_all = ""
    with open(pdf_path, "rb") as f:
        reader = PyPDF2.PdfReader(f)
        for page in reader.pages:
            text_all += (page.extract_text() or "") + "\n"
    return parse_names_and_yts_from_copied_text(text_all)


# =========================
#   YTJ (selenium)
# =========================
def start_new_driver():
    options = webdriver.ChromeOptions()
    options.add_argument("--start-maximized")
    options.add_argument("--disable-gpu")
    options.add_argument("--disable-dev-shm-usage")
    driver_path = ChromeDriverManager().install()
    return webdriver.Chrome(service=Service(driver_path), options=options)


def wait_body(driver, timeout=25):
    WebDriverWait(driver, timeout).until(EC.presence_of_element_located((By.TAG_NAME, "body")))


def ytj_open_company_by_name(driver, name: str, status_cb):
    status_cb(f"YTJ: haku nimellä: {name}")
    driver.get(YTJ_HOME)
    wait_body(driver, 25)
    time.sleep(0.4)

    candidates = []
    for css in ["input[type='search']", "input[type='text']"]:
        try:
            candidates += driver.find_elements(By.CSS_SELECTOR, css)
        except Exception:
            pass

    search_input = None
    for inp in candidates:
        try:
            if inp.is_displayed() and inp.is_enabled():
                search_input = inp
                ph = (inp.get_attribute("placeholder") or "").lower()
                if "hae" in ph or "y-tunnus" in ph or "toiminimi" in ph:
                    break
        except Exception:
            continue

    if not search_input:
        return False

    try:
        search_input.clear()
    except Exception:
        pass
    search_input.send_keys(name)
    search_input.send_keys(Keys.ENTER)
    time.sleep(1.2)

    if "/yritys/" in (driver.current_url or ""):
        return True

    try:
        links = driver.find_elements(By.XPATH, "//a[contains(@href,'/yritys/')]")
    except Exception:
        links = []

    for a in links:
        try:
            if a.is_displayed():
                a.click()
                time.sleep(0.9)
                if "/yritys/" in (driver.current_url or ""):
                    return True
        except Exception:
            continue
    return False


def extract_email_from_ytj_page(driver):
    try:
        for a in driver.find_elements(By.TAG_NAME, "a"):
            href = (a.get_attribute("href") or "")
            if href.lower().startswith("mailto:"):
                return href.split(":", 1)[1].strip()
    except Exception:
        pass

    try:
        body = driver.find_element(By.TAG_NAME, "body").text or ""
        return pick_email_from_text(body)
    except Exception:
        return ""


def fetch_emails_from_ytj_by_yts_and_names(driver, yts, names, status_cb, progress_cb, log_cb, stop_evt):
    emails = []
    seen = set()
    total = len(yts) + len(names)
    progress_cb(0, max(1, total))
    done = 0

    for yt in yts:
        if stop_evt.is_set():
            break
        done += 1
        status_cb(f"YTJ (Y-tunnus): {yt} ({done}/{total})")
        progress_cb(done, total)

        driver.get(YTJ_COMPANY_URL.format(yt))
        wait_body(driver, 25)
        time.sleep(0.6)

        email = ""
        for _ in range(10):
            email = extract_email_from_ytj_page(driver)
            if email:
                break
            time.sleep(0.2)

        if email:
            k = email.lower()
            if k not in seen:
                seen.add(k)
                emails.append(email)
                log_cb(email)

    for nm in names:
        if stop_evt.is_set():
            break
        done += 1
        progress_cb(done, total)

        ok = ytj_open_company_by_name(driver, nm, status_cb)
        if not ok:
            log_cb(f"YTJ: ei osumaa nimelle: {nm}")
            continue

        email = ""
        for _ in range(10):
            email = extract_email_from_ytj_page(driver)
            if email:
                break
            time.sleep(0.2)

        if email:
            k = email.lower()
            if k not in seen:
                seen.add(k)
                emails.append(email)
                log_cb(email)

    return emails


# =========================
#   KL "OPEN ALL ARROWS" TOOL (embedded)
# =========================
def safe_click(driver, elem):
    try:
        driver.execute_script("arguments[0].scrollIntoView({block:'center'});", elem)
        time.sleep(0.05)
        try:
            elem.click()
        except Exception:
            driver.execute_script("arguments[0].click();", elem)
        return True
    except Exception:
        return False


def try_accept_cookies(driver):
    texts = ["Hyväksy", "Hyväksy kaikki", "Salli kaikki", "Accept", "Accept all", "I agree", "OK", "Selvä"]
    try:
        buttons = driver.find_elements(By.XPATH, "//button|//a|//*[@role='button']")
    except Exception:
        return
    for b in buttons:
        try:
            t = (b.text or "").strip()
            if not t:
                continue
            low = t.lower()
            if any(x.lower() in low for x in texts):
                if b.is_displayed() and b.is_enabled():
                    safe_click(driver, b)
                    time.sleep(0.25)
        except Exception:
            continue


def click_all_arrows_once(driver):
    elems = driver.find_elements(By.CSS_SELECTOR, "[aria-expanded='false']")
    count = 0
    for e in elems:
        try:
            if e.is_displayed() and e.is_enabled():
                if safe_click(driver, e):
                    count += 1
                    time.sleep(0.01)
        except Exception:
            continue
    return count


def click_show_more(driver):
    try:
        buttons = driver.find_elements(By.XPATH, "//button|//*[@role='button']")
    except Exception:
        return False
    for b in buttons:
        try:
            if not b.is_displayed() or not b.is_enabled():
                continue
            if (b.text or "").strip().lower() == "näytä lisää":
                safe_click(driver, b)
                return True
        except Exception:
            continue
    return False


def scroll_to_bottom(driver):
    try:
        driver.execute_script("window.scrollTo(0, document.body.scrollHeight);")
    except Exception:
        pass


def find_chrome_path():
    candidates = [
        r"C:\Program Files\Google\Chrome\Application\chrome.exe",
        r"C:\Program Files (x86)\Google\Chrome\Application\chrome.exe",
    ]
    for c in candidates:
        if os.path.exists(c):
            return c
    return None


# =========================
#   GUI APP
# =========================
class App(tk.Tk):
    def __init__(self):
        super().__init__()
        reset_log()

        self.stop_evt = threading.Event()

        # KL tool state
        self.kl_proc = None
        self.kl_driver = None

        self.title("ProtestiBotti (All-in-one: KL tool + KL copy/paste + PDF→YTJ)")
        self.geometry("1120x900")

        tk.Label(self, text="ProtestiBotti", font=("Arial", 18, "bold")).pack(pady=8)
        tk.Label(
            self,
            text="Sisältää:\n"
                 "• KL-työkalu: Avaa kaikki nuolet + Näytä lisää (ohjattu Chrome)\n"
                 "• KL copy/paste → yritysnimet + Y-tunnukset → PDF/DOCX/TXT\n"
                 "• PDF → YTJ (Y-tunnus ensin, muuten nimi) → sähköpostit Wordiin\n",
            justify="center"
        ).pack(pady=4)

        # ---------- KL TOOL BOX ----------
        tool = tk.LabelFrame(self, text="KL-työkalu: Avaa kaikki nuolet", padx=10, pady=10)
        tool.pack(fill="x", padx=12, pady=10)

        tk.Label(
            tool,
            text="Tämä käynnistää oman ohjatun Chromen ja avaa protestilistalla kaikki siniset nuolet + 'Näytä lisää'.\n"
                 "Vinkki: Sulje normaali Chrome ennen kuin painat Käynnistä (ettei Chrome-profiili lukitu).",
            justify="left"
        ).pack(anchor="w")

        tool_row = tk.Frame(tool)
        tool_row.pack(fill="x", pady=6)

        tk.Button(tool_row, text="1) Käynnistä ohjattu Chrome", font=("Arial", 11, "bold"),
                  command=self.kl_start_guided_chrome).pack(side="left", padx=6)
        tk.Button(tool_row, text="2) Avaa protestilista", font=("Arial", 11, "bold"),
                  command=self.kl_open_page).pack(side="left", padx=6)
        tk.Button(tool_row, text="3) Avaa kaikki nuolet", font=("Arial", 11, "bold"),
                  command=self.kl_open_all_arrows).pack(side="left", padx=6)
        tk.Button(tool_row, text="Sulje KL Chrome", command=self.kl_shutdown).pack(side="left", padx=6)

        # ---------- KL COPY/PASTE BOX ----------
        box = tk.LabelFrame(self, text="Kauppalehti: copy/paste → tee tiedostot", padx=10, pady=10)
        box.pack(fill="x", padx=12, pady=10)

        tk.Label(
            box,
            text="Kun nuolet on auki, tee protestilistassa Ctrl+A → Ctrl+C ja liitä alle → 'Tee tiedostot'.",
            justify="left"
        ).pack(anchor="w")

        self.text = tk.Text(box, height=10)
        self.text.pack(fill="x", pady=6)

        row = tk.Frame(box)
        row.pack(fill="x")
        tk.Button(row, text="Tee tiedostot (PDF + DOCX + TXT)", font=("Arial", 12, "bold"),
                  command=self.make_files).pack(side="left")
        tk.Button(row, text="Tyhjennä", command=lambda: self.text.delete("1.0", tk.END)).pack(side="left", padx=8)

        # ---------- PDF → YTJ BOX ----------
        box2 = tk.LabelFrame(self, text="PDF → YTJ sähköpostit", padx=10, pady=10)
        box2.pack(fill="x", padx=12, pady=10)

        tk.Button(box2, text="Valitse PDF ja hae sähköpostit YTJ:stä", font=("Arial", 12, "bold"),
                  command=self.start_pdf_to_ytj).pack(anchor="w")

        # ---------- STATUS / LOG ----------
        self.status = tk.Label(self, text="Valmiina.", font=("Arial", 11))
        self.status.pack(pady=6)

        self.progress = ttk.Progressbar(self, orient="horizontal", mode="determinate", length=1060)
        self.progress.pack(pady=6)

        frame = tk.Frame(self)
        frame.pack(fill="both", expand=True, padx=14, pady=10)

        tk.Label(frame, text="Live-logi (uusimmat alimmaisena):").pack(anchor="w")
        self.listbox = tk.Listbox(frame, height=18)
        self.listbox.pack(side="left", fill="both", expand=True)

        sb = tk.Scrollbar(frame, orient="vertical", command=self.listbox.yview)
        sb.pack(side="right", fill="y")
        self.listbox.configure(yscrollcommand=sb.set)

        tk.Label(self, text=f"Tallennus: {OUT_DIR}", wraplength=1080, justify="center").pack(pady=6)

        self.protocol("WM_DELETE_WINDOW", self.on_close)

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
        self.progress["maximum"] = max(1, maximum)
        self.progress["value"] = value
        self.update_idletasks()

    # =========================
    #   KL TOOL actions
    # =========================
    def kl_attach_driver(self):
        if self.kl_driver:
            return True
        try:
            options = webdriver.ChromeOptions()
            options.add_experimental_option("debuggerAddress", "127.0.0.1:9222")
            driver_path = ChromeDriverManager().install()
            self.kl_driver = webdriver.Chrome(service=Service(driver_path), options=options)
            return True
        except Exception as e:
            self.ui_log(f"KL driver attach epäonnistui: {e}")
            return False

    def kl_start_guided_chrome(self):
        if self.kl_proc:
            self.set_status("KL Chrome on jo käynnissä.")
            return

        chrome_path = find_chrome_path()
        if not chrome_path:
            messagebox.showerror("Chrome puuttuu", "En löytänyt chrome.exe oletuspoluista. Asenna Google Chrome.")
            return

        base = get_exe_dir()
        profile_dir = os.path.join(base, "KLToolProfile")
        os.makedirs(profile_dir, exist_ok=True)

        self.set_status("Käynnistän KL-ohjatun Chromen (9222)…")
        args = [
            chrome_path,
            "--remote-debugging-port=9222",
            f"--user-data-dir={profile_dir}",
            "--no-first-run",
            "--no-default-browser-check",
            KL_URL
        ]
        try:
            self.kl_proc = subprocess.Popen(args, stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)
        except Exception as e:
            self.kl_proc = None
            messagebox.showerror("Ei käynnisty", f"Chrome ei käynnistynyt:\n{e}")
            return

        time.sleep(1.5)
        if not self.kl_attach_driver():
            messagebox.showwarning(
                "Liitäntä ei onnistunut",
                "Chrome käynnistyi, mutta Selenium ei saanut yhteyttä.\n"
                "Varmista että normaali Chrome on suljettu ja yritä uudelleen."
            )
            return

        self.set_status("KL Chrome käynnissä. Kirjaudu Kauppalehteen Chromessa jos pyytää.")

    def kl_open_page(self):
        if not self.kl_attach_driver():
            messagebox.showwarning("Ei yhteyttä", "Käynnistä KL Chrome ensin.")
            return
        try:
            self.kl_driver.get(KL_URL)
            wait_body(self.kl_driver, 25)
            try_accept_cookies(self.kl_driver)
            self.set_status("Protestilista avattu. Kirjaudu jos pyytää.")
        except Exception as e:
            self.ui_log(f"KL avaus epäonnistui: {e}")

    def kl_open_all_arrows(self):
        if not self.kl_attach_driver():
            messagebox.showwarning("Ei yhteyttä", "Käynnistä KL Chrome ensin.")
            return

        def worker():
            try:
                self.set_status("KL: avataan nuolet + Näytä lisää…")
                try_accept_cookies(self.kl_driver)

                rounds = 0
                total_clicked = 0

                while rounds < 80:
                    rounds += 1
                    try_accept_cookies(self.kl_driver)

                    clicked = click_all_arrows_once(self.kl_driver)
                    total_clicked += clicked

                    scroll_to_bottom(self.kl_driver)
                    time.sleep(0.6)

                    more = click_show_more(self.kl_driver)
                    if more:
                        self.ui_log(f"KL kierros {rounds}: nuolia {clicked}, Näytä lisää: kyllä")
                        time.sleep(1.2)
                        continue

                    self.ui_log(f"KL kierros {rounds}: nuolia {clicked}, Näytä lisää: ei")
                    if clicked == 0:
                        break

                self.set_status(f"KL valmis! Avattu nuolia ~{total_clicked}. Tee nyt Ctrl+A → Ctrl+C ja liitä bottiin.")
                messagebox.showinfo("Valmis", "Kaikki mahdolliset nuolet avattu.\n\nTee nyt Ctrl+A → Ctrl+C ja liitä bottiin.")
            except Exception as e:
                self.ui_log(f"KL virhe: {e}")
                messagebox.showerror("Virhe", str(e))

        threading.Thread(target=worker, daemon=True).start()

    def kl_shutdown(self):
        try:
            if self.kl_driver:
                try:
                    self.kl_driver.quit()
                except Exception:
                    pass
                self.kl_driver = None
        finally:
            if self.kl_proc:
                try:
                    self.kl_proc.terminate()
                except Exception:
                    pass
                self.kl_proc = None
        self.set_status("KL Chrome suljettu.")

    # =========================
    #   KL copy/paste -> files
    # =========================
    def make_files(self):
        raw = self.text.get("1.0", tk.END)
        names, yts = parse_names_and_yts_from_copied_text(raw)

        if not names and not yts:
            self.set_status("Ei löytynyt yritysnimiä tai Y-tunnuksia.")
            messagebox.showwarning("Ei löytynyt", "Liitä protestilistan sisältö (Ctrl+A/Ctrl+C) ja yritä uudestaan.")
            return

        self.set_status(f"Poimittu: nimet={len(names)} | y-tunnukset={len(yts)}")

        if names:
            p1 = save_word_plain_lines(names, "yritysnimet_kauppalehti.docx")
            self.ui_log(f"Tallennettu: {p1}")
            t1 = save_txt_lines(names, "yritysnimet_kauppalehti.txt")
            self.ui_log(f"Tallennettu: {t1}")
            pdf1 = save_pdf_lines(names, "yritysnimet_kauppalehti.pdf", title="Yritysnimet (Kauppalehti)")
            self.ui_log(f"Tallennettu: {pdf1}")

        if yts:
            p2 = save_word_plain_lines(yts, "ytunnukset_kauppalehti.docx")
            self.ui_log(f"Tallennettu: {p2}")
            t2 = save_txt_lines(yts, "ytunnukset_kauppalehti.txt")
            self.ui_log(f"Tallennettu: {t2}")

        combined = []
        if names:
            combined.append("YRITYSNIMET")
            combined += names
            combined.append("")
        if yts:
            combined.append("Y-TUNNUKSET")
            combined += yts

        pdf2 = save_pdf_lines(combined, "yritysnimet_ja_ytunnukset.pdf", title="Kauppalehti -> poimitut tiedot")
        self.ui_log(f"Tallennettu: {pdf2}")

        self.set_status("Valmis (tiedostot luotu).")
        messagebox.showinfo(
            "Valmis",
            f"Valmis!\n\nYritysnimiä: {len(names)}\nY-tunnuksia: {len(yts)}\n\nKansio:\n{OUT_DIR}"
        )

    # =========================
    #   PDF -> YTJ
    # =========================
    def start_pdf_to_ytj(self):
        pdf_path = filedialog.askopenfilename(filetypes=[("PDF files", "*.pdf")])
        if not pdf_path:
            return
        threading.Thread(target=self.run_pdf_to_ytj, args=(pdf_path,), daemon=True).start()

    def run_pdf_to_ytj(self, pdf_path):
        driver = None
        try:
            self.set_status("Luetaan PDF: yritysnimet + Y-tunnukset…")
            names, yts = extract_names_and_yts_from_pdf(pdf_path)

            if not names and not yts:
                self.set_status("PDF:stä ei löytynyt nimiä tai Y-tunnuksia.")
                messagebox.showwarning("Ei löytynyt", "PDF:stä ei löytynyt nimiä tai Y-tunnuksia.")
                return

            if names:
                p1 = save_word_plain_lines(names, "pdf_poimitut_yritysnimet.docx")
                self.ui_log(f"Tallennettu: {p1}")
            if yts:
                p2 = save_word_plain_lines(yts, "pdf_poimitut_ytunnukset.docx")
                self.ui_log(f"Tallennettu: {p2}")

            self.set_status("Käynnistetään Chrome ja haetaan sähköpostit YTJ:stä…")
            driver = start_new_driver()

            emails = fetch_emails_from_ytj_by_yts_and_names(
                driver,
                yts=yts,
                names=names,
                status_cb=self.set_status,
                progress_cb=self.set_progress,
                log_cb=self.ui_log,
                stop_evt=self.stop_evt
            )

            em_path = save_word_plain_lines(emails, "sahkopostit_ytj.docx")
            self.ui_log(f"Tallennettu: {em_path}")

            self.set_status("Valmis!")
            messagebox.showinfo(
                "Valmis",
                f"Valmis!\n\nNimiä PDF:stä: {len(names)}\nY-tunnuksia PDF:stä: {len(yts)}\nSähköposteja löytyi: {len(emails)}\n\nKansio:\n{OUT_DIR}"
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

    def on_close(self):
        # sulje KL Chrome jos auki
        try:
            self.kl_shutdown()
        except Exception:
            pass
        self.destroy()


if __name__ == "__main__":
    App().mainloop()
