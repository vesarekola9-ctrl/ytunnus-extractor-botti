import os
import re
import sys
import time
import csv
import json
import random
import threading
import difflib
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
#   REGEX
# =========================
YT_RE = re.compile(r"\b\d{7}-\d\b|\b\d{8}\b")
EMAIL_RE = re.compile(r"[A-Za-z0-9_.+-]+@[A-Za-z0-9-]+\.[A-Za-z0-9-.]+")
EMAIL_A_RE = re.compile(r"[A-Za-z0-9_.+-]+\s*\(a\)\s*[A-Za-z0-9-]+\.[A-Za-z0-9-.]+", re.I)

YTJ_HOME = "https://tietopalvelu.ytj.fi/"
YTJ_COMPANY_URL = "https://tietopalvelu.ytj.fi/yritys/{}"


# =========================
#   OUTPUT / LOG
# =========================
def get_exe_dir():
    if getattr(sys, "frozen", False):
        return os.path.dirname(sys.executable)
    return os.path.dirname(os.path.abspath(__file__))


def get_output_dir():
    base = get_exe_dir()
    try:
        test = os.path.join(base, "_write_test.tmp")
        with open(test, "w", encoding="utf-8") as f:
            f.write("ok")
        os.remove(test)
    except Exception:
        base = os.path.join(os.path.expanduser("~"), "Documents", "ProtestiBotti")

    out = os.path.join(base, time.strftime("%Y-%m-%d"))
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
    log_to_file(f"Log: {LOG_PATH}")
    log_to_file(f"Cache: {CACHE_PATH}")


# =========================
#   CACHE
# =========================
def cache_load():
    try:
        if os.path.exists(CACHE_PATH):
            with open(CACHE_PATH, "r", encoding="utf-8") as f:
                d = json.load(f)
                if isinstance(d, dict):
                    d.setdefault("yt", {})
                    d.setdefault("name", {})
                    return d
    except Exception:
        pass
    return {"yt": {}, "name": {}}


def cache_save(d):
    try:
        tmp = CACHE_PATH + ".tmp"
        with open(tmp, "w", encoding="utf-8") as f:
            json.dump(d, f, ensure_ascii=False, indent=2)
        os.replace(tmp, CACHE_PATH)
    except Exception:
        pass


def cache_clear():
    try:
        if os.path.exists(CACHE_PATH):
            os.remove(CACHE_PATH)
    except Exception:
        pass


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


def normalize_company_name(name: str) -> str:
    s = (name or "").strip().casefold()
    s = re.sub(r"\s+", " ", s)
    s = re.sub(r"[.,;:()\[\]{}\"'´`]", "", s)
    s = s.replace("&", " and ")
    s = re.sub(r"\s+", " ", s).strip()
    s = re.sub(r"\b(osakeyhtiö|osakeyhtio)\b", "oy", s)
    return s


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
        t = (str(line).strip() if line is not None else "")
        if t:
            doc.add_paragraph(t)
    doc.save(path)
    return path


def save_txt_lines(lines, filename):
    path = os.path.join(OUT_DIR, filename)
    with open(path, "w", encoding="utf-8") as f:
        for line in lines:
            t = (str(line).strip() if line is not None else "")
            if t:
                f.write(t + "\n")
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
#   PARSERS
# =========================
BAD_LINES = {
    "yritys", "sijainti", "summa", "häiriöpäivä", "tyyppi", "lähde",
    "viimeisimmät protestit", "protestilista",
    "y-tunnus", "julkaisupäivä", "alue", "velkoja",
}
BAD_CONTAINS = [
    "velkomustuomiot", "ulosotto", "konkurssi", "dun & bradstreet", "bisnode",
    "protestit", "protestia",
]


def parse_yts_only(raw: str):
    raw = raw or ""
    seen = set()
    out = []

    for m in YT_RE.findall(raw):
        n = normalize_yt(m)
        if n and n not in seen:
            seen.add(n)
            out.append(n)

    tokens = re.split(r"[,\s;]+", raw.strip())
    for t in tokens:
        n = normalize_yt(t)
        if n and n not in seen:
            seen.add(n)
            out.append(n)

    return out


def _is_money_line(s: str) -> bool:
    return bool(re.fullmatch(r"\d{1,3}(\s?\d{3})*\s?€", s.strip()))


def _is_date_line(s: str) -> bool:
    return bool(re.fullmatch(r"\d{1,2}\.\d{1,2}\.\d{4}", s.strip()))


def _looks_like_company_name(s: str) -> bool:
    s = (s or "").strip()
    if len(s) < 3:
        return False
    if not re.search(r"[A-Za-zÅÄÖåäö]", s):
        return False
    low = s.casefold()
    if low in BAD_LINES:
        return False
    if any(b in low for b in BAD_CONTAINS):
        return False
    if _is_money_line(s) or _is_date_line(s):
        return False
    if "y-tunnus" in low or YT_RE.search(s):
        return False
    # KL-yrityksissä usein Oy/Ab/Ry jne, mutta ei aina (esim. yhdistys)
    suffix_ok = any(x in low for x in [" oy", " ab", " ky", " ay", " ry", " tmi", " oyj", " lkv"])
    many_words = len(s.split()) >= 2
    return suffix_ok or many_words


def parse_names_and_yts_from_copied_text(raw: str):
    if not raw:
        return [], []

    yts = parse_yts_only(raw)

    lines = [ln.strip() for ln in raw.splitlines() if ln and ln.strip()]
    names = []
    seen = set()

    # “raha-rivi” heuristiikka: yritysnimi yleensä 1–3 riviä ennen €-riviä
    for i, ln in enumerate(lines):
        if not _is_money_line(ln):
            continue
        for back in (1, 2, 3):
            j = i - back
            if j >= 0:
                cand = re.sub(r"\s+", " ", lines[j]).strip(" -•\u2022")
                if _looks_like_company_name(cand):
                    key = cand.casefold()
                    if key not in seen:
                        seen.add(key)
                        names.append(cand)
                    break

    # fallback: kaikki yritys-näköiset rivit
    if not names:
        for ln in lines:
            cand = re.sub(r"\s+", " ", ln).strip(" -•\u2022")
            if _looks_like_company_name(cand):
                key = cand.casefold()
                if key not in seen:
                    seen.add(key)
                    names.append(cand)

    return names, yts


def extract_names_and_yts_from_pdf(pdf_path: str):
    text_all = ""
    with open(pdf_path, "rb") as f:
        reader = PyPDF2.PdfReader(f)
        for p in reader.pages:
            text_all += (p.extract_text() or "") + "\n"
    return parse_names_and_yts_from_copied_text(text_all)


# =========================
#   SELENIUM DRIVER
# =========================
def start_new_driver():
    # TÄRKEÄ: EI mitään selenium.webdriver.chrome.options importteja.
    options = webdriver.ChromeOptions()
    options.add_argument("--start-maximized")
    options.add_argument("--disable-gpu")
    options.add_argument("--disable-dev-shm-usage")
    # helpottaa “turvallinen selain” -valituksia joissain login flow’ssa
    options.add_argument("--disable-blink-features=AutomationControlled")
    options.add_experimental_option("excludeSwitches", ["enable-automation"])
    options.add_experimental_option("useAutomationExtension", False)

    driver_path = ChromeDriverManager().install()
    driver = webdriver.Chrome(service=Service(driver_path), options=options)

    # pieni stealth JS (ei rikota sivuja, mutta vähentää “automation detected”)
    try:
        driver.execute_cdp_cmd("Page.addScriptToEvaluateOnNewDocument", {
            "source": """
                Object.defineProperty(navigator, 'webdriver', {get: () => undefined});
            """
        })
    except Exception:
        pass
    return driver


def wait_body(driver, timeout=25):
    WebDriverWait(driver, timeout).until(EC.presence_of_element_located((By.TAG_NAME, "body")))


# =========================
#   YTJ: THROTTLE + SAFE_GET
# =========================
class Throttle:
    def __init__(self):
        self.base_delay = 0.35
        self.max_delay = 7.0
        self.fail_streak = 0

    def sleep_between(self):
        d = min(self.max_delay, self.base_delay + self.fail_streak * 0.7)
        d += random.uniform(0.0, 0.45)
        time.sleep(d)

    def on_success(self):
        self.fail_streak = max(0, self.fail_streak - 1)

    def on_fail(self):
        self.fail_streak = min(10, self.fail_streak + 1)

    def backoff(self, attempt: int):
        time.sleep(min(self.max_delay, (1.0 * (1.7 ** attempt)) + random.uniform(0.0, 0.9)))


def page_looks_blocked_or_captcha(driver) -> bool:
    try:
        t = (driver.find_element(By.TAG_NAME, "body").text or "").lower()
        bad = ["captcha", "verify you are human", "varmista että et ole robotti", "too many requests", "429"]
        return any(x in t for x in bad)
    except Exception:
        return False


def safe_get(driver, url: str, throttle: Throttle, stop_evt, status_cb, max_attempts=6) -> bool:
    for attempt in range(max_attempts):
        if stop_evt.is_set():
            return False
        try:
            throttle.sleep_between()
            driver.get(url)
            wait_body(driver, 25)

            if page_looks_blocked_or_captcha(driver):
                status_cb("YTJ: CAPTCHA/esto → hidastetaan…")
                throttle.on_fail()
                throttle.backoff(attempt)
                continue

            throttle.on_success()
            return True
        except TimeoutException:
            status_cb("YTJ: timeout → retry…")
            throttle.on_fail()
            throttle.backoff(attempt)
        except WebDriverException:
            status_cb("YTJ: webdriver error → retry…")
            throttle.on_fail()
            throttle.backoff(attempt)
    return False


# =========================
#   YTJ: Avaa "Näytä" varmasti (AGGRESSIVE)
# =========================
def js_click(driver, el) -> bool:
    try:
        driver.execute_script("arguments[0].scrollIntoView({block:'center'});", el)
        time.sleep(0.05)
    except Exception:
        pass
    try:
        el.click()
        return True
    except Exception:
        pass
    try:
        driver.execute_script("arguments[0].click();", el)
        return True
    except Exception:
        return False


def ytj_open_all_nayta(driver, status_cb=None, max_rounds=10):
    """
    Avaa kaikki 'Näytä' (Sähköposti/Matkapuhelin yms).
    Tehdään useita kierroksia, koska DOM päivittyy.
    Lisäksi yritetään myös shadowDOM läpi JS:llä.
    """
    total_clicked = 0

    # 1) normaali DOM - XPATH matchit (täsmä + contains)
    xpaths = [
        "//button[normalize-space()='Näytä']",
        "//a[normalize-space()='Näytä']",
        "//*[@role='button' and normalize-space()='Näytä']",
        "//button[contains(normalize-space(.), 'Näytä')]",
        "//*[@role='button' and contains(normalize-space(.), 'Näytä')]",
        # jos teksti ei ole napissa mutta label/title on
        "//button[contains(@aria-label,'Näytä') or contains(@title,'Näytä')]",
        "//*[@role='button' and (contains(@aria-label,'Näytä') or contains(@title,'Näytä'))]",
    ]

    def collect_buttons():
        found = []
        for xp in xpaths:
            try:
                found.extend(driver.find_elements(By.XPATH, xp))
            except Exception:
                pass
        # filteröi näkyvät
        out = []
        for e in found:
            try:
                if e.is_displayed() and e.is_enabled():
                    out.append(e)
            except Exception:
                pass
        return out

    for r in range(max_rounds):
        btns = collect_buttons()
        clicked_round = 0

        # järjestä y-sijainnin mukaan
        def y_pos(el):
            try:
                return el.location_once_scrolled_into_view.get("y", 10**9)
            except Exception:
                try:
                    return el.location.get("y", 10**9)
                except Exception:
                    return 10**9

        btns.sort(key=y_pos)

        for b in btns:
            try:
                if not b.is_displayed() or not b.is_enabled():
                    continue
                txt = (b.text or "").strip().casefold()
                aria = (b.get_attribute("aria-label") or "").casefold()
                title = (b.get_attribute("title") or "").casefold()

                # sallitaan: teksti on Näytä TAI aria/title sisältää näytä
                if not (txt == "näytä" or "näytä" in txt or "näytä" in aria or "näytä" in title):
                    continue

                if js_click(driver, b):
                    clicked_round += 1
                    total_clicked += 1
                    time.sleep(0.12)
            except StaleElementReferenceException:
                continue
            except Exception:
                continue

        # 2) Shadow DOM - varmistus (jos joku nappi on web componentissa)
        # Tämä etsii kaikki elementit joilla innerText/aria-label sisältää "Näytä" ja klikkaa.
        try:
            shadow_clicked = driver.execute_script(
                """
                const results = [];
                function deepQuery(node) {
                  if (!node) return;
                  // element
                  if (node.nodeType === 1) {
                    const el = node;
                    const t = (el.innerText || '').trim();
                    const aria = (el.getAttribute && el.getAttribute('aria-label')) || '';
                    const title = (el.getAttribute && el.getAttribute('title')) || '';
                    const role = (el.getAttribute && el.getAttribute('role')) || '';
                    const isBtn = (el.tagName === 'BUTTON' || role === 'button');

                    if (isBtn && (t === 'Näytä' || t.includes('Näytä') || aria.includes('Näytä') || title.includes('Näytä'))) {
                      results.push(el);
                    }

                    // shadow root
                    if (el.shadowRoot) {
                      deepQuery(el.shadowRoot);
                    }
                  }
                  // children
                  const kids = node.children || [];
                  for (let i=0; i<kids.length; i++) deepQuery(kids[i]);
                }
                deepQuery(document.body);
                let c = 0;
                for (const el of results) {
                  try { el.scrollIntoView({block:'center'}); el.click(); c++; } catch(e) {}
                }
                return c;
                """
            )
            if isinstance(shadow_clicked, int) and shadow_clicked > 0:
                clicked_round += shadow_clicked
                total_clicked += shadow_clicked
        except Exception:
            pass

        if status_cb:
            status_cb(f"YTJ: 'Näytä' avattu kierros {r+1}/{max_rounds} (klikattu {clicked_round})")

        if clicked_round == 0:
            break

        time.sleep(0.25)

    return total_clicked


def extract_email_from_ytj_page(driver) -> str:
    # 1) mailto:
    try:
        for a in driver.find_elements(By.TAG_NAME, "a"):
            href = (a.get_attribute("href") or "")
            if href.lower().startswith("mailto:"):
                return href.split(":", 1)[1].strip()
    except Exception:
        pass
    # 2) regex body
    try:
        body = driver.find_element(By.TAG_NAME, "body").text or ""
        return pick_email_from_text(body)
    except Exception:
        return ""


def wait_email_appears(driver, timeout=7.5) -> str:
    end = time.time() + timeout
    last = ""
    while time.time() < end:
        try:
            body = driver.find_element(By.TAG_NAME, "body").text or ""
            last = pick_email_from_text(body)
            if last:
                return last
        except Exception:
            pass
        time.sleep(0.25)
    return last


def extract_company_name_from_ytj_page(driver) -> str:
    for tag in ("h1", "h2"):
        try:
            el = driver.find_element(By.TAG_NAME, tag)
            txt = (el.text or "").strip()
            if txt and len(txt) >= 3:
                return txt
        except Exception:
            pass
    try:
        t = (driver.title or "").replace("YTJ", "").strip(" -|")
        return t.strip()
    except Exception:
        return ""


def find_input_by_label_text(driver, label_text: str):
    # ensisijainen: <label for="...">
    try:
        labels = driver.find_elements(By.XPATH, f"//label[contains(normalize-space(.), '{label_text}')]")
        for lab in labels:
            fid = (lab.get_attribute("for") or "").strip()
            if fid:
                el = driver.find_element(By.ID, fid)
                if el.is_displayed() and el.is_enabled():
                    return el
    except Exception:
        pass

    # fallback: "label_text" jälkeen seuraava input
    try:
        el = driver.find_element(By.XPATH, f"(//*[contains(normalize-space(.), '{label_text}')])[1]/following::input[not(@type='hidden')][1]")
        if el.is_displayed() and el.is_enabled():
            return el
    except Exception:
        pass
    return None


def ytj_rank_result_links(driver, company_name: str, limit=10):
    links = []
    end = time.time() + 8
    while time.time() < end:
        try:
            links = driver.find_elements(By.XPATH, "//a[contains(@href,'/yritys/')]")
            links = [a for a in links if "/yritys/" in (a.get_attribute("href") or "")]
            if links:
                break
        except Exception:
            pass
        time.sleep(0.25)

    cands = []
    for a in links[:60]:
        try:
            href = a.get_attribute("href") or ""
            shown = (a.text or "").strip()
            ctx = ""
            try:
                ctx = (a.find_element(By.XPATH, "ancestor::li[1]").text or "").strip()
            except Exception:
                try:
                    ctx = (a.find_element(By.XPATH, "ancestor::div[1]").text or "").strip()
                except Exception:
                    ctx = shown

            cand_name = shown or (ctx.splitlines()[0].strip() if ctx else "")
            score = int(similarity(company_name, cand_name) * 100)

            m = re.search(r"/yritys/([^/?#]+)", href)
            yt = m.group(1).strip() if m else ""
            cands.append((score, yt, href, cand_name))
        except Exception:
            continue

    cands.sort(key=lambda x: x[0], reverse=True)
    return cands[:limit]


def ytj_open_best_company_by_name(driver, company_name: str, throttle: Throttle, stop_evt, status_cb):
    ok = safe_get(driver, YTJ_HOME, throttle, stop_evt, status_cb, max_attempts=6)
    if not ok:
        return (False, "", "")

    time.sleep(0.6)
    name_input = find_input_by_label_text(driver, "Yrityksen tai yhteisön nimi")
    if not name_input:
        status_cb("YTJ: en löytänyt kenttää 'Yrityksen tai yhteisön nimi'.")
        return (False, "", "")

    try:
        name_input.click()
        name_input.clear()
    except Exception:
        pass

    status_cb(f"YTJ: haku nimellä: {company_name}")
    name_input.send_keys(company_name)
    name_input.send_keys(Keys.ENTER)
    time.sleep(1.2)

    # jos ohjautui suoraan yrityssivulle
    if "/yritys/" in (driver.current_url or ""):
        m = re.search(r"/yritys/([^/?#]+)", driver.current_url or "")
        yt = m.group(1).strip() if m else ""
        page_name = extract_company_name_from_ytj_page(driver)
        return (True, yt, page_name)

    cands = ytj_rank_result_links(driver, company_name, limit=10)
    if not cands:
        return (False, "", "")

    best = None
    for idx, (score, yt, url, cand_name) in enumerate(cands[:6], start=1):
        if stop_evt.is_set():
            return (False, "", "")
        target = YTJ_COMPANY_URL.format(yt) if yt else url
        status_cb(f"YTJ: kokeilen tulosta {idx}/6 (score={score})")
        ok2 = safe_get(driver, target, throttle, stop_evt, status_cb, max_attempts=6)
        if not ok2:
            continue
        page_name = extract_company_name_from_ytj_page(driver)
        sim = similarity(company_name, page_name or cand_name)
        if sim >= 0.72:
            return (True, yt, page_name)
        if best is None or sim > best[0]:
            best = (sim, yt, page_name)

    if best and best[0] >= 0.60:
        return (True, best[1], best[2])

    return (False, "", "")


# =========================
#   FETCH EMAILS
# =========================
def fetch_emails_from_ytj_by_yts_and_names(driver, yts, names, status_cb, progress_cb, log_cb, stop_evt):
    cache = cache_load()
    throttle = Throttle()

    master_rows = []  # yritysnimi, yt, email, status, source, ts
    emails = []
    seen_emails = set()

    def add_email(email: str):
        e = (email or "").strip()
        if not e:
            return
        k = e.lower()
        if k in seen_emails:
            return
        seen_emails.add(k)
        emails.append(e)

    total = len(yts or []) + len(names or [])
    progress_cb(0, max(1, total))
    done = 0

    # ---- 1) YT ----
    for yt in (yts or []):
        if stop_evt.is_set():
            break
        done += 1
        progress_cb(done, total)

        yt = normalize_yt(yt) or (yt or "").strip()
        if not yt:
            continue

        status_cb(f"YTJ (Y-tunnus): {yt} ({done}/{total})")

        hit = cache.get("yt", {}).get(yt)
        if hit is not None:
            email = (hit.get("email") or "").strip()
            if email:
                add_email(email)
                log_cb(f"[CACHE YT] {yt} -> {email}")
                master_rows.append(["", yt, email, "FOUND", "YT(CACHE)", int(time.time())])
            else:
                master_rows.append(["", yt, "", "NO_EMAIL", "YT(CACHE)", int(time.time())])
            continue

        ok = safe_get(driver, YTJ_COMPANY_URL.format(yt), throttle, stop_evt, status_cb, max_attempts=6)
        if not ok:
            cache.setdefault("yt", {})[yt] = {"email": "", "status": "ERROR", "ts": int(time.time())}
            cache_save(cache)
            master_rows.append(["", yt, "", "ERROR", "YT", int(time.time())])
            continue

        # ✅ avaa NÄYTÄ
        ytj_open_all_nayta(driver, status_cb=None, max_rounds=10)

        email = extract_email_from_ytj_page(driver) or wait_email_appears(driver, timeout=7.5)

        if email:
            add_email(email)
            cache.setdefault("yt", {})[yt] = {"email": email, "status": "FOUND", "ts": int(time.time())}
            cache_save(cache)
            master_rows.append(["", yt, email, "FOUND", "YT", int(time.time())])
            log_cb(f"[LIVE YT] {yt} -> {email}")
        else:
            cache.setdefault("yt", {})[yt] = {"email": "", "status": "NO_EMAIL", "ts": int(time.time())}
            cache_save(cache)
            master_rows.append(["", yt, "", "NO_EMAIL", "YT", int(time.time())])
            log_cb(f"[LIVE YT] {yt} -> (ei sähköpostia)")

    # ---- 2) NAMES ----
    for nm in (names or []):
        if stop_evt.is_set():
            break
        done += 1
        progress_cb(done, total)

        nm = (nm or "").strip()
        if not nm:
            continue

        status_cb(f"YTJ (nimi): {nm} ({done}/{total})")

        key = normalize_company_name(nm)
        hit = cache.get("name", {}).get(key)
        if hit is not None:
            email = (hit.get("email") or "").strip()
            yt_cached = (hit.get("yt") or "").strip()
            if email:
                add_email(email)
                log_cb(f"[CACHE NAME] {nm} -> {email}")
                master_rows.append([nm, yt_cached, email, "FOUND", "NAME(CACHE)", int(time.time())])
            else:
                master_rows.append([nm, yt_cached, "", "NO_EMAIL", "NAME(CACHE)", int(time.time())])
            continue

        ok, yt_found, page_name = ytj_open_best_company_by_name(driver, nm, throttle, stop_evt, status_cb)
        if not ok:
            cache.setdefault("name", {})[key] = {"email": "", "yt": "", "status": "NO_MATCH", "ts": int(time.time())}
            cache_save(cache)
            master_rows.append([nm, "", "", "NO_MATCH", "NAME", int(time.time())])
            continue

        yt_found = normalize_yt(yt_found) or (yt_found or "").strip()
        if yt_found:
            ok2 = safe_get(driver, YTJ_COMPANY_URL.format(yt_found), throttle, stop_evt, status_cb, max_attempts=6)
            if not ok2:
                cache.setdefault("name", {})[key] = {"email": "", "yt": yt_found, "status": "ERROR", "ts": int(time.time())}
                cache_save(cache)
                master_rows.append([nm, yt_found, "", "ERROR", "NAME", int(time.time())])
                continue

        # ✅ avaa NÄYTÄ
        ytj_open_all_nayta(driver, status_cb=None, max_rounds=10)

        email = extract_email_from_ytj_page(driver) or wait_email_appears(driver, timeout=7.5)

        if email:
            add_email(email)
            cache.setdefault("name", {})[key] = {"email": email, "yt": yt_found, "status": "FOUND", "ts": int(time.time()), "raw": nm, "page": page_name}
            if yt_found:
                cache.setdefault("yt", {})[yt_found] = {"email": email, "status": "FOUND", "ts": int(time.time())}
            cache_save(cache)
            master_rows.append([nm, yt_found, email, "FOUND", "NAME", int(time.time())])
            log_cb(f"[LIVE NAME] {nm} -> {email}")
        else:
            cache.setdefault("name", {})[key] = {"email": "", "yt": yt_found, "status": "NO_EMAIL", "ts": int(time.time()), "raw": nm, "page": page_name}
            cache_save(cache)
            master_rows.append([nm, yt_found, "", "NO_EMAIL", "NAME", int(time.time())])
            log_cb(f"[LIVE NAME] {nm} -> (ei sähköpostia)")

    progress_cb(total, max(1, total))

    # ✅ tärkeä: palautetaan vain aidot emailit
    emails = sorted(set([e for e in emails if e.strip()]), key=lambda x: x.lower())
    return emails, master_rows


# =========================
#   SCROLLABLE FRAME (UI)
# =========================
class ScrollableFrame(ttk.Frame):
    def __init__(self, container):
        super().__init__(container)
        self.canvas = tk.Canvas(self, highlightthickness=0)
        self.scrollbar = ttk.Scrollbar(self, orient="vertical", command=self.canvas.yview)
        self.inner = ttk.Frame(self.canvas)
        self.inner.bind("<Configure>", lambda e: self.canvas.configure(scrollregion=self.canvas.bbox("all")))
        self.win = self.canvas.create_window((0, 0), window=self.inner, anchor="nw")
        self.canvas.configure(yscrollcommand=self.scrollbar.set)
        self.canvas.pack(side="left", fill="both", expand=True)
        self.scrollbar.pack(side="right", fill="y")
        self.canvas.bind("<Configure>", self._on_canvas_configure)
        self.canvas.bind_all("<MouseWheel>", self._on_mousewheel)

    def _on_canvas_configure(self, event):
        try:
            self.canvas.itemconfig(self.win, width=event.width)
        except Exception:
            pass

    def _on_mousewheel(self, event):
        try:
            self.canvas.yview_scroll(int(-1 * (event.delta / 120)), "units")
        except Exception:
            pass


# =========================
#   APP
# =========================
class App(tk.Tk):
    def __init__(self):
        super().__init__()
        reset_log()

        self.stop_evt = threading.Event()

        self.title("ProtestiBotti (YTJ Näytä FIX)")
        self.geometry("1100x900")

        root = ScrollableFrame(self)
        root.pack(fill="both", expand=True)
        self.ui = root.inner

        tk.Label(self.ui, text="ProtestiBotti (YTJ Näytä FIX)", font=("Arial", 18, "bold")).pack(pady=8)
        tk.Label(
            self.ui,
            text=(
                "• Kauppalehti: Ctrl+A/Ctrl+C → liitä → poimi yritysnimet + Y-tunnukset\n"
                "• PDF: poimii nimet + YT → hakee emailit YTJ:stä\n"
                "• YTJ: avaa KAIKKI 'Näytä' aggressiivisesti ennen email-lukua\n"
                "• Output: DOCX/TXT/CSV + master.csv\n"
            ),
            justify="center"
        ).pack(pady=4)

        ctrl = tk.Frame(self.ui)
        ctrl.pack(fill="x", padx=12, pady=6)
        tk.Button(ctrl, text="STOP (keskeytä)", fg="white", bg="#aa0000", command=self.stop_now).pack(side="left", padx=6)
        tk.Button(ctrl, text="Avaa output-kansio", command=self.open_output_folder).pack(side="left", padx=6)
        tk.Button(ctrl, text="Tyhjennä cache", command=self.clear_cache_ui).pack(side="left", padx=6)

        box_yt = tk.LabelFrame(self.ui, text="Liitä Y-tunnukset → YTJ emailit", padx=10, pady=10)
        box_yt.pack(fill="x", padx=12, pady=10)
        tk.Label(box_yt, text="Liitä Y-tunnukset (rivit/pilkut/välit käy).").pack(anchor="w")
        self.yt_text = tk.Text(box_yt, height=6)
        self.yt_text.pack(fill="x", pady=6)
        tk.Button(box_yt, text="Hae emailit YTJ (YT)", font=("Arial", 11, "bold"), command=self.start_fetch_yts_only).pack(anchor="w", padx=6)

        box_kl = tk.LabelFrame(self.ui, text="Kauppalehti: liitä koko sivu → poimi nimet/YT → YTJ emailit", padx=10, pady=10)
        box_kl.pack(fill="x", padx=12, pady=10)
        tk.Label(box_kl, text="KL protestilistassa: Ctrl+A → Ctrl+C → liitä tähän:").pack(anchor="w")
        self.kl_text = tk.Text(box_kl, height=10)
        self.kl_text.pack(fill="x", pady=6)
        tk.Button(box_kl, text="Poimi + Hae emailit YTJ (nimet + YT)", font=("Arial", 11, "bold"), command=self.start_fetch_from_kl).pack(anchor="w", padx=6)

        box_pdf = tk.LabelFrame(self.ui, text="PDF → (nimet+YT) → YTJ emailit", padx=10, pady=10)
        box_pdf.pack(fill="x", padx=12, pady=10)
        tk.Button(box_pdf, text="Valitse PDF ja hae emailit", font=("Arial", 11, "bold"), command=self.start_pdf_to_ytj).pack(anchor="w", padx=6)

        self.status = tk.Label(self.ui, text="Valmiina.", font=("Arial", 11))
        self.status.pack(pady=6)

        self.progress = ttk.Progressbar(self.ui, orient="horizontal", mode="determinate", length=1000)
        self.progress.pack(pady=6)

        frame = tk.Frame(self.ui)
        frame.pack(fill="both", expand=True, padx=14, pady=10)
        tk.Label(frame, text="Live-logi:").pack(anchor="w")
        self.listbox = tk.Listbox(frame, height=14)
        self.listbox.pack(side="left", fill="both", expand=True)
        sb = tk.Scrollbar(frame, orient="vertical", command=self.listbox.yview)
        sb.pack(side="right", fill="y")
        self.listbox.configure(yscrollcommand=sb.set)

        tk.Label(self.ui, text=f"Tallennus: {OUT_DIR}\nCache: {CACHE_PATH}", wraplength=1020, justify="center").pack(pady=10)

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

    def stop_now(self):
        self.stop_evt.set()
        self.set_status("STOP pyydetty…")

    def open_output_folder(self):
        try:
            if sys.platform.startswith("win"):
                os.startfile(OUT_DIR)  # noqa
            else:
                messagebox.showinfo("Polku", OUT_DIR)
        except Exception as e:
            messagebox.showerror("Virhe", f"Ei voitu avata kansiota.\n{e}")

    def clear_cache_ui(self):
        if messagebox.askyesno("Tyhjennä cache", "Haluatko varmasti tyhjentää cachet?"):
            cache_clear()
            self.ui_log("Cache tyhjennetty.")
            messagebox.showinfo("OK", "Cache tyhjennetty.")

    def _read_yts_from_box(self):
        return parse_yts_only(self.yt_text.get("1.0", tk.END))

    def start_fetch_yts_only(self):
        yts = self._read_yts_from_box()
        if not yts:
            messagebox.showwarning("Ei löytynyt", "Liitä Y-tunnukset ensin.")
            return
        self.stop_evt.clear()
        threading.Thread(target=self._run_fetch_ytj, args=(yts, []), daemon=True).start()

    def start_fetch_from_kl(self):
        raw = self.kl_text.get("1.0", tk.END)
        names, yts = parse_names_and_yts_from_copied_text(raw)
        if not names and not yts:
            messagebox.showwarning("Ei löytynyt", "Liitä KL-sivun teksti ensin.")
            return
        # tallennetaan nimet erikseen myös
        save_word_plain_lines(names, "yritysnimet_kl.docx")
        save_txt_lines(names, "yritysnimet_kl.txt")
        self.ui_log(f"Poimittu nimet={len(names)} | y-tunnukset={len(yts)} (tallennettu yritysnimet_kl.*)")
        self.stop_evt.clear()
        threading.Thread(target=self._run_fetch_ytj, args=(yts, names), daemon=True).start()

    def start_pdf_to_ytj(self):
        pdf_path = filedialog.askopenfilename(filetypes=[("PDF files", "*.pdf")])
        if not pdf_path:
            return
        self.stop_evt.clear()
        threading.Thread(target=self._run_pdf_to_ytj, args=(pdf_path,), daemon=True).start()

    def _run_pdf_to_ytj(self, pdf_path):
        try:
            self.set_status("Luetaan PDF: nimet + Y-tunnukset…")
            names, yts = extract_names_and_yts_from_pdf(pdf_path)
            if not names and not yts:
                messagebox.showwarning("Ei löytynyt", "PDF:stä ei löytynyt nimiä tai Y-tunnuksia.")
                return

            # tallennetaan nimet + yts ensin
            save_word_plain_lines(names, "yritysnimet_pdf.docx")
            save_txt_lines(names, "yritysnimet_pdf.txt")
            save_word_plain_lines(yts, "ytunnukset_pdf.docx")
            save_txt_lines(yts, "ytunnukset_pdf.txt")

            self.ui_log(f"PDF poimittu: nimet={len(names)} | y-tunnukset={len(yts)}")
            self.set_status("Haetaan emailit YTJ:stä…")
            self._run_fetch_ytj(yts, names)
        except Exception as e:
            self.ui_log(f"VIRHE PDF: {e}")
            messagebox.showerror("Virhe", str(e))

    def _run_fetch_ytj(self, yts, names):
        driver = None
        try:
            self.set_status("Käynnistetään Chrome (YTJ)…")
            driver = start_new_driver()

            emails, master_rows = fetch_emails_from_ytj_by_yts_and_names(
                driver=driver,
                yts=yts or [],
                names=names or [],
                status_cb=self.set_status,
                progress_cb=self.set_progress,
                log_cb=self.ui_log,
                stop_evt=self.stop_evt,
            )

            # ✅ VAIN emailit tiedostoihin (ei “ei emailia” rivejä)
            save_word_plain_lines(emails, "sahkopostit_ytj.docx")
            save_txt_lines(emails, "sahkopostit_ytj.txt")
            save_csv_rows([[e] for e in emails], ["email"], "sahkopostit_ytj.csv")
            save_csv_rows(master_rows, ["yritysnimi", "ytunnus", "email", "status", "source", "ts"], "master.csv")

            self.set_status("Valmis!")
            messagebox.showinfo("Valmis", f"Sähköposteja: {len(emails)}\n\nKansio:\n{OUT_DIR}")

        except Exception as e:
            self.ui_log(f"VIRHE: {e}")
            self.ui_log("Katso log.txt (traceback alhaalla).")
            try:
                import traceback
                tb = traceback.format_exc()
                for ln in tb.splitlines():
                    self.ui_log(ln)
            except Exception:
                pass
            self.set_status("Virhe. Katso log.txt")
            messagebox.showerror("Virhe", f"{e}\n\nKatso log.txt:\n{LOG_PATH}")
        finally:
            if driver:
                try:
                    driver.quit()
                except Exception:
                    pass

    def on_close(self):
        try:
            self.stop_evt.set()
        except Exception:
            pass
        self.destroy()


if __name__ == "__main__":
    App().mainloop()
