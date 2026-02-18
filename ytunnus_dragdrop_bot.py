import os
import re
import sys
import time
import threading
import tkinter as tk
from tkinter import messagebox, filedialog
from tkinter import ttk
from dataclasses import dataclass

import PyPDF2
from docx import Document

from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import (
    StaleElementReferenceException,
    WebDriverException,
    TimeoutException,
)
from webdriver_manager.chrome import ChromeDriverManager


# =========================
#   CONFIG / REGEX
# =========================
YT_RE = re.compile(r"\b\d{7}-\d\b|\b\d{8}\b")
EMAIL_RE = re.compile(r"[A-Za-z0-9_.+-]+@[A-Za-z0-9-]+\.[A-Za-z0-9-.]+")
EMAIL_A_RE = re.compile(r"[A-Za-z0-9_.+-]+\s*\(a\)\s*[A-Za-z0-9-]+\.[A-Za-z0-9-.]+", re.I)

KAUPPALEHTI_URL = "https://www.kauppalehti.fi/yritykset/protestilista"
YTJ_HOME = "https://tietopalvelu.ytj.fi/"
YTJ_COMPANY_URL = "https://tietopalvelu.ytj.fi/yritys/{}"


# =========================
#   PATHS
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


def dump_debug(driver, tag="debug"):
    html_path = os.path.join(OUT_DIR, f"{tag}.html")
    png_path = os.path.join(OUT_DIR, f"{tag}.png")
    try:
        with open(html_path, "w", encoding="utf-8") as f:
            f.write(driver.page_source or "")
    except Exception:
        html_path = None
    try:
        driver.save_screenshot(png_path)
    except Exception:
        png_path = None
    return html_path, png_path


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


def normalize_name(s: str) -> str:
    s = (s or "").strip().casefold()
    # kevyt normalisointi
    s = re.sub(r"\s+", " ", s)
    s = s.replace("  ", " ")
    # yhtenäistä muutamia yleisiä muotoja
    s = s.replace("asunto oy", "as oy")
    s = s.replace("asunto-osakeyhtiö", "as oy")
    return s


def token_set(s: str):
    s = normalize_name(s)
    s = re.sub(r"[^0-9a-zåäö\s-]", " ", s)
    parts = [p for p in re.split(r"\s+", s) if p]
    return set(parts)


def score_match(query_name: str, query_location: str, candidate_text: str) -> int:
    """
    Pisteytä YTJ-hakutulos: nimi + sijainti.
    candidate_text = koko kortin/rivin teksti.
    """
    qn = normalize_name(query_name)
    qtoks = token_set(query_name)
    cl = normalize_name(candidate_text)

    score = 0

    # Nimi täsmää
    if qn and qn == cl:
        score += 120
    if qn and qn in cl:
        score += 60

    # Token-overlap
    ctoks = token_set(candidate_text)
    overlap = len(qtoks & ctoks)
    score += overlap * 8

    # Sijainti
    loc = normalize_name(query_location)
    if loc and loc in cl:
        score += 35

    # Pieni bonus oikeille yhtiömuodoille jos löytyy
    for suf in [" oy", " ab", " ky", " oyj", " tmi", " ry"]:
        if suf.strip() in qn and suf.strip() in cl:
            score += 10

    return score


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


def save_word_table(rows, headers, filename):
    """
    rows: list of lists (same length as headers)
    """
    path = os.path.join(OUT_DIR, filename)
    doc = Document()
    table = doc.add_table(rows=1, cols=len(headers))
    hdr_cells = table.rows[0].cells
    for i, h in enumerate(headers):
        hdr_cells[i].text = str(h)

    for r in rows:
        row_cells = table.add_row().cells
        for i, v in enumerate(r):
            row_cells[i].text = "" if v is None else str(v)

    doc.save(path)
    return path


def safe_scroll_into_view(driver, elem):
    try:
        driver.execute_script("arguments[0].scrollIntoView({block:'center'});", elem)
    except Exception:
        pass


def safe_click(driver, elem) -> bool:
    try:
        safe_scroll_into_view(driver, elem)
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
                        safe_click(driver, e)
                        time.sleep(0.2)
                        found = True
                        break
            except Exception:
                continue
        if not found:
            break


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
#   SELENIUM START
# =========================
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


def list_tabs(driver):
    tabs = []
    for h in driver.window_handles:
        try:
            driver.switch_to.window(h)
            tabs.append((driver.title or "", driver.current_url or ""))
        except Exception:
            tabs.append(("", ""))
    return tabs


def focus_protestilista_tab(driver, log_cb=None) -> bool:
    target_handle = None
    for handle in driver.window_handles:
        try:
            driver.switch_to.window(handle)
            url = (driver.current_url or "")
            if "kauppalehti.fi" in url and "protestilista" in url:
                target_handle = handle
                break
        except Exception:
            continue

    if log_cb:
        log_cb("Chrome TAB LISTA (title | url):")
        for title, url in list_tabs(driver):
            log_cb(f"  {title} | {url}")

    if target_handle:
        driver.switch_to.window(target_handle)
        return True
    return False


def open_new_tab(driver, url="about:blank"):
    driver.execute_script("window.open(arguments[0], '_blank');", url)
    driver.switch_to.window(driver.window_handles[-1])


# =========================
#   KAUPPALEHTI PAGE CHECKS
# =========================
def page_looks_like_login_or_paywall(driver) -> bool:
    try:
        text = (driver.find_element(By.TAG_NAME, "body").text or "").lower()
        bad_words = [
            "kirjaudu", "tilaa", "tilaajille", "vahvista henkilöllisyytesi",
            "sign in", "subscribe", "login",
            "pääsy evätty", "access denied",
            "jokin meni pieleen", "something went wrong",
        ]
        return any(w in text for w in bad_words)
    except Exception:
        return False


def ensure_protestilista_open(driver, status_cb, log_cb, max_wait_seconds=900) -> bool:
    if focus_protestilista_tab(driver, log_cb):
        status_cb("Löytyi protestilista-tab.")
    else:
        status_cb("Protestilista-tab ei löytynyt -> avaan protestilistan uuteen tabiin…")
        open_new_tab(driver, KAUPPALEHTI_URL)

    WebDriverWait(driver, 25).until(EC.presence_of_element_located((By.TAG_NAME, "body")))
    try_accept_cookies(driver)

    start = time.time()
    warned = False

    while True:
        try_accept_cookies(driver)

        ok = "protestilista" in (driver.current_url or "")
        if ok and not page_looks_like_login_or_paywall(driver):
            # anna JS:lle hetki
            time.sleep(0.8)
            return True

        if page_looks_like_login_or_paywall(driver) and not warned:
            warned = True
            status_cb("KL näyttää kirjautumisen/tilaajamuurin. Kirjaudu Chrome(9222)-ikkunassa ja pidä protestilista näkyvissä.")
            log_cb("Waiting for user login/unblock on Kauppalehti…")
            try:
                messagebox.showinfo(
                    "Kirjaudu Kauppalehteen",
                    "Botti ei näe vielä protestilistaa.\n\n"
                    "Kirjaudu Kauppalehteen siinä Chrome-ikkunassa joka on käynnistetty 9222-portilla.\n"
                    "Kun protestilista näkyy, botti jatkaa automaattisesti."
                )
            except Exception:
                pass

        if time.time() - start > max_wait_seconds:
            status_cb("Aikakatkaisu: protestilista ei tullut näkyviin.")
            log_cb("ERROR: timeout waiting protestilista")
            html, png = dump_debug(driver, "kl_timeout")
            log_cb(f"DEBUG dump: {html} | {png}")
            return False

        time.sleep(2)


def click_nayta_lisaa(driver) -> bool:
    try:
        driver.execute_script("window.scrollTo(0, document.body.scrollHeight);")
    except Exception:
        pass
    time.sleep(0.25)

    for b in driver.find_elements(By.XPATH, "//button|//*[@role='button']|//a"):
        try:
            if not b.is_displayed() or not b.is_enabled():
                continue
            txt = (b.text or "").strip().lower()
            if ("näytä" in txt and "lisää" in txt) or ("lataa" in txt and "lisää" in txt):
                return safe_click(driver, b)
        except Exception:
            continue
    return False


# =========================
#   HYBRID: KL rows -> YTJ search by name
# =========================
@dataclass
class KLRow:
    company: str
    location: str
    amount: str
    date: str
    ptype: str
    source: str


def read_kl_visible_rows(driver):
    """
    Lue näkyvät KL protestilistataulukon pää-rivit.
    Tämä ei avaa mitään detailia, vaan lukee rivin sarakkeet.
    """
    rows = []
    trs = driver.find_elements(By.XPATH, "//table//tbody//tr")
    for tr in trs:
        try:
            if not tr.is_displayed():
                continue
            txt = (tr.text or "").strip()
            if not txt:
                continue
            # skip detail rows
            if "Y-tunnus" in txt or "Y-TUNNUS" in txt:
                continue

            tds = tr.find_elements(By.XPATH, ".//td")
            if len(tds) < 6:
                continue

            company = (tds[0].text or "").strip()
            location = (tds[1].text or "").strip()
            amount = (tds[2].text or "").strip()
            date = (tds[3].text or "").strip()
            ptype = (tds[4].text or "").strip()
            source = (tds[5].text or "").strip()

            if company:
                rows.append(KLRow(company, location, amount, date, ptype, source))
        except Exception:
            continue
    return rows


def collect_kl_rows_all(driver, status_cb, log_cb, max_pages=999):
    """
    Kerää KL:stä kaikki rivit (Yritys + Sijainti + jne) käyttäen Näytä lisää.
    Dedupataan (company, location, date, amount, type, source).
    """
    if not ensure_protestilista_open(driver, status_cb, log_cb, max_wait_seconds=900):
        return []

    collected = []
    seen = set()
    rounds_no_new = 0
    page_clicks = 0

    while True:
        try_accept_cookies(driver)

        if page_looks_like_login_or_paywall(driver):
            status_cb("KL: näyttää blokilta/paywallilta -> debug-dump ja odotan…")
            html, png = dump_debug(driver, "kl_blocked")
            log_cb(f"DEBUG dump: {html} | {png}")
            time.sleep(3)
            continue

        visible = read_kl_visible_rows(driver)
        new = 0
        for r in visible:
            key = (r.company, r.location, r.amount, r.date, r.ptype, r.source)
            if key in seen:
                continue
            seen.add(key)
            collected.append(r)
            new += 1

        status_cb(f"KL: kerätty {len(collected)} (uusii {new})")

        if new == 0:
            rounds_no_new += 1
        else:
            rounds_no_new = 0

        if page_clicks >= max_pages:
            status_cb("KL: max_pages täynnä -> lopetan.")
            break

        if click_nayta_lisaa(driver):
            page_clicks += 1
            time.sleep(1.2)
            continue

        # jos ei uutta eikä lisää-nappia pariin kierrokseen -> done
        if rounds_no_new >= 2:
            break

    if not collected:
        html, png = dump_debug(driver, "kl_zero_rows")
        log_cb(f"DEBUG dump: {html} | {png}")

    return collected


# =========================
#   YTJ: search by name -> pick best -> extract YT + email
# =========================
def wait_ytj_loaded(driver):
    wait = WebDriverWait(driver, 25)
    wait.until(EC.presence_of_element_located((By.TAG_NAME, "body")))
    # ei pakoteta tiettyä tekstiä – YTJ vaihtuu joskus
    time.sleep(0.2)


def extract_email_from_ytj(driver):
    # 1) mailto
    try:
        for a in driver.find_elements(By.TAG_NAME, "a"):
            href = (a.get_attribute("href") or "")
            if href.lower().startswith("mailto:"):
                return href.split(":", 1)[1].strip()
    except Exception:
        pass

    # 2) taulukon "Sähköposti"-rivi
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

    # 3) koko body
    try:
        return pick_email_from_text(driver.find_element(By.TAG_NAME, "body").text or "")
    except Exception:
        return ""


def extract_yt_from_ytj_page(driver) -> str:
    """
    Poimi Y-tunnus YTJ-yrityssivulta.
    """
    try:
        text = driver.find_element(By.TAG_NAME, "body").text or ""
        # YTJ:ssä näkyy usein muodossa 1234567-8 tai 12345678
        for m in YT_RE.findall(text):
            n = normalize_yt(m)
            if n:
                return n
    except Exception:
        pass
    return ""


def ytj_open_home(driver):
    driver.get(YTJ_HOME)
    wait_ytj_loaded(driver)
    try_accept_cookies(driver)


def ytj_search_and_pick_best(driver, company_name: str, location: str, log_cb=None):
    """
    Hakee YTJ:stä yrityksen nimellä.
    Palauttaa (best_url, best_score, best_text).
    """
    # Mene etusivulle ja käytä hakukenttää (resilient)
    ytj_open_home(driver)

    # Etsi hakukenttä
    search_input = None
    candidates = driver.find_elements(By.XPATH, "//input")
    for inp in candidates:
        try:
            if not inp.is_displayed() or not inp.is_enabled():
                continue
            typ = (inp.get_attribute("type") or "").lower()
            aria = (inp.get_attribute("aria-label") or "").lower()
            ph = (inp.get_attribute("placeholder") or "").lower()
            # YTJ:llä hakukenttä on yleensä search/text ja sisältää "hae" / "yritys"
            if typ in ("search", "text") and ("hae" in aria or "hae" in ph or "yritys" in aria or "yritys" in ph):
                search_input = inp
                break
        except Exception:
            continue

    if not search_input:
        # fallback: ensimmäinen näkyvä input
        for inp in candidates:
            try:
                if inp.is_displayed() and inp.is_enabled():
                    search_input = inp
                    break
            except Exception:
                continue

    if not search_input:
        if log_cb:
            log_cb("YTJ: hakukenttää ei löytynyt")
        return "", 0, ""

    # Syötä haku
    try:
        search_input.click()
        time.sleep(0.05)
        search_input.clear()
    except Exception:
        pass

    try:
        search_input.send_keys(company_name)
        time.sleep(0.05)
        search_input.send_keys(Keys.ENTER)
    except Exception:
        # fallback: enter bodylle
        try:
            driver.find_element(By.TAG_NAME, "body").send_keys(Keys.ENTER)
        except Exception:
            pass

    # Odota että tuloksia/renderöitymistä tapahtuu
    time.sleep(1.0)
    try_accept_cookies(driver)

    # Kerää tuloslinkit (/yritys/)
    links = driver.find_elements(By.XPATH, "//a[contains(@href, '/yritys/')]")
    best = ("", 0, "")
    seen_urls = set()

    for a in links[:40]:
        try:
            url = (a.get_attribute("href") or "").strip()
            if not url or "/yritys/" not in url:
                continue
            if url in seen_urls:
                continue
            seen_urls.add(url)

            # Yritä ottaa "kortin" teksti (a:n lähin container)
            cand_text = ""
            try:
                container = a.find_element(By.XPATH, "ancestor::*[self::li or self::div or self::article][1]")
                cand_text = (container.text or "").strip()
            except Exception:
                cand_text = (a.text or "").strip()

            if not cand_text:
                continue

            sc = score_match(company_name, location, cand_text)
            if sc > best[1]:
                best = (url, sc, cand_text)
        except Exception:
            continue

    return best


def ytj_get_email_by_company_name(driver, company_name: str, location: str, log_cb=None):
    """
    Palauttaa dict: {yt, email, url, score}
    """
    url, score, text = ytj_search_and_pick_best(driver, company_name, location, log_cb=log_cb)
    if not url:
        return {"yt": "", "email": "", "url": "", "score": 0}

    driver.get(url)
    wait_ytj_loaded(driver)
    try_accept_cookies(driver)

    # avaa mahdolliset "Näytä" -napit (jos YTJ piilottaa tietoja)
    for _ in range(2):
        clicked = False
        for b in driver.find_elements(By.XPATH, "//button|//*[@role='button']|//a"):
            try:
                t = (b.text or "").strip().lower()
                if t == "näytä" and b.is_displayed() and b.is_enabled():
                    safe_click(driver, b)
                    clicked = True
                    time.sleep(0.15)
            except Exception:
                continue
        if not clicked:
            break

    yt = extract_yt_from_ytj_page(driver)
    email = extract_email_from_ytj(driver)

    return {"yt": yt, "email": email, "url": url, "score": score}


def run_hybrid_kl_to_ytj(driver, status_cb, progress_cb, log_cb):
    """
    1) Kerää KL rivit (yritys+sijainti+...)
    2) Hae YTJ:stä y-tunnus ja email nimellä
    3) Palauta (kl_rows, results_rows)
    """
    status_cb("KL: kerätään yritysrivit (nimi+sijainti)…")
    kl_rows = collect_kl_rows_all(driver, status_cb, log_cb)

    if not kl_rows:
        return [], []

    # Tallenna KL rivit wordiin heti
    kl_table = []
    for r in kl_rows:
        kl_table.append([r.company, r.location, r.date, r.amount, r.ptype, r.source])
    kl_path = save_word_table(
        kl_table,
        headers=["Yritys", "Sijainti", "Häiriöpäivä", "Summa", "Tyyppi", "Lähde"],
        filename="kl_rivit.docx",
    )
    log_cb(f"Tallennettu: {kl_path}")

    # Avaa YTJ uuteen tabiin, jotta KL pysyy ehjänä
    status_cb("Avataan YTJ uuteen tabiin…")
    open_new_tab(driver, "about:blank")
    time.sleep(0.2)

    results = []
    progress_cb(0, len(kl_rows))

    for i, r in enumerate(kl_rows, start=1):
        status_cb(f"YTJ-haku: {i}/{len(kl_rows)} — {r.company} ({r.location})")
        progress_cb(i - 1, len(kl_rows))

        try:
            data = ytj_get_email_by_company_name(driver, r.company, r.location, log_cb=log_cb)
        except Exception as e:
            log_cb(f"YTJ ERROR {r.company}: {e}")
            data = {"yt": "", "email": "", "url": "", "score": 0}

        results.append([
            r.company,
            r.location,
            data.get("yt", ""),
            data.get("email", ""),
            data.get("url", ""),
            str(data.get("score", 0)),
        ])

        # pieni hengähdys ettei hakusivu hermostu
        time.sleep(0.15)

    progress_cb(len(kl_rows), len(kl_rows))

    out_path = save_word_table(
        results,
        headers=["Yritys", "Sijainti", "Y-tunnus", "Sähköposti", "YTJ-linkki", "Score"],
        filename="tulokset.docx",
    )
    log_cb(f"Tallennettu: {out_path}")

    return kl_rows, results


# =========================
#   YTJ (YT list -> emails) old path for PDF-mode
# =========================
def fetch_emails_from_ytj_by_yt(driver, yt_list, status_cb, progress_cb, log_cb):
    emails = []
    seen = set()
    progress_cb(0, max(1, len(yt_list)))

    for i, yt in enumerate(yt_list, start=1):
        status_cb(f"YTJ: {i}/{len(yt_list)} {yt}")
        progress_cb(i - 1, len(yt_list))

        driver.get(YTJ_COMPANY_URL.format(yt))
        wait_ytj_loaded(driver)
        try_accept_cookies(driver)

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

        time.sleep(0.08)

    progress_cb(len(yt_list), max(1, len(yt_list)))
    return emails


# =========================
#   GUI
# =========================
class App(tk.Tk):
    def __init__(self):
        super().__init__()
        reset_log()

        self.title("ProtestiBotti (HYBRID: KL -> YTJ nimellä)")
        self.geometry("980x620")

        tk.Label(self, text="ProtestiBotti (Hybrid)", font=("Arial", 18, "bold")).pack(pady=10)
        tk.Label(
            self,
            text="Moodit:\n"
                 "1) Kauppalehti → kerää yritys+sijainti → hakee YTJ:stä Y-tunnus + sähköposti nimellä\n"
                 "2) PDF → Y-tunnukset → YTJ sähköpostit",
            justify="center"
        ).pack(pady=4)

        btn_row = tk.Frame(self)
        btn_row.pack(pady=8)

        tk.Button(btn_row, text="Kauppalehti (HYBRID) → Tulokset", font=("Arial", 12), command=self.start_hybrid_mode).grid(row=0, column=0, padx=8)
        tk.Button(btn_row, text="PDF → YTJ", font=("Arial", 12), command=self.start_pdf_mode).grid(row=0, column=1, padx=8)

        self.status = tk.Label(self, text="Valmiina.", font=("Arial", 11))
        self.status.pack(pady=6)

        self.progress = ttk.Progressbar(self, orient="horizontal", mode="determinate", length=920)
        self.progress.pack(pady=6)

        frame = tk.Frame(self)
        frame.pack(fill="both", expand=True, padx=14, pady=10)

        tk.Label(frame, text="Live-logi (uusimmat alimmaisena):").pack(anchor="w")
        self.listbox = tk.Listbox(frame, height=20)
        self.listbox.pack(side="left", fill="both", expand=True)

        sb = tk.Scrollbar(frame, orient="vertical", command=self.listbox.yview)
        sb.pack(side="right", fill="y")
        self.listbox.configure(yscrollcommand=sb.set)

        tk.Label(self, text=f"Tallennus: {OUT_DIR}", wraplength=960, justify="center").pack(pady=6)

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

    # ----- Hybrid mode -----
    def start_hybrid_mode(self):
        threading.Thread(target=self.run_hybrid_mode, daemon=True).start()

    def run_hybrid_mode(self):
        driver = None
        try:
            self.set_status("Liitytään Chrome-bottiin (9222)… (pidä KL protestilista auki ja kirjautuneena)")
            driver = attach_to_existing_chrome()

            self.set_status("Aloitetaan HYBRID: KL rivit → YTJ nimihaku…")
            kl_rows, results = run_hybrid_kl_to_ytj(driver, self.set_status, self.set_progress, self.ui_log)

            if not kl_rows:
                self.set_status("KL: Ei saatu rivejä.")
                messagebox.showwarning("Ei rivejä", "Kauppalehdestä ei saatu rivejä. Katso log.txt / debug-dumpit.")
                return

            # Laske sähköpostien määrä
            emails = [r[3] for r in results if r and r[3]]
            uniq_emails = len(set([e.lower() for e in emails]))

            self.set_status("Valmis!")
            messagebox.showinfo(
                "Valmis",
                f"Valmis!\n\nKansio:\n{OUT_DIR}\n\n"
                f"KL rivejä: {len(kl_rows)}\n"
                f"Sähköposteja (uniikki): {uniq_emails}\n\n"
                f"Tiedostot:\n- kl_rivit.docx\n- tulokset.docx"
            )

        except WebDriverException as e:
            self.ui_log(f"SELENIUM VIRHE: {e}")
            self.set_status("Virhe. Katso log.txt")
            messagebox.showerror("Virhe", f"Selenium/Chrome virhe.\nKatso log.txt:\n{LOG_PATH}\n\n{e}")
        except Exception as e:
            self.ui_log(f"VIRHE: {e}")
            self.set_status("Virhe. Katso log.txt")
            messagebox.showerror("Virhe", f"Tuli virhe.\nKatso log.txt:\n{LOG_PATH}\n\n{e}")
        finally:
            pass  # ei suljeta käyttäjän Chromea

    # ----- PDF mode -----
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

            emails = fetch_emails_from_ytj_by_yt(driver, yt_list, self.set_status, self.set_progress, self.ui_log)
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
