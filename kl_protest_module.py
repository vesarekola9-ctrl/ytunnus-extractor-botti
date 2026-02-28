# kl_protest_module.py
# Kauppalehti Protestilista helperit:
# - ensure_on_page: varmistaa että ollaan URL:ssa
# - click_show_more_until_end: overlay/popup closers + "Näytä lisää" loop
# - extract_ytunnukset_via_js: kerää koko DOM text / HTML ja regex Y-tunnukset

import re
import time
from typing import Callable, List, Optional

from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.common.exceptions import (
    NoSuchElementException,
    WebDriverException,
    StaleElementReferenceException,
    ElementClickInterceptedException,
)

# hyväksyy myös 8-numeroisen muodon (muutetaan 7-1 muotoon lopussa)
YT_RE = re.compile(r"\b\d{7}-\d\b|\b\d{8}\b")

DEFAULT_URL = "https://www.kauppalehti.fi/yritykset/protestilista"


def _log(status_cb: Optional[Callable[[str], None]], msg: str):
    if status_cb:
        try:
            status_cb(msg)
        except Exception:
            pass


def _normalize_yt(m: str) -> Optional[str]:
    m = (m or "").strip().replace(" ", "")
    if re.fullmatch(r"\d{7}-\d", m):
        return m
    if re.fullmatch(r"\d{8}", m):
        return m[:7] + "-" + m[7]
    return None


def _safe_click(driver, el) -> bool:
    if not el:
        return False
    try:
        driver.execute_script("arguments[0].scrollIntoView({block:'center'});", el)
        time.sleep(0.03)
        try:
            el.click()
        except Exception:
            driver.execute_script("arguments[0].click();", el)
        return True
    except Exception:
        return False


def _try_press_esc(driver, n=2):
    try:
        body = driver.find_element(By.TAG_NAME, "body")
        for _ in range(n):
            body.send_keys(Keys.ESCAPE)
            time.sleep(0.08)
    except Exception:
        pass


def close_overlays(driver, status_cb=None):
    """
    Best-effort: sulje consent/modal/overlay:
    - ESC
    - etsi napit/ankkurit joissa close/sulje/hyväksy/accept/ok ja klikkaa
    - piilota yleisiä overlay-elementtejä JS:llä
    """
    _try_press_esc(driver, n=2)

    keywords = [
        "sulje", "close", "hyväksy", "accept", "ok", "selvä", "agree", "i agree",
        "x", "×", "✕"
    ]

    try:
        elems = driver.find_elements(By.XPATH, "//button|//a|//*[@role='button']")
    except Exception:
        elems = []

    for el in elems[:120]:
        try:
            if not el.is_displayed() or not el.is_enabled():
                continue
            t = (el.text or "").strip()
            if not t:
                t = (el.get_attribute("aria-label") or "").strip()
            if not t:
                continue
            low = t.lower()
            if any(k in low for k in keywords):
                if _safe_click(driver, el):
                    time.sleep(0.12)
        except Exception:
            continue

    try:
        driver.execute_script(
            """
            const sels = [
              '[role="dialog"]', '.modal', '.overlay', '.cookie', '.consent',
              '[style*="position: fixed"]'
            ];
            const nodes = [];
            sels.forEach(s => nodes.push(...document.querySelectorAll(s)));
            nodes.slice(0, 40).forEach(n => { try { n.style.display='none'; } catch(e){} });
            """
        )
    except Exception:
        pass


def ensure_on_page(driver, url: str, status_cb=None):
    target = (url or "").strip() or DEFAULT_URL
    _log(status_cb, f"KL: Navigoidaan: {target}")
    try:
        driver.get(target)
        time.sleep(1.6)
    except Exception:
        pass
    close_overlays(driver, status_cb=status_cb)


def _scroll_to_bottom(driver):
    try:
        driver.execute_script("window.scrollTo(0, document.body.scrollHeight);")
    except Exception:
        pass


def _find_show_more(driver):
    xps = [
        "//button[contains(translate(.,'ABCDEFGHIJKLMNOPQRSTUVWXYZÅÄÖ','abcdefghijklmnopqrstuvwxyzåäö'),'näytä lisää')]",
        "//a[contains(translate(.,'ABCDEFGHIJKLMNOPQRSTUVWXYZÅÄÖ','abcdefghijklmnopqrstuvwxyzåäö'),'näytä lisää')]",
        "//button[contains(translate(.,'ABCDEFGHIJKLMNOPQRSTUVWXYZ','abcdefghijklmnopqrstuvwxyz'),'show more')]",
        "//a[contains(translate(.,'ABCDEFGHIJKLMNOPQRSTUVWXYZ','abcdefghijklmnopqrstuvwxyz'),'show more')]",
    ]
    for xp in xps:
        try:
            el = driver.find_element(By.XPATH, xp)
            if el and el.is_displayed():
                return el
        except NoSuchElementException:
            continue
        except Exception:
            continue
    return None


def click_show_more_until_end(
    driver,
    stop_flag,
    status_cb=None,
    max_passes: int = 500,
    scroll_sleep: float = 0.25,
    post_click_sleep: float = 0.35,
):
    clicks = 0
    for i in range(max_passes):
        if stop_flag and stop_flag.is_set():
            _log(status_cb, "KL: STOP havaittu – lopetetaan Näytä lisää loop.")
            break

        close_overlays(driver, status_cb=status_cb)
        _scroll_to_bottom(driver)
        time.sleep(max(0.05, scroll_sleep))

        btn = _find_show_more(driver)
        if not btn:
            _log(status_cb, f"KL: Ei enää 'Näytä lisää' (pass {i+1}/{max_passes}).")
            break

        try:
            driver.execute_script("arguments[0].scrollIntoView({block:'center'});", btn)
            time.sleep(0.05)
        except Exception:
            pass

        try:
            ok = _safe_click(driver, btn)
            if not ok:
                break
            clicks += 1
            if clicks % 10 == 0:
                _log(status_cb, f"KL: Näytä lisää klikattu {clicks} kertaa…")
            time.sleep(max(0.05, post_click_sleep))
        except (ElementClickInterceptedException, StaleElementReferenceException, WebDriverException):
            close_overlays(driver, status_cb=status_cb)
            time.sleep(0.1)
            continue
        except Exception:
            break

    _log(status_cb, f"KL: Näytä lisää loop valmis. Klikkauksia: {clicks}")


def extract_ytunnukset_via_js(driver) -> List[str]:
    text = ""
    html = ""
    try:
        text = driver.execute_script("return document.body ? document.body.innerText : '';") or ""
    except Exception:
        text = ""

    try:
        html = driver.execute_script("return document.documentElement ? document.documentElement.outerHTML : '';") or ""
    except Exception:
        html = ""

    blob = (text or "") + "\n" + (html or "")
    found = set()
    for m in YT_RE.findall(blob):
        n = _normalize_yt(m)
        if n:
            found.add(n)
    return sorted(found)
