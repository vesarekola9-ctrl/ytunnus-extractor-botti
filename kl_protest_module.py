# kl_protest_module.py
# Kauppalehti Protestilista module
# FIX: prevents accidental navigation to /porssi/... and other wrong pages
# - Ensures we're on protest list URL
# - "Show more" loop robust
# - Extract Y-tunnus via JS/regex from current DOM
#
# NOTE: This module does NOT open devtools/F12. It doesn't need to.
# Selenium can run JS directly via execute_script.

import re
import time
from typing import Callable, List, Optional
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys

DEFAULT_URL = "https://www.kauppalehti.fi/yritykset/protestilista"
YT_RE = re.compile(r"\b\d{7}-\d\b|\b\d{8}\b")


def _log(cb: Optional[Callable[[str], None]], msg: str):
    if cb:
        try:
            cb(msg)
        except Exception:
            pass


def _normalize_yt(m: str) -> Optional[str]:
    m = (m or "").strip().replace(" ", "")
    if re.fullmatch(r"\d{7}-\d", m):
        return m
    if re.fullmatch(r"\d{8}", m):
        return m[:7] + "-" + m[7]
    return None


def _is_protest_url(url: str) -> bool:
    u = (url or "").lower()
    return "/yritykset/protestilista" in u


def _is_bad_url(url: str) -> bool:
    u = (url or "").lower()
    bad = ["/porssi", "/indeksit", "/indeks", "/uutiset", "/mainos", "/tilaa"]
    return any(b in u for b in bad)


def close_overlays(driver):
    # ESC + quick hide common overlays
    try:
        driver.find_element(By.TAG_NAME, "body").send_keys(Keys.ESCAPE)
    except Exception:
        pass
    try:
        driver.execute_script(
            """
            const sels = [
              '[role="dialog"]','[aria-modal="true"]',
              '.modal','.overlay',
              '[class*="overlay"]','[class*="modal"]',
              '[id*="overlay"]','[id*="modal"]'
            ];
            for (const s of sels){
              document.querySelectorAll(s).forEach(n=>{ n.style.display='none'; });
            }
            """
        )
    except Exception:
        pass


def ensure_on_page(driver, url: str, status_cb=None):
    """
    Hard guard: if not on protest list, force navigate to target url (or DEFAULT_URL).
    """
    target = url or DEFAULT_URL
    cur = ""
    try:
        cur = driver.current_url or ""
    except Exception:
        pass

    if not _is_protest_url(cur):
        _log(status_cb, f"KL: Not on protest page (current={cur}) -> opening {target}")
        try:
            driver.get(target)
            time.sleep(1.0)
        except Exception:
            driver.get(DEFAULT_URL)
            time.sleep(1.0)

    # If still wrong (redirect), try default
    try:
        if not _is_protest_url(driver.current_url or ""):
            driver.get(DEFAULT_URL)
            time.sleep(1.0)
    except Exception:
        pass

    close_overlays(driver)


def _find_show_more_button(driver):
    xps = [
        "//button[contains(.,'Näytä lisää')]",
        "//a[contains(.,'Näytä lisää')]",
        "//button[contains(.,'Show more')]",
        "//a[contains(.,'Show more')]",
    ]
    for xp in xps:
        try:
            el = driver.find_element(By.XPATH, xp)
            if el.is_displayed():
                return el
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
    last_url = ""

    for i in range(max_passes):
        if stop_flag.is_set():
            _log(status_cb, "KL: Stop requested.")
            break

        # Guard
        try:
            cur = driver.current_url or ""
        except Exception:
            cur = ""
        if _is_bad_url(cur) or (last_url and cur != last_url and not _is_protest_url(cur)):
            _log(status_cb, f"KL: Detected wrong navigation -> recover to protest list (cur={cur})")
            ensure_on_page(driver, DEFAULT_URL, status_cb=status_cb)

        ensure_on_page(driver, DEFAULT_URL, status_cb=status_cb)

        # Small scroll to trigger lazy loading
        try:
            driver.execute_script("window.scrollBy(0, Math.max(300, window.innerHeight*0.6));")
        except Exception:
            pass
        time.sleep(scroll_sleep)

        close_overlays(driver)

        btn = _find_show_more_button(driver)
        if not btn:
            _log(status_cb, "KL: No show-more button found -> end.")
            break

        # Safe click without following unwanted links
        try:
            driver.execute_script("arguments[0].scrollIntoView({block:'center'});", btn)
            time.sleep(0.08)
            btn.click()
        except Exception:
            try:
                driver.execute_script("arguments[0].click();", btn)
            except Exception:
                _log(status_cb, "KL: Click failed -> end.")
                break

        clicks += 1
        if clicks % 10 == 0:
            _log(status_cb, f"KL: Clicked show-more {clicks}x")

        time.sleep(post_click_sleep)

        try:
            last_url = driver.current_url or ""
        except Exception:
            last_url = ""

    _log(status_cb, f"KL: Show-more finished ({clicks} clicks)")


def extract_ytunnukset_via_js(driver) -> List[str]:
    """
    JS pulls both innerText and HTML; regex finds YTs.
    """
    try:
        blob = driver.execute_script(
            """
            let out = [];
            out.push(document.body ? document.body.innerText : "");
            out.push(document.documentElement ? document.documentElement.outerHTML : "");
            return out.join("\\n");
            """
        )
    except Exception:
        blob = ""

    yts = set()
    for m in YT_RE.findall(blob or ""):
        n = _normalize_yt(m)
        if n:
            yts.add(n)

    return sorted(yts)
