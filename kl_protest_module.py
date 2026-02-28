# kl_protest_module.py
# Kauppalehti Protestilista -> Y-tunnukset (JS/regex) -> list
# NOTE: Requires Chrome session (logged in) via Remote Debugging (attach selenium)

import re
import time
from typing import List, Tuple, Optional

from selenium.webdriver.common.by import By
from selenium.common.exceptions import WebDriverException

YT_RE = re.compile(r"\b\d{7}-\d\b")

DEFAULT_REGEX_JS = r"""\b\d{7}-\d\b"""

def _safe_find_buttons(driver):
    try:
        return driver.find_elements(By.XPATH, "//button|//a|//*[@role='button']")
    except Exception:
        return []

def click_show_more_until_end(driver, stop_flag=None, status_cb=None, max_passes: int = 400) -> int:
    """
    Click 'Näytä lisää' until it disappears or max_passes reached.
    Returns number of clicks.
    """
    clicks = 0
    last_count = -1

    for i in range(max_passes):
        if stop_flag and stop_flag.is_set():
            break

        # Scroll near bottom to reveal lazy-load / button
        try:
            driver.execute_script("window.scrollTo(0, document.body.scrollHeight);")
        except Exception:
            pass
        time.sleep(0.25)

        btn = None
        for el in _safe_find_buttons(driver):
            if stop_flag and stop_flag.is_set():
                break
            try:
                t = (el.text or "").strip().lower()
                if not t:
                    continue
                if "näytä lisää" in t:
                    if el.is_displayed() and el.is_enabled():
                        btn = el
                        break
            except Exception:
                continue

        if not btn:
            # no more show more
            break

        try:
            driver.execute_script("arguments[0].scrollIntoView({block:'center'});", btn)
            time.sleep(0.05)
            try:
                btn.click()
            except Exception:
                driver.execute_script("arguments[0].click();", btn)
            clicks += 1
            if status_cb:
                status_cb(f"KL: Klikattu 'Näytä lisää' ({clicks})")
            time.sleep(0.35)
        except Exception:
            # if click fails, try to continue a bit then break if nothing changes
            time.sleep(0.25)

        # Optional: detect if nothing changes (stuck)
        try:
            current = driver.execute_script(
                "return (document.body.innerText.match(/\\b\\d{7}-\\d\\b/g) || []).length;"
            )
            if isinstance(current, int):
                if current == last_count and i > 5:
                    # likely no more new content coming
                    pass
                last_count = current
        except Exception:
            pass

    return clicks

def extract_ytunnukset_via_js(driver) -> List[str]:
    """
    Executes the proven JS approach:
      (document.body.innerText.match(/\b\d{7}-\d\b/g) || [])
    Then dedups + sorts.
    """
    js = r"return (document.body.innerText.match(/\b\d{7}-\d\b/g) || []);"
    try:
        raw = driver.execute_script(js)
    except WebDriverException:
        raw = []
    if not raw:
        return []

    # Normalize, dedup, sort
    yts = []
    seen = set()
    for x in raw:
        if not x:
            continue
        s = str(x).strip()
        if not YT_RE.fullmatch(s):
            continue
        if s in seen:
            continue
        seen.add(s)
        yts.append(s)
    yts.sort()
    return yts

def open_url(driver, url: str):
    try:
        driver.get(url)
    except Exception:
        pass

def ensure_on_protest_page(driver, url: str, status_cb=None):
    """
    Navigates to URL. If user isn't logged in, page may show a paywall.
    We don't bypass it; we just inform status.
    """
    open_url(driver, url)
    time.sleep(0.6)
    try:
        t = (driver.title or "").lower()
        if status_cb:
            status_cb(f"KL: Auki: {driver.current_url}")
        # no strict checks; UI can show 'Kirjaudu' etc.
    except Exception:
        pass
