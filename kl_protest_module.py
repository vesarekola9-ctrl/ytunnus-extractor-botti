# kl_protest_module.py
# Kauppalehti Protestilista -> "Näytä lisää" -> Y-tunnukset (JS regex) -> list[str]
# Used by app.py as a NEW module (does not modify existing logic elsewhere).

import re
import time
from typing import List

from selenium.webdriver.common.by import By
from selenium.common.exceptions import WebDriverException

YT_RE = re.compile(r"\b\d{7}-\d\b")

def ensure_on_page(driver, url: str, status_cb=None):
    try:
        driver.get(url)
    except Exception:
        pass
    time.sleep(0.6)
    try:
        if status_cb:
            status_cb(f"KL: Auki: {driver.current_url}")
    except Exception:
        pass

def _safe_buttons(driver):
    try:
        return driver.find_elements(By.XPATH, "//button|//a|//*[@role='button']")
    except Exception:
        return []

def click_show_more_until_end(
    driver,
    stop_flag=None,
    status_cb=None,
    max_passes: int = 500,
    scroll_sleep: float = 0.25,
    post_click_sleep: float = 0.35,
) -> int:
    """
    Click 'Näytä lisää' until it disappears or max_passes reached.
    Returns number of clicks performed.
    """
    clicks = 0
    for _ in range(max_passes):
        if stop_flag and stop_flag.is_set():
            break

        # scroll bottom to reveal button/lazy-load
        try:
            driver.execute_script("window.scrollTo(0, document.body.scrollHeight);")
        except Exception:
            pass
        time.sleep(scroll_sleep)

        btn = None
        for el in _safe_buttons(driver):
            if stop_flag and stop_flag.is_set():
                break
            try:
                t = (el.text or "").strip().lower()
                if not t:
                    continue
                if "näytä lisää" in t and el.is_displayed() and el.is_enabled():
                    btn = el
                    break
            except Exception:
                continue

        if not btn:
            if status_cb:
                status_cb(f"KL: 'Näytä lisää' loppui. Klikkauksia: {clicks}")
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
            time.sleep(post_click_sleep)
        except Exception:
            time.sleep(0.2)

    return clicks

def extract_ytunnukset_via_js(driver) -> List[str]:
    """
    Executes proven JS:
      (document.body.innerText.match(/\\b\\d{7}-\\d\\b/g) || [])
    Then dedup + sort.
    """
    js = r"return (document.body.innerText.match(/\b\d{7}-\d\b/g) || []);"
    try:
        raw = driver.execute_script(js)
    except WebDriverException:
        raw = []

    if not raw:
        return []

    seen = set()
    yts = []
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
