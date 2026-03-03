# kl_protest_module.py
# Kauppalehti Protestilista helper
# - pysyy oikealla URL:lla (ei harhaudu OMXHPI tms)
# - sulkee overlayt / popupit / modaalit
# - klikkaa "Näytä lisää" loopissa

import re
import time
from typing import Callable, Optional, List

from selenium.webdriver.common.by import By
from selenium.common.exceptions import WebDriverException

YT_RE = re.compile(r"\b\d{7}-\d\b|\b\d{8}\b")

KL_ALLOWED_PREFIX = "https://www.kauppalehti.fi/yritykset/protestilista"
KL_DOMAIN = "www.kauppalehti.fi"


def _normalize_yt(x: str) -> Optional[str]:
    x = (x or "").strip().replace(" ", "")
    if re.fullmatch(r"\d{7}-\d", x):
        return x
    if re.fullmatch(r"\d{8}", x):
        return x[:7] + "-" + x[7]
    return None


def _status(cb: Optional[Callable[[str], None]], msg: str):
    if cb:
        cb(msg)


def ensure_on_page(driver, url: str, status_cb: Optional[Callable[[str], None]] = None):
    try:
        cur = (driver.current_url or "")
    except Exception:
        cur = ""

    if not cur.startswith(KL_ALLOWED_PREFIX):
        _status(status_cb, "KL: Navigoidaan protestilistaan…")
        driver.get(url)
        time.sleep(0.6)

    # If we got redirected somewhere else, force back
    for _ in range(3):
        cur = driver.current_url or ""
        if cur.startswith(KL_ALLOWED_PREFIX):
            return
        _status(status_cb, f"KL: Huom! Olet väärällä sivulla ({cur}). Palataan protestilistaan…")
        driver.get(url)
        time.sleep(0.8)


def close_overlays(driver, status_cb: Optional[Callable[[str], None]] = None):
    # Generic overlay closers
    texts = ["Sulje", "Close", "×", "X", "Ok", "Hyväksy", "Accept", "Accept all", "Hyväksy kaikki"]
    xpaths = [
        "//button",
        "//a",
        "//*[@role='button']",
        "//*[contains(@class,'close')]",
        "//*[contains(@aria-label,'close')]",
        "//*[contains(@aria-label,'Close')]",
        "//*[contains(@aria-label,'sulje')]",
        "//*[contains(@aria-label,'Sulje')]",
    ]

    def safe_click(el) -> bool:
        try:
            if not el.is_displayed() or not el.is_enabled():
                return False
            driver.execute_script("arguments[0].scrollIntoView({block:'center'});", el)
            time.sleep(0.02)
            try:
                el.click()
            except Exception:
                driver.execute_script("arguments[0].click();", el)
            return True
        except Exception:
            return False

    # ESC a couple times
    try:
        from selenium.webdriver.common.keys import Keys

        body = driver.find_element(By.TAG_NAME, "body")
        body.send_keys(Keys.ESCAPE)
        time.sleep(0.05)
        body.send_keys(Keys.ESCAPE)
        time.sleep(0.05)
    except Exception:
        pass

    for _ in range(2):
        clicked_any = False
        for xp in xpaths:
            try:
                els = driver.find_elements(By.XPATH, xp)
            except Exception:
                els = []
            for e in els[:120]:
                try:
                    t = (e.text or "").strip()
                    al = (e.get_attribute("aria-label") or "").strip()
                    title = (e.get_attribute("title") or "").strip()
                    blob = " ".join([t, al, title]).strip()
                    if not blob:
                        continue

                    low = blob.lower()
                    if any(tok.lower() in low for tok in [x.lower() for x in texts]) or low in ("x", "×"):
                        if safe_click(e):
                            clicked_any = True
                            _status(status_cb, "KL: Suljettiin overlay/pop-up.")
                            time.sleep(0.10)
                            break
                except Exception:
                    continue
            if clicked_any:
                break
        if not clicked_any:
            break


def _is_still_protest(driver) -> bool:
    try:
        cur = driver.current_url or ""
    except Exception:
        return False
    return (KL_DOMAIN in cur) and cur.startswith(KL_ALLOWED_PREFIX)


def click_show_more_until_end(
    driver,
    stop_flag,
    status_cb: Optional[Callable[[str], None]] = None,
    max_passes: int = 600,
    scroll_sleep: float = 0.20,
    post_click_sleep: float = 0.25,
):
    """
    Scroll + click "Näytä lisää" until it disappears / max passes.
    Guard: if navigation goes away -> return to protest list.
    """
    from selenium.webdriver.common.keys import Keys

    def find_show_more():
        # exact-ish button text
        candidates = []
        for xp in ("//button", "//a", "//*[@role='button']"):
            try:
                candidates.extend(driver.find_elements(By.XPATH, xp))
            except Exception:
                pass
        for el in candidates:
            try:
                t = (el.text or "").strip().lower()
                if not t:
                    continue
                if "näytä lisää" in t or "show more" in t:
                    if el.is_displayed() and el.is_enabled():
                        return el
            except Exception:
                continue
        return None

    body = None
    try:
        body = driver.find_element(By.TAG_NAME, "body")
    except Exception:
        pass

    passes = 0
    while passes < max_passes and (not stop_flag.is_set()):
        passes += 1

        # close overlays frequently
        close_overlays(driver, status_cb=status_cb)

        # guard against wrong navigation
        if not _is_still_protest(driver):
            _status(status_cb, f"KL: Navigointi karkasi ({driver.current_url}). Palautetaan protestilistaan…")
            driver.get(KL_ALLOWED_PREFIX)
            time.sleep(0.8)
            continue

        btn = find_show_more()
        if btn:
            try:
                driver.execute_script("arguments[0].scrollIntoView({block:'center'});", btn)
                time.sleep(0.02)
                try:
                    btn.click()
                except Exception:
                    driver.execute_script("arguments[0].click();", btn)
                _status(status_cb, f"KL: Klikattu 'Näytä lisää' ({passes}/{max_passes})")
                time.sleep(post_click_sleep)
                continue
            except WebDriverException:
                time.sleep(0.15)

        # no button -> scroll down a bit and try again
        try:
            if body:
                body.send_keys(Keys.END)
            else:
                driver.execute_script("window.scrollTo(0, document.body.scrollHeight);")
        except Exception:
            pass
        time.sleep(scroll_sleep)

        # try one more time; if still none for several cycles -> likely end
        if passes > 25 and btn is None:
            # quick heuristic: if button absent and we're near bottom -> exit
            # (we keep it simple to avoid false stops)
            pass

    _status(status_cb, "KL: Latauslooppi valmis.")


def extract_ytunnukset_via_js(driver) -> List[str]:
    """
    Extract Y-tunnus from full DOM text (fast).
    """
    try:
        txt = driver.execute_script("return document.body ? document.body.innerText : '';") or ""
    except Exception:
        try:
            txt = driver.find_element(By.TAG_NAME, "body").text or ""
        except Exception:
            txt = ""

    yts = set()
    for m in YT_RE.findall(txt):
        n = _normalize_yt(m)
        if n:
            yts.add(n)
    return sorted(yts)
