# kl_protest_module.py (HARDENED)
# - ensure_on_page(driver, url, status_cb)
# - click_show_more_until_end(driver, stop_flag, status_cb, ...)
# - extract_ytunnukset_via_js(driver)

import re
import time
from typing import Callable, List, Optional

from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.action_chains import ActionChains

from selenium.common.exceptions import (
    WebDriverException,
    StaleElementReferenceException,
    ElementClickInterceptedException,
)

DEFAULT_URL = "https://www.kauppalehti.fi/yritykset/protestilista"
YT_RE = re.compile(r"\b\d{7}-\d\b|\b\d{8}\b")


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
        pass

    try:
        ActionChains(driver).move_to_element(el).pause(0.05).click(el).perform()
        return True
    except Exception:
        return False


def _try_press_esc(driver, n=2):
    try:
        body = driver.find_element(By.TAG_NAME, "body")
        for _ in range(n):
            body.send_keys(Keys.ESCAPE)
            time.sleep(0.06)
    except Exception:
        pass


def close_overlays(driver, status_cb=None):
    _try_press_esc(driver, n=2)

    keywords = [
        "sulje",
        "close",
        "hyväksy",
        "accept",
        "accept all",
        "hyväksy kaikki",
        "ok",
        "selvä",
        "agree",
        "i agree",
        "×",
        "✕",
    ]

    try:
        elems = driver.find_elements(By.XPATH, "//button|//a|//*[@role='button']")
    except Exception:
        elems = []

    clicked = 0
    for el in elems[:180]:
        try:
            if not el.is_displayed():
                continue
            txt = (el.text or "").strip() or (el.get_attribute("aria-label") or "").strip()
            if not txt:
                continue
            if any(k in txt.lower() for k in keywords):
                if _safe_click(driver, el):
                    clicked += 1
                    time.sleep(0.08)
        except Exception:
            continue

    if clicked:
        _log(status_cb, f"KL: overlay/consent click: {clicked}")

    try:
        driver.execute_script(
            """
            const sels = [
              '[role="dialog"]', '.modal', '.overlay', '.cookie', '.consent',
              '[data-testid*="overlay"]', '[class*="overlay"]', '[class*="modal"]',
              '[style*="position: fixed"]'
            ];
            const nodes = [];
            sels.forEach(s => nodes.push(...document.querySelectorAll(s)));
            nodes.slice(0, 60).forEach(n => { try { n.style.display='none'; n.style.visibility='hidden'; } catch(e){} });
            """
        )
    except Exception:
        pass


def ensure_on_page(driver, url: str, status_cb=None):
    target = (url or "").strip() or DEFAULT_URL
    _log(status_cb, f"KL: Navigoidaan: {target}")
    try:
        driver.get(target)
        time.sleep(1.3)
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
        # FI
        "//button[contains(translate(.,'ABCDEFGHIJKLMNOPQRSTUVWXYZÅÄÖ','abcdefghijklmnopqrstuvwxyzåäö'),'näytä lisää')]",
        "//a[contains(translate(.,'ABCDEFGHIJKLMNOPQRSTUVWXYZÅÄÖ','abcdefghijklmnopqrstuvwxyzåäö'),'näytä lisää')]",
        "//*[self::button or self::a][.//*[contains(translate(.,'ABCDEFGHIJKLMNOPQRSTUVWXYZÅÄÖ','abcdefghijklmnopqrstuvwxyzåäö'),'näytä lisää')]]",
        # EN
        "//button[contains(translate(.,'ABCDEFGHIJKLMNOPQRSTUVWXYZ','abcdefghijklmnopqrstuvwxyz'),'show more')]",
        "//a[contains(translate(.,'ABCDEFGHIJKLMNOPQRSTUVWXYZ','abcdefghijklmnopqrstuvwxyz'),'show more')]",
        "//*[self::button or self::a][.//*[contains(translate(.,'ABCDEFGHIJKLMNOPQRSTUVWXYZ','abcdefghijklmnopqrstuvwxyz'),'show more')]]",
        # generic
        "//*[@data-testid='show-more' or @data-testid='load-more']",
        "//*[contains(@class,'show-more') or contains(@class,'load-more')]",
    ]
    for xp in xps:
        try:
            el = driver.find_element(By.XPATH, xp)
            if el and el.is_displayed():
                return el
        except Exception:
            continue
    return None


def click_show_more_until_end(
    driver,
    stop_flag,
    status_cb=None,
    max_passes: int = 700,
    scroll_sleep: float = 0.25,
    post_click_sleep: float = 0.35,
    stuck_rounds_limit: int = 8,
):
    clicks = 0
    stuck = 0

    prev_len = 0
    prev_yt = 0

    for i in range(max_passes):
        if stop_flag and stop_flag.is_set():
            _log(status_cb, "KL: STOP havaittu – lopetetaan loop.")
            break

        close_overlays(driver, status_cb=status_cb)
        _scroll_to_bottom(driver)
        time.sleep(max(0.05, scroll_sleep))

        try:
            src = driver.page_source or ""
            cur_len = len(src)
            cur_yt = len({y for y in (_normalize_yt(x) for x in YT_RE.findall(src)) if y})
        except Exception:
            cur_len = prev_len
            cur_yt = prev_yt

        btn = _find_show_more(driver)
        if not btn:
            _log(status_cb, f"KL: Ei löytynyt 'Näytä lisää' (pass {i+1}/{max_passes}).")
            break

        ok = False
        try:
            ok = _safe_click(driver, btn)
        except (ElementClickInterceptedException, StaleElementReferenceException, WebDriverException):
            close_overlays(driver, status_cb=status_cb)
            ok = _safe_click(driver, btn)
        except Exception:
            ok = False

        if not ok:
            _log(status_cb, "KL: Klikkaus epäonnistui (btn löytyi mutta ei klikattu).")
            break

        clicks += 1
        if clicks % 10 == 0:
            _log(status_cb, f"KL: Näytä lisää klikattu {clicks} kertaa…")

        time.sleep(max(0.05, post_click_sleep))

        try:
            src2 = driver.page_source or ""
            new_len = len(src2)
            new_yt = len({y for y in (_normalize_yt(x) for x in YT_RE.findall(src2)) if y})
        except Exception:
            new_len = cur_len
            new_yt = cur_yt

        grew = (new_len > cur_len + 50) or (new_yt > cur_yt)
        if not grew:
            stuck += 1
            _log(status_cb, f"KL: Ei kasvanut (stuck {stuck}/{stuck_rounds_limit}) len {cur_len}->{new_len}, yt {cur_yt}->{new_yt}")
            if stuck >= stuck_rounds_limit:
                _log(status_cb, "KL: Sisältö ei kasva -> lopetetaan loop (todennäköisesti loppu / blokki).")
                break
        else:
            stuck = 0

        prev_len, prev_yt = new_len, new_yt

    _log(status_cb, f"KL: Loop valmis. Klikkauksia: {clicks}")


def extract_ytunnukset_via_js(driver) -> List[str]:
    parts = []

    try:
        parts.append(driver.execute_script("return document.body ? document.body.innerText : ''") or "")
    except Exception:
        pass

    try:
        parts.append(driver.execute_script("return document.documentElement ? document.documentElement.outerHTML : ''") or "")
    except Exception:
        pass

    try:
        parts.append(
            driver.execute_script(
                """
                const out = [];
                const as = Array.from(document.querySelectorAll('a')).slice(0, 4000);
                for (const a of as) {
                  const h = (a.getAttribute('href') || '').trim();
                  if (h) out.push(h);
                  const dt = (a.getAttribute('data-id') || '') + ' ' + (a.getAttribute('data-ytunnus') || '');
                  if (dt.trim()) out.push(dt.trim());
                }
                return out.join('\\n');
                """
            )
            or ""
        )
    except Exception:
        pass

    blob = "\n".join(parts)
    found = set()
    for m in YT_RE.findall(blob):
        n = _normalize_yt(m)
        if n:
            found.add(n)
    return sorted(found)
