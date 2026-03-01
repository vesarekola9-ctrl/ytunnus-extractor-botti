# kl_protest_module.py (ULTRA HARDENED + NAV GUARD)
#
# Fixes:
# - Prevents accidental navigation away from protest list (e.g. to /porssi/...)
# - Strictly filters deep-crawl links to only /yritykset/...
# - After every "Näytä lisää" click, validates URL still looks like protest list.
# - If URL escapes, goes back to protest URL and continues.

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
    NoSuchWindowException,
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
    """Try hard to remove popups, modals, cookie consents, overlays."""
    _try_press_esc(driver, n=2)

    keywords = [
        "sulje", "close",
        "hyväksy", "accept", "accept all", "hyväksy kaikki",
        "ok", "selvä",
        "agree", "i agree",
        "continue", "jatka",
        "×", "✕",
    ]

    try:
        elems = driver.find_elements(By.XPATH, "//button|//a|//*[@role='button']")
    except Exception:
        elems = []

    clicked = 0
    for el in elems[:320]:
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

    # Hide overlay-like nodes
    try:
        driver.execute_script(
            """
            const sels = [
              '[role="dialog"]',
              '.modal', '.overlay', '.cookie', '.consent',
              '[data-testid*="overlay"]',
              '[class*="overlay"]',
              '[class*="modal"]',
              '[class*="paywall"]',
              '[id*="paywall"]'
            ];
            const nodes = [];
            sels.forEach(s => nodes.push(...document.querySelectorAll(s)));
            nodes.slice(0, 160).forEach(n => {
              try { n.style.display='none'; n.style.visibility='hidden'; } catch(e){}
            });
            """
        )
    except Exception:
        pass


def _switch_to_matching_tab(driver, url_substring: str, status_cb=None) -> bool:
    """If multiple tabs are open in debug Chrome, switch to one that matches protest URL."""
    try:
        handles = driver.window_handles
    except Exception:
        return False

    url_substring = (url_substring or "").strip()
    if not url_substring:
        return False

    try:
        cur = driver.current_url or ""
        if url_substring in cur:
            return True
    except Exception:
        pass

    for h in handles:
        try:
            driver.switch_to.window(h)
            time.sleep(0.15)
            u = driver.current_url or ""
            if url_substring in u:
                _log(status_cb, f"KL: Vaihdettu välilehteen: {u}")
                return True
        except NoSuchWindowException:
            continue
        except Exception:
            continue
    return False


def _is_protest_url(u: str) -> bool:
    if not u:
        return False
    u = u.lower()
    return ("/yritykset/protestilista" in u)


def _is_bad_escape_url(u: str) -> bool:
    """Pages we definitely do NOT want to end up on."""
    if not u:
        return False
    u = u.lower()
    bad = [
        "/porssi/", "/pörssi/", "/indeksit", "/indeks",
        "/uutiset", "/news",
        "/mainos", "/ads",
    ]
    return any(b in u for b in bad)


def ensure_on_page(driver, url: str, status_cb=None):
    target = (url or "").strip() or DEFAULT_URL
    _log(status_cb, f"KL: Navigoidaan: {target}")

    _switch_to_matching_tab(driver, "/yritykset/protestilista", status_cb=status_cb)

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


def _count_yts_fast(driver) -> int:
    try:
        txt = driver.execute_script("return document.body ? document.body.innerText : ''") or ""
    except Exception:
        txt = ""
    found = set()
    for m in YT_RE.findall(txt):
        n = _normalize_yt(m)
        if n:
            found.add(n)
    return len(found)


def _find_show_more_selenium(driver):
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


def _click_show_more_js(driver) -> bool:
    """Fallback: click show more via JS scanning innerText."""
    try:
        return bool(
            driver.execute_script(
                """
                function match(t){
                  t=(t||'').toLowerCase();
                  return t.includes('näytä lisää') || t.includes('show more');
                }
                const candidates = []
                  .concat(Array.from(document.querySelectorAll('button')))
                  .concat(Array.from(document.querySelectorAll('a')))
                  .concat(Array.from(document.querySelectorAll('[role="button"]')));
                for (const el of candidates) {
                  const txt = (el.innerText || el.textContent || '').trim();
                  if (!txt) continue;
                  if (!match(txt)) continue;

                  // Try to avoid nav links by preferring <button> or role=button
                  const tag = (el.tagName || '').toLowerCase();
                  if (tag === 'a') {
                    // if anchor has href and isn't javascript:void(0), it might navigate
                    const href = (el.getAttribute('href') || '').trim();
                    if (href && !href.startsWith('#') && !href.toLowerCase().startsWith('javascript')) {
                      // still allow if it looks like a load-more anchor (common pattern)
                      // but we'll keep nav guard in Python anyway
                    }
                  }

                  try { el.scrollIntoView({block:'center'}); } catch(e){}
                  try { el.click(); return true; } catch(e){}
                }
                return false;
                """
            )
        )
    except Exception:
        return False


def _return_to_protest_if_escaped(driver, protest_url: str, status_cb=None) -> bool:
    """If we accidentally navigated away, go back / reload protest URL."""
    try:
        cur = driver.current_url or ""
    except Exception:
        cur = ""

    if _is_protest_url(cur):
        return False

    if _is_bad_escape_url(cur) or (cur and protest_url and protest_url not in cur):
        _log(status_cb, f"KL: VAROITUS: karkasi sivulle: {cur}")
        # try back first
        try:
            driver.back()
            time.sleep(0.7)
        except Exception:
            pass
        try:
            cur2 = driver.current_url or ""
        except Exception:
            cur2 = ""

        if _is_protest_url(cur2):
            _log(status_cb, "KL: Palattiin takaisin (back).")
            close_overlays(driver, status_cb=status_cb)
            return True

        # hard reload
        try:
            _log(status_cb, "KL: Palataan protestilistaan (reload).")
            driver.get(protest_url or DEFAULT_URL)
            time.sleep(1.1)
            close_overlays(driver, status_cb=status_cb)
            return True
        except Exception:
            return True

    return False


def click_show_more_until_end(
    driver,
    stop_flag,
    status_cb=None,
    max_passes: int = 700,
    scroll_sleep: float = 0.25,
    post_click_sleep: float = 0.35,
    stuck_rounds_limit: int = 10,
    protest_url: str = DEFAULT_URL,
):
    clicks = 0
    stuck = 0
    prev_yts = _count_yts_fast(driver)

    for i in range(max_passes):
        if stop_flag and stop_flag.is_set():
            _log(status_cb, "KL: STOP havaittu – lopetetaan loop.")
            break

        # if escaped, recover
        _return_to_protest_if_escaped(driver, protest_url, status_cb=status_cb)

        close_overlays(driver, status_cb=status_cb)
        _scroll_to_bottom(driver)
        time.sleep(max(0.05, scroll_sleep))

        btn = _find_show_more_selenium(driver)
        ok = False
        if btn:
            try:
                ok = _safe_click(driver, btn)
            except (ElementClickInterceptedException, StaleElementReferenceException, WebDriverException):
                close_overlays(driver, status_cb=status_cb)
                ok = _safe_click(driver, btn)
            except Exception:
                ok = False

        if not ok:
            close_overlays(driver, status_cb=status_cb)
            ok = _click_show_more_js(driver)

        # After click, ensure we didn't navigate away
        time.sleep(max(0.05, post_click_sleep))
        if _return_to_protest_if_escaped(driver, protest_url, status_cb=status_cb):
            # we escaped; do NOT count this as progress click, continue loop
            continue

        if not ok:
            _log(status_cb, f"KL: 'Näytä lisää' ei klikattavissa (pass {i+1}/{max_passes}).")
            break

        clicks += 1
        if clicks % 8 == 0:
            _log(status_cb, f"KL: Näytä lisää klikattu {clicks} kertaa…")

        close_overlays(driver, status_cb=status_cb)

        cur_yts = _count_yts_fast(driver)
        grew = cur_yts > prev_yts
        if not grew:
            stuck += 1
            _log(status_cb, f"KL: Ei kasvanut (stuck {stuck}/{stuck_rounds_limit}) yts {prev_yts}->{cur_yts}")
            if stuck >= stuck_rounds_limit:
                _log(status_cb, "KL: Sisältö ei kasva -> lopetetaan loop.")
                break
        else:
            stuck = 0
            prev_yts = cur_yts

    _log(status_cb, f"KL: Loop valmis. Klikkauksia: {clicks}. Y-tunnuksia näkyvissä: {prev_yts}")


def _collect_big_blob_js(driver) -> str:
    """Collect a huge text blob from many sources (text + html + scripts + attributes)."""
    try:
        return driver.execute_script(
            """
            const out = [];

            try { out.push(document.body ? document.body.innerText : ''); } catch(e){}

            try {
              const html = document.documentElement ? document.documentElement.outerHTML : '';
              out.push(html.slice(0, 2_000_000));
            } catch(e){}

            try {
              const scripts = Array.from(document.querySelectorAll('script')).slice(0, 220);
              for (const s of scripts) {
                const t = (s.textContent || '').trim();
                if (!t) continue;
                out.push(t.slice(0, 20000));
              }
            } catch(e){}

            try {
              const as = Array.from(document.querySelectorAll('a')).slice(0, 9000);
              for (const a of as) {
                const h = (a.getAttribute('href') || '').trim();
                if (h) out.push(h);
                const dt =
                  (a.getAttribute('data-id') || '') + ' ' +
                  (a.getAttribute('data-ytunnus') || '') + ' ' +
                  (a.getAttribute('data-y-tunnus') || '');
                if (dt.trim()) out.push(dt.trim());
              }
            } catch(e){}

            try {
              const nodes = Array.from(document.querySelectorAll('*')).slice(0, 12000);
              for (const n of nodes) {
                const attrs = n.attributes;
                if (!attrs) continue;
                for (let i=0;i<attrs.length;i++) {
                  const a = attrs[i];
                  const name = (a.name || '').toLowerCase();
                  if (!name) continue;
                  if (name.startsWith('data-') || name.startsWith('aria-') || name==='id' || name==='class' || name==='href') {
                    const v = (a.value || '').toString();
                    if (v && v.length < 300) out.push(v);
                  }
                }
              }
            } catch(e){}

            return out.join('\\n');
            """
        ) or ""
    except Exception:
        return ""


def _extract_yts_from_blob(blob: str) -> List[str]:
    found = set()
    for m in YT_RE.findall(blob or ""):
        n = _normalize_yt(m)
        if n:
            found.add(n)
    return sorted(found)


def _collect_company_links_from_list(driver, limit: int = 200) -> List[str]:
    """
    STRICT filter: only internal KL company paths.
    Prevents grabbing /porssi/ etc.
    """
    try:
        links = driver.execute_script(
            """
            const out = [];
            const as = Array.from(document.querySelectorAll('a')).slice(0, 14000);
            for (const a of as) {
              const href = (a.getAttribute('href') || '').trim();
              if (!href) continue;
              out.push(href);
            }
            return Array.from(new Set(out)).slice(0, 5000);
            """
        ) or []
    except Exception:
        links = []

    good = []
    for h in links:
        hh = (h or "").lower()
        if "protestilista" in hh:
            continue
        # STRICT: only /yritykset/ links (company pages)
        if "/yritykset/" not in hh:
            continue
        # block known bad
        if "/porssi/" in hh or "/indeks" in hh or "/uutiset" in hh:
            continue
        good.append(h)

    # Normalize to absolute
    abs_links = []
    for h in good:
        if h.startswith("http"):
            # only accept kauppalehti domain
            if "kauppalehti.fi" in h.lower():
                abs_links.append(h)
        else:
            abs_links.append("https://www.kauppalehti.fi" + h)

    # de-dup + limit
    seen = set()
    out = []
    for u in abs_links:
        if u in seen:
            continue
        seen.add(u)
        out.append(u)
        if len(out) >= limit:
            break
    return out


def _deep_crawl_for_yts(driver, status_cb=None, limit_pages: int = 60, protest_url: str = DEFAULT_URL) -> List[str]:
    links = _collect_company_links_from_list(driver, limit=limit_pages * 3)
    if not links:
        _log(status_cb, "KL: Deep crawl: ei löytynyt /yritykset/ yrityslinkkejä listasta.")
        return []

    _log(status_cb, f"KL: Deep crawl fallback: avataan {min(limit_pages, len(links))} yrityssivua…")
    yts = set()

    opened = 0
    for u in links[:limit_pages]:
        try:
            close_overlays(driver, status_cb=status_cb)
            driver.get(u)
            opened += 1
            if opened % 10 == 0:
                _log(status_cb, f"KL: Deep crawl {opened}/{min(limit_pages,len(links))}…")
            time.sleep(0.9)
            close_overlays(driver, status_cb=status_cb)

            blob = _collect_big_blob_js(driver)
            for yt in _extract_yts_from_blob(blob):
                yts.add(yt)

            time.sleep(0.12)
        except Exception:
            continue

    # Return to protest list
    try:
        driver.get(protest_url or DEFAULT_URL)
        time.sleep(0.8)
        close_overlays(driver, status_cb=status_cb)
    except Exception:
        pass

    return sorted(yts)


def extract_ytunnukset_via_js(driver, status_cb=None, protest_url: str = DEFAULT_URL) -> List[str]:
    """
    Primary: ultra blob extraction from current document.
    Fallback: deep crawl company pages if 0.
    Also guards: if currently on wrong page, go back to protest URL first.
    """
    # Guard: if we're not on protest list, return there first
    try:
        cur = driver.current_url or ""
    except Exception:
        cur = ""
    if not _is_protest_url(cur):
        _log(status_cb, f"KL: Guard: ei protestisivulla ({cur}) -> palataan {protest_url}")
        try:
            driver.get(protest_url or DEFAULT_URL)
            time.sleep(1.1)
        except Exception:
            pass

    close_overlays(driver, status_cb=status_cb)

    blob = _collect_big_blob_js(driver)
    yts = _extract_yts_from_blob(blob)

    if yts:
        _log(status_cb, f"KL: Extract: löytyi {len(yts)} Y-tunnusta suoraan listasta.")
        return yts

    _log(status_cb, "KL: Extract: 0 Y-tunnusta listasta -> deep crawl fallback…")
    yts2 = _deep_crawl_for_yts(driver, status_cb=status_cb, limit_pages=60, protest_url=protest_url)
    _log(status_cb, f"KL: Deep crawl: löytyi {len(yts2)} Y-tunnusta.")
    return yts2
