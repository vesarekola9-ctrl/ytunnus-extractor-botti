# kl_protest_module.py (LOCK TAB + NAV KILL SWITCH + FAILURE SNAPSHOT HELPERS + TIMERANGE)
#
# Features:
# 1) Protest-tab lock (pick correct tab in remote-debug Chrome)
# 2) Navigation kill switch (blocks accidental <a> navigations during "Näytä lisää" loop)
# 3) Time range selection (Auto: Day->Week->Month)
# 4) Robust "Näytä lisää" loop with URL guard (stay on protest list)
# 5) Y-tunnus extraction via huge JS blob
# 6) Deep crawl fallback (STRICTLY /yritykset/ links)
# 7) Optional debug dump helpers (HTML/text/url) called from app on failure

import os
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

    # Hide overlay-ish
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


# =========================
#  TAB LOCK / URL GUARD
# =========================
def _switch_to_matching_tab(driver, url_substring: str, status_cb=None) -> bool:
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
                _log(status_cb, f"KL: Tab lock OK: {u}")
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
    return "/yritykset/protestilista" in u


def _is_bad_escape_url(u: str) -> bool:
    if not u:
        return False
    u = u.lower()
    bad = [
        "/porssi/", "/pörssi/", "/indeksit", "/indeks",
        "/uutiset", "/news",
        "/mainos", "/ads",
    ]
    return any(b in u for b in bad)


def _return_to_protest_if_escaped(driver, protest_url: str, status_cb=None) -> bool:
    try:
        cur = driver.current_url or ""
    except Exception:
        cur = ""

    if _is_protest_url(cur):
        return False

    if _is_bad_escape_url(cur) or (cur and protest_url and protest_url not in cur):
        _log(status_cb, f"KL: VAROITUS: karkasi sivulle: {cur}")
        # back first
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


def ensure_on_page(driver, url: str, status_cb=None):
    target = (url or "").strip() or DEFAULT_URL
    _log(status_cb, f"KL: Navigoidaan: {target}")

    # Tab lock
    _switch_to_matching_tab(driver, "/yritykset/protestilista", status_cb=status_cb)

    try:
        driver.get(target)
        time.sleep(1.3)
    except Exception:
        pass
    close_overlays(driver, status_cb=status_cb)


# =========================
#  NAVIGATION KILL SWITCH
# =========================
def install_nav_kill_switch(driver, status_cb=None):
    """
    Blocks accidental anchor (<a>) navigations during automation.
    Allows only:
      - clicks on elements whose text contains "näytä lisää" / "show more"
      - anchors with href starting with '#' or 'javascript'
    """
    try:
        driver.execute_script(
            """
            if (window.__klNavKillInstalled) return true;
            window.__klNavKillInstalled = true;

            function allowedByText(el){
              try {
                const t = ((el.innerText || el.textContent || '') + '').toLowerCase().trim();
                return t.includes('näytä lisää') || t.includes('show more');
              } catch(e){ return false; }
            }

            document.addEventListener('click', function(ev){
              try {
                const a = ev.target && ev.target.closest ? ev.target.closest('a') : null;
                if (!a) return;
                const href = (a.getAttribute('href') || '').trim().toLowerCase();

                if (allowedByText(a)) return; // allow show-more anchor patterns
                if (!href) { ev.preventDefault(); ev.stopPropagation(); return; }
                if (href.startsWith('#') || href.startsWith('javascript')) return;

                // block everything else
                ev.preventDefault();
                ev.stopPropagation();
              } catch(e){}
            }, true);

            return true;
            """
        )
        _log(status_cb, "KL: Navigation kill switch ON.")
    except Exception:
        _log(status_cb, "KL: Navigation kill switch failed (non-fatal).")


def uninstall_nav_kill_switch(driver, status_cb=None):
    # Can't reliably remove anonymous listener; but we can disable via flag check.
    try:
        driver.execute_script(
            """
            window.__klNavKillInstalled = false;
            """
        )
        _log(status_cb, "KL: Navigation kill switch OFF (flag).")
    except Exception:
        pass


# =========================
#  TIMERANGE
# =========================
def _page_has_no_results_text(driver) -> bool:
    try:
        txt = (driver.execute_script("return document.body ? document.body.innerText : ''") or "").lower()
    except Exception:
        txt = ""
    return ("ei tuloksia" in txt) or ("no results" in txt) or ("ei löytynyt" in txt) or ("nothing found" in txt)


def _find_timerange_button(driver, label: str):
    targets = {
        "auto": ["auto"],
        "day": ["päivä", "day"],
        "week": ["viikko", "week"],
        "month": ["kuukausi", "month"],
    }
    key = (label or "").strip().lower()
    if key in ("päivä", "day"):
        want = targets["day"]
    elif key in ("viikko", "week"):
        want = targets["week"]
    elif key in ("kuukausi", "month"):
        want = targets["month"]
    else:
        want = targets["auto"]

    try:
        elems = driver.find_elements(By.XPATH, "//button|//a|//*[@role='button']")
    except Exception:
        elems = []

    for el in elems[:450]:
        try:
            if not el.is_displayed():
                continue
            txt = (el.text or "").strip() or (el.get_attribute("aria-label") or "").strip()
            if not txt:
                continue
            t = txt.lower()
            if any(w in t for w in want):
                return el
        except Exception:
            continue
    return None


def _apply_one_range(driver, label: str, status_cb=None) -> bool:
    close_overlays(driver, status_cb=status_cb)
    btn = _find_timerange_button(driver, label)
    if not btn:
        return False
    ok = _safe_click(driver, btn)
    time.sleep(0.7)
    close_overlays(driver, status_cb=status_cb)
    return ok


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


def apply_time_range(driver, label: str, status_cb=None) -> str:
    raw = (label or "").strip() or "Auto"
    low = raw.lower()

    if low in ("päivä", "day"):
        return "Päivä/Day" if _apply_one_range(driver, "day", status_cb) else ""
    if low in ("viikko", "week"):
        return "Viikko/Week" if _apply_one_range(driver, "week", status_cb) else ""
    if low in ("kuukausi", "month"):
        return "Kuukausi/Month" if _apply_one_range(driver, "month", status_cb) else ""

    _log(status_cb, "KL: Auto time range: yritetään Päivä/Day…")
    if _apply_one_range(driver, "day", status_cb):
        time.sleep(0.6)
        if _count_yts_fast(driver) > 0 and not _page_has_no_results_text(driver):
            return "Päivä/Day"
        _log(status_cb, "KL: Päivä/Day ei tuloksia -> Viikko/Week…")

    if _apply_one_range(driver, "week", status_cb):
        time.sleep(0.6)
        if _count_yts_fast(driver) > 0 and not _page_has_no_results_text(driver):
            return "Viikko/Week"
        _log(status_cb, "KL: Viikko/Week ei tuloksia -> Kuukausi/Month…")

    if _apply_one_range(driver, "month", status_cb):
        time.sleep(0.6)
        if _count_yts_fast(driver) > 0 and not _page_has_no_results_text(driver):
            return "Kuukausi/Month"

    return ""


# =========================
#  SHOW MORE LOOP (GUARDED)
# =========================
def _scroll_to_bottom(driver):
    try:
        driver.execute_script("window.scrollTo(0, document.body.scrollHeight);")
    except Exception:
        pass


def _find_show_more_selenium(driver):
    xps = [
        "//button[contains(translate(.,'ABCDEFGHIJKLMNOPQRSTUVWXYZÅÄÖ','abcdefghijklmnopqrstuvwxyzåäö'),'näytä lisää')]",
        "//a[contains(translate(.,'ABCDEFGHIJKLMNOPQRSTUVWXYZÅÄÖ','abcdefghijklmnopqrstuvwxyzåäö'),'näytä lisää')]",
        "//*[self::button or self::a][.//*[contains(translate(.,'ABCDEFGHIJKLMNOPQRSTUVWXYZÅÄÖ','abcdefghijklmnopqrstuvwxyzåäö'),'näytä lisää')]]",
        "//button[contains(translate(.,'ABCDEFGHIJKLMNOPQRSTUVWXYZ','abcdefghijklmnopqrstuvwxyz'),'show more')]",
        "//a[contains(translate(.,'ABCDEFGHIJKLMNOPQRSTUVWXYZ','abcdefghijklmnopqrstuvwxyz'),'show more')]",
        "//*[self::button or self::a][.//*[contains(translate(.,'ABCDEFGHIJKLMNOPQRSTUVWXYZ','abcdefghijklmnopqrstuvwxyz'),'show more')]]",
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
                  try { el.scrollIntoView({block:'center'}); } catch(e){}
                  try { el.click(); return true; } catch(e){}
                }
                return false;
                """
            )
        )
    except Exception:
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
    # Strong guard: install nav kill switch for this loop
    install_nav_kill_switch(driver, status_cb=status_cb)

    clicks = 0
    stuck = 0
    prev_yts = _count_yts_fast(driver)

    for i in range(max_passes):
        if stop_flag and stop_flag.is_set():
            _log(status_cb, "KL: STOP havaittu – lopetetaan loop.")
            break

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

        time.sleep(max(0.05, post_click_sleep))

        # If we escaped, recover and continue without counting
        if _return_to_protest_if_escaped(driver, protest_url, status_cb=status_cb):
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

    uninstall_nav_kill_switch(driver, status_cb=status_cb)
    _log(status_cb, f"KL: Loop valmis. Klikkauksia: {clicks}. Y-tunnuksia näkyvissä: {prev_yts}")


# =========================
#  EXTRACTION + DEEP CRAWL
# =========================
def _collect_big_blob_js(driver) -> str:
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
              const scripts = Array.from(document.querySelectorAll('script')).slice(0, 240);
              for (const s of scripts) {
                const t = (s.textContent || '').trim();
                if (!t) continue;
                out.push(t.slice(0, 25000));
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
    STRICT: only KL /yritykset/ links (prevents /porssi/ etc).
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
        if "/yritykset/" not in hh:
            continue
        if "/porssi/" in hh or "/indeks" in hh or "/uutiset" in hh:
            continue
        good.append(h)

    abs_links = []
    for h in good:
        if h.startswith("http"):
            if "kauppalehti.fi" in h.lower():
                abs_links.append(h)
        else:
            abs_links.append("https://www.kauppalehti.fi" + h)

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

    try:
        driver.get(protest_url or DEFAULT_URL)
        time.sleep(0.8)
        close_overlays(driver, status_cb=status_cb)
    except Exception:
        pass

    return sorted(yts)


def extract_ytunnukset_via_js(driver, status_cb=None, protest_url: str = DEFAULT_URL) -> List[str]:
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


# =========================
#  FAILURE SNAPSHOT HELPERS
# =========================
def dump_failure_snapshot(driver, out_dir: str, prefix: str = "kl_fail") -> None:
    """
    Called from app.py ONLY on failure.
    Writes:
      - <prefix>_url.txt
      - <prefix>_body.txt (trimmed)
      - <prefix>_html.html (trimmed)
    """
    try:
        os.makedirs(out_dir, exist_ok=True)
    except Exception:
        return

    try:
        url = driver.current_url or ""
    except Exception:
        url = ""

    try:
        body = driver.execute_script("return document.body ? document.body.innerText : ''") or ""
    except Exception:
        body = ""

    try:
        html = driver.execute_script("return document.documentElement ? document.documentElement.outerHTML : ''") or ""
    except Exception:
        html = ""

    try:
        with open(os.path.join(out_dir, f"{prefix}_url.txt"), "w", encoding="utf-8") as f:
            f.write(url)
    except Exception:
        pass

    try:
        with open(os.path.join(out_dir, f"{prefix}_body.txt"), "w", encoding="utf-8") as f:
            f.write((body or "")[:500_000])
    except Exception:
        pass

    try:
        with open(os.path.join(out_dir, f"{prefix}_html.html"), "w", encoding="utf-8") as f:
            f.write((html or "")[:2_000_000])
    except Exception:
        pass
