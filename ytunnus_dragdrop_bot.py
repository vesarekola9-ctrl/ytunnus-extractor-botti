def get_company_rows_tbody(driver):
    """
    Palauttaa vain varsinaiset yritysrivit (tbody/tr), ei headeria.
    Jättää detail-rivit pois.
    """
    rows = []
    candidates = driver.find_elements(By.XPATH, "//table//tbody//tr")
    for r in candidates:
        try:
            if not r.is_displayed():
                continue

            txt = (r.text or "")
            # detail-rivissä näkyy "Y-TUNNUS" -> skip
            if "Y-TUNNUS" in txt:
                continue

            # yritysrivissä on yrityksen nimi linkkinä 1. sarakkeessa
            links = r.find_elements(By.XPATH, ".//td[1]//a[contains(@href,'/yritykset/') and normalize-space(.)!='']")
            if not links:
                continue

            rows.append(r)
        except Exception:
            continue
    return rows


def open_row_by_summa_cell(row):
    """
    Klikkaa SUMMA-solua (3. sarake) -> avaa detail-rivin alle.
    Tämä EI klikkaa yritysnimeä (td[1]).
    """
    try:
        cell = row.find_element(By.XPATH, ".//td[3]")
        return cell
    except Exception:
        return None


def extract_yt_from_detail_row(row):
    """
    Detail-rivi on yleensä heti tämän rivin jälkeen: following-sibling::tr[1]
    Siitä etsitään Y-TUNNUS ja poimitaan y-tunnus.
    """
    try:
        detail = row.find_element(By.XPATH, "following-sibling::tr[1]")
        txt = detail.text or ""
        if "Y-TUNNUS" not in txt:
            return ""
        found = YT_RE.findall(txt)
        for m in found:
            n = normalize_yt(m)
            if n:
                return n
    except Exception:
        pass
    return ""


def collect_yts_from_kauppalehti(driver, status_cb, log_cb):
    wait = WebDriverWait(driver, 25)
    wait.until(EC.presence_of_element_located((By.TAG_NAME, "body")))
    try_accept_cookies(driver)

    collected = set()
    processed = set()  # rivin "fingerprint"

    def row_fingerprint(r):
        # fingerprint: yritysnimi + sijainti + summa (EI päivämäärää)
        try:
            name = r.find_element(By.XPATH, ".//td[1]//a[contains(@href,'/yritykset/') and normalize-space(.)!='']").text.strip()
        except Exception:
            name = ""
        try:
            location = r.find_element(By.XPATH, ".//td[2]").text.strip()
        except Exception:
            location = ""
        try:
            amount = r.find_element(By.XPATH, ".//td[3]").text.strip()
        except Exception:
            amount = ""
        return f"{name}|{location}|{amount}"

    while True:
        rows = get_company_rows_tbody(driver)
        if not rows:
            status_cb("Kauppalehti: en löydä yritysrivejä (tbody). Oletko kirjautuneena ja lista näkyy?")
            break

        status_cb(f"Kauppalehti: näkyviä rivejä {len(rows)} | kerätty {len(collected)}")

        new_in_pass = 0

        for row in rows:
            try:
                fp = row_fingerprint(row)
                if fp in processed:
                    continue
                processed.add(fp)

                # 1) klikkaa SUMMA-solua (EI yritysnimeä)
                click_cell = open_row_by_summa_cell(row)
                if click_cell is None:
                    continue

                driver.execute_script("arguments[0].scrollIntoView({block:'center'});", click_cell)
                time.sleep(0.03)

                # click (JS fallback)
                try:
                    click_cell.click()
                except Exception:
                    driver.execute_script("arguments[0].click();", click_cell)

                # 2) odota että detail-rivi tulee ja poimi Y-TUNNUS
                yt = ""
                for _ in range(25):  # max ~2.5s
                    yt = extract_yt_from_detail_row(row)
                    if yt:
                        break
                    time.sleep(0.1)

                if yt and yt not in collected:
                    collected.add(yt)
                    new_in_pass += 1
                    log_cb(f"+ {yt} (yht {len(collected)})")

                # 3) sulje takaisin klikkaamalla samaa SUMMA-solua uudelleen
                try:
                    click_cell.click()
                except Exception:
                    try:
                        driver.execute_script("arguments[0].click();", click_cell)
                    except Exception:
                        pass

                time.sleep(0.02)

            except StaleElementReferenceException:
                continue
            except Exception:
                continue

        # 4) Näytä lisää ja jatka
        if click_nayta_lisaa(driver):
            status_cb("Kauppalehti: Näytä lisää…")
            time.sleep(1.2)
            try:
                driver.execute_script("window.scrollTo(0, document.body.scrollHeight);")
            except Exception:
                pass
            time.sleep(0.8)
            continue

        # 5) jos ei enää uusia eikä Näytä lisää -> loppu
        if new_in_pass == 0:
            status_cb("Kauppalehti: ei uusia + ei Näytä lisää -> valmis.")
            break

    return sorted(collected)
