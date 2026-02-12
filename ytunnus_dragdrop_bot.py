import os
import re
import threading
import tkinter as tk
from tkinter import messagebox

from tkinterdnd2 import DND_FILES, TkinterDnD

import PyPDF2
from docx import Document
import openpyxl

# Selenium + driver manager
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from webdriver_manager.chrome import ChromeDriverManager


# -----------------------------
# Y-tunnus logiikka
# -----------------------------
def fix_ytunnus(yt):
    yt = yt.strip()

    if re.match(r"^\d{7}-\d$", yt):
        return yt

    if re.match(r"^\d{8}$", yt):
        return yt[:7] + "-" + yt[7]

    return None


def extract_ytunnukset_from_pdf(pdf_path):
    ytunnukset = set()

    with open(pdf_path, "rb") as f:
        reader = PyPDF2.PdfReader(f)
        for page in reader.pages:
            text = page.extract_text()
            if not text:
                continue

            matches = re.findall(r"\b\d{7}-\d\b|\b\d{8}\b", text)

            for m in matches:
                fixed = fix_ytunnus(m)
                if fixed:
                    ytunnukset.add(fixed)

    return sorted(list(ytunnukset))


# -----------------------------
# Tallennus Word + Excel
# -----------------------------
def save_ytunnukset_word(ytunnukset, filename="ytunnukset.docx"):
    doc = Document()
    doc.add_heading("Y-tunnukset", level=1)

    for yt in ytunnukset:
        doc.add_paragraph(yt)

    doc.save(filename)


def save_ytunnukset_excel(ytunnukset, filename="ytunnukset.xlsx"):
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Y-tunnukset"

    ws["A1"] = "Y-tunnus"

    for i, yt in enumerate(ytunnukset, start=2):
        ws[f"A{i}"] = yt

    wb.save(filename)


def save_emails_word(results, filename="virre_sahkopostit.docx"):
    doc = Document()
    doc.add_heading("Virre - Sähköpostit", level=1)

    for yt, email in results:
        doc.add_paragraph(f"{yt}  ->  {email}")

    doc.save(filename)


def save_emails_excel(results, filename="virre_sahkopostit.xlsx"):
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Sähköpostit"

    ws["A1"] = "Y-tunnus"
    ws["B1"] = "Sähköposti"

    for i, (yt, email) in enumerate(results, start=2):
        ws[f"A{i}"] = yt
        ws[f"B{i}"] = email

    wb.save(filename)


# -----------------------------
# Virre sähköpostien haku
# -----------------------------
def fetch_email_from_virre(driver, ytunnus, log_func):
    driver.get("https://virre.prh.fi/novus/home")

    wait = WebDriverWait(driver, 20)

    try:
        log_func("Etsitään hakukenttä...")

        # Tämä hakee ekana tekstikentän (Virre käyttää tätä yleensä)
        search_input = wait.until(
            EC.presence_of_element_located((By.CSS_SELECTOR, "input[type='text']"))
        )

        search_input.clear()
        search_input.send_keys(ytunnus)

        log_func("Painetaan Hae...")

        # Etsitään nappi joka sisältää "Hae"
        buttons = driver.find_elements(By.TAG_NAME, "button")
        clicked = False

        for b in buttons:
            if "hae" in b.text.strip().lower():
                b.click()
                clicked = True
                break

        if not clicked:
            return "EI HAKUNAPPIA"

        # Odotetaan että sivu muuttuu / tulos latautuu
        wait.until(EC.presence_of_element_located((By.TAG_NAME, "body")))

        page_source = driver.page_source

        email_match = re.search(r"[a-zA-Z0-9_.+-]+@[a-zA-Z0-9-]+\.[a-zA-Z0-9-.]+", page_source)
        if email_match:
            return email_match.group(0)

        return "EI SÄHKÖPOSTIA"

    except Exception as e:
        return f"VIRHE: {str(e)}"


def fetch_all_emails(ytunnukset, log_func):
    log_func("Käynnistetään Chrome (Selenium)...")

    try:
        options = webdriver.ChromeOptions()

        # Näkyvä Chrome (debug)
        options.add_argument("--start-maximized")

        # tärkeä exe:ssä
        options.add_argument("--disable-gpu")

        # tämä auttaa jos ongelmia profiilin kanssa
        options.add_argument("--disable-dev-shm-usage")

        driver_pat_
