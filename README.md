# Finnish Business Email Finder

**Modes**
1) **PLAY (Kauppalehti protestilista → YTJ)**  
   - Start Chrome in remote debugging mode (app button).
   - Login manually to Kauppalehti.
   - Open protest list.
   - Click PLAY in the app.
   - App loads all via "Näytä lisää", extracts Y-tunnus, fetches emails from YTJ.

2) **Paste/Clipboard → YTJ**
   - Paste page content / list.
   - App extracts:
     - direct emails
     - Y-tunnus
     - if no Y-tunnus: optional name fallback (YTJ search by company name).
   - Fetch emails from YTJ (FAST parallel requests + Selenium fallback).

3) **PDF → YTJ**
   - Extract Y-tunnus from PDF and fetch emails.

**Output (created only when run completes)**
FinnishBusinessEmailFinder/YYYY-MM-DD/run_HH-MM-SS/
- results.xlsx (Results + Missing + Summary)
- results.csv
- emails.docx

## Chrome debug attach (PLAY mode)
The app can start Chrome like:
chrome.exe --remote-debugging-port=9222 --user-data-dir="...\\ChromeDebugProfile"

You must:
1) Login to Kauppalehti in that Chrome
2) Open https://www.kauppalehti.fi/yritykset/protestilista
3) Press PLAY

## Build EXE
### Local
pip install -r requirements.txt
pyinstaller --onefile --windowed --name FinnishBusinessEmailFinder app.py

### GitHub Actions
Push to main -> workflow builds Windows EXE artifact.
