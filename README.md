# Finnish Business Email Finder

Sellable Windows EXE that:
- Reads Y-tunnus from:
  - Kauppalehti Protestilista (PLAY mode; requires user login in Chrome)
  - Paste/Clipboard
  - PDF
- Fetches emails from YTJ for each Y-tunnus
- Exports:
  - results.xlsx (Results + Missing + Summary)
  - results.csv
  - emails.docx

## Build (GitHub Actions)
Push to main → Actions builds Windows EXE artifact:
- FinnishBusinessEmailFinder-win64 / FinnishBusinessEmailFinder.exe

## Using PLAY: Protestilista → YTJ
1) Click **Käynnistä Chrome debug** in the app (or start manually)
2) Login to Kauppalehti in that Chrome window
3) Open protest list URL
4) Click **PLAY: Protestilista → YTJ**

### Manual start (optional)
PowerShell:
```powershell
& "C:\Program Files\Google\Chrome\Application\chrome.exe" `
  --remote-debugging-port=9222 `
  --user-data-dir="C:\ChromeDebug"
