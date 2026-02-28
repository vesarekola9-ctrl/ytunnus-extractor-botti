# LeadForge FI (LeadForgeFI.exe)

## What it does
- PDF -> extract Y-tunnus -> fetch email from YTJ
- Clipboard -> extract company names (+location) -> YTJ -> Y-tunnus -> email
- If YTJ has no email, optional website fallback

## Output
Created only when saving results:
- results.xlsx (Results, Found Only, Not Found)
- results.csv
- emails.docx

## License
- DEMO: max 20 companies per run
- PRO: unlimited
Save a PRO key in the app (LF-XXXX-XXXX-XXXX-CC)

## Build
GitHub Actions builds a Windows EXE artifact: LeadForgeFI.exe
