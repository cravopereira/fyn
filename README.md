# Fyn

A personal finance app for the phone that works on top of a Google Sheet. The sheet stays the database; Fyn reads it and writes to it through a small Apps Script web app.

## Install on iPhone
1. Open https://cravopereira.github.io/fyn/ in Safari.
2. Share › Add to Home Screen.
3. Open Fyn and paste your Web app URL and key (see below), or try it with sample data.

## Connect your sheet (once, on a computer)
1. In the sheet: Extensions › Apps Script.
2. Replace the "Close month" file with `backend/CloseMonth.gs`.
3. Add a file called `Api` and paste `backend/Api.gs`. Save.
4. Run `createAppKey` and copy the key it shows.
5. Deploy › New deployment › Web app. Execute as: Me. Who has access: Anyone. Copy the URL.

The key is the only thing that protects the sheet. It lives only on your phone. Running `createAppKey` again makes a new key and blocks the old one.

No personal data is stored in this repository. "Try with sample data" uses invented numbers.
