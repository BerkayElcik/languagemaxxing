# Sprachpaar

A shared weekly language-learning tracker for two — Berkay (German) and Yağmur (French).
Live at: https://berkayelcik.github.io/languagemaxxing/

- `index.html` — the website. Works standalone (saves to your browser only) until it's
  connected to a shared Google Sheet — see `apps-script/README.md` for that setup.
- `apps-script/` — the small backend (Google Apps Script) that lets the site read and
  write a shared Google Sheet, so a checkbox ticked on one device shows up on the other.

## Enabling GitHub Pages

Settings → Pages → Source: **Deploy from a branch** → Branch: **main**, folder **/ (root)** → Save.
The site will be live at `https://<your-username>.github.io/<repo-name>/` a minute or two later.
