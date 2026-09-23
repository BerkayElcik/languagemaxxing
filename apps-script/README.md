# Connecting the site to a Google Sheet

This makes checking a box on the website (from either of you, on any device) show up for
the other person too, and gates sign-in, checkboxes, and each person's sticker board to
just the two of you — the Sheet is the shared database both of you write through.

## 1. Create the Sheet

Go to [sheets.google.com](https://sheets.google.com) and create a new blank spreadsheet,
e.g. named "Sprachpaar Data". You don't need to add any columns — the script creates its
own tabs automatically the first time each is used.

## 2. Add the script

1. In that Sheet, open **Extensions → Apps Script**.
2. Delete the placeholder `function myFunction() {}` code.
3. Paste in the contents of [`Code.gs`](./Code.gs) from this folder.
4. Click the disk icon (or Ctrl+S) to save. Name the project anything, e.g. "Sprachpaar sync".
5. Near the top of the file, check the `ALLOWED` map has the right two email addresses —
   only these can sign in, check boxes, or edit a board:
   ```js
   var ALLOWED = {
     'kemalberkayelcik@hotmail.com': 'berkay',
     'yagmurkaran6@gmail.com': 'yagmur',
   };
   ```

## 3. Deploy it as a Web App

1. Click **Deploy → New deployment**.
2. Click the gear icon next to "Select type" and choose **Web app**.
3. Set **Execute as: Me**, and **Who has access: Anyone**.
4. Click **Deploy**.
5. Google will ask you to authorize the script (it's your own script acting on your own
   Sheet) — click through the consent screen (you may see an "unverified app" warning
   since this is a personal script; click **Advanced → Go to \[project name\] (unsafe)**
   to proceed — this is expected for scripts you write yourself).
6. Copy the **Web app URL** it gives you (ends in `/exec`).

## 4. Wire it into the site

Open `index.html` at the repo root, find this line near the top of the `<script>` block:

```js
var SYNC_URL = '';
```

Paste your Web App URL between the quotes, e.g.:

```js
var SYNC_URL = 'https://script.google.com/macros/s/AKfycb.../exec';
```

Commit and push that change (or ask Claude to do it). Once it's live:

- **Sign in** (top-right "Sign in" button) sends a 4-digit code to your email — enter it
  to sign in. Only checking/unchecking boxes and editing your own sticker board require
  being signed in; viewing the tracker doesn't.
- **Boards** — each side has "+ Sticker" (paste an image URL), "+ Text" (draggable,
  rotatable, recolorable/re-fontable note — drag the item to move it, the small ⟲ handle
  to rotate, click the text to edit it in place), and "+ Music" (paste a direct audio file
  URL). Berkay can only add/remove on the left board, Yağmur only on the right — enforced
  by the backend, not just hidden in the UI.
- Yağmur's board is seeded with one track the first time the Sheet runs: Dvořák's
  *Serenade for Strings, Op. 22 — II. Tempo di Valse* (CC BY-SA 4.0, Wikimedia Commons).

## Notes

- Tabs created automatically: **Completions** (checkbox state), **LoginCodes** (temporary
  sign-in codes, 10-minute expiry), **Sessions** (signed-in tokens), **BoardItems**
  (stickers/text/music per side). You can open any of them to see the raw data.
- Emails send via `MailApp`, which runs under whichever Google account you deployed the
  script with, and shares that account's daily sending quota (100/day on a free account —
  far more than two people signing in occasionally will ever need).
- If you ever redeploy the script (not just edit — an actual **New deployment**), you'll
  get a new URL and need to update `SYNC_URL` again. Editing the code and using
  **Deploy → Manage deployments → Edit → Deploy** keeps the same URL.
- Reading the tracker (viewing progress) stays open to anyone with the link; only writes
  (checkboxes, board edits, sign-in) are restricted to the two allowed emails.
