# Connecting the site to a Google Sheet

This makes checking a box on the website (from either of you, on any device) show up for
the other person too — the Sheet is the shared database both of you write through.

## 1. Create the Sheet

Go to [sheets.google.com](https://sheets.google.com) and create a new blank spreadsheet,
e.g. named "Sprachpaar Data". You don't need to add any columns — the script creates its
own "Completions" tab automatically the first time it runs.

## 2. Add the script

1. In that Sheet, open **Extensions → Apps Script**.
2. Delete the placeholder `function myFunction() {}` code.
3. Paste in the contents of [`Code.gs`](./Code.gs) from this folder.
4. Click the disk icon (or Ctrl+S) to save. Name the project anything, e.g. "Sprachpaar sync".

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

Commit and push that change (or ask Claude to do it) — once it's live, both of you will
be reading and writing the same Sheet.

## Notes

- Every checkbox toggle writes one row to the "Completions" tab (`Key`, `Done`,
  `UpdatedAt`). You can open that tab any time to see the raw data, or build your own
  formulas/charts off of it.
- If you ever redeploy the script (not just edit — an actual **New deployment**), you'll
  get a new URL and need to update `SYNC_URL` again. Editing the code and using
  **Deploy → Manage deployments → Edit → Deploy** keeps the same URL.
- This is a lightweight personal setup (no login system) — anyone with the URL can read
  or write to it. Don't share the URL outside the two of you.
