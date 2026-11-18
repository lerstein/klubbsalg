# Klubbsalg

Web app (PWA) for registering club merchandise sales at stands: sellers pick seller, stand, product, size,
quantity and payment; sales sync to a Google Sheet. Admin actions (delete, prices, stock, remove seller/product/stand)
need a password.
Live at **lerstein.github.io/klubbsalg** (GitHub Pages from this repo).

**This repo is public.** Never put the `/exec` URL, the sheet ID/URL, the admin password or sales data here.

## Layout

- `index.html` – the whole app (HTML, CSS, JS in one file). Texts in Norwegian.
  Local state in `localStorage` under `klubbsalg_v3*` (`KEY` in the script).
- The backend `/exec` URL is **not** in the code: each device pastes it once under «Google Sheet»
  (stored in `klubbsalg_v3_sheet`). The app works offline and pushes queued sales when it has a connection.
- `sw.js` – service worker (offline/PWA install). Bump its cache name when `index.html` changes.
- `velkommen.mp3` – welcome sound.
- `backend/Code.gs` – copy of the Apps Script backend (see below).

## Backend: Apps Script

| | |
|---|---|
| Project | «Klubbsalg» in Apps Script (owner Magnus), bound to the Klubbsalg sheet (`getActiveSpreadsheet`) |
| Deployment | «Klubbsalg v1», web app, execute as Me, access Anyone |
| Secrets | `ADMIN_PASSWORD` (and optional `EDIT_WINDOW_MINUTES`) in **Script properties** – never in code |
| Sheets | `Salg`, `Produkter`, `Selgere`, `Stand`, `Lager` (created on first use) |

`backend/Code.gs` holds no IDs or secrets, so it is fine in this public repo. Keep it that way.

### The Apps Script editor is the live source

This repo is a copy, not a deploy pipeline (no clasp).
Magnus or Claude (via Claude in Chrome) may edit directly in the Apps Script editor.

Workflow for a backend change:
1. **Before** editing here, check that `backend/Code.gs` matches the editor (someone may have changed it there).
   If it differs, copy the editor version into this repo first and commit it as "Sync from Apps Script".
2. Make the change in this repo, commit.
3. Paste it into the Apps Script editor (Monaco: `monaco.editor.getModels()`), save.
4. **Deploy → Manage deployments → ✏️ Edit → Version: New version → Deploy.**
   Never "New deployment" – that changes the `/exec` URL, and every seller's phone would have to be set up again.
5. Verify in the app on lerstein.github.io/klubbsalg.

### Frontend change

1. Edit `index.html`, check the script parses (e.g. `node -e "new Function(...)"` on the `<script>` block).
2. If it needs a new backend action, add and deploy it first – the frontend must not call an action that isn't deployed.
3. Bump the cache name in `sw.js`, commit and push to `main`; GitHub Pages updates within a minute or two.

## Rules to keep

- Keep the app simple – sellers use it on their phones at a stand, often in a hurry.
- Deleting is soft (timestamp in «Slettet»); sellers may edit/delete their own sales for 30 minutes without password,
  checked on the server. Do not move password checks into the app.
- Don't change the `localStorage` keys without migrating existing data – devices may hold unsynced sales.
