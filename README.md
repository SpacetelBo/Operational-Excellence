# Spacetel – Operational Excellence dashboard

Static dashboard (`public/index.html`) + one Netlify Function (`netlify/functions/api.mjs`) + Netlify Blobs for storage. No build step.

| Who | Can do |
|---|---|
| **Anyone with the link (viewers)** | See the latest published report, use search / filters, tab switching, copy image, PNG, Excel export |
| **Master login** | Everything above, plus upload the workbook (publishes it for everyone), view / open / delete the **Log** |

The master password is checked **on the server**. Upload and Log endpoints reject requests without a valid master session, so hiding buttons is not the only protection.

## Deploy (GitHub → Netlify)

1. Create a **private** GitHub repo and push this folder:
   ```bash
   cd spacetel-operational-excellence
   git init && git add . && git commit -m "Operational Excellence dashboard"
   git branch -M main
   git remote add origin https://github.com/<you>/<repo>.git
   git push -u origin main
   ```
2. In Netlify: **Add new site → Import an existing project → GitHub →** pick the repo.
   Build settings are read from `netlify.toml` (publish `public`, functions `netlify/functions`). Leave build command empty. Deploy.
3. **Set the password as an environment variable** (recommended, so it is not only in the code):
   Site configuration → Environment variables →
   - `MASTER_PASSWORD` = `spacetel123` (or a stronger one)
   - `SESSION_SECRET` = any long random string (optional, signs login sessions)

   Then **Deploys → Trigger deploy → Deploy site**.
   If `MASTER_PASSWORD` is not set, the function falls back to `spacetel123`.
4. Open the site URL → click **🔒 Master login** (bottom-right) → upload the workbook.

## Daily use
- Master logs in, uploads the day's Excel workbook. It is processed in the browser, then published. Everyone who opens (or refreshes) the site sees it.
- Each upload is also saved to the **Log** tab (the latest 100 are kept). Re-uploading identical data updates the report but is not logged twice.
- Master sessions last 12 hours or until the tab is closed / Log out.

## Changing the password
Change `MASTER_PASSWORD` in Netlify → redeploy. All existing master sessions stop working (the session key is derived from it unless `SESSION_SECRET` is set).

## Local testing
```bash
npm i -g netlify-cli
npm install
netlify dev        # http://localhost:8888
```

## Notes
- Libraries (SheetJS, ExcelJS, html2canvas) are bundled in `public/vendor`, so the page doesn't depend on a CDN. Only the Inter font is loaded from Google Fonts (falls back to system fonts if blocked).
- Reports are visible to anyone who has the URL. For a restricted audience, use Netlify's site password / SSO protection (paid plans) or keep the URL private.
- After 8 wrong passwords from one IP within 10 minutes, login is temporarily blocked.
