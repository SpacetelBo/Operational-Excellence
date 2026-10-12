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
3. In Netlify → **Site configuration → Environment variables**, set these four initial passwords:
   - `LOKESH_PASSWORD`
   - `ANAND_PASSWORD`
   - `UPENDER_PASSWORD`
   - `VENU_PASSWORD`

   Also set `SESSION_SECRET` to a long, random, stable secret. These four password variables initialize the credential store on first login only. After that, changed passwords are saved as salted scrypt hashes in Netlify Blobs. Do not commit actual passwords to GitHub.
4. Trigger a deploy, then open the site URL → click **🔒 Master login** (bottom-right) → log in and upload the workbook.

## Daily use
- Master logs in, uploads the day's Excel workbook. It is processed in the browser, then published. Everyone who opens (or refreshes) the site sees it.
- Each upload is also saved to the **Log** tab (the latest 100 are kept). Re-uploading identical data updates the report but is not logged twice.
- Master sessions last 12 hours or until the tab is closed / Log out.

## Changing a password
After logging in, click the bottom-right **Master · Log out** button. In the login panel, click **Change password**. Enter the current password and the new password twice. New passwords must be at least 8 characters. Changes take effect immediately and do not require a redeploy. Each user can change only the password for the account they are currently logged into.

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


## Individual editor passwords and audit log

Configure these four environment variables in Netlify site settings (Site configuration → Environment variables), then redeploy:

- `LOKESH_PASSWORD` — password for Lokesh
- `ANAND_PASSWORD` — password for Anand
- `UPENDER_PASSWORD` — password for Upender
- `VENU_PASSWORD` — password for Venu

Use four different, strong passwords and share each only with its owner. Do not put passwords in the HTML or commit them to GitHub. When a password is used to publish changed data, the Log tab records the authenticated user's name in the **Updated By** column. Existing log entries created before this feature may show **Legacy / unknown**.


## In-dashboard password changes

Configure `LOKESH_PASSWORD`, `ANAND_PASSWORD`, `UPENDER_PASSWORD`, and `VENU_PASSWORD` in Netlify before the first login. Set a long, stable `SESSION_SECRET` in Netlify. These initial passwords bootstrap the account store on first login; after that, credentials are saved as salted scrypt hashes in Netlify Blobs. Do not commit actual passwords to GitHub.

After logging in, click the bottom-right **Master · Log out** button, then choose **Change password**. Enter the current password and the new password twice. New passwords must be at least 8 characters. Changes take effect immediately and do not require redeployment. Each user can change only the password for the account they are logged into.
