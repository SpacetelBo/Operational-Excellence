# Spacetel – Operational Excellence dashboard

Static dashboard (`public/index.html`) + one Netlify Function (`netlify/functions/api.mjs`) + Netlify Blobs for storage. No build step.

| Who | Can do |
|---|---|
| **Anyone with the link (viewers)** | See the latest published report, use search / filters, tab switching, copy image, PNG, Excel export |
| **Signed-in user** | Everything above, plus upload the workbook (publishes it for everyone), view / open / delete the **Log**, and **change their own password** |

There are **no built-in or shared accounts**. Every user creates their own account (see below). The Log's **Updated By** column shows the user's **name** only.

## Creating an account
1. Open the site → **🔒 Login** → **Create a new account**.
2. Enter **your name** and your **username**: 2 letters + 4 numbers + `@cricketlogin.com`. Anything else is rejected.
3. Click **Send OTP**. A 6-digit code is emailed **only to `backoffice@spacetelllc.com`**, together with the name and username that asked. The requester never sees the code; the back office gives it to the person after checking who is asking.
4. The person enters the code and chooses a password (at least 8 characters) → **Create account**. They are signed in straight away.
5. Next time: **Login** → username + password. **Change password** is under the same button once signed in.

Rules enforced on the server:
- The code is valid 10 minutes, works once, allows 5 wrong tries (parallel guesses included), and only one code per username per minute.
- 5 codes per visitor per hour, 30 per hour overall (protects the back-office mailbox). The code is stored only as a hash.
- **Names are unique** (ignoring capitals, spaces and dots), so nobody can appear in the Log as someone else. If a name is taken the person adds a surname or initial.
- A username that exists can never be re-created or overwritten.
- Every account has the same rights (publish, see and delete Log entries). The OTP is the approval step, so the back office should give a code only to people who are allowed to publish.

## One-time setup in Netlify

### 1. Environment variables (Site configuration → Environment variables)
| Variable | Value |
|---|---|
| `SESSION_SECRET` | a long random string (40+ characters). Keep it the same forever; changing it signs everyone out |
| `RESEND_API_KEY` | API key from [Resend](https://resend.com) (see below). Needed for the OTP email |
| `OTP_FROM` | *optional* sender, only after you verify your domain in Resend, e.g. `Operational Excellence <noreply@spacetelllc.com>` |
| `SIGNUP_DISABLED` | *optional*, set to `1` to stop new accounts (existing accounts keep working) |

No passwords are stored in Netlify or in the code. (Older versions used `LOKESH_PASSWORD`, `ANAND_PASSWORD`, etc. and `MASTER_PASSWORD`; they are no longer used and can be deleted.)

### 2. Email through Resend
Netlify cannot send email by itself.
1. Create a free Resend account **using `backoffice@spacetelllc.com`** and create an API key.
2. Add it in Netlify as `RESEND_API_KEY`, then redeploy.
3. Without more setup, Resend sends from its test address, which as far as I know only delivers to the email the Resend account was created with, which is why step 1 matters. To send from your own address, verify `spacetelllc.com` in Resend and set `OTP_FROM`. (Check Resend's current documentation for its sender rules.)

### 3. Deploy (GitHub → Netlify)
Publish directory `public`, functions directory `netlify/functions`, no build command (all read from `netlify.toml`). Every change pushed to GitHub redeploys. On the free plan each deploy uses credits, so combine changes into **one** upload.

## Daily use
- Log in and upload the day's Excel workbook. It is processed in the browser, then published. Everyone who opens or refreshes the site sees it.
- Each publish is saved to the **Log** tab under your name (the latest 100 are kept). Re-uploading identical data updates the report but is not logged twice.
- Sessions last 12 hours or until the tab is closed / Log out. After changing your password, other browsers signed in as you are signed out.

## Removing a user, or a forgotten password
In the site's Netlify Blobs store `operational-excellence`, delete the key `users/<their username>`. It takes effect immediately, even for a browser still signed in, and their name becomes free for reuse after about 2 minutes. (Use the Blobs page of the site, or `netlify blobs:delete operational-excellence users/<username>`; check `netlify blobs:delete --help`.) A person with a forgotten password simply creates the account again with a new OTP.

## Security notes
- Passwords: salted scrypt, verified on the server only; minimum 8 characters.
- Sessions: signed with `SESSION_SECRET`; the account is re-checked on every request. (If `SESSION_SECRET` is missing a random one is generated once and kept in Blobs, but please set it.)
- Brute-force protection is stored in Blobs, so it applies across all function instances: 8 wrong passwords per visitor per account, 30 per account overall, per 15 minutes; the change-password form is limited the same way.
- Reports are visible to anyone with the URL. For a restricted audience use Netlify's site password / SSO (paid plans) or keep the URL private.
- Libraries (SheetJS, ExcelJS, html2canvas) are bundled in `public/vendor`. Only the Inter font is loaded from Google Fonts.

## Local testing
```bash
npm i -g netlify-cli
npm install
# local .env (git-ignored): SESSION_SECRET, RESEND_API_KEY
netlify dev        # http://localhost:8888
```

## Test plan after deploying
1. Viewer (private window): report with filters shows; no upload box, no Log tab; `/api/logs` answers `{"error":"Login required."}`.
2. Login box shows only **Username** and **Password**, with no name list and no example.
3. **Create a new account** with a test name and username → Send OTP → the back-office inbox (and spam) gets the code, with the name and username in the email.
4. Enter a wrong code (error), then the right code and a password → signed in under your name.
5. Upload a workbook → Log shows your **name** under Updated By.
6. Try the same name for a second username → refused ("already used").
7. A wrong username format and a second OTP request within a minute are refused.
8. Change password: wrong current (error), short new (error), mismatch (error), valid → still signed in. Log out → old password fails, new one works.
9. Trigger a deploy, then log in with the new password (must still work).
10. Delete the test user (see "Removing a user").
