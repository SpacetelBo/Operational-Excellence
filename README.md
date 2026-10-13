# Spacetel – Operational Excellence dashboard

Static dashboard (`public/index.html`) + one Netlify Function (`netlify/functions/api.mjs`) + Netlify Blobs for storage. No build step.

| Who | Can do |
|---|---|
| **Anyone with the link (viewers)** | See the latest published report, use search / filters, tab switching, copy image, PNG, Excel export |
| **Signed-in user** (Lokesh, Anand, Upender, Venu, or anyone who created a Cricket-login account) | Everything above, plus upload the workbook (publishes it for everyone), view / open / delete the **Log**, and **change their own password** |

Every user has their own account and password. Who you are is decided by the server from your login; the browser never tells the server a name. Upload, Log and password endpoints reject requests without a valid session, so hiding buttons is not the only protection.

## What changed in this version
- **Four individual accounts** replace the single shared master login.
- **Log → Updated By** column: the name of the user who published each upload (older entries show `Legacy / unknown`).
- **Change Password** inside the app (click your name at the bottom right → Change password). No Netlify or GitHub access needed.

Files changed: `public/index.html`, `netlify/functions/api.mjs`, `package.json` (Blobs library `^8` → `^10`), `README.md`. `netlify.toml`, `.gitignore` and `public/vendor/*` are unchanged.

## Creating your own account (Cricket login + OTP)
Anyone with a Cricket login can create their own account:

1. Open the site → **🔒 Login** → **Create a new account (Cricket login)**.
2. Enter the username in exactly this form: **2 letters + 4 digits + `@cricketlogin.com`**, for example `vr6584@cricketlogin.com`. Anything else is rejected.
3. Click **Send OTP**. A 6-digit code is emailed **only to `backoffice@spacetelllc.com`**. The requester never sees it on screen; the back office gives it to the person once they have checked who is asking. (The email shows the username that was requested.)
4. The person enters the OTP, chooses a password (at least 8 characters), and clicks **Create account**. They are signed in straight away.
5. Later they log in with **Login → Cricket login user… →** their username and password, and can use **Change password** like everyone else.

Rules enforced on the server: the code is valid for 10 minutes, works once, allows 5 wrong tries (parallel guesses included), at least 60 s between codes for one username, 5 codes per visitor per hour and 30 per hour overall (so nobody can flood the back-office mailbox). The code is stored only as a hash. A username that already exists can never be re-created or overwritten.

In the Log, **Updated By** shows the part before the `@` in capitals (`vr6584@cricketlogin.com` → `VR6584`).

**Important:** every account, including self-created ones, has the same rights (upload and publish, see and delete Log entries). The OTP is the approval step, so the back office should give a code only to people who are allowed to publish.

### One-time email setup (needed for the OTP email)
Netlify cannot send email by itself, so this uses [Resend](https://resend.com) (free plan is enough for this use).
1. Create a Resend account **using `backoffice@spacetelllc.com`** and create an API key.
2. In Netlify add the environment variable `RESEND_API_KEY` with that key, then redeploy.
3. Without any further setup, emails are sent from Resend's test address, which Resend only allows to deliver to the email the account was created with, which is why step 1 matters. To send from your own address instead, verify the domain `spacetelllc.com` in Resend and also add `OTP_FROM`, for example `Operational Excellence <noreply@spacetelllc.com>`.
4. Optional: `SIGNUP_DISABLED` = `1` switches self-service sign-up off (existing accounts keep working).

(Check Resend's current documentation for its sender and domain rules; they are outside this project.)

### Removing a self-created user / forgotten password
Delete that user's record from the Netlify Blobs store `operational-excellence`, key `users/vr6584@cricketlogin.com` (Blobs page of the site, or `netlify blobs:delete operational-excellence users/vr6584@cricketlogin.com`; check `netlify blobs:delete --help`). It takes effect immediately, even for a browser that is still signed in. For a forgotten password the person simply creates the account again with a new OTP.

## Deploy / update (GitHub → Netlify)

### A. First setup, or moving to this version
1. Copy these files over your existing repo (keep the same folder structure), then:
   ```bash
   git add .
   git commit -m "Individual logins, Updated By log column, change password"
   git push
   ```
   Netlify redeploys automatically (publish `public`, functions `netlify/functions`, no build command).
2. In Netlify: **Site configuration → Environment variables** add these (they are **not** in the code or the ZIP):

   | Variable | Value |
   |---|---|
   | `LOKESH_PASSWORD` | Lokesh's initial password |
   | `ANAND_PASSWORD` | Anand's initial password |
   | `UPENDER_PASSWORD` | Upender's initial password |
   | `VENU_PASSWORD` | Venu's initial password |
   | `SESSION_SECRET` | a long random string (e.g. 40+ random characters). Keep it the same forever; changing it signs everyone out |
   | `RESEND_API_KEY` | only needed for self-service sign-up, see "One-time email setup" above |

   Set the variables **before** the first login. Then **Deploys → Trigger deploy → Deploy site** so the function sees them.
3. Open the site → **🔒 Login** → pick your name → enter your initial password. Then use **Change password** to pick your own.

### B. How accounts are stored
- The first time an account logs in, its initial password (from the environment variable) is hashed with salted scrypt and saved in Netlify Blobs (`users/<name>`). Plain passwords are never stored.
- After that, **Blobs is the source of truth**. Redeploying, or changing/removing the environment variables, does **not** change an existing account's password.
- If two people hit a brand-new account at the same moment, only one record is created (conditional "only if new" write).

### C. Forgot a password
There is no email reset. To reset one account: delete that account's record from the site's Blobs store (`operational-excellence`, key `users/<name>`, e.g. `users/anand`) using the Netlify CLI (`netlify blobs:delete operational-excellence users/anand`) or the Blobs page of your site in the Netlify dashboard, make sure the matching `*_PASSWORD` variable holds the temporary password you want, and have the user log in with it, then change it. (Check the exact CLI syntax with `netlify blobs:delete --help`.)

## Daily use
- Log in, upload the day's Excel workbook. It is processed in the browser, then published. Everyone who opens (or refreshes) the site sees it.
- Each publish is saved to the **Log** tab with your name (the latest 100 are kept). Re-uploading identical data updates the report but is not logged twice.
- Sessions last 12 hours or until the tab is closed / Log out. Changing your password keeps *this* browser signed in; any other browser signed in as you is signed out.

## Security notes
- Passwords: salted scrypt, verified on the server only. Minimum 8 characters for new passwords.
- Sessions: signed with `SESSION_SECRET`; the account is re-checked on every request. (If `SESSION_SECRET` is missing, a random secret is generated once and kept in Blobs, but please set it.)
- Brute-force protection is stored in Blobs, so it applies across all function instances: 8 wrong passwords per visitor per account, 30 per account overall, within 15 minutes. Wrong "current password" attempts on the change form are limited the same way.
- Reports are visible to anyone with the URL. For a restricted audience use Netlify's site password / SSO (paid plans) or keep the URL private.
- Libraries (SheetJS, ExcelJS, html2canvas) are bundled in `public/vendor`. Only the Inter font is loaded from Google Fonts.

## Local testing
```bash
npm i -g netlify-cli
npm install
# create a local .env (git-ignored) with LOKESH_PASSWORD, ANAND_PASSWORD, UPENDER_PASSWORD, VENU_PASSWORD, SESSION_SECRET
netlify dev        # http://localhost:8888
```

## Test plan after deploying (about 10 minutes)
1. Log in as each of the four users with their initial password → each sees upload + Log.
2. Wrong password, and another person's password, → "Wrong user or password."; the 9th quick failure shows "Too many attempts".
3. As Lokesh upload a workbook → Log shows Updated By = Lokesh. Repeat as Anand and Venu → names match.
4. Open a log, go back, delete a test log.
5. As Anand: **Change password** with a wrong current password (error), a 5-character new password (error), mismatched confirmation (error), then a valid change → success, still signed in.
6. Log out → old password fails, new password works.
7. **Persistence:** Deploys → Trigger deploy. After it finishes, log in with the *new* password (it must work) and the *original* password must fail.
8. Open the site in a private window (not logged in): the report shows with filters, no upload box, no Log tab. Visit `/api/logs` in that window → `{"error":"Login required."}`.
9. **New account:** Login → Create a new account → enter a test username such as `ab1234@cricketlogin.com` → Send OTP. Check the back-office mailbox (and spam) for the email, enter the code and a password → you are signed in as `AB1234`. Publish a workbook and check the Log shows `AB1234`. Then delete the test record (see "Removing a self-created user").
10. A wrong username format (for example `abc123@cricketlogin.com` or `ab1234@gmail.com`) must be rejected, and a second OTP request within 60 seconds must be refused.
