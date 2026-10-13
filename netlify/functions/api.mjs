// Spacetel Operational Excellence - API
//
// Public  : GET  /api/report            latest published report (anyone can view)
//           POST /api/register/request  { username, name } -> emails a one-time code (OTP) to the back office
//           POST /api/register/complete { username, otp, password } -> creates that account and signs in
//           POST /api/login             { user, password } -> session token
// Signed in:
//           GET  /api/session           who am I (the name comes from the signed session)
//           POST /api/password          { current, next } change MY password
//           POST /api/upload            publish a new report (+ save to the log, recorded under MY name)
//           GET  /api/logs[?ts=...]     list the log / open one saved log
//           DELETE /api/logs?ts=...     delete one log entry
//
// Every account is created by its owner: username = 2 letters + 4 digits + @cricketlogin.com, plus a name.
// The one-time code goes ONLY to the back office, who gives it to the person. There are no built-in accounts
// and no shared passwords. Everything marked "Signed in" is enforced here on the server; the identity comes
// from the signed session token, never from anything the browser sends in the request body.

import { getStore } from "@netlify/blobs";
import crypto from "node:crypto";
import { promisify } from "node:util";

const scrypt = promisify(crypto.scrypt);

// Self-service accounts: username must be exactly 2 letters + 4 digits + @cricketlogin.com.
// The one-time code is emailed ONLY to this address. The requester never sees it; the back office gives it to them.
const OTP_TO = "backoffice@spacetelllc.com";
const REG_ID = /^[a-z]{2}[0-9]{4}@cricketlogin\.com$/;
const OTP_TTL_MS = 10 * 60 * 1000;      // a code is valid for 10 minutes
const OTP_RESEND_MS = 60 * 1000;        // at least 60 s between codes for the same username
const MAX_OTP_TRIES = 5;                // wrong codes allowed per code
const MAX_OTP_REQ_PER_IP_HOUR = 5;      // codes one visitor can trigger per hour
const MAX_OTP_REQ_PER_HOUR = 30;        // codes anyone can trigger per hour (protects the back-office mailbox)
const HOUR_MS = 60 * 60 * 1000;

const TOKEN_TTL_MS = 12 * 60 * 60 * 1000; // sessions last 12 hours
const MAX_BODY_BYTES = 5 * 1024 * 1024;
const MAX_LOGS = 100;
const MIN_PASSWORD = 8;
const MAX_PASSWORD = 200;
const SCRYPT = { N: 16384, r: 8, p: 1, len: 64 };
const THROTTLE_WINDOW_MS = 15 * 60 * 1000;
const MAX_FAILS_PER_IP_USER = 8;  // wrong passwords per visitor per account
const MAX_FAILS_PER_USER = 30;    // wrong passwords per account, from anywhere
const MAX_FAILS_PW_CHANGE = 8;    // wrong "current password" attempts while signed in

const store = () => getStore({ name: "operational-excellence", consistency: "strong" });

const json = (obj, status = 200) =>
  new Response(JSON.stringify(obj), {
    status,
    headers: { "content-type": "application/json; charset=utf-8", "cache-control": "no-store" },
  });

const sleep = (ms) => new Promise((r) => setTimeout(r, ms));
const sha = (s) => crypto.createHash("sha256").update(String(s)).digest("hex");
const b64u = (buf) => Buffer.from(buf).toString("base64url");
const eqBuf = (a, b) => a.length === b.length && crypto.timingSafeEqual(a, b);
const num = (v) => (Number.isFinite(Number(v)) ? Number(v) : 0);

// ---------- password hashing (salted scrypt) ----------
async function hashPassword(pw) {
  const salt = crypto.randomBytes(16);
  const key = await scrypt(pw, salt, SCRYPT.len, { N: SCRYPT.N, r: SCRYPT.r, p: SCRYPT.p, maxmem: 64 * 1024 * 1024 });
  return { algo: "scrypt", N: SCRYPT.N, r: SCRYPT.r, p: SCRYPT.p, salt: salt.toString("base64"), hash: key.toString("base64") };
}
async function verifyPassword(pw, cred) {
  if (!cred || cred.algo !== "scrypt") return false;
  const salt = Buffer.from(cred.salt, "base64");
  const want = Buffer.from(cred.hash, "base64");
  const key = await scrypt(pw, salt, want.length, { N: cred.N, r: cred.r, p: cred.p, maxmem: 64 * 1024 * 1024 });
  return eqBuf(key, want);
}
let DUMMY = null; // so an unknown account name takes as long to reject as a real one
const dummyCheck = async (pw) => {
  DUMMY = DUMMY || (await hashPassword("not-a-real-password"));
  await verifyPassword(pw, DUMMY);
};

// ---------- account records in Netlify Blobs (users/<username>) ----------
const userKey = (id) => "users/" + id;

const loadUser = (s, id) => s.get(userKey(id), { type: "json" });

// ---------- sessions (signed tokens) ----------
let cachedSecret = null;
async function sessionSecret(s) {
  if (process.env.SESSION_SECRET) return process.env.SESSION_SECRET;
  if (cachedSecret) return cachedSecret;
  // Fallback if SESSION_SECRET was not set: a random secret kept in Blobs, so it stays stable across deploys.
  let v = await s.get("config/session-secret");
  if (!v) {
    await s.set("config/session-secret", crypto.randomBytes(32).toString("hex"), { onlyIfNew: true });
    v = await s.get("config/session-secret");
  }
  return (cachedSecret = v);
}
const sign = (secret, payload) => crypto.createHmac("sha256", secret).update(payload).digest();

async function makeToken(s, rec) {
  const exp = Date.now() + TOKEN_TTL_MS;
  const p = b64u(JSON.stringify({ u: rec.id, v: rec.ver, exp }));
  return { token: p + "." + b64u(sign(await sessionSecret(s), p)), exp };
}

// Returns { id, name } for a valid session, otherwise null. The account record is re-read on every call,
// so a session stops working as soon as that account's password changes on another device.
async function authenticate(req, s) {
  const h = req.headers.get("authorization") || "";
  const t = h.startsWith("Bearer ") ? h.slice(7) : "";
  const [p, sig] = t.split(".");
  if (!p || !sig) return null;
  const given = Buffer.from(sig, "base64url");
  if (!eqBuf(given, sign(await sessionSecret(s), p))) return null;
  let data;
  try { data = JSON.parse(Buffer.from(p, "base64url").toString()); } catch { return null; }
  if (!data || typeof data.u !== "string" || !(data.exp > Date.now())) return null;
  const rec = await loadUser(s, data.u);
  if (!rec || rec.ver !== data.v) return null;
  return { id: rec.id, name: rec.name };
}

// ---------- brute-force protection (shared by every function instance, kept in Blobs) ----------
async function isBlocked(s, key, max, windowMs = THROTTLE_WINDOW_MS) {
  const r = await s.get(key, { type: "json" });
  return !!r && Date.now() - r.t < windowMs && r.n >= max;
}
async function recordFail(s, key, windowMs = THROTTLE_WINDOW_MS) {
  const now = Date.now();
  let r = await s.get(key, { type: "json" });
  if (!r || now - r.t >= windowMs) r = { n: 0, t: now };
  r.n++;
  await s.setJSON(key, r);
}
const clientIp = (req, context) => context?.ip || req.headers.get("x-nf-client-connection-ip") || "unknown";

// ---------- one-time codes (OTP) for self-service sign-up ----------
const normId = (v) => (typeof v === "string" ? v.trim().toLowerCase() : "");
const pendingKey = (id) => "pending/" + sha(id).slice(0, 32);
const otpHash = (secret, id, otp) => crypto.createHmac("sha256", secret).update("otp|" + id + "|" + otp).digest("hex");
// The person's name is what appears in the Log ("Updated By"). Names are unique so nobody can sign in as,
// or be mistaken for, someone else.
const cleanName = (v) => (typeof v === "string" ? v.normalize("NFC").replace(/\s+/g, " ").trim() : "");
const NAME_RE = /^[\p{L}][\p{L} .'\u2019-]{1,39}$/u;
const nameOk = (n) => NAME_RE.test(n) && (n.match(/\p{L}/gu) || []).length >= 2;
const nameKey = (n) => "names/" + sha(n.toLowerCase().replace(/[^\p{L}]/gu, "")).slice(0, 32);
// True if another account already uses this name, or is being created with it right now. A reservation whose
// account no longer exists (it was deleted, or creation failed) is only cleared once it is a couple of minutes
// old, so a sign-up that is half way through can never be taken over.
const NAME_HOLD_MS = 2 * 60 * 1000;
async function nameTaken(s, n) {
  const r = await s.get(nameKey(n), { type: "json" });
  if (!r) return false;
  if (await s.get(userKey(r.id), { type: "json" })) return true;
  if (Date.now() - (r.at || 0) < NAME_HOLD_MS) return true;
  await s.delete(nameKey(n)).catch(() => {});
  return false;
}

// Sends the code to the back office through Resend (https://resend.com). Needs RESEND_API_KEY.
// OTP_FROM is optional; without it Resend's test sender is used, which can only deliver to the address the
// Resend account was created with. To send from your own domain, verify the domain in Resend and set OTP_FROM.
async function sendOtpEmail({ id, name, otp, ip, now }) {
  const key = process.env.RESEND_API_KEY;
  if (!key) throw Object.assign(new Error("email-not-configured"), { code: "noemail" });
  const from = process.env.OTP_FROM || "Operational Excellence <onboarding@resend.dev>";
  const when = new Date(now).toUTCString();
  const text =
    `Someone asked to create an account on the Operational Excellence dashboard.\n\n` +
    `Name: ${name}\nUsername: ${id}\nOTP: ${otp}\nValid for 10 minutes.\n\n` +
    `Requested ${when} from IP ${ip}.\n\n` +
    `Give this OTP only to the person who owns this Cricket login. If you do not recognise the request, ignore this email.`;
  const esc = (t) => String(t).replace(/[&<>"']/g, (c) => ({ "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;", "'": "&#39;" }[c]));
  const html =
    `<p>Someone asked to create an account on the <b>Operational Excellence</b> dashboard.</p>` +
    `<p>Name: <b>${esc(name)}</b><br>Username: <b>${id}</b><br>OTP: <b style="font-size:20px;letter-spacing:3px">${otp}</b><br>Valid for 10 minutes.</p>` +
    `<p style="color:#555">Requested ${when} from IP ${ip}.<br>Give this OTP only to the person who owns this Cricket login. If you do not recognise the request, ignore this email.</p>`;
  const r = await fetch("https://api.resend.com/emails", {
    method: "POST",
    headers: { authorization: "Bearer " + key, "content-type": "application/json" },
    body: JSON.stringify({ from, to: [OTP_TO], subject: `Operational Excellence OTP for ${name} (${id})`, text, html }),
  });
  if (!r.ok) throw new Error("email-failed-" + r.status);
}

// ---------- small helpers ----------
async function readBody(req) {
  const text = await req.text();
  if (text.length > MAX_BODY_BYTES) throw new Error("too-large");
  return JSON.parse(text || "{}");
}
const validPayload = (b) =>
  b && (b.enc === "gz" || b.enc === "raw") && typeof b.z === "string" && b.z.length > 0 && b.z.length < 4.5e6;

// Read-modify-write on one JSON blob without losing a concurrent write (retries on conflict).
async function updateJSON(s, key, fn) {
  for (let i = 0; i < 6; i++) {
    const cur = await s.getWithMetadata(key, { type: "json" });
    const next = fn(cur ? cur.data : null);
    const r = cur ? await s.setJSON(key, next, { onlyIfMatch: cur.etag }) : await s.setJSON(key, next, { onlyIfNew: true });
    if (r.modified) return next;
  }
  throw new Error("write-conflict");
}

export default async (req, context) => {
  const url = new URL(req.url);
  const route = url.pathname.replace(/^\/api\/?/, "").replace(/\/$/, "");
  const method = req.method;

  try {
    const s = store();

    // ---------- public ----------
    if (route === "report" && method === "GET") {
      const r = await s.get("report/latest", { type: "json" });
      if (!r) return json({ empty: true });
      const out = { ts: r.ts, enc: r.enc, z: r.z };
      if (await authenticate(req, s)) { out.file = r.file; out.by = r.by || null; }
      return json(out);
    }

    // ---------- self-service sign-up, step 1: email a code to the back office ----------
    if (route === "register/request" && method === "POST") {
      if (process.env.SIGNUP_DISABLED === "1") return json({ error: "Creating accounts is turned off. Please contact the administrator." }, 403);
      let body;
      try { body = await readBody(req); } catch { return json({ error: "Bad request." }, 400); }
      const id = normId(body.username);
      const name = cleanName(body.name);
      if (!REG_ID.test(id)) {
        return json({ error: "Username must be 2 letters, 4 numbers, then @cricketlogin.com." }, 400);
      }
      if (!nameOk(name)) {
        return json({ error: "Enter your name (2 to 40 characters: letters, spaces, dots, hyphens or apostrophes; no numbers)." }, 400);
      }
      if (await s.get(userKey(id), { type: "json" })) return json({ error: "This username is already registered. Please log in instead." }, 409);
      if (await nameTaken(s, name)) return json({ error: "This name is already used by another account. Add your surname or an initial." }, 409);

      const ip = clientIp(req, context);
      const ipKey = "throttle/otpip/" + sha(ip).slice(0, 32);
      const allKey = "throttle/otp-all";
      if ((await isBlocked(s, ipKey, MAX_OTP_REQ_PER_IP_HOUR, HOUR_MS)) || (await isBlocked(s, allKey, MAX_OTP_REQ_PER_HOUR, HOUR_MS))) {
        return json({ error: "Too many OTP requests. Please try again later." }, 429);
      }
      const now = Date.now();
      const old = await s.get(pendingKey(id), { type: "json" });
      if (old && now - old.createdAt < OTP_RESEND_MS) {
        const wait = Math.ceil((OTP_RESEND_MS - (now - old.createdAt)) / 1000);
        return json({ error: `An OTP was just sent. Please wait ${wait} seconds before asking for another.` }, 429);
      }
      await recordFail(s, ipKey, HOUR_MS); // every request counts, whether or not the email goes out
      await recordFail(s, allKey, HOUR_MS);

      const otp = String(crypto.randomInt(0, 1000000)).padStart(6, "0");
      await s.setJSON(pendingKey(id), { createdAt: now, exp: now + OTP_TTL_MS, tries: 0, name, h: otpHash(await sessionSecret(s), id, otp) });
      try {
        await sendOtpEmail({ id, name, otp, ip, now });
      } catch (e) {
        await s.delete(pendingKey(id)).catch(() => {});
        console.error("otp email failed:", e && e.message);
        return json({ error: e && e.code === "noemail" ? "Email is not set up yet. Please contact the administrator." : "The OTP email could not be sent. Please try again, or contact the administrator." }, 503);
      }
      return json({ ok: true, expiresInMinutes: Math.round(OTP_TTL_MS / 60000) }); // the code is never returned to the browser
    }

    // ---------- self-service sign-up, step 2: code + chosen password -> account ----------
    if (route === "register/complete" && method === "POST") {
      if (process.env.SIGNUP_DISABLED === "1") return json({ error: "Creating accounts is turned off. Please contact the administrator." }, 403);
      let body;
      try { body = await readBody(req); } catch { return json({ error: "Bad request." }, 400); }
      const id = normId(body.username);
      const otp = typeof body.otp === "string" ? body.otp.trim() : "";
      const pw = typeof body.password === "string" ? body.password : "";
      if (!REG_ID.test(id)) return json({ error: "Username must be 2 letters, 4 numbers, then @cricketlogin.com." }, 400);
      if (!/^[0-9]{6}$/.test(otp)) return json({ error: "Enter the 6-digit OTP." }, 400);
      if (pw.length < MIN_PASSWORD) return json({ error: `The password must be at least ${MIN_PASSWORD} characters.` }, 400);
      if (pw.length > MAX_PASSWORD) return json({ error: "The password is too long." }, 400);

      const gKey = "throttle/regip/" + sha(clientIp(req, context) + "|" + id).slice(0, 32);
      if (await isBlocked(s, gKey, 12)) return json({ error: "Too many attempts. Try again in a few minutes." }, 429);

      // Reserve one attempt BEFORE checking the code. Each guess therefore costs a slot even when requests
      // arrive in parallel, so at most MAX_OTP_TRIES guesses are ever checked against one code.
      let pend = null;
      for (let i = 0; i < 8 && !pend; i++) {
        const got = await s.getWithMetadata(pendingKey(id), { type: "json" });
        if (!got) return json({ error: "No OTP was requested for this username, or it was already used. Request a new OTP." }, 400);
        if (got.data.exp < Date.now()) { await s.delete(pendingKey(id)).catch(() => {}); return json({ error: "This OTP has expired. Request a new one." }, 410); }
        if (got.data.tries >= MAX_OTP_TRIES) { await s.delete(pendingKey(id)).catch(() => {}); return json({ error: "Too many wrong OTP attempts. Request a new OTP." }, 429); }
        const w = await s.setJSON(pendingKey(id), { ...got.data, tries: got.data.tries + 1 }, { onlyIfMatch: got.etag });
        if (w.modified) pend = { ...got.data, tries: got.data.tries + 1 };
      }
      if (!pend) return json({ error: "The server is busy. Please try again." }, 409);

      const want = Buffer.from(pend.h, "hex");
      const have = Buffer.from(otpHash(await sessionSecret(s), id, otp), "hex");
      if (!eqBuf(have, want)) {
        await recordFail(s, gKey);
        await sleep(700);
        const left = MAX_OTP_TRIES - pend.tries;
        return json({ error: left > 0 ? `Incorrect OTP. ${left} attempt${left === 1 ? "" : "s"} left.` : "Incorrect OTP. Request a new OTP." }, 403);
      }
      // The name was given when the OTP was requested (the back office saw it in the email); it cannot be changed here.
      const name = pend.name;
      let held = await s.setJSON(nameKey(name), { id, name, at: Date.now() }, { onlyIfNew: true }); // reserve the name atomically
      if (!held.modified && !(await nameTaken(s, name))) held = await s.setJSON(nameKey(name), { id, name, at: Date.now() }, { onlyIfNew: true });
      if (!held.modified) return json({ error: "This name was just taken by another account. Request a new OTP with your surname or an initial added." }, 409);
      const rec = { id, name, cred: await hashPassword(pw), ver: 1, createdAt: Date.now(), changedAt: null };
      const w = await s.setJSON(userKey(id), rec, { onlyIfNew: true }); // never overwrites an existing account
      if (!w.modified) await s.delete(nameKey(name)).catch(() => {});
      await s.delete(pendingKey(id)).catch(() => {});
      if (!w.modified) return json({ error: "This username is already registered. Please log in instead." }, 409);
      const { token, exp } = await makeToken(s, rec);
      return json({ ok: true, token, exp, user: { id: rec.id, name: rec.name } });
    }

    // ---------- login ----------
    if (route === "login" && method === "POST") {
      let body;
      try { body = await readBody(req); } catch { return json({ error: "Bad request." }, 400); }
      const id = typeof body.user === "string" ? body.user.trim().toLowerCase() : "";
      const pw = typeof body.password === "string" ? body.password : "";
      if (!REG_ID.test(id) || !pw || pw.length > MAX_PASSWORD) {
        await dummyCheck(pw.slice(0, MAX_PASSWORD));
        await sleep(700);
        return json({ error: "Wrong username or password." }, 401);
      }
      const ipKey = "throttle/ipuser/" + sha(clientIp(req, context) + "|" + id).slice(0, 32);
      const acctKey = "throttle/user/" + id;
      if ((await isBlocked(s, ipKey, MAX_FAILS_PER_IP_USER)) || (await isBlocked(s, acctKey, MAX_FAILS_PER_USER))) {
        return json({ error: "Too many attempts. Try again in a few minutes." }, 429);
      }
      const rec = await loadUser(s, id);
      if (!rec) { // well-formed username that was never registered: same answer and timing as a wrong password
        await dummyCheck(pw);
        await recordFail(s, ipKey);
        await recordFail(s, acctKey);
        await sleep(700);
        return json({ error: "Wrong username or password." }, 401);
      }
      if (!(await verifyPassword(pw, rec.cred))) {
        await recordFail(s, ipKey);
        await recordFail(s, acctKey);
        await sleep(700);
        return json({ error: "Wrong username or password." }, 401);
      }
      await s.delete(ipKey).catch(() => {});
      const { token, exp } = await makeToken(s, rec);
      return json({ token, exp, user: { id: rec.id, name: rec.name } });
    }

    // ---------- everything below needs a signed-in account ----------
    if (!["session", "password", "upload", "logs"].includes(route)) return json({ error: "Not found." }, 404);
    const me = await authenticate(req, s);
    if (!me) return json({ error: "Login required." }, 401);

    if (route === "session" && method === "GET") return json({ ok: true, user: me });

    // ---------- change my own password ----------
    if (route === "password" && method === "POST") {
      let b;
      try { b = await readBody(req); } catch { return json({ error: "Bad request." }, 400); }
      const cur = typeof b.current === "string" ? b.current : "";
      const next = typeof b.next === "string" ? b.next : "";
      if (!cur) return json({ error: "Enter your current password." }, 400);
      if (!next) return json({ error: "Enter a new password." }, 400);
      if (next.length < MIN_PASSWORD) return json({ error: `The new password must be at least ${MIN_PASSWORD} characters.` }, 400);
      if (next.length > MAX_PASSWORD) return json({ error: "The new password is too long." }, 400);
      if (next === cur) return json({ error: "The new password must be different from the current one." }, 400);
      if (cur.length > MAX_PASSWORD) return json({ error: "Current password is incorrect." }, 403);

      const tKey = "throttle/pw/" + me.id; // only ever the signed-in account, never a name from the request
      if (await isBlocked(s, tKey, MAX_FAILS_PW_CHANGE)) {
        return json({ error: "Too many attempts. Try again in a few minutes." }, 429);
      }
      const newCred = await hashPassword(next);
      for (let i = 0; i < 6; i++) {
        const got = await s.getWithMetadata(userKey(me.id), { type: "json" });
        if (!got) return json({ error: "Account not found." }, 404);
        if (!(await verifyPassword(cur, got.data.cred))) {
          await recordFail(s, tKey);
          await sleep(700);
          return json({ error: "Current password is incorrect." }, 403);
        }
        const updated = { ...got.data, cred: newCred, ver: (got.data.ver || 1) + 1, changedAt: Date.now() };
        const w = await s.setJSON(userKey(me.id), updated, { onlyIfMatch: got.etag });
        if (w.modified) {
          await s.delete(tKey).catch(() => {});
          const { token, exp } = await makeToken(s, updated); // keeps this device signed in; other devices must log in again
          return json({ ok: true, token, exp, user: { id: updated.id, name: updated.name } });
        }
      }
      return json({ error: "Could not save the new password. Please try again." }, 409);
    }

    // ---------- publish ----------
    if (route === "upload" && method === "POST") {
      let b;
      try { b = await readBody(req); } catch (e) {
        return json({ error: e.message === "too-large" ? "Upload is too large." : "Bad request." }, e.message === "too-large" ? 413 : 400);
      }
      if (!validPayload(b)) return json({ error: "Invalid report data." }, 400);
      const ts = Date.now();
      const file = String(b.file || "").slice(0, 200);
      // `me` comes from the verified session. Any user name in the request body is ignored.
      await s.setJSON("report/latest", { ts, file, by: me.name, byId: me.id, enc: b.enc, z: b.z });

      let logged = false;
      let drop = [];
      await updateJSON(s, "logs/index", (cur) => {
        const index = Array.isArray(cur) ? cur.slice() : [];
        drop = [];
        logged = false;
        if (index[0] && b.h && index[0].h === b.h) return index; // same data as the last log: report updated, not logged again
        const m = b.meta || {};
        index.unshift({
          ts, file, by: me.name, byId: me.id, h: String(b.h || ""),
          demo: num(m.demo), hand: num(m.hand), lines: num(m.lines), tr: num(m.tr),
          tst: num(m.tst), emp: num(m.emp), trn: num(m.trn),
        });
        drop = index.splice(MAX_LOGS);
        logged = true;
        return index;
      });
      if (logged) {
        await s.setJSON("logs/" + ts, { enc: b.enc, z: b.z });
        for (const d of drop) await s.delete("logs/" + d.ts).catch(() => {});
      }
      return json({ ok: true, ts, logged, by: me.name });
    }

    // ---------- log ----------
    if (route === "logs" && method === "GET") {
      const ts = url.searchParams.get("ts");
      if (ts) {
        if (!/^\d{10,15}$/.test(ts)) return json({ error: "Bad request." }, 400);
        const d = await s.get("logs/" + ts, { type: "json" });
        return d ? json(d) : json({ error: "Log not found." }, 404);
      }
      return json({ logs: (await s.get("logs/index", { type: "json" })) || [] });
    }

    if (route === "logs" && method === "DELETE") {
      const ts = url.searchParams.get("ts");
      if (!ts || !/^\d{10,15}$/.test(ts)) return json({ error: "Bad request." }, 400);
      await updateJSON(s, "logs/index", (cur) => (Array.isArray(cur) ? cur : []).filter((e) => String(e.ts) !== ts));
      await s.delete("logs/" + ts).catch(() => {});
      return json({ ok: true });
    }

    return json({ error: "Method not allowed." }, 405);
  } catch (err) {
    console.error("api error:", err && err.message);
    return json({ error: "Server error." }, 500);
  }
};

export const config = { path: "/api/*" };
