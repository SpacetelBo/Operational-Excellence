// Spacetel Operational Excellence - API
//
// Public  : GET  /api/report            latest published report (anyone can view)
//           GET  /api/users             the account names shown in the login list (names only)
//           POST /api/login             { user, password } -> session token
// Signed in (any of the four accounts):
//           GET  /api/session           who am I (name comes from the signed session)
//           POST /api/password          { current, next } change MY password
//           POST /api/upload            publish a new report (+ save to the log, recorded under MY name)
//           GET  /api/logs[?ts=...]     list the log / open one saved log
//           DELETE /api/logs?ts=...     delete one log entry
//
// Everything marked "Signed in" is enforced here on the server. The user's identity is read from the
// signed session token, never from anything the browser sends in the request body.
//
// Accounts: the first time an account is used, its initial password is taken from a Netlify environment
// variable, hashed with scrypt and saved in Netlify Blobs. From then on Blobs is the source of truth:
// changing a password in the app is saved there, and redeploying never overwrites it.

import { getStore } from "@netlify/blobs";
import crypto from "node:crypto";
import { promisify } from "node:util";

const scrypt = promisify(crypto.scrypt);

// The four accounts. `env` names the environment variable that supplies the INITIAL password only.
const USERS = [
  { id: "lokesh", name: "Lokesh", env: "LOKESH_PASSWORD" },
  { id: "anand", name: "Anand", env: "ANAND_PASSWORD" },
  { id: "upender", name: "Upender", env: "UPENDER_PASSWORD" },
  { id: "venu", name: "Venu", env: "VENU_PASSWORD" },
];

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
const findUser = (id) => USERS.find((u) => u.id === id);

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

// ---------- account records in Netlify Blobs ----------
const userKey = (id) => "users/" + id;

// Returns the stored record. If the account has never been used, creates it from its environment
// variable. The write is "only if new", so simultaneous first logins cannot overwrite each other and a
// redeploy can never overwrite a password that was changed in the app.
async function loadUser(s, id, { init }) {
  const rec = await s.get(userKey(id), { type: "json" });
  if (rec || !init) return rec;
  const u = findUser(id);
  const initial = u && process.env[u.env];
  if (!initial) return null;
  const fresh = { id: u.id, name: u.name, cred: await hashPassword(initial), ver: 1, createdAt: Date.now(), changedAt: null };
  await s.setJSON(userKey(id), fresh, { onlyIfNew: true });
  return await s.get(userKey(id), { type: "json" }); // whichever request won the race
}

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
  const rec = await loadUser(s, data.u, { init: false });
  if (!rec || rec.ver !== data.v) return null;
  return { id: rec.id, name: rec.name };
}

// ---------- brute-force protection (shared by every function instance, kept in Blobs) ----------
async function isBlocked(s, key, max) {
  const r = await s.get(key, { type: "json" });
  return !!r && Date.now() - r.t < THROTTLE_WINDOW_MS && r.n >= max;
}
async function recordFail(s, key) {
  const now = Date.now();
  let r = await s.get(key, { type: "json" });
  if (!r || now - r.t >= THROTTLE_WINDOW_MS) r = { n: 0, t: now };
  r.n++;
  await s.setJSON(key, r);
}
const clientIp = (req, context) => context?.ip || req.headers.get("x-nf-client-connection-ip") || "unknown";

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

    if (route === "users" && method === "GET") {
      return json({ users: USERS.map((u) => ({ id: u.id, name: u.name })) });
    }

    // ---------- login ----------
    if (route === "login" && method === "POST") {
      let body;
      try { body = await readBody(req); } catch { return json({ error: "Bad request." }, 400); }
      const id = typeof body.user === "string" ? body.user.toLowerCase() : "";
      const pw = typeof body.password === "string" ? body.password : "";
      const u = findUser(id);
      if (!u || !pw || pw.length > MAX_PASSWORD) {
        await dummyCheck(pw.slice(0, MAX_PASSWORD));
        await sleep(700);
        return json({ error: "Wrong user or password." }, 401);
      }
      const ipKey = "throttle/ipuser/" + sha(clientIp(req, context) + "|" + id).slice(0, 32);
      const acctKey = "throttle/user/" + id;
      if ((await isBlocked(s, ipKey, MAX_FAILS_PER_IP_USER)) || (await isBlocked(s, acctKey, MAX_FAILS_PER_USER))) {
        return json({ error: "Too many attempts. Try again in a few minutes." }, 429);
      }
      const rec = await loadUser(s, id, { init: true });
      if (!rec) return json({ error: "This account has not been set up yet. Please contact the administrator." }, 503);
      if (!(await verifyPassword(pw, rec.cred))) {
        await recordFail(s, ipKey);
        await recordFail(s, acctKey);
        await sleep(700);
        return json({ error: "Wrong user or password." }, 401);
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
