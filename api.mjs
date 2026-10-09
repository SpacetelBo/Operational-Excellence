// Spacetel Operational Excellence - API
// Public : GET  /api/report            latest published report (anyone can view)
// Master : POST /api/login             exchange the master password for a session token
//          GET  /api/session           check a session token
//          POST /api/upload            publish a new report (+ save to the log)
//          GET  /api/logs[?ts=...]     list the log / open one saved log
//          DELETE /api/logs?ts=...     delete one log entry
// Everything marked "Master" is enforced here on the server, not in the browser.

import { getStore } from "@netlify/blobs";
import crypto from "node:crypto";

const PASSWORD = process.env.MASTER_PASSWORD || "spacetel123";
const SECRET =
  process.env.SESSION_SECRET ||
  crypto.createHash("sha256").update("oe-session|" + PASSWORD).digest("hex");
const TOKEN_TTL_MS = 12 * 60 * 60 * 1000; // 12 hours
const MAX_BODY_BYTES = 5 * 1024 * 1024;
const MAX_LOGS = 100;

const store = () => getStore({ name: "operational-excellence", consistency: "strong" });

const json = (obj, status = 200) =>
  new Response(JSON.stringify(obj), {
    status,
    headers: { "content-type": "application/json; charset=utf-8", "cache-control": "no-store" },
  });

const sha = (s) => crypto.createHash("sha256").update(String(s)).digest();
const sign = (p) => crypto.createHmac("sha256", SECRET).update(p).digest("base64url");
const safeEq = (a, b) => crypto.timingSafeEqual(sha(a), sha(b));

function makeToken() {
  const exp = Date.now() + TOKEN_TTL_MS;
  const p = Buffer.from(JSON.stringify({ exp })).toString("base64url");
  return { token: p + "." + sign(p), exp };
}
function isMaster(req) {
  const h = req.headers.get("authorization") || "";
  const t = h.startsWith("Bearer ") ? h.slice(7) : "";
  const [p, s] = t.split(".");
  if (!p || !s || !safeEq(s, sign(p))) return false;
  try {
    return JSON.parse(Buffer.from(p, "base64url").toString()).exp > Date.now();
  } catch {
    return false;
  }
}

// Best-effort brute-force guard (per warm function instance)
const fails = new Map();
const WINDOW = 10 * 60 * 1000, MAX_FAILS = 8;
function blocked(ip) {
  const a = (fails.get(ip) || []).filter((t) => Date.now() - t < WINDOW);
  fails.set(ip, a);
  return a.length >= MAX_FAILS;
}
const sleep = (ms) => new Promise((r) => setTimeout(r, ms));

const num = (v) => (Number.isFinite(Number(v)) ? Number(v) : 0);
const validPayload = (b) =>
  b && (b.enc === "gz" || b.enc === "raw") && typeof b.z === "string" && b.z.length > 0 && b.z.length < 4.5e6;

async function readBody(req) {
  const text = await req.text();
  if (text.length > MAX_BODY_BYTES) throw new Error("too-large");
  return JSON.parse(text || "{}");
}

export default async (req, context) => {
  const url = new URL(req.url);
  const route = url.pathname.replace(/^\/api\/?/, "").replace(/\/$/, "");
  const method = req.method;

  try {
    // ---------- public ----------
    if (route === "report" && method === "GET") {
      const r = await store().get("report/latest", { type: "json" });
      if (!r) return json({ empty: true });
      const out = { ts: r.ts, enc: r.enc, z: r.z };
      if (isMaster(req)) out.file = r.file;
      return json(out);
    }

    // ---------- login ----------
    if (route === "login" && method === "POST") {
      const ip = context?.ip || req.headers.get("x-nf-client-connection-ip") || "unknown";
      if (blocked(ip)) return json({ error: "Too many attempts. Try again in a few minutes." }, 429);
      let body;
      try { body = await readBody(req); } catch { return json({ error: "Bad request." }, 400); }
      if (typeof body.password === "string" && safeEq(body.password, PASSWORD)) {
        fails.delete(ip);
        const { token, exp } = makeToken();
        return json({ token, exp });
      }
      fails.set(ip, [...(fails.get(ip) || []), Date.now()]);
      await sleep(800);
      return json({ error: "Wrong password." }, 401);
    }

    // ---------- everything below needs the master login ----------
    if (!["session", "upload", "logs"].includes(route)) return json({ error: "Not found." }, 404);
    if (!isMaster(req)) return json({ error: "Master login required." }, 401);

    if (route === "session" && method === "GET") return json({ ok: true });

    if (route === "upload" && method === "POST") {
      let b;
      try { b = await readBody(req); } catch (e) {
        return json({ error: e.message === "too-large" ? "Upload is too large." : "Bad request." }, e.message === "too-large" ? 413 : 400);
      }
      if (!validPayload(b)) return json({ error: "Invalid report data." }, 400);
      const s = store();
      const ts = Date.now();
      const file = String(b.file || "").slice(0, 200);
      await s.setJSON("report/latest", { ts, file, enc: b.enc, z: b.z });

      let logged = false;
      const index = (await s.get("logs/index", { type: "json" })) || [];
      if (!(index[0] && b.h && index[0].h === b.h)) {
        const m = b.meta || {};
        const entry = {
          ts, file, h: String(b.h || ""),
          demo: num(m.demo), hand: num(m.hand), lines: num(m.lines), tr: num(m.tr),
          tst: num(m.tst), emp: num(m.emp), trn: num(m.trn),
        };
        await s.setJSON("logs/" + ts, { enc: b.enc, z: b.z });
        index.unshift(entry);
        const drop = index.splice(MAX_LOGS);
        await s.setJSON("logs/index", index);
        for (const d of drop) await s.delete("logs/" + d.ts).catch(() => {});
        logged = true;
      }
      return json({ ok: true, ts, logged });
    }

    if (route === "logs" && method === "GET") {
      const s = store();
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
      const s = store();
      const index = ((await s.get("logs/index", { type: "json" })) || []).filter((e) => String(e.ts) !== ts);
      await s.setJSON("logs/index", index);
      await s.delete("logs/" + ts).catch(() => {});
      return json({ ok: true });
    }

    return json({ error: "Method not allowed." }, 405);
  } catch (err) {
    console.error(err);
    return json({ error: "Server error." }, 500);
  }
};

export const config = { path: "/api/*" };
