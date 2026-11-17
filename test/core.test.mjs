import { test } from "node:test";
import assert from "node:assert/strict";
import { readFileSync } from "node:fs";
import { createHash } from "node:crypto";
import vm from "node:vm";

const source = readFileSync(new URL("../src/core.js", import.meta.url), "utf8");
const KEY = "iban_00000000000000000000000000000000";
const VALID = {
  valid: true,
  iban: "DE89370400440532013000",
  country_name: "Germany",
  bank_name: "Commerzbank",
  bic: "COBADEFFXXX",
};

function load() {
  const ctx = { console };
  vm.createContext(ctx);
  vm.runInContext(source, ctx);
  return ctx.IBC;
}

function response(status, body, remaining) {
  return {
    status,
    headers: { get: (name) => (name.toLowerCase() === "x-ratelimit-remaining" && remaining !== undefined ? String(remaining) : null) },
    json: async () => body,
  };
}

function api(answer) {
  const calls = [];
  const fetch = async (url, init) => {
    const iban = JSON.parse(init.body).iban;
    calls.push({ url, iban, auth: init.headers.Authorization });
    return answer(iban, calls.length);
  };
  return { calls, fetch };
}

function byMap(map) {
  return (iban, n) =>
    map[iban]
      ? response(200, map[iban], 90 - n)
      : response(200, { valid: false, iban, country_name: "Germany", error: "Check digits do not match." }, 90 - n);
}

function setup({ answer = byMap({ DE89370400440532013000: VALID }), key = KEY, clock } = {}) {
  const IBC = load();
  const store = new Map();
  const remaining = [];
  const server = api(answer);
  const checker = IBC.createChecker({
    fetch: server.fetch,
    getKey: async () => key,
    cache: { get: async (k) => store.get(k) || null, set: async (k, v) => void store.set(k, v) },
    hash: async (t) => createHash("sha256").update(t).digest("hex"),
    now: clock || (() => 1_000_000),
    onRemaining: async (n) => remaining.push(n),
  });
  return { IBC, checker, server, store, remaining };
}

const plain = (v) => JSON.parse(JSON.stringify(v));

test("a cell gives a 1x1 result, cleaned of spaces and case", async () => {
  const { checker, server } = setup();
  assert.deepEqual(plain(await checker.field(" de89 3704 0044 0532 0130 00 ", "valid")), [[true]]);
  assert.deepEqual(plain(await checker.field([["DE89370400440532013000"]], "bank")), [["Commerzbank"]]);
  assert.deepEqual(plain(await checker.field([["DE89370400440532013000"]], "bic")), [["COBADEFFXXX"]]);
  assert.equal(server.calls.length, 1, "the same IBAN within six hours comes from the cache");
  assert.equal(server.calls[0].auth, `Bearer ${KEY}`);
  assert.equal(server.calls[0].url, "https://ibanchecker.cash/api/v1/validate");
});

test("a range keeps its shape; empty cells and text without digits are not sent", async () => {
  const { checker, server } = setup();
  const out = await checker.field([["IBAN"], ["DE89370400440532013000"], [""], ["DE00370400440532013000"], ["de89370400440532013000"]], "valid");
  assert.deepEqual(plain(out), [[""], [true], [""], [false], [true]]);
  assert.deepEqual(server.calls.map((c) => c.iban), ["DE89370400440532013000", "DE00370400440532013000"]);
});

test("the cache is keyed by a hash, never by the IBAN, and expires after six hours", async () => {
  let t = 1_000_000;
  const { checker, server, store } = setup({ clock: () => t });
  await checker.field("DE89370400440532013000", "valid");
  for (const k of store.keys()) assert.ok(!k.includes("DE89"), "no IBAN in a cache key");
  for (const v of store.values()) assert.ok(!JSON.stringify(v).includes("DE89"), "no IBAN in a cached value");
  t += 6 * 60 * 60 * 1000 - 1;
  await checker.field("DE89370400440532013000", "valid");
  assert.equal(server.calls.length, 1);
  t += 2;
  await checker.field("DE89370400440532013000", "valid");
  assert.equal(server.calls.length, 2);
});

test("an invalid IBAN says so in every column, a bank we do not hold is named as our gap", async () => {
  const { checker } = setup({ answer: byMap({ GB82WEST12345698765432: { ...VALID, bank_name: null, bic: null, country_name: "United Kingdom" } }) });
  assert.deepEqual(plain(await checker.field("GB82WEST12345698765432", "bank")), [["Not in our bank directory"]]);
  assert.deepEqual(plain(await checker.check([["DE00370400440532013000"]])), [[false, "Check digits do not match.", "Germany", "Invalid IBAN", "Invalid IBAN"]]);
});

test("IBAN.CHECK gives five columns per row", async () => {
  const { checker } = setup();
  assert.deepEqual(plain(await checker.check([["DE89370400440532013000"], [""]])), [
    [true, "", "Germany", "Commerzbank", "COBADEFFXXX"],
    ["", "", "", "", ""],
  ]);
});

test("without a key nothing is sent and the error says where to add one", async () => {
  const { checker, server } = setup({ key: null });
  await assert.rejects(checker.field("DE89370400440532013000", "valid"), (err) => err.code === "noKey" && /Home > IBAN Checker > Settings/.test(err.message));
  assert.equal(server.calls.length, 0);
});

test("a refused key fails the formula and stops after the first batch", async () => {
  const { checker, server } = setup({ answer: () => response(401, { error: "Invalid or inactive API key." }) });
  const ibans = Array.from({ length: 120 }, (_, i) => [`DE${String(i).padStart(20, "0")}`]);
  await assert.rejects(checker.field(ibans, "valid"), (err) => err.code === "badKey");
  assert.equal(server.calls.length, 10, "no request after the batch that was refused");
});

test("a used-up quota fails the formula when nothing was answered, and marks the rest otherwise", async () => {
  const capped = setup({ answer: () => response(429, { error: "Monthly limit reached." }, 0) });
  await assert.rejects(capped.checker.field("DE89370400440532013000", "valid"), (err) => err.code === "quota");

  const partial = setup({ answer: (iban, n) => (n <= 3 ? response(200, { ...VALID, iban }, 3 - n) : response(429, {}, 0)) });
  const ibans = Array.from({ length: 25 }, (_, i) => [`DE${String(i).padStart(20, "0")}`]);
  const out = plain(await partial.checker.field(ibans, "valid"));
  assert.equal(out.filter((r) => r[0] === true).length, 3);
  assert.ok(out.slice(3).every((r) => r[0] === partial.IBC.MSG.quota));
  assert.equal(partial.server.calls.length, 10, "the batch after the 429 is not sent");
});

test("more than 300 new IBANs in one formula are refused before anything is sent", async () => {
  const { checker, server } = setup();
  const ibans = Array.from({ length: 301 }, (_, i) => [`DE${String(i).padStart(20, "0")}`]);
  await assert.rejects(checker.field(ibans, "valid"), (err) => err.code === "tooMany");
  assert.equal(server.calls.length, 0);
});

test("the ribbon run has no cap, sends each IBAN once and reports the lowest remaining count", async () => {
  const { IBC, checker, server, remaining } = setup();
  const ibans = Array.from({ length: 350 }, (_, i) => `DE${String(i % 340).padStart(20, "0")}`);
  const result = await checker.lookup(IBC.unique([ibans]), 0);
  assert.equal(server.calls.length, 340);
  assert.equal(result.charged, 340);
  assert.equal(result.remaining, 90 - 340);
  assert.deepEqual(remaining, [90 - 340]);
});

test("a network failure is reported as unavailable and not cached", async () => {
  let down = true;
  const { checker, server } = setup({
    answer: (iban, n) => {
      if (down) throw new Error("offline");
      return response(200, VALID, 90 - n);
    },
  });
  assert.deepEqual(plain(await checker.field("DE89370400440532013000", "valid")), [["ibanchecker.cash could not be reached, try again later"]]);
  down = false;
  assert.deepEqual(plain(await checker.field("DE89370400440532013000", "valid")), [[true]]);
  assert.equal(server.calls.length, 2);
});

test("identical IBANs from formulas running at the same time share one request", async () => {
  const { checker, server } = setup();
  await Promise.all([checker.field("DE89370400440532013000", "valid"), checker.field("DE89370400440532013000", "bank")]);
  assert.equal(server.calls.length, 1);
});

test("helpers: header detection, key shape and hint, empty block", () => {
  const IBC = load();
  assert.equal(IBC.clean("IBAN"), "");
  assert.equal(IBC.keyShape(KEY), true);
  assert.equal(IBC.keyShape("sk_live_x"), false);
  assert.equal(IBC.keyHint(KEY), "iban_…0000");
  assert.equal(IBC.isEmptyBlock([["", ""]], [["", ""]]), true);
  assert.equal(IBC.isEmptyBlock([["", ""]], [["=A1", ""]]), false);
});

test("metadata: five functions, each with a helpUrl (policy 1120.5.1), matrix in and out", () => {
  const meta = JSON.parse(readFileSync(new URL("../src/functions.json", import.meta.url), "utf8"));
  assert.deepEqual(meta.functions.map((f) => f.id), ["VALID", "ERROR", "BANK", "BIC", "CHECK"]);
  for (const f of meta.functions) {
    assert.match(f.helpUrl, /^https:\/\/ibanchecker\.cash\//);
    assert.equal(f.result.dimensionality, "matrix");
    assert.equal(f.parameters[0].dimensionality, "matrix");
  }
});

test("manifest: every URL is on ibanchecker.cash over https (policy 1120.5.2) and the version matches", () => {
  const xml = readFileSync(new URL("../manifest.xml", import.meta.url), "utf8");
  const urls = [...xml.matchAll(/DefaultValue="(https?:[^"]+)"/g)].map((m) => m[1]);
  assert.ok(urls.length > 8);
  for (const u of urls) assert.match(u, /^https:\/\/ibanchecker\.cash\/(excel\/|integrations\/excel)/);
  const IBC = load();
  assert.match(xml, new RegExp(`<Version>${IBC.VERSION.replace(/\./g, "\\.")}\\.0</Version>`));
  const pkg = JSON.parse(readFileSync(new URL("../package.json", import.meta.url), "utf8"));
  assert.equal(pkg.version, IBC.VERSION);
});
