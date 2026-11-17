// Pure logic shared by the formulas and the task pane. No Office API is touched here, so the
// tests load this file as it is.
(function (root) {
  const VERSION = "0.1.0";
  const API = "https://ibanchecker.cash/api/v1/validate";
  const CACHE_MS = 6 * 60 * 60 * 1000;
  const PARALLEL = 10;
  const FORMULA_MAX = 300;
  const MENU_MAX_ROWS = 5000;
  const PANE = "Home > IBAN Checker";

  const MSG = {
    noKey: `Add your ibanchecker.cash API key: ${PANE} > Settings. A free key is available at ibanchecker.cash/api-docs`,
    badKey: `ibanchecker.cash did not accept the API key. Check it in ${PANE} > Settings`,
    quota: "Monthly limit of your API key reached: ibanchecker.cash/pricing",
    unavailable: "ibanchecker.cash could not be reached, try again later",
    tooMany: `More than ${FORMULA_MAX} new IBANs in one formula: select them and use ${PANE} > Check selected IBANs`,
    invalid: "Invalid IBAN",
    notInDirectory: "Not in our bank directory",
  };

  const CHECK_HEADERS = ["Valid", "Error", "Country", "Bank", "BIC"];
  const FIELDS = ["valid", "error", "country", "bank", "bic"];

  function clean(value) {
    if (value === null || value === undefined) return "";
    const text = String(value).replace(/\s+/g, "").toUpperCase();
    // Every IBAN carries check digits, so text without a digit (a header such as "IBAN") is not sent and not charged.
    return /\d/.test(text) ? text : "";
  }

  function grid(input) {
    if (Array.isArray(input)) return input.map((row) => (Array.isArray(row) ? row : [row]));
    return [[input]];
  }

  function unique(rows) {
    const seen = new Set();
    const out = [];
    rows.forEach((row) =>
      row.forEach((cell) => {
        const iban = clean(cell);
        if (iban && !seen.has(iban)) {
          seen.add(iban);
          out.push(iban);
        }
      })
    );
    return out;
  }

  function record(status, body) {
    if (status === 200 && body && typeof body.valid === "boolean") {
      return {
        kind: "ok",
        valid: body.valid,
        error: body.valid ? "" : body.error || MSG.invalid,
        country: body.country_name || "",
        bank: body.bank_name || "",
        bic: body.bic || "",
      };
    }
    if (status === 401) return { kind: "auth" };
    if (status === 429) return { kind: "quota" };
    return { kind: "unavailable" };
  }

  function cell(field, rec) {
    if (!rec) return "";
    if (rec.kind === "quota") return MSG.quota;
    if (rec.kind === "auth") return MSG.badKey;
    if (rec.kind !== "ok") return MSG.unavailable;
    switch (field) {
      case "valid":
        return rec.valid;
      case "error":
        return rec.error;
      case "country":
        return rec.country;
      case "bank":
        if (!rec.valid) return MSG.invalid;
        return rec.bank || MSG.notInDirectory;
      case "bic":
        if (!rec.valid) return MSG.invalid;
        return rec.bic || MSG.notInDirectory;
      default:
        return "";
    }
  }

  function checkRow(rec) {
    if (!rec) return ["", "", "", "", ""];
    if (rec.kind !== "ok") return [cell("valid", rec), "", "", "", ""];
    return FIELDS.map((field) => cell(field, rec));
  }

  function keyShape(key) {
    return /^iban_[A-Za-z0-9]+$/.test(key);
  }

  function keyHint(key) {
    return key ? `iban_…${key.slice(-4)}` : "";
  }

  function isEmptyBlock(values, formulas) {
    const blank = (rows) => rows.every((row) => row.every((v) => v === "" || v === null));
    return blank(values) && blank(formulas || values);
  }

  function fail(code) {
    const err = new Error(MSG[code]);
    err.code = code;
    return err;
  }

  // deps: fetch, getKey() -> key|null, cache {get(k), set(k, v)} async, hash(text) -> hex, now() -> ms,
  // onRemaining(n) to remember the last count the API reported.
  function createChecker(deps) {
    const now = deps.now || (() => Date.now());
    const inflight = new Map();

    async function cacheKey(iban) {
      return `ibc1:${await deps.hash(`ibc1:${iban}`)}`;
    }

    async function fromCache(iban) {
      try {
        const hit = await deps.cache.get(await cacheKey(iban));
        if (hit && hit.at && now() - hit.at < CACHE_MS && hit.rec && hit.rec.kind === "ok") return hit.rec;
      } catch (err) {
        // A broken cache entry is a miss, never a failure.
      }
      return null;
    }

    function fetchOne(iban, key) {
      const id = `${key}:${iban}`;
      if (inflight.has(id)) return inflight.get(id);
      const run = (async () => {
        let status = 0;
        let body = null;
        let left = null;
        try {
          const res = await deps.fetch(API, {
            method: "POST",
            headers: { "Content-Type": "application/json", Authorization: `Bearer ${key}` },
            body: JSON.stringify({ iban }),
          });
          status = res.status;
          left = res.headers.get("X-RateLimit-Remaining");
          body = await res.json().catch(() => null);
        } catch (err) {
          status = 0;
        }
        const rec = record(status, body);
        if (rec.kind === "ok") await deps.cache.set(await cacheKey(iban), { at: now(), rec }).catch(() => {});
        return { rec, left: left === null ? null : Number(left) };
      })();
      inflight.set(id, run);
      run.finally(() => inflight.delete(id));
      return run;
    }

    async function lookup(ibans, maxUncached) {
      const key = await deps.getKey();
      if (!key) throw fail("noKey");
      const records = {};
      const pending = [];
      for (const iban of ibans) {
        const rec = await fromCache(iban);
        if (rec) records[iban] = rec;
        else pending.push(iban);
      }
      const result = { records, cached: ibans.length - pending.length, charged: 0, auth: false, quota: false, tooMany: false, remaining: null };
      if (maxUncached && pending.length > maxUncached) {
        result.tooMany = true;
        return result;
      }
      for (let i = 0; i < pending.length; i += PARALLEL) {
        const batch = pending.slice(i, i + PARALLEL);
        if (result.auth || result.quota) {
          batch.forEach((iban) => (records[iban] = { kind: result.auth ? "auth" : "quota" }));
          continue;
        }
        const answers = await Promise.all(batch.map((iban) => fetchOne(iban, key)));
        answers.forEach(({ rec, left }, j) => {
          records[batch[j]] = rec;
          if (rec.kind === "ok") {
            result.charged += 1;
            if (left !== null && !Number.isNaN(left) && (result.remaining === null || left < result.remaining)) result.remaining = left;
          }
          if (rec.kind === "auth") result.auth = true;
          if (rec.kind === "quota") {
            result.quota = true;
            result.remaining = 0;
          }
        });
      }
      if (result.remaining !== null && deps.onRemaining) {
        try {
          await deps.onRemaining(result.remaining);
        } catch (err) {
          // The count is only a hint for the task pane.
        }
      }
      return result;
    }

    async function formulaLookup(rows) {
      const result = await lookup(unique(rows), FORMULA_MAX);
      if (result.tooMany) throw fail("tooMany");
      if (result.auth) throw fail("badKey");
      if (result.quota && result.charged === 0 && result.cached === 0) throw fail("quota");
      return result.records;
    }

    async function field(input, name) {
      const rows = grid(input);
      const recs = await formulaLookup(rows);
      return rows.map((row) => row.map((c) => cell(name, recs[clean(c)])));
    }

    async function check(input) {
      const rows = grid(input);
      const recs = await formulaLookup(rows);
      return rows.map((row) => checkRow(recs[clean(row[0])]));
    }

    return { lookup, field, check };
  }

  root.IBC = {
    VERSION,
    API,
    CACHE_MS,
    FORMULA_MAX,
    MENU_MAX_ROWS,
    MSG,
    CHECK_HEADERS,
    clean,
    grid,
    unique,
    record,
    cell,
    checkRow,
    keyShape,
    keyHint,
    isEmptyBlock,
    createChecker,
  };
})(typeof globalThis !== "undefined" ? globalThis : this);
