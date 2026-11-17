// Office glue for the formulas. Runs in the shared runtime, the same page as the task pane.
(function () {
  const KEY_ITEM = "ibanchecker.apiKey";
  const REMAINING_ITEM = "ibanchecker.remaining";
  const memory = new Map();

  function local() {
    try {
      return window.localStorage;
    } catch (err) {
      return null;
    }
  }

  function prefix() {
    try {
      return (Office.context && Office.context.partitionKey) || "";
    } catch (err) {
      return "";
    }
  }

  // Results are kept under a SHA-256 hash of the IBAN, never the IBAN itself.
  const cache = {
    async get(id) {
      if (memory.has(id)) return memory.get(id);
      const store = local();
      const raw = store && store.getItem(`${prefix()}${id}`);
      return raw ? JSON.parse(raw) : null;
    },
    async set(id, value) {
      memory.set(id, value);
      const store = local();
      if (!store) return;
      try {
        store.setItem(`${prefix()}${id}`, JSON.stringify(value));
      } catch (err) {
        sweep(store);
      }
    },
  };

  function sweep(store) {
    const cutoff = Date.now() - IBC.CACHE_MS;
    for (let i = store.length - 1; i >= 0; i -= 1) {
      const id = store.key(i);
      if (!id || id.indexOf("ibc1:") === -1) continue;
      try {
        const v = JSON.parse(store.getItem(id));
        if (!v || v.at < cutoff) store.removeItem(id);
      } catch (err) {
        store.removeItem(id);
      }
    }
  }

  async function hash(text) {
    const bytes = await crypto.subtle.digest("SHA-256", new TextEncoder().encode(text));
    return Array.from(new Uint8Array(bytes), (b) => b.toString(16).padStart(2, "0")).join("");
  }

  const settings = {
    async getKey() {
      return (await OfficeRuntime.storage.getItem(KEY_ITEM)) || null;
    },
    async setKey(key) {
      await OfficeRuntime.storage.setItem(KEY_ITEM, key);
      await OfficeRuntime.storage.removeItem(REMAINING_ITEM);
    },
    async removeKey() {
      await OfficeRuntime.storage.removeItems([KEY_ITEM, REMAINING_ITEM]);
    },
    async getRemaining() {
      const raw = await OfficeRuntime.storage.getItem(REMAINING_ITEM);
      try {
        return raw ? JSON.parse(raw) : null;
      } catch (err) {
        return null;
      }
    },
    async setRemaining(n) {
      await OfficeRuntime.storage.setItem(REMAINING_ITEM, JSON.stringify({ remaining: n, at: new Date().toISOString() }));
      if (window.ibcOnRemaining) window.ibcOnRemaining();
    },
  };

  const checker = IBC.createChecker({
    fetch: (url, init) => fetch(url, init),
    getKey: settings.getKey,
    cache,
    hash,
    onRemaining: settings.setRemaining,
  });

  const startStore = local();
  if (startStore) sweep(startStore);

  window.ibcSettings = settings;
  window.ibcChecker = checker;

  function asExcelError(err) {
    const message = err && err.code ? err.message : IBC.MSG.unavailable;
    return new CustomFunctions.Error(CustomFunctions.ErrorCode.notAvailable, message);
  }

  function field(name) {
    return async (ibans) => {
      try {
        return await checker.field(ibans, name);
      } catch (err) {
        throw asExcelError(err);
      }
    };
  }

  async function check(ibans) {
    try {
      return await checker.check(ibans);
    } catch (err) {
      throw asExcelError(err);
    }
  }

  CustomFunctions.associate("VALID", field("valid"));
  CustomFunctions.associate("ERROR", field("error"));
  CustomFunctions.associate("BANK", field("bank"));
  CustomFunctions.associate("BIC", field("bic"));
  CustomFunctions.associate("CHECK", check);
})();
