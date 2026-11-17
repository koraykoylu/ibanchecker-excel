(function () {
  const $ = (id) => document.getElementById(id);
  const fmt = (n) => Number(n).toLocaleString("en-US");
  let running = false;
  let answer = null;

  async function renderSettings() {
    const key = await window.ibcSettings.getKey();
    $("has-key").classList.toggle("hidden", !key);
    $("no-key").classList.toggle("hidden", !!key);
    $("hint").textContent = IBC.keyHint(key);
    $("key").value = "";
    $("save").disabled = false;
    $("remove").disabled = false;
    await renderUsage();
  }

  async function renderUsage() {
    const usage = await window.ibcSettings.getRemaining();
    $("usage").textContent = usage
      ? `${fmt(usage.remaining)} requests left this month at the last check (${new Date(usage.at).toLocaleString()}).`
      : "";
  }

  function say(text) {
    $("status").textContent = text;
  }

  function ask() {
    $("confirm").classList.remove("hidden");
    $("confirm-yes").focus();
    return new Promise((resolve) => (answer = resolve));
  }

  function answered(yes) {
    $("confirm").classList.add("hidden");
    if (answer) answer(yes);
    answer = null;
  }

  function writeHeaderIfLabelled(above, header) {
    const label = above.values[0][0];
    if (label === "" || label === null || /\d/.test(String(label))) return false;
    if (!IBC.isEmptyBlock(header.values, header.formulas)) return false;
    return true;
  }

  async function checkSelection() {
    if (running) return;
    $("message").textContent = "";
    const key = await window.ibcSettings.getKey();
    if (!key) {
      say("");
      $("message").textContent = "Save your API key below first.";
      $("key").focus();
      return;
    }
    running = true;
    $("check").disabled = true;
    try {
      await Excel.run(async (ctx) => {
        const sel = ctx.workbook.getSelectedRange();
        sel.load("rowIndex,columnIndex,rowCount,columnCount");
        const sheet = sel.worksheet;
        await ctx.sync();

        const rows = sel.rowCount;
        if (rows > IBC.MENU_MAX_ROWS) {
          say(`Select at most ${fmt(IBC.MENU_MAX_ROWS)} rows at a time.`);
          return;
        }
        const top = sel.rowIndex;
        const ibanCol = sel.columnIndex;
        let outCol = ibanCol + sel.columnCount;

        const source = sheet.getRangeByIndexes(top, ibanCol, rows, 1);
        source.load("values");
        let target = sheet.getRangeByIndexes(top, outCol, rows, IBC.CHECK_HEADERS.length);
        target.load("values,formulas");
        await ctx.sync();

        const values = source.values;
        if (!values.some((row) => IBC.clean(row[0]))) {
          say("No IBANs in the first selected column.");
          return;
        }

        if (!IBC.isEmptyBlock(target.values, target.formulas)) {
          say("");
          const yes = await ask();
          if (!yes) {
            say("Nothing was changed.");
            return;
          }
          sheet.getRangeByIndexes(0, outCol, 1, IBC.CHECK_HEADERS.length).getEntireColumn().insert(Excel.InsertShiftDirection.right);
          await ctx.sync();
          target = sheet.getRangeByIndexes(top, outCol, rows, IBC.CHECK_HEADERS.length);
        }

        say(`Checking ${fmt(values.filter((row) => IBC.clean(row[0])).length)} IBANs…`);
        const result = await window.ibcChecker.lookup(IBC.unique(values), 0);
        if (result.auth) {
          say("");
          $("message").textContent = IBC.MSG.badKey;
          return;
        }

        if (top > 0) {
          const above = sheet.getRangeByIndexes(top - 1, ibanCol, 1, 1);
          const header = sheet.getRangeByIndexes(top - 1, outCol, 1, IBC.CHECK_HEADERS.length);
          above.load("values");
          header.load("values,formulas");
          await ctx.sync();
          if (writeHeaderIfLabelled(above, header)) header.values = [IBC.CHECK_HEADERS];
        }
        target.values = values.map((row) => IBC.checkRow(result.records[IBC.clean(row[0])]));
        await ctx.sync();

        const checked = values.filter((row) => IBC.clean(row[0])).length;
        const lines = [`${fmt(checked)} ${checked === 1 ? "IBAN" : "IBANs"} checked.`];
        if (result.cached) {
          lines.push(`${fmt(result.cached)} ${result.cached === 1 ? "was" : "were"} checked in the last six hours and not charged again.`);
        }
        if (result.remaining !== null) lines.push(`${fmt(result.remaining)} requests left on your key this month.`);
        if (result.quota) lines.push(`${IBC.MSG.quota}. The rows after the limit are marked.`);
        say(lines.join(" "));
      });
    } catch (err) {
      say("");
      $("message").textContent = err && err.code ? err.message : `The check stopped: ${err && err.message ? err.message : err}`;
    } finally {
      running = false;
      $("check").disabled = false;
      answered(false);
    }
  }

  async function saveKey() {
    const key = $("key").value.trim();
    $("message").textContent = "";
    if (!IBC.keyShape(key)) {
      $("message").textContent = "That is not an ibanchecker.cash API key: keys start with iban_ followed by letters and digits.";
      return;
    }
    $("save").disabled = true;
    try {
      await window.ibcSettings.setKey(key);
      await renderSettings();
    } catch (err) {
      $("message").textContent = `The key could not be saved: ${err && err.message ? err.message : err}`;
      $("save").disabled = false;
    }
  }

  async function removeKey() {
    $("remove").disabled = true;
    try {
      await window.ibcSettings.removeKey();
      await renderSettings();
    } catch (err) {
      $("message").textContent = `The key could not be removed: ${err && err.message ? err.message : err}`;
      $("remove").disabled = false;
    }
  }

  window.ibcOnRemaining = () => renderUsage().catch(() => {});

  Office.onReady(() => {
    $("version").textContent = `Version ${IBC.VERSION}`;
    $("check").addEventListener("click", checkSelection);
    $("confirm-yes").addEventListener("click", () => answered(true));
    $("confirm-no").addEventListener("click", () => answered(false));
    $("save").addEventListener("click", saveKey);
    $("key").addEventListener("keydown", (e) => {
      if (e.key === "Enter") saveKey();
    });
    $("remove").addEventListener("click", removeKey);
    renderSettings().catch((err) => ($("message").textContent = String(err && err.message ? err.message : err)));

    Office.actions.associate("checkSelection", async (event) => {
      try {
        await Office.addin.showAsTaskpane();
      } catch (err) {
        // The pane may already be open.
      }
      event.completed();
      checkSelection();
    });
  });
})();
