(function () {
  const STORAGE_KEYS = {
    apiUrl: "householdBudgetApiUrl",
    apiToken: "householdBudgetApiToken",
  };

  function readWindowConfig() {
    try {
      return JSON.parse(window.name || "{}").householdBudget || {};
    } catch (error) {
      return {};
    }
  }

  function writeWindowConfig(values) {
    let state = {};
    try {
      state = JSON.parse(window.name || "{}");
    } catch (error) {
      state = {};
    }
    state.householdBudget = { ...(state.householdBudget || {}), ...values };
    window.name = JSON.stringify(state);
  }

  function readStorage(key) {
    const windowConfig = readWindowConfig();
    try {
      return localStorage.getItem(key) || sessionStorage.getItem(key) || windowConfig[key] || "";
    } catch (error) {
      return windowConfig[key] || "";
    }
  }

  function writeStorage(key, value) {
    const normalizedValue = String(value || "").trim();
    try {
      if (normalizedValue) {
        localStorage.setItem(key, normalizedValue);
        sessionStorage.setItem(key, normalizedValue);
      } else {
        localStorage.removeItem(key);
        sessionStorage.removeItem(key);
      }
    } catch (error) {
      // file:// origins can be fussy about storage; window.name is the fallback.
    }
    writeWindowConfig({ [key]: normalizedValue });
  }

  const CONFIG = {
    spreadsheetId: "1Hkd6trAZlkSaGFxOtcwyXa_1UDTqBqUqA_DQ7Y8zeV0",
    apiUrl: readStorage(STORAGE_KEYS.apiUrl),
    apiToken: readStorage(STORAGE_KEYS.apiToken),
    currentSettlementMonth: "2026-07",
  };

  const TABLES = {
    bankAccounts: "bank_accounts",
    cards: "cards",
    loanCompanies: "loan_companies",
    budgets: "monthly_budgets",
    fixedExpenses: "fixed_expenses",
    expenses: "expenses",
    installmentSchedules: "installment_schedules",
    loans: "loans",
    reserveEntries: "reserve_entries",
  };

  function setApiUrl(url) {
    CONFIG.apiUrl = String(url || "").trim();
    writeStorage(STORAGE_KEYS.apiUrl, CONFIG.apiUrl);
  }

  function getApiUrl() {
    return CONFIG.apiUrl;
  }

  function setApiToken(token) {
    CONFIG.apiToken = String(token || "").trim();
    writeStorage(STORAGE_KEYS.apiToken, CONFIG.apiToken);
  }

  function getApiToken() {
    return CONFIG.apiToken;
  }

  function assertApiUrl() {
    if (!CONFIG.apiUrl) {
      throw new Error("Google Apps Script 배포 URL이 필요합니다. HouseholdBudgetDB.setApiUrl(url)로 설정하세요.");
    }
  }

  async function get(action, params = {}) {
    assertApiUrl();
    const url = new URL(CONFIG.apiUrl);
    url.searchParams.set("action", action);
    if (CONFIG.apiToken) url.searchParams.set("token", CONFIG.apiToken);
    Object.entries(params).forEach(([key, value]) => {
      if (value !== undefined && value !== null && value !== "") url.searchParams.set(key, value);
    });
    if (location.protocol === "file:") return jsonp(url);
    try {
      return await request(url.toString());
    } catch (error) {
      return jsonp(url);
    }
  }

  async function post(action, payload = {}) {
    assertApiUrl();
    if (location.protocol === "file:" && ["cloneFixedMonth", "resetVariableExpenses"].includes(action)) {
      return get(action, payload);
    }
    return request(CONFIG.apiUrl, {
      method: "POST",
      headers: { "Content-Type": "text/plain;charset=utf-8" },
      body: JSON.stringify({ action, token: CONFIG.apiToken, ...payload }),
    });
  }

  async function request(url, options) {
    const response = await fetch(url, options);
    const payload = await response.json();
    if (!payload.ok) throw new Error(payload.error || "DB 요청에 실패했습니다.");
    return payload.data !== undefined ? payload.data : payload.rows;
  }

  function jsonp(url) {
    return new Promise((resolve, reject) => {
      const callbackName = `householdBudgetCallback_${Date.now()}_${Math.random().toString(36).slice(2)}`;
      const script = document.createElement("script");
      const cleanup = () => {
        delete window[callbackName];
        script.remove();
      };
      url.searchParams.set("callback", callbackName);
      window[callbackName] = (payload) => {
        cleanup();
        if (!payload.ok) reject(new Error(payload.error || "DB 요청에 실패했습니다."));
        else resolve(payload.data !== undefined ? payload.data : payload.rows);
      };
      script.onerror = () => {
        cleanup();
        reject(new Error("DB JSONP 요청에 실패했습니다."));
      };
      script.src = url.toString();
      document.head.appendChild(script);
    });
  }

  function snapshot(settlementMonth = getSelectedSettlementMonthValue()) {
    return get("snapshot", { settlement_month: settlementMonth });
  }

  function snapshotCacheKey(settlementMonth = getSelectedSettlementMonthValue()) {
    return `householdBudgetSnapshot:${CONFIG.spreadsheetId}:${settlementMonth}`;
  }

  function readCachedSnapshot(settlementMonth = getSelectedSettlementMonthValue()) {
    try {
      const cached = localStorage.getItem(snapshotCacheKey(settlementMonth));
      return cached ? JSON.parse(cached) : null;
    } catch (error) {
      return null;
    }
  }

  function writeCachedSnapshot(data, settlementMonth = getSelectedSettlementMonthValue()) {
    try {
      localStorage.setItem(snapshotCacheKey(settlementMonth), JSON.stringify({
        saved_at: new Date().toISOString(),
        data,
      }));
    } catch (error) {
      // Cache is only a speed boost. Ignore storage limits or private-mode failures.
    }
  }

  function clearCachedSnapshot(settlementMonth = getSelectedSettlementMonthValue()) {
    try {
      localStorage.removeItem(snapshotCacheKey(settlementMonth));
    } catch (error) {
      // Ignore cache cleanup failures.
    }
  }

  function list(table, filters = {}) {
    return get("list", { table, ...filters });
  }

  function upsert(table, row) {
    if (row?.settlement_month) clearCachedSnapshot(row.settlement_month);
    return post("upsert", { table, row });
  }

  function remove(table, id) {
    clearCachedSnapshot();
    return post("delete", { table, id });
  }

  function cloneFixedMonth(sourceMonth, targetMonth) {
    clearCachedSnapshot(targetMonth);
    return post("cloneFixedMonth", { source_month: sourceMonth, target_month: targetMonth });
  }

  function resetVariableExpenses(settlementMonth) {
    clearCachedSnapshot(settlementMonth);
    return post("resetVariableExpenses", { settlement_month: settlementMonth });
  }

  function parseKoreanMonth(label) {
    const match = String(label || "").match(/(\d{4})년\s*(\d{1,2})월/);
    return match ? `${match[1]}-${String(match[2]).padStart(2, "0")}` : "";
  }

  function getSelectedSettlementMonthValue() {
    const select = document.querySelector("[name='globalSettlementMonth']");
    return parseKoreanMonth(select?.value) || CONFIG.currentSettlementMonth;
  }

  function formatWon(value) {
    return `₩${Math.round(numberValue(value)).toLocaleString("ko-KR")}`;
  }

  function numberValue(value) {
    return Number(String(value ?? 0).replace(/[^0-9.-]/g, "")) || 0;
  }

  function escapeHtml(value) {
    return String(value ?? "")
      .replace(/&/g, "&amp;")
      .replace(/</g, "&lt;")
      .replace(/>/g, "&gt;")
      .replace(/"/g, "&quot;");
  }

  function setPageStatus(status, message) {
    const topbar = document.querySelector(".topbar");
    if (!topbar) return;
    let panel = document.getElementById("dbPageStatus");
    if (!panel) {
      panel = document.createElement("div");
      panel.id = "dbPageStatus";
      panel.className = "db-page-status";
      topbar.appendChild(panel);
    }
    panel.dataset.status = status;
    if (status === "connected") {
      panel.innerHTML = `<span>${escapeHtml(message)}</span>`;
      return;
    }
    panel.innerHTML = `
      <span>${escapeHtml(message)}</span>
      <div class="db-page-connect">
        <input class="control" id="dbPageApiUrl" placeholder="Apps Script Web App URL" value="${escapeHtml(CONFIG.apiUrl)}">
        <button class="button primary" type="button" id="dbPageApiSave">연결</button>
      </div>
    `;
    panel.querySelector("#dbPageApiSave")?.addEventListener("click", async () => {
      setApiUrl(panel.querySelector("#dbPageApiUrl")?.value || "");
      await hydratePage(window.__householdBudgetSnapshotHandler);
    });
  }

  function emitStatus(status, message) {
    setPageStatus(status, message);
    window.dispatchEvent(new CustomEvent("household-budget-db-status", {
      detail: { status, message },
    }));
  }

  async function hydratePage(onSnapshot) {
    if (typeof onSnapshot === "function") window.__householdBudgetSnapshotHandler = onSnapshot;
    if (!CONFIG.apiUrl) {
      document.documentElement.dataset.dbStatus = "not-configured";
      emitStatus("not-configured", "Google Sheets 연동 URL이 저장되지 않았습니다.");
      return null;
    }

    const settlementMonth = getSelectedSettlementMonthValue();
    const cached = readCachedSnapshot(settlementMonth);
    if (cached?.data && typeof onSnapshot === "function") {
      document.documentElement.dataset.dbStatus = "cached";
      emitStatus("cached", "저장된 DB 데이터 표시 중");
      onSnapshot(cached.data);
    }

    try {
      const data = await snapshot(settlementMonth);
      writeCachedSnapshot(data, settlementMonth);
      document.documentElement.dataset.dbStatus = "connected";
      emitStatus("connected", "Google Sheets DB 연결됨");
      if (typeof onSnapshot === "function") onSnapshot(data);
      return data;
    } catch (error) {
      document.documentElement.dataset.dbStatus = "error";
      emitStatus("error", error.message || "Google Sheets DB 요청에 실패했습니다.");
      console.error(error);
      return null;
    }
  }

  function refreshPage() {
    clearCachedSnapshot(getSelectedSettlementMonthValue());
    return hydratePage(window.__householdBudgetSnapshotHandler);
  }

  window.HouseholdBudgetDB = {
    CONFIG,
    TABLES,
    setApiUrl,
    getApiUrl,
    setApiToken,
    getApiToken,
    snapshot,
    list,
    upsert,
    remove,
    cloneFixedMonth,
    resetVariableExpenses,
    refreshPage,
    parseKoreanMonth,
    getSelectedSettlementMonthValue,
    formatWon,
    numberValue,
    hydratePage,
  };

  if ("serviceWorker" in navigator && location.protocol !== "file:") {
    window.addEventListener("load", () => {
      navigator.serviceWorker.register("service-worker.js").catch((error) => {
        console.warn("Service worker registration failed", error);
      });
    });
  }
})();
