(function () {
  const STORAGE_KEYS = {
    apiUrl: "householdBudgetApiUrl",
    apiToken: "householdBudgetApiToken",
    selectedSettlementMonth: "householdBudgetSelectedSettlementMonth",
    settlementMonths: "householdBudgetSettlementMonths",
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
      if (value !== undefined && value !== null && value !== "") {
        url.searchParams.set(key, typeof value === "object" ? JSON.stringify(value) : value);
      }
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
    if (location.protocol === "file:") {
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

  function remove(table, id, settlementMonth) {
    clearCachedSnapshot(settlementMonth);
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

  function deleteFixedExpenseScope(fixedExpenseId, settlementMonth, scope) {
    clearCachedSnapshot(settlementMonth);
    return post("deleteFixedExpenseScope", {
      fixed_expense_id: fixedExpenseId,
      settlement_month: settlementMonth,
      scope,
    });
  }

  function parseMonthValue(value) {
    const koreanMatch = String(value || "").match(/(\d{4})년\s*(\d{1,2})월/);
    if (koreanMatch) return `${koreanMatch[1]}-${String(koreanMatch[2]).padStart(2, "0")}`;
    const isoMatch = String(value || "").match(/^(\d{4})-(\d{1,2})$/);
    return isoMatch ? `${isoMatch[1]}-${String(isoMatch[2]).padStart(2, "0")}` : "";
  }

  function formatKoreanMonth(value) {
    const month = parseMonthValue(value);
    const match = month.match(/^(\d{4})-(\d{2})$/);
    return match ? `${match[1]}년 ${Number(match[2])}월` : "";
  }

  function uniqueSortedMonths(months) {
    return Array.from(new Set(months.map(parseMonthValue).filter(Boolean)))
      .sort((a, b) => b.localeCompare(a));
  }

  function readSettlementMonthsFromDom() {
    return Array.from(document.querySelectorAll("[name='globalSettlementMonth'] option"))
      .map((option) => parseMonthValue(option.value || option.textContent));
  }

  function getSettlementMonths() {
    const saved = readStorage(STORAGE_KEYS.settlementMonths);
    if (saved) {
      try {
        const parsed = JSON.parse(saved);
        if (Array.isArray(parsed) && parsed.length) return uniqueSortedMonths(parsed);
      } catch (error) {
        // Ignore malformed stored month lists and rebuild from the page.
      }
    }
    return uniqueSortedMonths([...readSettlementMonthsFromDom(), CONFIG.currentSettlementMonth]);
  }

  function setSettlementMonths(months) {
    writeStorage(STORAGE_KEYS.settlementMonths, JSON.stringify(uniqueSortedMonths(months)));
  }

  function addOneMonth(month) {
    const match = parseMonthValue(month).match(/^(\d{4})-(\d{2})$/);
    if (!match) return CONFIG.currentSettlementMonth;
    const date = new Date(Number(match[1]), Number(match[2]), 1);
    return `${date.getFullYear()}-${String(date.getMonth() + 1).padStart(2, "0")}`;
  }

  function previousMonth(month) {
    const match = parseMonthValue(month).match(/^(\d{4})-(\d{2})$/);
    if (!match) return "";
    const date = new Date(Number(match[1]), Number(match[2]) - 2, 1);
    return `${date.getFullYear()}-${String(date.getMonth() + 1).padStart(2, "0")}`;
  }

  function monthNumber(month) {
    const match = parseMonthValue(month).match(/^(\d{4})-(\d{2})$/);
    return match ? Number(match[1]) * 12 + Number(match[2]) : 0;
  }

  function shouldAutoInitializeMonth(data, settlementMonth) {
    if (monthNumber(settlementMonth) <= monthNumber(CONFIG.currentSettlementMonth)) return false;
    const rows = data?.rows || {};
    return !(rows.budgets || []).length || !(rows.fixed_expenses || []).length;
  }

  async function ensureSettlementMonthInitialized(settlementMonth = getSelectedSettlementMonthValue()) {
    const month = parseMonthValue(settlementMonth);
    if (!month || monthNumber(month) <= monthNumber(CONFIG.currentSettlementMonth)) return false;
    const data = await snapshot(month);
    if (!shouldAutoInitializeMonth(data, month)) return false;
    emitStatus("cached", `${formatKoreanMonth(month)} 고정 항목 세팅 중`);
    await cloneFixedMonth(previousMonth(month), month);
    clearCachedSnapshot(month);
    return true;
  }

  function getSelectedSettlementMonthStored() {
    return parseMonthValue(readStorage(STORAGE_KEYS.selectedSettlementMonth));
  }

  function setSelectedSettlementMonth(month) {
    const normalized = parseMonthValue(month);
    if (normalized) writeStorage(STORAGE_KEYS.selectedSettlementMonth, normalized);
  }

  function renderSettlementMonthSelect(select, months, selectedMonth) {
    const currentMonth = parseMonthValue(select.value || select.selectedOptions?.[0]?.textContent) || selectedMonth;
    const nextSelected = months.includes(selectedMonth) ? selectedMonth : currentMonth;
    select.innerHTML = months
      .map((month) => `<option value="${formatKoreanMonth(month)}"${month === nextSelected ? " selected" : ""}>${formatKoreanMonth(month)}</option>`)
      .join("");
    select.dataset.monthManaged = "true";
  }

  function installSettlementMonthControls(label) {
    if (label.querySelector(".global-month-actions")) return;
    const actions = document.createElement("div");
    actions.className = "global-month-actions";
    actions.innerHTML = `
      <button class="month-action-button" type="button" data-month-action="add" aria-label="정산월 추가">+</button>
      <button class="month-action-button" type="button" data-month-action="delete" aria-label="정산월 삭제">-</button>
    `;
    label.appendChild(actions);
  }

  function syncSettlementMonthControls() {
    const selects = Array.from(document.querySelectorAll("[name='globalSettlementMonth']"));
    if (!selects.length) return;
    const selectedMonth = getSelectedSettlementMonthStored()
      || parseMonthValue(selects[0].value || selects[0].selectedOptions?.[0]?.textContent)
      || CONFIG.currentSettlementMonth;
    const months = uniqueSortedMonths([...getSettlementMonths(), selectedMonth]);
    setSettlementMonths(months);
    selects.forEach((select) => {
      renderSettlementMonthSelect(select, months, selectedMonth);
      const label = select.closest(".global-month");
      if (label) installSettlementMonthControls(label);
      if (!select.dataset.monthListenerBound) {
        select.dataset.monthListenerBound = "true";
        select.addEventListener("change", () => setSelectedSettlementMonth(parseMonthValue(select.value)));
      }
    });
  }

  function addSettlementMonth() {
    const currentMonths = getSettlementMonths();
    const latestMonth = currentMonths.slice().sort((a, b) => a.localeCompare(b)).at(-1) || CONFIG.currentSettlementMonth;
    const month = addOneMonth(latestMonth);
    const months = uniqueSortedMonths([...getSettlementMonths(), month]);
    setSettlementMonths(months);
    setSelectedSettlementMonth(month);
    syncSettlementMonthControls();
    document.querySelector("[name='globalSettlementMonth']")?.dispatchEvent(new Event("change", { bubbles: true }));
  }

  function deleteSettlementMonth() {
    const selectedMonth = getSelectedSettlementMonthValue();
    const months = getSettlementMonths();
    if (months.length <= 1) {
      alert("정산월은 최소 1개가 필요합니다.");
      return;
    }
    if (!confirm(`${formatKoreanMonth(selectedMonth)}을 정산월 목록에서 삭제할까요? DB 데이터는 삭제되지 않습니다.`)) return;
    const nextMonths = months.filter((month) => month !== selectedMonth);
    setSettlementMonths(nextMonths);
    setSelectedSettlementMonth(nextMonths[0]);
    syncSettlementMonthControls();
    document.querySelector("[name='globalSettlementMonth']")?.dispatchEvent(new Event("change", { bubbles: true }));
  }

  function parseKoreanMonth(label) {
    return parseMonthValue(label);
  }

  function getSelectedSettlementMonthValue() {
    const select = document.querySelector("[name='globalSettlementMonth']");
    if (select?.dataset.monthManaged !== "true") {
      return getSelectedSettlementMonthStored() || parseKoreanMonth(select?.value) || CONFIG.currentSettlementMonth;
    }
    return parseKoreanMonth(select?.value) || getSelectedSettlementMonthStored() || CONFIG.currentSettlementMonth;
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
    syncLoadingState(status, message);
    window.dispatchEvent(new CustomEvent("household-budget-db-status", {
      detail: { status, message },
    }));
  }

  function loadingPanel() {
    let panel = document.getElementById("dbLoadingPanel");
    if (!panel) {
      panel = document.createElement("div");
      panel.id = "dbLoadingPanel";
      panel.className = "db-loading-panel";
      panel.innerHTML = `<div class="db-loading-bar"></div><span>DB 불러오는 중</span>`;
      document.body.appendChild(panel);
    }
    return panel;
  }

  function syncLoadingState(status, message = "") {
    const panel = loadingPanel();
    const isLoading = status === "loading" || status === "cached";
    panel.classList.toggle("is-active", isLoading);
    panel.querySelector("span").textContent = message || "DB 불러오는 중";
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
    emitStatus("loading", `${formatKoreanMonth(settlementMonth)} 데이터 불러오는 중`);
    if (cached?.data && typeof onSnapshot === "function") {
      document.documentElement.dataset.dbStatus = "cached";
      emitStatus("cached", "저장된 DB 데이터 표시 중");
      onSnapshot(cached.data);
    }

    try {
      let data = await snapshot(settlementMonth);
      if (shouldAutoInitializeMonth(data, settlementMonth)) {
        await ensureSettlementMonthInitialized(settlementMonth);
        data = await snapshot(settlementMonth);
      }
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
    deleteFixedExpenseScope,
    ensureSettlementMonthInitialized,
    refreshPage,
    parseKoreanMonth,
    getSelectedSettlementMonthValue,
    getSettlementMonths,
    setSettlementMonths,
    syncSettlementMonthControls,
    formatWon,
    numberValue,
    hydratePage,
  };

  document.addEventListener("click", (event) => {
    const button = event.target.closest("[data-month-action]");
    if (!button) return;
    event.preventDefault();
    event.stopPropagation();
    if (button.dataset.monthAction === "add") addSettlementMonth();
    if (button.dataset.monthAction === "delete") deleteSettlementMonth();
  });

  if (document.readyState === "loading") {
    document.addEventListener("DOMContentLoaded", syncSettlementMonthControls);
  } else {
    syncSettlementMonthControls();
  }

  if ("serviceWorker" in navigator && location.protocol !== "file:") {
    window.addEventListener("load", () => {
      navigator.serviceWorker.register("service-worker.js").catch((error) => {
        console.warn("Service worker registration failed", error);
      });
    });
  }
})();
