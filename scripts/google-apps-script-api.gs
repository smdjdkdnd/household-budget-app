const SPREADSHEET_ID = '1Hkd6trAZlkSaGFxOtcwyXa_1UDTqBqUqA_DQ7Y8zeV0';
const CURRENT_SETTLEMENT_MONTH = '2026-07';
const API_TOKEN = '';

const TABLES = {
  bank_accounts: { id: 'bank_account_id' },
  cards: { id: 'card_id' },
  loan_companies: { id: 'loan_company_id' },
  budget_templates: { id: 'budget_template_id' },
  fixed_expense_templates: { id: 'fixed_expense_template_id' },
  monthly_budgets: { id: 'budget_id', month: 'settlement_month' },
  fixed_expenses: { id: 'fixed_expense_id', month: 'settlement_month' },
  expenses: { id: 'expense_id', month: 'settlement_month' },
  installment_schedules: { id: 'schedule_id', month: 'settlement_month' },
  loans: { id: 'loan_id' },
  reserve_entries: { id: 'reserve_entry_id' },
};

function doGet(event) {
  try {
    const params = event.parameter || {};
    checkToken(params.token);
    const action = params.action || 'snapshot';
    if (action === 'schema') return jsonResponse({ ok: true, tables: TABLES }, params.callback);
    if (action === 'list') return jsonResponse({ ok: true, rows: listRows(params.table, params) }, params.callback);
    if (action === 'snapshot') return jsonResponse({ ok: true, data: getSnapshot(params.settlement_month || CURRENT_SETTLEMENT_MONTH) }, params.callback);
    if (action === 'upsert') return jsonResponse({ ok: true, data: upsertRow(params.table, parseJson(params.row)) }, params.callback);
    if (action === 'delete') return jsonResponse({ ok: true, data: deleteRow(params.table, params.id) }, params.callback);
    if (action === 'deleteFixedExpenseScope') return jsonResponse({ ok: true, data: deleteFixedExpenseScope(params.fixed_expense_id, params.settlement_month, params.scope) }, params.callback);
    if (action === 'resetVariableExpenses') return jsonResponse({ ok: true, data: resetVariableExpenses(params.settlement_month) }, params.callback);
    if (action === 'cloneFixedMonth') return jsonResponse({ ok: true, data: cloneFixedMonth(params.source_month, params.target_month) }, params.callback);
    return jsonResponse({ ok: false, error: `Unknown action: ${action}` }, params.callback);
  } catch (error) {
    return jsonResponse({ ok: false, error: String(error && error.message ? error.message : error) }, (event.parameter || {}).callback);
  }
}

function doPost(event) {
  try {
    const body = JSON.parse((event.postData && event.postData.contents) || '{}');
    checkToken(body.token);
    const action = body.action;
    let result;

    if (action === 'upsert') result = upsertRow(body.table, body.row || {});
    else if (action === 'delete') result = deleteRow(body.table, body.id);
    else if (action === 'deleteFixedExpenseScope') result = deleteFixedExpenseScope(body.fixed_expense_id, body.settlement_month, body.scope);
    else if (action === 'resetVariableExpenses') result = resetVariableExpenses(body.settlement_month);
    else if (action === 'cloneFixedMonth') result = cloneFixedMonth(body.source_month, body.target_month);
    else return jsonResponse({ ok: false, error: `Unknown action: ${action}` }, 400);

    return jsonResponse({ ok: true, data: result });
  } catch (error) {
    return jsonResponse({ ok: false, error: String(error && error.message ? error.message : error) }, 500);
  }
}

function parseJson(value) {
  if (!value) return {};
  return typeof value === 'object' ? value : JSON.parse(value);
}

function jsonResponse(payload, callback) {
  const json = JSON.stringify(payload);
  if (callback) {
    return ContentService
      .createTextOutput(String(callback).replace(/[^\w.$]/g, '') + '(' + json + ');')
      .setMimeType(ContentService.MimeType.JAVASCRIPT);
  }
  return ContentService
    .createTextOutput(json)
    .setMimeType(ContentService.MimeType.JSON);
}

function checkToken(token) {
  if (API_TOKEN && token !== API_TOKEN) throw new Error('Invalid API token');
}

function getSpreadsheet() {
  return SpreadsheetApp.openById(SPREADSHEET_ID);
}

function getSheet(tableName) {
  if (!TABLES[tableName]) throw new Error(`Unknown table: ${tableName}`);
  const sheet = getSpreadsheet().getSheetByName(tableName);
  if (!sheet) throw new Error(`Sheet not found: ${tableName}`);
  return sheet;
}

function getHeaders(sheet) {
  const lastColumn = sheet.getLastColumn();
  return sheet.getRange(1, 1, 1, lastColumn).getValues()[0].map(String);
}

function readTable(tableName) {
  const sheet = getSheet(tableName);
  const headers = getHeaders(sheet);
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return [];
  return sheet.getRange(2, 1, lastRow - 1, headers.length).getValues()
    .filter(row => row.some(value => value !== '' && value !== null))
    .map(row => rowToObject(headers, row));
}

function rowToObject(headers, row) {
  return headers.reduce((object, header, index) => {
    object[header] = normalizeCell(row[index], header);
    return object;
  }, {});
}

function normalizeCell(value, header) {
  const isDate = Object.prototype.toString.call(value) === '[object Date]';
  if (header === 'settlement_month' || /_month$/.test(header)) {
    if (isDate) return Utilities.formatDate(value, 'Asia/Seoul', 'yyyy-MM');
    if (typeof value === 'number') return serialDate(value, 'yyyy-MM');
    const match = String(value || '').match(/^(\d{4})-(\d{2})/);
    return match ? `${match[1]}-${match[2]}` : value;
  }
  if (/_date$|_at$|start_date|used_date|entry_date/.test(header)) {
    if (isDate) return Utilities.formatDate(value, 'Asia/Seoul', 'yyyy-MM-dd');
    if (typeof value === 'number') return serialDate(value, 'yyyy-MM-dd');
  }
  return value;
}

function serialDate(value, format) {
  return Utilities.formatDate(new Date(Math.round((value - 25569) * 86400 * 1000)), 'Asia/Seoul', format);
}

function listRows(tableName, filters) {
  const table = TABLES[tableName];
  if (!table) throw new Error(`Unknown table: ${tableName}`);
  return readTable(tableName).filter(row => {
    if (filters.settlement_month && table.month && row[table.month] !== filters.settlement_month) return false;
    if (filters.status && row.status !== filters.status) return false;
    return true;
  });
}

function upsertRow(tableName, row) {
  const table = TABLES[tableName];
  if (!table) throw new Error(`Unknown table: ${tableName}`);
  const sheet = getSheet(tableName);
  const headers = getHeaders(sheet);
  const idColumn = table.id;
  const id = row[idColumn] || createId(tableName);
  const normalizedRow = Object.assign({}, row, {
    [idColumn]: id,
    updated_at: today(),
  });
  if (!normalizedRow.created_at) normalizedRow.created_at = today();

  const rowValues = headers.map(header => normalizedRow[header] === undefined ? '' : normalizedRow[header]);
  const rowIndex = findRowIndex(sheet, headers, idColumn, id);
  if (rowIndex) sheet.getRange(rowIndex, 1, 1, headers.length).setValues([rowValues]);
  else sheet.appendRow(rowValues);
  return { id, row: normalizedRow };
}

function deleteRow(tableName, id) {
  const table = TABLES[tableName];
  if (!table) throw new Error(`Unknown table: ${tableName}`);
  const sheet = getSheet(tableName);
  const headers = getHeaders(sheet);
  const rowIndex = findRowIndex(sheet, headers, table.id, id);
  if (!rowIndex) return { deleted: false, id };
  sheet.deleteRow(rowIndex);
  return { deleted: true, id };
}

function deleteFixedExpenseScope(id, settlementMonth, scope) {
  const fixedRows = readTable('fixed_expenses');
  const target = fixedRows.find(row => String(row.fixed_expense_id) === String(id));
  if (!target) return { deleted: 0, templates_updated: 0 };
  if (scope !== 'future') return deleteRow('fixed_expenses', id);

  const sheet = getSheet('fixed_expenses');
  const headers = getHeaders(sheet);
  const monthIndex = headers.indexOf('settlement_month');
  const titleIndex = headers.indexOf('title');
  const methodTypeIndex = headers.indexOf('payment_method_type');
  const methodIdIndex = headers.indexOf('payment_method_id');
  let deleted = 0;

  for (let rowIndex = sheet.getLastRow(); rowIndex >= 2; rowIndex -= 1) {
    const values = sheet.getRange(rowIndex, 1, 1, headers.length).getValues()[0];
    const sameTitle = String(values[titleIndex]) === String(target.title);
    const sameMethod = String(values[methodTypeIndex] || '') === String(target.payment_method_type || '')
      && String(values[methodIdIndex] || '') === String(target.payment_method_id || '');
    if (monthValue(normalizeCell(values[monthIndex], 'settlement_month')) >= monthValue(settlementMonth) && sameTitle && sameMethod) {
      sheet.deleteRow(rowIndex);
      deleted += 1;
    }
  }

  return {
    deleted,
    templates_updated: stopFixedExpenseTemplates(target, settlementMonth),
  };
}

function stopFixedExpenseTemplates(target, settlementMonth) {
  const sheet = getSheet('fixed_expense_templates');
  const headers = getHeaders(sheet);
  const titleIndex = headers.indexOf('title');
  const methodTypeIndex = headers.indexOf('payment_method_type');
  const methodIdIndex = headers.indexOf('payment_method_id');
  const applyToIndex = headers.indexOf('apply_to_month');
  const updatedAtIndex = headers.indexOf('updated_at');
  if (titleIndex < 0 || applyToIndex < 0) return 0;

  let updated = 0;
  for (let rowIndex = 2; rowIndex <= sheet.getLastRow(); rowIndex += 1) {
    const values = sheet.getRange(rowIndex, 1, 1, headers.length).getValues()[0];
    const sameTitle = String(values[titleIndex]) === String(target.title);
    const sameMethod = methodTypeIndex < 0 || methodIdIndex < 0
      || (String(values[methodTypeIndex] || '') === String(target.payment_method_type || '')
        && String(values[methodIdIndex] || '') === String(target.payment_method_id || ''));
    if (sameTitle && sameMethod) {
      sheet.getRange(rowIndex, applyToIndex + 1).setValue(previousMonth(settlementMonth));
      if (updatedAtIndex >= 0) sheet.getRange(rowIndex, updatedAtIndex + 1).setValue(today());
      updated += 1;
    }
  }
  return updated;
}

function findRowIndex(sheet, headers, idColumn, id) {
  const idIndex = headers.indexOf(idColumn);
  if (idIndex < 0) throw new Error(`Missing id column: ${idColumn}`);
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return null;
  const ids = sheet.getRange(2, idIndex + 1, lastRow - 1, 1).getValues();
  const found = ids.findIndex(row => String(row[0]) === String(id));
  return found >= 0 ? found + 2 : null;
}

function resetVariableExpenses(settlementMonth) {
  if (!isFutureMonth(settlementMonth, CURRENT_SETTLEMENT_MONTH)) {
    throw new Error('현재월 또는 과거월의 변동지출은 초기화할 수 없습니다.');
  }

  const sheet = getSheet('expenses');
  const headers = getHeaders(sheet);
  const monthIndex = headers.indexOf('settlement_month');
  const typeIndex = headers.indexOf('expense_type');
  const lastRow = sheet.getLastRow();
  let deleted = 0;
  for (let rowIndex = lastRow; rowIndex >= 2; rowIndex -= 1) {
    const values = sheet.getRange(rowIndex, 1, 1, headers.length).getValues()[0];
    if (normalizeCell(values[monthIndex], 'settlement_month') === settlementMonth && values[typeIndex] === 'variable') {
      sheet.deleteRow(rowIndex);
      deleted += 1;
    }
  }
  return { settlement_month: settlementMonth, deleted };
}

function cloneFixedMonth(sourceMonth, targetMonth) {
  if (!sourceMonth || !targetMonth) throw new Error('source_month and target_month are required.');
  if (sourceMonth === targetMonth) throw new Error('같은 정산월로는 복사할 수 없습니다.');

  let budgetRows = readTable('budget_templates')
    .filter(row => activeTemplate(row, targetMonth))
    .map(row => budgetFromTemplate(row, targetMonth));
  let fixedExpenseRows = readTable('fixed_expense_templates')
    .filter(row => activeTemplate(row, targetMonth))
    .map(row => fixedExpenseFromTemplate(row, targetMonth));

  if (!budgetRows.length && !fixedExpenseRows.length) {
    budgetRows = readTable('monthly_budgets')
      .filter(row => row.settlement_month === sourceMonth && row.budget_type === 'fixed')
      .map(row => cloneMonthlyRow(row, 'budget_id', targetMonth, 'budget'));
    fixedExpenseRows = readTable('fixed_expenses')
      .filter(row => row.settlement_month === sourceMonth)
      .map(row => cloneMonthlyRow(row, 'fixed_expense_id', targetMonth, 'fixed'));
  }

  budgetRows.forEach(row => upsertRow('monthly_budgets', row));
  fixedExpenseRows.forEach(row => upsertRow('fixed_expenses', row));

  return {
    source_month: sourceMonth,
    target_month: targetMonth,
    monthly_budgets: budgetRows.length,
    fixed_expenses: fixedExpenseRows.length,
  };
}

function activeTemplate(row, targetMonth) {
  const applyFrom = row.apply_from_month || '1900-01';
  const applyTo = row.apply_to_month || '';
  return (row.status || '사용') === '사용'
    && monthValue(applyFrom) <= monthValue(targetMonth)
    && (!applyTo || monthValue(applyTo) >= monthValue(targetMonth));
}

function budgetFromTemplate(row, targetMonth) {
  return {
    budget_id: `budget_${targetMonth.replace('-', '')}_${slug(row.title)}`,
    settlement_month: targetMonth,
    budget_type: row.budget_type || 'fixed',
    title: row.title,
    amount: row.amount,
    extra_type: row.extra_type || '',
    prepaid_amount: '',
    collected_amount: '',
    recovery_status: row.recovery_status || 'none',
    recovered_date: '',
    memo: row.memo || '템플릿에서 생성',
    created_at: today(),
    updated_at: today(),
  };
}

function fixedExpenseFromTemplate(row, targetMonth) {
  return {
    fixed_expense_id: `fixed_${targetMonth.replace('-', '')}_${slug(row.title)}`,
    settlement_month: targetMonth,
    title: row.title,
    amount: row.amount,
    payment_method_type: row.payment_method_type || '',
    payment_method_id: row.payment_method_id || '',
    withdraw_day: row.withdraw_day || '',
    status: row.status || '사용',
    memo: row.memo || '템플릿에서 생성',
    created_at: today(),
    updated_at: today(),
  };
}

function cloneMonthlyRow(row, idColumn, targetMonth, prefix) {
  const cloned = Object.assign({}, row);
  cloned[idColumn] = `${prefix}_${targetMonth.replace('-', '')}_${slug(row.title || row[idColumn])}`;
  cloned.settlement_month = targetMonth;
  cloned.created_at = today();
  cloned.updated_at = today();
  return cloned;
}

function getSnapshot(settlementMonth) {
  const bankAccounts = readTable('bank_accounts').filter(row => row.status === '사용');
  const cards = readTable('cards').filter(row => row.status === '사용');
  const loanCompanies = readTable('loan_companies').filter(row => row.status === '사용');
  const budgets = listRows('monthly_budgets', { settlement_month: settlementMonth });
  const fixedExpenses = listRows('fixed_expenses', { settlement_month: settlementMonth });
  const expenses = listRows('expenses', { settlement_month: settlementMonth });
  const installments = listRows('installment_schedules', { settlement_month: settlementMonth });
  const loans = readTable('loans').filter(row => row.status === '사용');
  const reserves = readTable('reserve_entries');

  const totalBudget = sum(budgets, 'amount');
  const allowanceBudget = sum(fixedExpenses.filter(row => row.title === '용돈'), 'amount');
  const publicFixedTotal = sum(fixedExpenses.filter(row => row.title !== '용돈'), 'amount');
  const fixedTotal = publicFixedTotal + allowanceBudget;
  const publicVariable = sum(expenses.filter(row => row.expense_type === 'variable' && row.ledger_type === 'public'), 'amount');
  const allowanceSpent = sum(expenses.filter(row => row.ledger_type === 'allowance'), 'amount');
  const livingBudget = totalBudget - fixedTotal;
  const publicSettlementSpent = fixedTotal + publicVariable;

  return {
    settlement_month: settlementMonth,
    master: { bank_accounts: bankAccounts, cards, loan_companies: loanCompanies },
    rows: { budgets, fixed_expenses: fixedExpenses, expenses, installment_schedules: installments, loans, reserve_entries: reserves },
    summary: {
      total_budget: totalBudget,
      fixed_total: fixedTotal,
      public_fixed_total: publicFixedTotal,
      living_budget: livingBudget,
      public_spent: publicVariable,
      public_settlement_spent: publicSettlementSpent,
      public_remaining: livingBudget - publicVariable,
      allowance_budget: allowanceBudget,
      allowance_spent: allowanceSpent,
      allowance_remaining: allowanceBudget - allowanceSpent,
      loan_payment_total: sumLoanPayments(loans),
      reserve_balance: reserveBalance(reserves),
    },
  };
}

function sum(rows, key) {
  return rows.reduce((total, row) => total + numberValue(row[key]), 0);
}

function reserveBalance(rows) {
  return rows.reduce((total, row) => {
    const amount = numberValue(row.amount);
    return row.entry_type === 'expense' ? total - amount : total + amount;
  }, 0);
}

function numberValue(value) {
  return Number(String(value || 0).replace(/[^0-9.-]/g, '')) || 0;
}

function sumLoanPayments(loans) {
  return loans.reduce((total, loan) => {
    const match = String(loan.memo || '').match(/(?:보정월상환|월상환)\s*([0-9,]+)/);
    return total + (match ? Number(match[1].replace(/,/g, '')) : 0);
  }, 0);
}

function isFutureMonth(month, baseMonth) {
  return monthValue(month) > monthValue(baseMonth);
}

function monthValue(month) {
  const match = normalizeMonth(month).match(/^(\d{4})-(\d{2})$/);
  if (!match) throw new Error(`Invalid month: ${month}`);
  return Number(match[1]) * 12 + Number(match[2]);
}

function previousMonth(month) {
  const match = normalizeMonth(month).match(/^(\d{4})-(\d{2})$/);
  if (!match) throw new Error(`Invalid month: ${month}`);
  const date = new Date(Number(match[1]), Number(match[2]) - 2, 1);
  return Utilities.formatDate(date, 'Asia/Seoul', 'yyyy-MM');
}

function normalizeMonth(month) {
  const isDate = Object.prototype.toString.call(month) === '[object Date]';
  if (isDate) return Utilities.formatDate(month, 'Asia/Seoul', 'yyyy-MM');
  if (typeof month === 'number') return serialDate(month, 'yyyy-MM');
  const match = String(month || '').match(/(\d{4})-(\d{2})/);
  if (match) return `${match[1]}-${match[2]}`;
  const date = new Date(month);
  return isNaN(date.getTime()) ? '' : Utilities.formatDate(date, 'Asia/Seoul', 'yyyy-MM');
}

function createId(tableName) {
  return `${tableName}_${Utilities.getUuid().slice(0, 8)}`;
}

function slug(value) {
  return String(value || 'item')
    .replace(/\s+/g, '_')
    .replace(/[^\w가-힣-]/g, '')
    .slice(0, 24);
}

function today() {
  return Utilities.formatDate(new Date(), 'Asia/Seoul', 'yyyy-MM-dd');
}
