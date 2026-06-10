const SPREADSHEET_ID = '1Hkd6trAZlkSaGFxOtcwyXa_1UDTqBqUqA_DQ7Y8zeV0';
const CURRENT_SETTLEMENT_MONTH = '2026-07';
const API_TOKEN = '';
const TABLES = {
  bank_accounts: ['bank_account_id'],
  cards: ['card_id'],
  loan_companies: ['loan_company_id'],
  budget_templates: ['budget_template_id'],
  fixed_expense_templates: ['fixed_expense_template_id'],
  monthly_budgets: ['budget_id', 'settlement_month'],
  fixed_expenses: ['fixed_expense_id', 'settlement_month'],
  expenses: ['expense_id', 'settlement_month'],
  installment_schedules: ['schedule_id', 'settlement_month'],
  loans: ['loan_id'],
  reserve_entries: ['reserve_entry_id'],
};
function doGet(e) {
  try {
    const p = e.parameter || {};
    checkToken(p.token);
    if ((p.action || 'snapshot') === 'snapshot') return out({ ok: true, data: snapshot(p.settlement_month || CURRENT_SETTLEMENT_MONTH) }, p.callback);
    if (p.action === 'list') return out({ ok: true, rows: listRows(p.table, p) }, p.callback);
    if (p.action === 'schema') return out({ ok: true, tables: TABLES }, p.callback);
    if (p.action === 'upsert') return out({ ok: true, data: upsertRow(p.table, parseJson(p.row)) }, p.callback);
    if (p.action === 'delete') return out({ ok: true, data: deleteRow(p.table, p.id) }, p.callback);
    if (p.action === 'deleteFixedExpenseScope') return out({ ok: true, data: deleteFixedExpenseScope(p.fixed_expense_id, p.settlement_month, p.scope) }, p.callback);
    if (p.action === 'resetVariableExpenses') return out({ ok: true, data: resetVariableExpenses(p.settlement_month) }, p.callback);
    if (p.action === 'cloneFixedMonth') return out({ ok: true, data: cloneFixedMonth(p.source_month, p.target_month) }, p.callback);
    return out({ ok: false, error: 'Unknown action: ' + p.action }, p.callback);
  } catch (err) {
    return out({ ok: false, error: String(err.message || err) }, (e.parameter || {}).callback);
  }
}
function doPost(e) {
  try {
    const b = JSON.parse((e.postData && e.postData.contents) || '{}');
    checkToken(b.token);
    if (b.action === 'upsert') return out({ ok: true, data: upsertRow(b.table, b.row || {}) });
    if (b.action === 'delete') return out({ ok: true, data: deleteRow(b.table, b.id) });
    if (b.action === 'deleteFixedExpenseScope') return out({ ok: true, data: deleteFixedExpenseScope(b.fixed_expense_id, b.settlement_month, b.scope) });
    if (b.action === 'resetVariableExpenses') return out({ ok: true, data: resetVariableExpenses(b.settlement_month) });
    if (b.action === 'cloneFixedMonth') return out({ ok: true, data: cloneFixedMonth(b.source_month, b.target_month) });
    return out({ ok: false, error: 'Unknown action: ' + b.action });
  } catch (err) {
    return out({ ok: false, error: String(err.message || err) });
  }
}
function parseJson(value) {
  if (!value) return {};
  return typeof value === 'object' ? value : JSON.parse(value);
}
function out(payload, callback) {
  const json = JSON.stringify(payload);
  if (callback) {
    return ContentService
      .createTextOutput(String(callback).replace(/[^\w.$]/g, '') + '(' + json + ');')
      .setMimeType(ContentService.MimeType.JAVASCRIPT);
  }
  return ContentService.createTextOutput(json).setMimeType(ContentService.MimeType.JSON);
}
function checkToken(token) {
  if (API_TOKEN && token !== API_TOKEN) throw new Error('Invalid API token');
}
function ss() {
  return SpreadsheetApp.openById(SPREADSHEET_ID);
}
function sheet(name) {
  if (!TABLES[name]) throw new Error('Unknown table: ' + name);
  const s = ss().getSheetByName(name);
  if (!s) throw new Error('Sheet not found: ' + name);
  return s;
}
function headers(s) {
  return s.getRange(1, 1, 1, s.getLastColumn()).getValues()[0].map(String);
}
function rows(name) {
  const s = sheet(name), h = headers(s), last = s.getLastRow();
  if (last < 2) return [];
  return s.getRange(2, 1, last - 1, h.length).getValues()
    .filter(r => r.some(v => v !== '' && v !== null))
    .map(r => h.reduce((o, k, i) => (o[k] = cell(r[i], k), o), {}));
}
function cell(v, key) {
  const isDate = Object.prototype.toString.call(v) === '[object Date]';
  if (key === 'settlement_month' || /_month$/.test(key)) {
    if (isDate) return Utilities.formatDate(v, 'Asia/Seoul', 'yyyy-MM');
    if (typeof v === 'number') return serialDate(v, 'yyyy-MM');
    const m = String(v || '').match(/^(\d{4})-(\d{2})/);
    return m ? m[1] + '-' + m[2] : v;
  }
  if (/_date$|_at$|start_date|used_date|entry_date/.test(key)) {
    if (isDate) return Utilities.formatDate(v, 'Asia/Seoul', 'yyyy-MM-dd');
    if (typeof v === 'number') return serialDate(v, 'yyyy-MM-dd');
  }
  return v;
}
function listRows(name, f) {
  const monthCol = TABLES[name] && TABLES[name][1];
  return rows(name).filter(r => {
    if (f.settlement_month && monthCol && r[monthCol] !== f.settlement_month) return false;
    if (f.status && r.status !== f.status) return false;
    return true;
  });
}
function upsertRow(name, row) {
  const idCol = TABLES[name][0], s = sheet(name), h = headers(s);
  const id = row[idCol] || (name + '_' + Utilities.getUuid().slice(0, 8));
  const saved = Object.assign({}, row, { [idCol]: id, updated_at: today() });
  if (!saved.created_at) saved.created_at = today();
  const values = h.map(k => saved[k] === undefined ? '' : saved[k]);
  const idx = findIndex(s, h, idCol, id);
  if (idx) s.getRange(idx, 1, 1, h.length).setValues([values]);
  else s.appendRow(values);
  return { id: id, row: saved };
}
function deleteRow(name, id) {
  const s = sheet(name), h = headers(s), idx = findIndex(s, h, TABLES[name][0], id);
  if (!idx) return { deleted: false, id: id };
  s.deleteRow(idx);
  return { deleted: true, id: id };
}
function deleteFixedExpenseScope(id, month, scope) {
  const fixedRows = rows('fixed_expenses');
  const target = fixedRows.find(r => String(r.fixed_expense_id) === String(id));
  if (!target) return { deleted: 0, templates_updated: 0 };
  if (scope !== 'future') return deleteRow('fixed_expenses', id);

  const s = sheet('fixed_expenses'), h = headers(s);
  const idCol = h.indexOf('fixed_expense_id'), monthCol = h.indexOf('settlement_month'), titleCol = h.indexOf('title');
  const methodTypeCol = h.indexOf('payment_method_type'), methodIdCol = h.indexOf('payment_method_id');
  let deleted = 0;
  for (let i = s.getLastRow(); i >= 2; i--) {
    const v = s.getRange(i, 1, 1, h.length).getValues()[0];
    const sameTitle = String(v[titleCol]) === String(target.title);
    const sameMethod = String(v[methodTypeCol] || '') === String(target.payment_method_type || '') && String(v[methodIdCol] || '') === String(target.payment_method_id || '');
    if (monthValue(cell(v[monthCol], 'settlement_month')) >= monthValue(month) && sameTitle && sameMethod) {
      s.deleteRow(i);
      deleted++;
    }
  }
  return { deleted: deleted, templates_updated: stopFixedExpenseTemplates(target, month) };
}
function stopFixedExpenseTemplates(target, month) {
  const s = sheet('fixed_expense_templates'), h = headers(s);
  const titleCol = h.indexOf('title'), methodTypeCol = h.indexOf('payment_method_type'), methodIdCol = h.indexOf('payment_method_id'), toCol = h.indexOf('apply_to_month'), updatedCol = h.indexOf('updated_at');
  if (titleCol < 0 || toCol < 0) return 0;
  let updated = 0;
  for (let i = 2; i <= s.getLastRow(); i++) {
    const v = s.getRange(i, 1, 1, h.length).getValues()[0];
    const sameTitle = String(v[titleCol]) === String(target.title);
    const sameMethod = methodTypeCol < 0 || methodIdCol < 0 || (String(v[methodTypeCol] || '') === String(target.payment_method_type || '') && String(v[methodIdCol] || '') === String(target.payment_method_id || ''));
    if (sameTitle && sameMethod) {
      s.getRange(i, toCol + 1).setValue(previousMonth(month));
      if (updatedCol >= 0) s.getRange(i, updatedCol + 1).setValue(today());
      updated++;
    }
  }
  return updated;
}
function findIndex(s, h, idCol, id) {
  const c = h.indexOf(idCol);
  if (c < 0 || s.getLastRow() < 2) return null;
  const ids = s.getRange(2, c + 1, s.getLastRow() - 1, 1).getValues();
  const found = ids.findIndex(r => String(r[0]) === String(id));
  return found >= 0 ? found + 2 : null;
}
function resetVariableExpenses(month) {
  if (monthValue(month) <= monthValue(CURRENT_SETTLEMENT_MONTH)) throw new Error('현재월/과거월 초기화 불가');
  const s = sheet('expenses'), h = headers(s), m = h.indexOf('settlement_month'), t = h.indexOf('expense_type');
  let deleted = 0;
  for (let i = s.getLastRow(); i >= 2; i--) {
    const v = s.getRange(i, 1, 1, h.length).getValues()[0];
    if (cell(v[m], 'settlement_month') === month && v[t] === 'variable') { s.deleteRow(i); deleted++; }
  }
  return { settlement_month: month, deleted: deleted };
}
function cloneFixedMonth(source, target) {
  if (!source || !target || source === target) throw new Error('정산월 확인 필요');
  let b = rows('budget_templates').filter(r => activeTemplate(r, target)).map(r => budgetFromTemplate(r, target));
  let f = rows('fixed_expense_templates').filter(r => activeTemplate(r, target)).map(r => fixedFromTemplate(r, target));
  if (!b.length && !f.length) {
    b = rows('monthly_budgets').filter(r => r.settlement_month === source && r.budget_type === 'fixed')
      .map(r => cloneRow(r, 'budget_id', target, 'budget'));
    f = rows('fixed_expenses').filter(r => r.settlement_month === source)
      .map(r => cloneRow(r, 'fixed_expense_id', target, 'fixed'));
  }
  b.forEach(r => upsertRow('monthly_budgets', r));
  f.forEach(r => upsertRow('fixed_expenses', r));
  return { source_month: source, target_month: target, monthly_budgets: b.length, fixed_expenses: f.length };
}
function activeTemplate(r, month) {
  const from = r.apply_from_month || '1900-01', to = r.apply_to_month || '';
  return (r.status || '사용') === '사용' && monthValue(from) <= monthValue(month) && (!to || monthValue(to) >= monthValue(month));
}
function budgetFromTemplate(r, month) {
  return {
    budget_id: 'budget_' + month.replace('-', '') + '_' + slug(r.title),
    settlement_month: month,
    budget_type: r.budget_type || 'fixed',
    title: r.title,
    amount: r.amount,
    extra_type: r.extra_type || '',
    prepaid_amount: '',
    collected_amount: '',
    recovery_status: r.recovery_status || 'none',
    recovered_date: '',
    memo: r.memo || '템플릿에서 생성',
    created_at: today(),
    updated_at: today()
  };
}
function fixedFromTemplate(r, month) {
  return {
    fixed_expense_id: 'fixed_' + month.replace('-', '') + '_' + slug(r.title),
    settlement_month: month,
    title: r.title,
    amount: r.amount,
    payment_method_type: r.payment_method_type || '',
    payment_method_id: r.payment_method_id || '',
    withdraw_day: r.withdraw_day || '',
    status: r.status || '사용',
    memo: r.memo || '템플릿에서 생성',
    created_at: today(),
    updated_at: today()
  };
}
function cloneRow(r, idCol, month, prefix) {
  const n = Object.assign({}, r);
  n[idCol] = prefix + '_' + month.replace('-', '') + '_' + slug(n.title || n[idCol]);
  n.settlement_month = month;
  n.created_at = today();
  n.updated_at = today();
  return n;
}
function snapshot(month) {
  const budgets = listRows('monthly_budgets', { settlement_month: month });
  const fixed = listRows('fixed_expenses', { settlement_month: month });
  const exp = listRows('expenses', { settlement_month: month });
  const loans = rows('loans').filter(r => r.status === '사용');
  const reserves = rows('reserve_entries');
  const total = sum(budgets, 'amount');
  const allowanceBudget = sum(fixed.filter(r => r.title === '용돈'), 'amount');
  const publicFixedTotal = sum(fixed.filter(r => r.title !== '용돈'), 'amount');
  const fixedTotal = publicFixedTotal + allowanceBudget;
  const living = total - fixedTotal;
  const publicSpent = sum(exp.filter(r => r.expense_type === 'variable' && r.ledger_type === 'public'), 'amount');
  const allowanceSpent = sum(exp.filter(r => r.ledger_type === 'allowance'), 'amount');
  const publicSettlementSpent = fixedTotal + publicSpent;
  return {
    settlement_month: month,
    master: { bank_accounts: rows('bank_accounts'), cards: rows('cards'), loan_companies: rows('loan_companies') },
    rows: { budgets: budgets, fixed_expenses: fixed, expenses: exp, installment_schedules: listRows('installment_schedules', { settlement_month: month }), loans: loans, reserve_entries: reserves },
    summary: { total_budget: total, fixed_total: fixedTotal, public_fixed_total: publicFixedTotal, living_budget: living, public_spent: publicSpent, public_settlement_spent: publicSettlementSpent, public_remaining: living - publicSpent, allowance_budget: allowanceBudget, allowance_spent: allowanceSpent, allowance_remaining: allowanceBudget - allowanceSpent, loan_payment_total: loanTotal(loans), reserve_balance: reserveBalance(reserves) }
  };
}
function serialDate(v, fmt) {
  return Utilities.formatDate(new Date(Math.round((v - 25569) * 86400 * 1000)), 'Asia/Seoul', fmt);
}
function num(v) {
  return Number(String(v || 0).replace(/[^0-9.-]/g, '')) || 0;
}
function sum(a, k) { return a.reduce((t, r) => t + num(r[k]), 0); }
function reserveBalance(a) { return a.reduce((t, r) => t + (r.entry_type === 'expense' ? -1 : 1) * num(r.amount), 0); }
function loanTotal(a) { return a.reduce((t, r) => { const m = String(r.memo || '').match(/(?:보정월상환|월상환)\s*([0-9,]+)/); return t + (m ? Number(m[1].replace(/,/g, '')) : 0); }, 0); }
function monthText(m) { const isDate = Object.prototype.toString.call(m) === '[object Date]'; if (isDate) return Utilities.formatDate(m, 'Asia/Seoul', 'yyyy-MM'); if (typeof m === 'number') return serialDate(m, 'yyyy-MM'); const x = String(m || '').match(/(\d{4})-(\d{2})/); if (x) return x[1] + '-' + x[2]; const d = new Date(m); return isNaN(d.getTime()) ? '' : Utilities.formatDate(d, 'Asia/Seoul', 'yyyy-MM'); }
function monthValue(m) { const x = monthText(m).match(/^(\d{4})-(\d{2})$/); if (!x) throw new Error('Invalid month: ' + m); return Number(x[1]) * 12 + Number(x[2]); }
function previousMonth(m) { const x = monthText(m).match(/^(\d{4})-(\d{2})$/); if (!x) throw new Error('Invalid month: ' + m); const d = new Date(Number(x[1]), Number(x[2]) - 2, 1); return Utilities.formatDate(d, 'Asia/Seoul', 'yyyy-MM'); }
function slug(v) { return String(v || 'item').replace(/\s+/g, '_').replace(/[^\w가-힣-]/g, '').slice(0, 24); }
function today() { return Utilities.formatDate(new Date(), 'Asia/Seoul', 'yyyy-MM-dd'); }
