import fs from "node:fs/promises";
import { SpreadsheetFile, Workbook } from "@oai/artifact-tool";

const outputDir = "outputs";
const outputPath = `${outputDir}/household-budget-db-schema.xlsx`;

const workbook = Workbook.create();

const theme = {
  navy: "#0F172A",
  blue: "#2563EB",
  paleBlue: "#EFF6FF",
  border: "#CBD5E1",
  header: "#E2E8F0",
  muted: "#64748B",
  green: "#16A34A",
  red: "#DC2626",
  yellow: "#FEF3C7",
};

const lists = {
  status: ["사용", "중지"],
  budget_type: ["fixed", "extra"],
  extra_type: ["cash_add", "reimbursement"],
  recovery_status: ["none", "pending", "collected"],
  ledger_type: ["public", "allowance"],
  expense_type: ["variable", "fixed"],
  payment_method_type: ["card", "bank", "cash"],
  payment_type: ["single", "installment"],
  installment_status: ["scheduled", "applied", "paid_off", "excluded"],
  repayment_type: ["equal_principal_interest", "equal_principal", "interest_only"],
  reserve_entry_type: ["income", "expense"],
  settlement_rule: ["used_next_month", "same_month", "manual"],
};

const sheets = [
  {
    name: "README",
    values: [
      ["가계부 DB 스프레드시트", ""],
      ["목적", "화면 설계가 끝난 뒤 실제 값 연동을 하기 위한 원장형 DB 구조입니다."],
      ["사용 원칙", "이 스프레드시트는 직접 입력용이 아니라 앱 화면에서 저장한 값을 쌓는 저장소입니다."],
      ["정산월 규칙", "카드 사용월과 실제 정산월이 다를 수 있으므로 모든 예산/지출 데이터에는 settlement_month를 별도로 둡니다."],
      ["추천 흐름", "DB관리 화면에서 master 값을 저장하고, 지출/예산/대출/비상금 화면에서 records 값을 저장합니다."],
      ["주의", "비상금은 독립 장부이므로 예산/지출/대시보드 계산에 자동 포함하지 않습니다."],
      ["탭 구분", "master: bank_accounts, cards, loan_companies / records: monthly_budgets, fixed_expenses, expenses, installment_schedules, loans, reserve_entries"],
    ],
    widths: [220, 740],
  },
  {
    name: "_lists",
    values: [
      ["list_name", "value", "label_ko", "memo"],
      ...Object.entries(lists).flatMap(([key, values]) =>
        values.map((value) => [key, value, labelFor(key, value), ""])
      ),
    ],
    widths: [180, 220, 220, 280],
  },
  {
    name: "schema_dictionary",
    values: [
      ["sheet_name", "column_name", "required", "description"],
      ...dictionaryRows(),
    ],
    widths: [190, 220, 90, 520],
  },
  tableSheet("bank_accounts", [
    ["bank_account_id", "name", "withdraw_day", "settlement_rule", "status", "memo", "created_at", "updated_at"],
    ["bank_001", "국민", 15, "used_next_month", "사용", "기본 출금 계좌", new Date("2026-06-01"), new Date("2026-06-01")],
    ["bank_002", "카카오", 15, "used_next_month", "사용", "", new Date("2026-06-01"), new Date("2026-06-01")],
    ["bank_003", "토스", 15, "used_next_month", "사용", "", new Date("2026-06-01"), new Date("2026-06-01")],
    ["bank_004", "현대", 15, "used_next_month", "사용", "현대카드 연결용", new Date("2026-06-01"), new Date("2026-06-01")],
    ["bank_005", "삼성", 15, "used_next_month", "사용", "삼성카드 연결용", new Date("2026-06-01"), new Date("2026-06-01")],
  ], [170, 140, 110, 160, 90, 260, 130, 130]),
  tableSheet("cards", [
    ["card_id", "name", "linked_bank_account_id", "payment_day", "settlement_rule", "status", "memo", "created_at", "updated_at"],
    ["card_001", "현대카드", "bank_004", 15, "used_next_month", "사용", "", new Date("2026-06-01"), new Date("2026-06-01")],
    ["card_002", "국민카드", "bank_001", 15, "used_next_month", "사용", "", new Date("2026-06-01"), new Date("2026-06-01")],
    ["card_003", "삼성카드", "bank_005", 15, "used_next_month", "사용", "", new Date("2026-06-01"), new Date("2026-06-01")],
  ], [150, 150, 190, 110, 160, 90, 240, 130, 130]),
  tableSheet("loan_companies", [
    ["loan_company_id", "name", "linked_bank_account_id", "withdraw_day", "status", "memo", "created_at", "updated_at"],
    ["loanco_001", "국민", "bank_001", 15, "사용", "", new Date("2026-06-01"), new Date("2026-06-01")],
    ["loanco_002", "토스", "bank_003", 15, "사용", "", new Date("2026-06-01"), new Date("2026-06-01")],
    ["loanco_003", "경남", "bank_001", 15, "사용", "", new Date("2026-06-01"), new Date("2026-06-01")],
  ], [170, 140, 190, 110, 90, 240, 130, 130]),
  tableSheet("monthly_budgets", [
    ["budget_id", "settlement_month", "budget_type", "title", "amount", "extra_type", "prepaid_amount", "collected_amount", "recovery_status", "recovered_date", "memo", "created_at", "updated_at"],
    ["budget_202607_001", "2026-07", "fixed", "기본 예산", 5400000, "", "", "", "none", "", "월 기본 수입/예산", new Date("2026-06-01"), new Date("2026-06-01")],
    ["budget_202607_002", "2026-07", "extra", "현금 추가", 200000, "cash_add", "", "", "none", "", "상태/선지급/회수 없음", new Date("2026-06-01"), new Date("2026-06-01")],
    ["budget_202607_003", "2026-07", "extra", "선지급 회수", 120000, "reimbursement", 120000, 0, "pending", "", "회수대기 클릭 시 회수 처리", new Date("2026-06-01"), new Date("2026-06-01")],
  ], [180, 130, 120, 180, 120, 140, 130, 140, 140, 130, 260, 130, 130]),
  tableSheet("fixed_expenses", [
    ["fixed_expense_id", "settlement_month", "title", "amount", "payment_method_type", "payment_method_id", "withdraw_day", "status", "memo", "created_at", "updated_at"],
    ["fixed_202607_001", "2026-07", "공과금", 420000, "bank", "bank_001", 15, "사용", "", new Date("2026-06-01"), new Date("2026-06-01")],
    ["fixed_202607_002", "2026-07", "용돈", 700000, "bank", "bank_002", 15, "사용", "용돈 내부 장부 지급액", new Date("2026-06-01"), new Date("2026-06-01")],
    ["fixed_202607_003", "2026-07", "보험", 300000, "card", "card_001", 15, "사용", "", new Date("2026-06-01"), new Date("2026-06-01")],
  ], [190, 130, 180, 120, 160, 160, 110, 90, 260, 130, 130]),
  tableSheet("expenses", [
    ["expense_id", "used_date", "settlement_month", "expense_type", "ledger_type", "title", "amount", "payment_method_type", "payment_method_id", "payment_type", "installment_group_id", "installment_months", "installment_round", "memo", "created_at", "updated_at"],
    ["exp_202606_001", new Date("2026-06-03"), "2026-07", "variable", "public", "마트", 64800, "card", "card_001", "single", "", "", "", "", new Date("2026-06-03"), new Date("2026-06-03")],
    ["exp_202606_002", new Date("2026-06-05"), "2026-07", "variable", "allowance", "커피", 5800, "card", "card_002", "single", "", "", "", "", new Date("2026-06-05"), new Date("2026-06-05")],
    ["exp_202606_003", new Date("2026-06-10"), "2026-07", "variable", "public", "가전", 360000, "card", "card_003", "installment", "inst_202606_001", 3, 1, "3개월 할부 1회차", new Date("2026-06-10"), new Date("2026-06-10")],
  ], [180, 130, 130, 130, 120, 180, 120, 160, 160, 130, 180, 150, 140, 240, 130, 130]),
  tableSheet("installment_schedules", [
    ["schedule_id", "installment_group_id", "expense_id", "settlement_month", "installment_round", "installment_months", "amount", "status", "early_payoff_date", "memo", "created_at", "updated_at"],
    ["sch_202607_001", "inst_202606_001", "exp_202606_003", "2026-07", 1, 3, 120000, "applied", "", "", new Date("2026-06-10"), new Date("2026-06-10")],
    ["sch_202608_001", "inst_202606_001", "exp_202606_003", "2026-08", 2, 3, 120000, "scheduled", "", "", new Date("2026-06-10"), new Date("2026-06-10")],
    ["sch_202609_001", "inst_202606_001", "exp_202606_003", "2026-09", 3, 3, 120000, "scheduled", "", "", new Date("2026-06-10"), new Date("2026-06-10")],
  ], [190, 180, 180, 130, 150, 150, 120, 120, 140, 240, 130, 130]),
  tableSheet("loans", [
    ["loan_id", "title", "loan_company_id", "withdraw_bank_account_id", "start_date", "principal", "annual_rate", "term_months", "repayment_type", "status", "memo", "created_at", "updated_at"],
    ["loan_001", "전세대출", "loanco_001", "bank_001", new Date("2024-03-15"), 120000000, 0.043, 120, "equal_principal_interest", "사용", "", new Date("2026-06-01"), new Date("2026-06-01")],
    ["loan_002", "신용대출", "loanco_002", "bank_003", new Date("2025-01-10"), 20000000, 0.057, 60, "equal_principal", "사용", "", new Date("2026-06-01"), new Date("2026-06-01")],
  ], [150, 180, 170, 200, 130, 140, 120, 120, 210, 90, 240, 130, 130]),
  tableSheet("reserve_entries", [
    ["reserve_entry_id", "entry_date", "entry_type", "title", "amount", "created_at", "updated_at"],
    ["reserve_001", new Date("2026-06-01"), "income", "비상금 입금", 500000, new Date("2026-06-01"), new Date("2026-06-01")],
    ["reserve_002", new Date("2026-06-04"), "expense", "긴급 지출", 100000, new Date("2026-06-04"), new Date("2026-06-04")],
  ], [180, 130, 120, 180, 120, 130, 130]),
];

function labelFor(listName, value) {
  const scopedLabels = {
    "budget_type:fixed": "고정 예산",
    "budget_type:extra": "추가 예산",
    "expense_type:variable": "변동 지출",
    "expense_type:fixed": "고정 지출",
  };
  if (scopedLabels[`${listName}:${value}`]) return scopedLabels[`${listName}:${value}`];

  const labels = {
    "사용": "사용",
    "중지": "중지",
    cash_add: "현금 추가",
    reimbursement: "선지급 회수",
    none: "대상 없음",
    pending: "회수대기",
    collected: "회수완료",
    public: "공금",
    allowance: "용돈",
    card: "카드",
    bank: "은행/계좌",
    cash: "현금",
    single: "일시불",
    installment: "할부",
    scheduled: "예정",
    applied: "반영",
    paid_off: "중도상환",
    excluded: "제외",
    equal_principal_interest: "원리금균등",
    equal_principal: "원금균등",
    interest_only: "만기일시/이자만",
    income: "수입",
    expense: "지출",
    used_next_month: "사용 다음 달 정산",
    same_month: "당월 정산",
    manual: "수동",
  };
  return labels[value] || value;
}

function tableSheet(name, values, widths) {
  return { name, values, widths, table: true };
}

function dictionaryRows() {
  const defs = {
    bank_accounts: [
      ["bank_account_id", "Y", "은행/계좌 고유 ID"],
      ["name", "Y", "은행 또는 계좌 표시명"],
      ["withdraw_day", "N", "기본 출금일"],
      ["settlement_rule", "N", "정산 규칙"],
      ["status", "Y", "사용/중지"],
    ],
    cards: [
      ["card_id", "Y", "카드사 고유 ID"],
      ["name", "Y", "카드사 표시명"],
      ["linked_bank_account_id", "Y", "연결 은행/계좌 ID"],
      ["payment_day", "Y", "카드 결제일"],
      ["settlement_rule", "Y", "사용월과 정산월 계산 규칙"],
    ],
    loan_companies: [
      ["loan_company_id", "Y", "대출 금융사 고유 ID"],
      ["name", "Y", "금융사 표시명"],
      ["linked_bank_account_id", "N", "기본 연결 은행/계좌 ID"],
      ["withdraw_day", "N", "기본 출금일"],
    ],
    monthly_budgets: [
      ["budget_id", "Y", "예산 기록 고유 ID"],
      ["settlement_month", "Y", "정산월, yyyy-mm"],
      ["budget_type", "Y", "fixed 또는 extra"],
      ["title", "Y", "예산 항목명"],
      ["amount", "Y", "예산 금액"],
      ["extra_type", "N", "추가 예산 유형"],
      ["recovery_status", "N", "회수 필요 예산의 회수 상태"],
    ],
    fixed_expenses: [
      ["fixed_expense_id", "Y", "고정 지출 고유 ID"],
      ["settlement_month", "Y", "정산월, yyyy-mm"],
      ["title", "Y", "고정 지출 항목명"],
      ["amount", "Y", "정산월에 반영될 금액"],
      ["payment_method_type", "N", "card/bank/cash"],
      ["payment_method_id", "N", "카드사 또는 은행/계좌 ID"],
    ],
    expenses: [
      ["expense_id", "Y", "지출 기록 고유 ID"],
      ["used_date", "Y", "실제 사용일"],
      ["settlement_month", "Y", "정산 반영월"],
      ["expense_type", "Y", "variable 또는 fixed"],
      ["ledger_type", "Y", "public 또는 allowance"],
      ["title", "Y", "지출 항목명"],
      ["amount", "Y", "정산 반영 금액"],
      ["payment_type", "Y", "single 또는 installment"],
    ],
    installment_schedules: [
      ["schedule_id", "Y", "할부 스케줄 고유 ID"],
      ["installment_group_id", "Y", "하나의 할부 결제를 묶는 ID"],
      ["settlement_month", "Y", "회차별 정산월"],
      ["installment_round", "Y", "할부 회차"],
      ["amount", "Y", "회차별 반영 금액"],
      ["status", "Y", "scheduled/applied/paid_off/excluded"],
    ],
    loans: [
      ["loan_id", "Y", "대출 고유 ID"],
      ["title", "Y", "대출명"],
      ["loan_company_id", "Y", "대출 금융사 ID"],
      ["withdraw_bank_account_id", "N", "상환 출금 은행/계좌 ID"],
      ["start_date", "Y", "실행일"],
      ["principal", "Y", "대출 원금"],
      ["annual_rate", "Y", "연 이율"],
      ["term_months", "Y", "대출 기간 개월 수"],
      ["repayment_type", "Y", "상환 방식"],
    ],
    reserve_entries: [
      ["reserve_entry_id", "Y", "비상금 기록 고유 ID"],
      ["entry_date", "Y", "거래 일자"],
      ["entry_type", "Y", "income 또는 expense"],
      ["title", "Y", "항목"],
      ["amount", "Y", "금액"],
    ],
  };

  return Object.entries(defs).flatMap(([sheet, rows]) =>
    rows.map(([column, required, description]) => [sheet, column, required, description])
  );
}

function colLetter(index) {
  let n = index + 1;
  let s = "";
  while (n > 0) {
    const mod = (n - 1) % 26;
    s = String.fromCharCode(65 + mod) + s;
    n = Math.floor((n - 1) / 26);
  }
  return s;
}

function formatSheet(sheet, values, widths, tableName) {
  const rowCount = values.length;
  const colCount = values[0].length;
  const lastCol = colLetter(colCount - 1);
  const used = sheet.getRange(`A1:${lastCol}${rowCount}`);

  sheet.showGridLines = false;
  sheet.freezePanes.freezeRows(1);
  used.values = values;
  used.format.borders = { preset: "all", style: "thin", color: theme.border };
  used.format.font = { name: "Arial", size: 10, color: theme.navy };
  used.format.wrapText = false;

  const header = sheet.getRange(`A1:${lastCol}1`);
  header.format.fill = { color: theme.header };
  header.format.font = { bold: true, color: theme.navy };
  header.format.rowHeightPx = 28;

  for (let c = 0; c < colCount; c += 1) {
    const width = widths[c] || 140;
    sheet.getRangeByIndexes(0, c, Math.max(rowCount, 20), 1).format.columnWidthPx = width;
  }

  if (tableName) {
    const table = sheet.tables.add(`A1:${lastCol}${Math.max(rowCount, 2)}`, true, tableName);
    table.style = "TableStyleMedium2";
    table.showFilterButton = true;
  }

  const dateColumns = values[0]
    .map((h, index) => String(h).includes("date") || String(h).endsWith("_at") ? index : -1)
    .filter((index) => index >= 0);
  for (const index of dateColumns) {
    sheet.getRangeByIndexes(1, index, 499, 1).setNumberFormat("yyyy-mm-dd");
  }

  const amountColumns = values[0]
    .map((h, index) => String(h).includes("amount") || h === "principal" ? index : -1)
    .filter((index) => index >= 0);
  for (const index of amountColumns) {
    sheet.getRangeByIndexes(1, index, 499, 1).setNumberFormat("#,##0");
  }

  const rateIndex = values[0].indexOf("annual_rate");
  if (rateIndex >= 0) {
    sheet.getRangeByIndexes(1, rateIndex, 499, 1).setNumberFormat("0.00%");
  }
}

function addValidation(sheet, columnName, listKey) {
  const header = sheet.getRange(`A1:${colLetter(sheet.getUsedRange().values[0].length - 1)}1`).values[0];
  const index = header.indexOf(columnName);
  if (index < 0) return;
  const values = lists[listKey];
  sheet.getRangeByIndexes(1, index, 499, 1).dataValidation = {
    rule: { type: "list", values },
  };
}

function addSheet(config) {
  const sheet = workbook.worksheets.add(config.name);
  formatSheet(sheet, config.values, config.widths, config.table ? `${config.name}_table` : null);
  return sheet;
}

for (const config of sheets) {
  addSheet(config);
}

const validationMap = {
  bank_accounts: [["settlement_rule", "settlement_rule"], ["status", "status"]],
  cards: [["settlement_rule", "settlement_rule"], ["status", "status"]],
  loan_companies: [["status", "status"]],
  monthly_budgets: [["budget_type", "budget_type"], ["extra_type", "extra_type"], ["recovery_status", "recovery_status"]],
  fixed_expenses: [["payment_method_type", "payment_method_type"], ["status", "status"]],
  expenses: [["expense_type", "expense_type"], ["ledger_type", "ledger_type"], ["payment_method_type", "payment_method_type"], ["payment_type", "payment_type"]],
  installment_schedules: [["status", "installment_status"]],
  loans: [["repayment_type", "repayment_type"], ["status", "status"]],
  reserve_entries: [["entry_type", "reserve_entry_type"]],
};

for (const [sheetName, rules] of Object.entries(validationMap)) {
  const sheet = workbook.worksheets.getItem(sheetName);
  for (const [columnName, listKey] of rules) addValidation(sheet, columnName, listKey);
}

const readme = workbook.worksheets.getItem("README");
readme.getRange("A1:B1").merge();
readme.getRange("A1").values = [["가계부 DB 스프레드시트"]];
readme.getRange("A1").format.fill = { color: theme.blue };
readme.getRange("A1").format.font = { bold: true, color: "#FFFFFF", size: 15 };
readme.getRange("A1").format.rowHeightPx = 34;
readme.getRange("A2:A7").format.fill = { color: theme.paleBlue };
readme.getRange("A2:A7").format.font = { bold: true, color: theme.navy };
readme.getRange("B2:B7").format.wrapText = true;
readme.getRange("A2:B7").format.rowHeightPx = 36;

const dict = workbook.worksheets.getItem("schema_dictionary");
dict.getRange("C2:C500").format.fill = { color: theme.yellow };

await fs.mkdir(outputDir, { recursive: true });

const overview = await workbook.inspect({
  kind: "sheet,table",
  maxChars: 6000,
  tableMaxRows: 4,
  tableMaxCols: 8,
});
console.log(overview.ndjson);

const errors = await workbook.inspect({
  kind: "match",
  searchTerm: "#REF!|#DIV/0!|#VALUE!|#NAME\\?|#N/A",
  options: { useRegex: true, maxResults: 100 },
  summary: "formula error scan",
});
console.log(errors.ndjson);

for (const sheetName of ["README", "bank_accounts", "monthly_budgets", "expenses", "loans"]) {
  const preview = await workbook.render({ sheetName, autoCrop: "all", scale: 1, format: "png" });
  const bytes = new Uint8Array(await preview.arrayBuffer());
  await fs.writeFile(`${outputDir}/${sheetName}-db-preview.png`, bytes);
}

const output = await SpreadsheetFile.exportXlsx(workbook);
await output.save(outputPath);
console.log(outputPath);
