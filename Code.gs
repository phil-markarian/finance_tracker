/**
 * Creates a new Google Spreadsheet with sheets for tracking income and expenses
 * over multiple years and months, including expense categories.
 *
 * @returns {Spreadsheet} The newly created spreadsheet.
 * @customfunction
 */
function createExpenseTrackingSpreadsheet() {
  const spreadsheetName = "Expense Tracker";
  const startYear = 2024;
  const endYear = 2024;

  const ss = SpreadsheetApp.create(spreadsheetName);
  const mainSheet = ss.getSheets()[0].setName("Yearly Income/Expenses");
  createMainIndexSheet(mainSheet, startYear, endYear);

  // Define expense categories
  const expenseCategories = [
    "Housing",
    "Utilities",
    "Food",
    "Transportation",
    "Insurance",
    "Healthcare",
    "Entertainment",
    "Education",
    "Personal Care",
    "Debt Repayment",
    "Savings & Investments",
    "Other"
  ];

  // Create year sheets
  for (let year = startYear; year <= endYear; year++) {
    const yearSheet = ss.insertSheet().setName(year.toString());
    createYearSheet(yearSheet, year, expenseCategories);

    // Create month sheets for the year
    for (let month = 1; month <= 12; month++) {
      const monthSheetName = `${year}.${month}`;
      const monthSheet = ss.insertSheet().setName(monthSheetName);
      createMonthSheet(monthSheet, expenseCategories);
    }
  }

  return ss;
}

/**
 * Generates a `SUM` formula string for a given range and column reference across 12 months.
 * @param {string} year The year to use in the formula.
 * @param {string} column The column reference in each month's sheet.
 * @param {number} rowIndex The row index for specific categories.
 * @returns {string} The generated `SUM` formula.
 */
function generateSumFormulaAcrossMonths(year, column, rowIndex) {
  const monthReferences = Array.from({ length: 12 }, (_, i) => `'${year}.${i + 1}'!${column}${rowIndex}`);
  const referenceList = monthReferences.join(',');

  return `=SUM(${referenceList})`;
}

/**
 * Sets up the structure and formatting for the Main Index sheet.
 *
 * @param {Sheet} sheet The sheet to set up as the Main Index sheet.
 * @param {number} startYear The starting year for creating year and month sheets.
 * @param {number} endYear The ending year for creating year and month sheets.
 */
function createMainIndexSheet(sheet, startYear, endYear) {
  const headerRange = sheet.getRange("A1:C1");
  headerRange.setValues([["Year", "Total Expenses", "Total Income"]]);
  headerRange.setFontWeight("bold");

  let rowIndex = 2;
  for (let year = startYear; year <= endYear; year++) {
    const totalExpensesFormula = `=SUM('${year}'!C2)`;
    const totalIncomeFormula = `=SUM('${year}'!F2)`;
    sheet.getRange(`A${rowIndex}:C${rowIndex}`).setValues([[year, totalExpensesFormula, totalIncomeFormula]]);
    rowIndex++;
  }

  const overallExpensesFormula = `=SUM(B2:B${rowIndex - 1})`;
  const overallIncomeFormula = `=SUM(C2:C${rowIndex - 1})`;
  sheet.getRange(`A${rowIndex}:C${rowIndex}`).setValues([["Total", overallExpensesFormula, overallIncomeFormula]]);
}

/**
 * Sets up the structure and formatting for a year sheet.
 *
 * @param {Sheet} sheet The sheet to set up as a year sheet.
 * @param {number} year The year for the sheet.
 * @param {Array<string>} expenseCategories The list of expense categories.
 */
function createYearSheet(sheet, year, expenseCategories) {
  setupYearHeader(sheet);
  setupExpenseCategories(sheet, expenseCategories, year);
  setupAdditionalIncomeFormulas(sheet, year);
  applyExpenseFormulas(sheet, expenseCategories, year);
}

function applyExpenseFormulas(sheet, expenseCategories, year) {
  // Apply formulas for each category in Column B after all month sheets are created
  for (let i = 0; i < expenseCategories.length; i++) {
    let formulaCell = sheet.getRange("B" + (i + 2));  // B2, B3, B4, ..., Bn
    let formula = generateSumFormulaAcrossMonths(year, "B", i + 2);
    formulaCell.setFormula(formula);
  }
}

/**
 * Sets up the main header row for the year sheet.
 */
function setupYearHeader(sheet) {
  const headerRange = sheet.getRange("A1:F1");
  headerRange.setValues([["Categories", "Amount", "Total Expenses", "Freelance Income", "Company Income", "Total Income"]]);
  headerRange.setFontWeight("bold");
}

/**
 * Sets up the expense categories and applies formulas to calculate expenses from all months.
 *
 * @param {Sheet} sheet The sheet to set up with formulas.
 * @param {Array<string>} expenseCategories The list of expense categories.
 * @param {number} year The year being processed.
 */
function setupExpenseCategories(sheet, expenseCategories, year) {
  // Setting up the categories in Column A
  const categoriesRange = sheet.getRange("A2:A" + (expenseCategories.length + 1));
  categoriesRange.setValues(expenseCategories.map(category => [category]));

  // Set up the "Total Expenses" formula in Column C
  const totalExpensesRange = sheet.getRange("C2");
  totalExpensesRange.setFormula("=SUM(B2:B)");
}

/**
 * Sets up the additional income formulas like freelance and company incomes.
 * @param {Sheet} sheet The sheet to update with the income formulas.
 * @param {number} year The year for which the formulas should be set up.
 */
/**
 * Sets up the additional income formulas like freelance and company incomes.
 * @param {Sheet} sheet The sheet to update with the income formulas.
 * @param {number} year The year for which the formulas should be set up.
 */
function setupAdditionalIncomeFormulas(sheet, year) {
  const freelanceIncomeRange = sheet.getRange("D2");
  const companyIncomeRange = sheet.getRange("E2");
  const freelanceIncomeFormula = generateSumFormulaAcrossMonths(year, "F", 3);
  const companyIncomeFormula = generateSumFormulaAcrossMonths(year, "H", 3);
  freelanceIncomeRange.setFormula(freelanceIncomeFormula);
  companyIncomeRange.setFormula(companyIncomeFormula);

  const totalExpensesRange = sheet.getRange("C2");
  totalExpensesRange.setFormula("=SUM(B2:B)");

  const totalIncomeRange = sheet.getRange("F2");
  totalIncomeRange.setFormula("=SUM(D2+E2)");
}

/**
 * Sets up the structure and formatting for a month sheet.
 *
 * @param {Sheet} sheet The sheet to set up as a month sheet.
 * @param {Array<string>} expenseCategories The list of expense categories.
 */
function createMonthSheet(sheet, expenseCategories) {
  setupMonthHeader(sheet);
  setupExpenseCategories(sheet, expenseCategories);
  setupFreelanceIncomeHeader(sheet);
  setupCompanyIncomeHeader(sheet);
  setupMonthlySummary(sheet, expenseCategories);
}

/**
 * Sets up the header row for the month sheet.
 */
function setupMonthHeader(sheet) {
  const headerRange = sheet.getRange("A1:B1");
  headerRange.setValues([["Expense Category", "Amount"]]);
  headerRange.setFontWeight("bold");
}

/**
 * Sets up the expense categories for the month sheet.
 */
function setupExpenseCategories(sheet, expenseCategories) {
  const expenseCategoriesRange = sheet.getRange("A2:A" + (expenseCategories.length + 1));
  expenseCategoriesRange.setValues(expenseCategories.map(category => [category]));
}

/**
 * Sets up the header for freelance income in the month sheet.
 */
function setupFreelanceIncomeHeader(sheet) {
  const freelanceIncomeHeaderRange = sheet.getRange("E1:F1");
  freelanceIncomeHeaderRange.setValues([["Freelance Income", ""]]);
  freelanceIncomeHeaderRange.setFontWeight("bold");

  const freelanceIncomeSubHeaderRange = sheet.getRange("E2:F2");
  freelanceIncomeSubHeaderRange.setValues([["Source", "Amount"]]);
  freelanceIncomeSubHeaderRange.setFontWeight("bold");
}

/**
 * Sets up the header for company income in the month sheet.
 */
function setupCompanyIncomeHeader(sheet) {
  const companyIncomeHeaderRange = sheet.getRange("G1:H1");
  companyIncomeHeaderRange.setValues([["Company Income", ""]]);
  companyIncomeHeaderRange.setFontWeight("bold");

  const companyIncomeSubHeaderRange = sheet.getRange("G2:H2");
  companyIncomeSubHeaderRange.setValues([["Company", "Amount"]]);
  companyIncomeSubHeaderRange.setFontWeight("bold");
}

/**
 * Sets up the monthly summary section in the month sheet.
 */
function setupMonthlySummary(sheet, expenseCategories) {
  const summaryHeaderRange = sheet.getRange("C" + (expenseCategories.length + 4) + ":D" + (expenseCategories.length + 4));
  summaryHeaderRange.setValues([["Total Income", "Total Expenses"]]);
  summaryHeaderRange.setFontWeight("bold");

  const totalIncomeFormula = "=SUM(F3:F) + SUM(H3:H)";
  const totalExpensesFormula = "=SUM(B2:B)";
  sheet.getRange("C" + (expenseCategories.length + 5) + ":D" + (expenseCategories.length + 5)).setFormulas([
    [totalIncomeFormula, totalExpensesFormula]
  ]);
}
