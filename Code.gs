/**
 * @OnlyCurrentDoc
 *
 * Reference: https://developers.google.com/apps-script/guides/services/authorization
 */

 function onOpen() {
  const ui = SpreadsheetApp.getUi();

  ui.createMenu('💲 Rebalance')
    .addItem('❓ Help', 'showHelp')
    .addItem('📁 Import data from Schwab', 'showImportDialog')
    .addItem('🧮 Calculate buy', 'calculateBuy')
    .addToUi();
}

function showHelp() {
  const ui = SpreadsheetApp.getUi();

  const helpMessage = `Welcome to rebalance-portfolio!
  
  DATA ENTRY:
  * If you use Schwab, press "💲 Rebalance > 📁 Import data from Schwab" to import a CSV file with your positions.
  * Otherwise, manually input your portfolio by adding and removing rows above the "Cash" row.
  
  NOTES:
  * When adding a new row, remember to copy all cells down from the row above it. No formula updates are necessary, just copy the entire row down.
  * Only update the fields highlighted in yellow: "Symbol", "Quantity", "Target %", and "Fractional?". All other fields will automatically update.
  * Put any value in the "Fractional?" column if that stock is able to be purchased in fractional shares. Otherwise, leave it blank.
  
  CALCULATE BUY:
  * Press "💲 Rebalance > 🧮 Calculate buy" to calculate the optimized purchases using a linear optimization algorithm.`;

  ui.alert('Help', helpMessage, ui.ButtonSet.OK);
}

function showImportDialog() {
  const ui = SpreadsheetApp.getUi();
  const html = HtmlService.createHtmlOutputFromFile('import').setWidth(500).setHeight(250);
  ui.showModalDialog(html, 'Import CSV from Schwab');
}

// This function is called client-side from the import dialog
function importSchwabData(data) {
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();

  // Extract the existing portfolio's target % values
  const oldPortfolio = spreadsheet.getSheetByName('Portfolio');
  const oldCashRow = oldPortfolio.createTextFinder('Cash').findNext().getRow(); // Assuming 'Cash' does not appear anywhere else on the portfolio
  const oldStocks = oldPortfolio.getRange(`B4:G${oldCashRow}`).getValues(); // Hardcoded range assuming a fixed template
  const oldTargets = new Map();
  oldStocks.forEach(r => oldTargets.set(r[0], r[5])); // Hardcoded Symbol and Target % indices (0 and 5)

  // Set up the new portfolio sheet and new values
  const template = spreadsheet.getSheetByName('Template'); // Hidden sheet (!!!) MAXIMUM 100 STOCKS (!!!)
  const newPortfolio = template.copyTo(spreadsheet); // Will create a new sheet. Necessary to duplicate formatting.
  const newStocks = []; // Will replace columns B and C
  const newTargets = []; // Will replace column G

  // Remove extraneous rows from the Schwab export
  const rows = data.split('\n');
  rows.splice(0, 3); // First 3 rows are headers
  rows.splice(-2, 2); // Last 2 rows are summary rows

  // Parse the remaining rows from the Schwab export
  const csv = Utilities.parseCsv(rows.join('\n'));
  for (const row of csv) {
    if (row[0] === 'Cash & Cash Investments') { // Schwab export treats the cash row uniquely
      newStocks.push(['Cash', row[6]]);
      newTargets.push([oldTargets.get('Cash')]); // Assumes Cash target was defined in oldPortfolio
    }
    else {
      newStocks.push([row[0], row[2]]);
      newTargets.push([oldTargets.has(row[0]) ? oldTargets.get(row[0]) : 0]);
    }
  }

  // Set new portfolio values and delete excess rows (!!!) MAXIMUM 100 STOCKS (!!!)
  newPortfolio.deleteRows(4, 100 - csv.length + 1); // MUST delete before inserting
  newPortfolio.getRange(`B4:C${4 + csv.length - 1}`).setValues(newStocks);
  newPortfolio.getRange(`G4:G${4 + csv.length - 1}`).setValues(newTargets);

  // Replace oldPortfolio with newPortfolio
  newPortfolio.showSheet();
  spreadsheet.deleteSheet(oldPortfolio);
  spreadsheet.setActiveSheet(newPortfolio);
  newPortfolio.setName('Portfolio');
}

function calculateBuy() {
  // Extract Portfolio data
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  const portfolio = spreadsheet.getSheetByName('Portfolio');
  const cashRow = portfolio.createTextFinder('Cash').findNext().getRow();
  const stocks = portfolio.getRange(`B4:J${cashRow - 1}`).getValues();
  const availableCash = portfolio.getRange(`C${cashRow}`).getValue();
  const totalPortfolioValue = portfolio.getRange(`E${cashRow + 1}`).getValue();

  // Set up LinearOptimizationEngine
  const engine = LinearOptimizationService.createEngine();
  const cashConstraint = engine.addConstraint(availableCash - 5, availableCash); // Hardcoded max remaining cash of $5
  const maxDeviation = 0.001; // Hardcoded max allowable deviation of 0.1% from each target allocation
  const names = []; // To fetch the picks after solving

  for (const row of stocks) {
    const name = row[0];
    const price = row[2];
    const fractional = row[6];
    const optimal = row[8];
    names.push([name]);

    const deviation = totalPortfolioValue * maxDeviation / price;
    const lowerBound = optimal - deviation;
    const upperBound = optimal + deviation;
    const type = fractional === '' ? LinearOptimizationService.VariableType.INTEGER : LinearOptimizationService.VariableType.CONTINUOUS;
    const coefficient = -price; // Objective: minimize remaining cash

    engine.addVariable(name, lowerBound, upperBound, type, coefficient);
    cashConstraint.setCoefficient(name, price);
  }

  engine.setMinimization();
  const solution = engine.solve();

  portfolio.getRange(`T4:T${cashRow - 1}`).setValues(names.map(r => [solution.getVariableValue(r[0])]));
}
