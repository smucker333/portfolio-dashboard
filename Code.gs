function doGet(e) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Investments - Portfolio");
  const data = sheet.getDataRange().getValues();

  const accounts = [];
  let currentAccount = null;
  let totalInvested = 0;
  let totalValue = 0;
  let inCashSection = false;

  const investmentAccountNames = ["Brokerage", "HSA", "Simple IRA", "Roth IRA", "Stocks", "Crypto"];
  const skipRows = ["Fund Name", "Name", "High Yield Savings Fund", "Investing", ""];

  for (let i = 0; i < data.length; i++) {
    const row = data[i];
    const col_b = String(row[1]).trim();
    const col_c = String(row[2]).trim();
    const col_d = row[3];
    const col_e = row[4];
    const col_f = row[5];
    const col_g = row[6];
    const col_h = row[7];

    // Detect Cash section header
    if (col_b === "Cash" || col_c === "Cash") {
      inCashSection = true;
      currentAccount = { name: "Cash", funds: [], invested: 0, value: 0, profit: 0, returnPct: 0 };
      accounts.push(currentAccount);
      continue;
    }

    // Detect investment account headers
    const matchedName = investmentAccountNames.find(n => col_b === n || col_c === n);
    if (matchedName) {
      inCashSection = false;
      currentAccount = { name: matchedName, funds: [], invested: 0, value: 0, profit: 0, returnPct: 0 };
      accounts.push(currentAccount);
      continue;
    }

    // Skip header/label rows
    if (skipRows.includes(col_b) || col_b.startsWith("Fund")) continue;

    // Handle Total rows
    if ((col_b === "Total" || col_c === "Total") && currentAccount) {
      if (inCashSection) {
        // Cash total: col_e = invested, col_f = total value (different layout)
        const inv = parseFloat(String(col_e).replace(/[$,]/g, "")) || 0;
        const val = parseFloat(String(col_f).replace(/[$,]/g, "")) || 0;
        currentAccount.invested = inv;
        currentAccount.value = val;
        currentAccount.profit = val - inv;
        currentAccount.returnPct = inv > 0 ? ((val - inv) / inv * 100) : 0;
        totalInvested += inv;
        totalValue += val;
      } else {
        const inv = parseFloat(String(col_e).replace(/[$,]/g, "")) || 0;
        const val = parseFloat(String(col_g).replace(/[$,]/g, "")) || 0;
        currentAccount.invested = inv;
        currentAccount.value = val;
        currentAccount.profit = val - inv;
        currentAccount.returnPct = inv > 0 ? ((val - inv) / inv * 100) : 0;
        totalInvested += inv;
        totalValue += val;
      }
      continue;
    }

    // Skip grand total row
    if (col_b === "Total" && col_c === "") continue;

    // Add fund rows
    if (currentAccount && col_b && col_b !== "") {
      if (inCashSection) {
        // Cash layout: B=ticker, E=invested, F=total value, G=profit, H=%
        const invested = parseFloat(String(col_e).replace(/[$,]/g, "")) || 0;
        const value    = parseFloat(String(col_f).replace(/[$,]/g, "")) || 0;
        const profit   = parseFloat(String(col_g).replace(/[$,]/g, "")) || (value - invested);
        const pct      = invested > 0 ? ((value - invested) / invested * 100) : 0;
        if (invested > 0 || value > 0) {
          currentAccount.funds.push({
            ticker: col_b,
            description: "High yield savings",
            price: 0,
            invested: invested,
            owned: 0,
            value: value,
            profit: profit,
            returnPct: pct
          });
        }
      } else {
        // Normal layout: B=ticker, C=desc, D=price, E=invested, F=owned, G=value, H=profit
        const price    = parseFloat(String(col_d).replace(/[$,]/g, "")) || 0;
        const invested = parseFloat(String(col_e).replace(/[$,]/g, "")) || 0;
        const owned    = parseFloat(String(col_f).replace(/[,$]/g, "")) || 0;
        const value    = parseFloat(String(col_g).replace(/[$,]/g, "")) || 0;
        const profit   = parseFloat(String(col_h).replace(/[$,]/g, "")) || (value - invested);
        const pct      = invested > 0 ? ((value - invested) / invested * 100) : 0;
        if (invested > 0 || value > 0) {
          currentAccount.funds.push({
            ticker: col_b,
            description: col_c,
            price: price,
            invested: invested,
            owned: owned,
            value: value,
            profit: profit,
            returnPct: pct
          });
        }
      }
    }
  }

  const result = {
    accounts: accounts,
    summary: {
      totalInvested: totalInvested,
      totalValue: totalValue,
      totalProfit: totalValue - totalInvested,
      returnPct: totalInvested > 0 ? ((totalValue - totalInvested) / totalInvested * 100) : 0
    },
    lastUpdated: new Date().toISOString()
  };

  const json = JSON.stringify(result);
  const callback = e && e.parameter && e.parameter.callback;
  const output = callback ? callback + "(" + json + ");" : json;
  const mimeType = callback ? ContentService.MimeType.JAVASCRIPT : ContentService.MimeType.JSON;

  return ContentService.createTextOutput(output).setMimeType(mimeType);
}
