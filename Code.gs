function doGet(e) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Sheet1");
  const data = sheet.getDataRange().getValues();

  const accounts = [];
  let currentAccount = null;
  let totalInvested = 0;
  let totalValue = 0;

  const accountNames = ["Brokerage", "HSA", "Simple IRA", "Roth IRA", "Stocks", "Crypto"];

  for (let i = 0; i < data.length; i++) {
    const row = data[i];
    const col_b = String(row[1]).trim();
    const col_c = String(row[2]).trim();
    const col_d = row[3];
    const col_e = row[4];
    const col_f = row[5];
    const col_g = row[6];
    const col_h = row[7];

    const matchedName = accountNames.find(n => col_b === n || col_c === n);
    if (matchedName) {
      currentAccount = { name: matchedName, funds: [], invested: 0, value: 0, profit: 0, returnPct: 0 };
      accounts.push(currentAccount);
      continue;
    }

    if (col_c === "Total" && currentAccount) {
      const inv = parseFloat(String(col_e).replace(/[$,]/g, "")) || 0;
      const val = parseFloat(String(col_g).replace(/[$,]/g, "")) || 0;
      currentAccount.invested = inv;
      currentAccount.value = val;
      currentAccount.profit = val - inv;
      currentAccount.returnPct = inv > 0 ? ((val - inv) / inv * 100) : 0;
      totalInvested += inv;
      totalValue += val;
      continue;
    }

    if (currentAccount && col_b && col_b !== "" && col_b !== "Fund Name" && col_c !== "Total" && col_b !== "Name") {
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
