function createCumulativeUnitsDataFrameAndDisplay() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Input');
  if (!sheet) {
    throw new Error("Sheet named 'Input' not found.");
  }

  const data = sheet.getDataRange().getValues();
  if (data.length < 2) {
    throw new Error("No data found in the sheet. Ensure the 'Input' sheet contains both headers and data rows");
  }

  const headers = data[0].map(header => header.trim());
  const emailIdIndex = headers.indexOf('Email_Id');
  const clientNameIndex = headers.indexOf('Client_Name');
  const tickerIndex = headers.indexOf('Ticker');
  const transactionTypeIndex = headers.indexOf('Transaction_Type');
  const unitsIndex = headers.indexOf('Units');
  const dateIndex = headers.indexOf('Date');
  const currencyIndex = headers.indexOf('Currency');
  const reportingCurrencyIndex = headers.indexOf('Reporting_Currency');

  if (
    [emailIdIndex, clientNameIndex, tickerIndex, transactionTypeIndex, unitsIndex, dateIndex].some(
      (index) => index === -1
    )
  ) {
    throw new Error("One or more required columns are missing in the 'Input' sheet.");
  }

  const cumulativeUnitsByClient = {}; // Object to store cumulative units for each client
  const earliestTransactionDates = {}; // Track the earliest transaction date for each client
  const clientTickerCurrencyMap = {}; // Track currency mapping per client and ticker

  const ui = SpreadsheetApp.getUi();
  const clientResponse = ui.prompt("Client Data", "Enter Client Name:", ui.ButtonSet.OK_CANCEL);
  const emailResponse = ui.prompt("Client Data", "Enter Client Email ID:", ui.ButtonSet.OK_CANCEL);

  if (clientResponse.getSelectedButton() !== ui.Button.OK || emailResponse.getSelectedButton() !== ui.Button.OK) {
    ui.alert("Operation cancelled.");
    return;
  }

  const clientNameFilter = clientResponse.getResponseText().trim().toLowerCase();
  const emailIdFilter = emailResponse.getResponseText().trim().toLowerCase();

  const clientKeyFilter = `${clientNameFilter}_${emailIdFilter}`;

  // Validate if client details exist in the data
  const validClient = data.slice(1).some(row => {
    const emailId = row[emailIdIndex]?.toString().trim().toLowerCase();
    const clientName = row[clientNameIndex]?.toString().trim().toLowerCase();
    return clientName === clientNameFilter && emailId === emailIdFilter;
  });

  if (!validClient) {
    ui.alert("No matching client details found. Please verify the Client Name and Email ID.");
    return;
  }

  // Step 1: Parse transactions and calculate cumulative units
  data.slice(1).forEach((row) => {
    const emailId = row[emailIdIndex]?.toString().trim().toLowerCase();
    const clientName = row[clientNameIndex]?.toString().trim().toLowerCase();
    const ticker = row[tickerIndex]?.toString().trim();
    const transactionType = row[transactionTypeIndex]?.toString().trim().toLowerCase();
    const units = parseFloat(row[unitsIndex]) || 0;
    const transactionDate = new Date(row[dateIndex]);
    const currency = row[currencyIndex]?.toString().trim();
    const reportingCurrency = row[reportingCurrencyIndex]?.toString().trim();

    if (isNaN(transactionDate)) {
      return; // Skip rows with invalid dates
    }

    const dateStr = Utilities.formatDate(
      transactionDate,
      Session.getScriptTimeZone(),
      'yyyy-MM-dd'
    );

    const clientKey = `${clientName}_${emailId}`;

    if (!cumulativeUnitsByClient[clientKey]) {
      cumulativeUnitsByClient[clientKey] = {};
      earliestTransactionDates[clientKey] = transactionDate;
    }

    if (!clientTickerCurrencyMap[clientKey]) {
      clientTickerCurrencyMap[clientKey] = {};
    }

    if (!clientTickerCurrencyMap[clientKey][ticker]) {
      clientTickerCurrencyMap[clientKey][ticker] = { base: currency, reporting: reportingCurrency };
    }

    if (!cumulativeUnitsByClient[clientKey][ticker]) {
      cumulativeUnitsByClient[clientKey][ticker] = {};
    }

    if (!cumulativeUnitsByClient[clientKey][ticker][dateStr]) {
      cumulativeUnitsByClient[clientKey][ticker][dateStr] = 0;
    }

    // Update cumulative units based on transaction type
    if (transactionType === 'buy') {
      cumulativeUnitsByClient[clientKey][ticker][dateStr] += units;
    } else if (transactionType === 'sell') {
      cumulativeUnitsByClient[clientKey][ticker][dateStr] -= units;
    } else if (transactionType === 'dividend payout') {
      // Dividend payouts are neutral in terms of units; no change
    }

    // Update earliest transaction date if necessary
    if (transactionDate < earliestTransactionDates[clientKey]) {
      earliestTransactionDates[clientKey] = transactionDate;
    }
  });

  // Step 2: Fill missing dates and carry forward cumulative units
  const today = new Date();
  for (const clientKey in cumulativeUnitsByClient) {
    if (clientKey !== clientKeyFilter) {
      continue;
    }

    const startDate = new Date(earliestTransactionDates[clientKey]);
    for (const ticker in cumulativeUnitsByClient[clientKey]) {
      const dates = Object.keys(cumulativeUnitsByClient[clientKey][ticker]).sort(
        (a, b) => new Date(a) - new Date(b)
      );

      let currentDate = new Date(startDate);
      let lastKnownUnits = 0;

      while (currentDate <= today) {
        const dateStr = Utilities.formatDate(
          currentDate,
          Session.getScriptTimeZone(),
          'yyyy-MM-dd'
        );

        if (dates.includes(dateStr)) {
          lastKnownUnits += cumulativeUnitsByClient[clientKey][ticker][dateStr];
          cumulativeUnitsByClient[clientKey][ticker][dateStr] = lastKnownUnits;
        } else {
          cumulativeUnitsByClient[clientKey][ticker][dateStr] = lastKnownUnits;
        }

        currentDate.setDate(currentDate.getDate() + 1);
      }
    }
  }

  // Step 3: Write cumulative units and calculate AUM
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();

  for (const clientKey in cumulativeUnitsByClient) {
    if (clientKey !== clientKeyFilter) {
      continue;
    }

    const [clientName, emailId] = clientKey.split('_');
    let clientSheet = spreadsheet.getSheetByName(clientName);

    if (!clientSheet) {
      clientSheet = spreadsheet.insertSheet(clientName);
    } else {
      clientSheet.clear();
    }

    const tickers = Object.keys(cumulativeUnitsByClient[clientKey]);
    const headers = ['Date', ...tickers, 'AUM'];
    clientSheet.appendRow(headers);

    // Set header style to bold
    const headerRange = clientSheet.getRange(1, 1, 1, headers.length);
    headerRange.setFontWeight("bold");

    // Collect all unique dates
    const allDates = new Set();
    tickers.forEach((ticker) => {
      Object.keys(cumulativeUnitsByClient[clientKey][ticker]).forEach((date) =>
        allDates.add(date)
      );
    });

    const sortedDates = Array.from(allDates).sort((a, b) => new Date(a) - new Date(b));

    // Fetch historical prices for tickers and exchange rate
    const historicalPrices = fetchHistoricalPricesUsingGoogleFinance(tickers, sortedDates);
    const exchangeRates = fetchHistoricalExchangeRates(clientTickerCurrencyMap[clientKey], sortedDates); 

    // Write data to the sheet
    sortedDates.forEach((date, idx) => {
      const row = [date];
      let totalAUM = 0;

      tickers.forEach((ticker) => {
        const units = cumulativeUnitsByClient[clientKey][ticker][date] || 0;
        let price = historicalPrices[ticker][date];
        const exchangeRate = exchangeRates[ticker][date] || 1; // Default to 1 if same currency or no data

        // Use the last available price if the current date's price is missing
        if (!price) {
          const availableDates = Object.keys(historicalPrices[ticker]).sort((a, b) => new Date(a) - new Date(b));
          const previousDate = availableDates.reverse().find(d => new Date(d) < new Date(date));
          price = historicalPrices[ticker][previousDate] || 0;
        }

        const value = units * price * exchangeRate;
        totalAUM += value;
        row.push(units);
      });

      row.push(totalAUM);
      clientSheet.appendRow(row);

      // Set font to Calibri for all rows
      const range = clientSheet.getRange(idx + 2, 1, 1, headers.length);
      range.setFontFamily("Calibri");
    });

    Logger.log(
      `Cumulative units and AUM displayed for client: ${clientName} (${emailId}) in sheet: ${clientName}`
    );
  }
}

function fetchHistoricalPricesUsingGoogleFinance(tickers, dates) {
  const prices = {};

  tickers.forEach((ticker) => {
    prices[ticker] = {};
    const tempSheet = SpreadsheetApp.getActiveSpreadsheet().insertSheet(`Temp_${ticker}`);
    const formula = `=GOOGLEFINANCE("${ticker}", "price", DATE(2023,1,1), TODAY(), "daily")`;
    tempSheet.getRange("A1").setFormula(formula);
    SpreadsheetApp.flush();

    const rows = tempSheet.getDataRange().getValues();
    const datePrices = {};

    // Parse data from Google Finance
    rows.slice(1).forEach((row) => {
      const date = Utilities.formatDate(new Date(row[0]), Session.getScriptTimeZone(), 'yyyy-MM-dd');
      const price = parseFloat(row[1]);
      if (!isNaN(price)) {
        datePrices[date] = price;
      }
    });

    // Process and fill missing dates with previous price
    let lastKnownPrice = 0; // Default to 0 if no price is available
    dates.forEach((date) => {
      if (datePrices[date] !== undefined) {
        lastKnownPrice = datePrices[date];
      }
      prices[ticker][date] = lastKnownPrice; // Use the last known price
    });

    SpreadsheetApp.getActiveSpreadsheet().deleteSheet(tempSheet);
  });

  return prices;
}


function fetchHistoricalExchangeRates(clientTickerCurrencyMap, dates) {
  const exchangeRates = {};
  const tempSheet = SpreadsheetApp.getActiveSpreadsheet().insertSheet("Temp_Exchange");

  for (const ticker in clientTickerCurrencyMap) {
    const { base, reporting } = clientTickerCurrencyMap[ticker];
    exchangeRates[ticker] = {};

    if (base === reporting) {
      dates.forEach(date => exchangeRates[ticker][date] = 1); // Default 1 for same currency
      continue;
    }

    const formula = `=GOOGLEFINANCE("CURRENCY:${base}${reporting}", "price", DATE(2023,1,1), TODAY(), "daily")`;
    tempSheet.getRange("A1").setFormula(formula);
    SpreadsheetApp.flush();

    const rows = tempSheet.getDataRange().getValues();
    const dateRates = {};

    // Parse and collect data
    rows.slice(1).forEach(row => {
      const date = Utilities.formatDate(new Date(row[0]), Session.getScriptTimeZone(), 'yyyy-MM-dd');
      const rate = parseFloat(row[1]);
      if (!isNaN(rate)) {
        dateRates[date] = rate;
      }
    });

    // Fill exchangeRates with all dates
    let lastAvailableRate = 0; // Default for missing dates
    dates.forEach(date => {
      if (dateRates[date] !== undefined) {
        lastAvailableRate = dateRates[date];
      }
      exchangeRates[ticker][date] = lastAvailableRate;
    });
  }

  SpreadsheetApp.getActiveSpreadsheet().deleteSheet(tempSheet);
  return exchangeRates;
}
