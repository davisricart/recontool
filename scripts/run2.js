/**
 * Excel Reconciliation Script - JavaScript implementation for comparing payment data
 * This function compares and reconciles data from two Excel files/sheets:
 * 1. Payments Hub Transaction
 * 2. Sales Totals
 */
async function compareAndDisplayData(XLSX, file1Data, file2Data) {
  try {
    // Load workbooks with all options to ensure proper reading
    const workbook1 = XLSX.read(file1Data, {
      cellStyles: true,
      cellFormulas: true,
      cellDates: true,
      cellNF: true,
      sheetStubs: true
    });
    
    const workbook2 = XLSX.read(file2Data, {
      cellStyles: true,
      cellFormulas: true,
      cellDates: true,
      cellNF: true,
      sheetStubs: true
    });

    // Get sheets from the workbooks
    const paymentsHubSheet = workbook1.Sheets[workbook1.SheetNames[0]];
    const salesTotalsSheet = workbook2.Sheets[workbook2.SheetNames[0]];

    // Convert sheets to JSON for easier processing
    const paymentsHubData = XLSX.utils.sheet_to_json(paymentsHubSheet, {
      header: 1,
      defval: "",
      raw: false
    });
    
    const salesTotalsData = XLSX.utils.sheet_to_json(salesTotalsSheet, {
      header: 1,
      defval: "",
      raw: false
    });

    // Clean and normalize headers
    const cleanHeaders = (headers) => {
      return headers.map(header => 
        header ? header.toString().trim() : ""
      );
    };

    if (paymentsHubData.length > 0) {
      paymentsHubData[0] = cleanHeaders(paymentsHubData[0]);
    }
    
    if (salesTotalsData.length > 0) {
      salesTotalsData[0] = cleanHeaders(salesTotalsData[0]);
    }

    // Find column indices in Payments Hub data
    const dateColIndex = findColumnIndex(paymentsHubData[0], "Date");
    const customerNameColIndex = findColumnIndex(paymentsHubData[0], "Customer Name");
    const totalAmountColIndex = findColumnIndex(paymentsHubData[0], "Total Transaction Amount");
    const discountingAmountColIndex = findColumnIndex(paymentsHubData[0], "Cash Discounting Amount");
    const cardBrandColIndex = findColumnIndex(paymentsHubData[0], "Card Brand");

    // Calculate K-R (Total Transaction Amount - Cash Discounting Amount)
    const paymentsHubWithKR = paymentsHubData.map((row, index) => {
      if (index === 0) {
        // Add header for K-R
        return [...row, "K-R"];
      } else if (row.length > 0) {
        // Calculate K-R value for data rows
        const totalAmount = parseFloat(row[totalAmountColIndex]) || 0;
        const discountAmount = parseFloat(row[discountingAmountColIndex]) || 0;
        const krValue = totalAmount - discountAmount;
        return [...row, krValue];
      }
      return row;
    });

    // Process Sales Totals data
    const salesDateColIndex = findColumnIndex(salesTotalsData[0], "Date");
    const salesCardBrandColIndex = findColumnIndex(salesTotalsData[0], "Card Brand");
    
    // Find or create Amount column in Sales Totals
    let salesAmountColIndex = findColumnIndex(salesTotalsData[0], "Amount");
    let processedSalesData = salesTotalsData;
    
    if (salesAmountColIndex === -1) {
      // Look for E column as a fallback, similar to VBA macro logic
      const salesEColIndex = findColumnIndex(salesTotalsData[0], "E");
      
      if (salesEColIndex !== -1) {
        // Add Amount column using E values
        processedSalesData = salesTotalsData.map((row, index) => {
          if (index === 0) {
            return [...row, "Amount"];
          } else if (row.length > 0) {
            const eValue = parseFloat(row[salesEColIndex]) || 0;
            return [...row, eValue];
          }
          return row;
        });
        salesAmountColIndex = processedSalesData[0].length - 1;
      }
    }

    // Add Count column to track matching records
    const krColIndex = paymentsHubWithKR[0].length - 1;
    const paymentsHubWithCount = paymentsHubWithKR.map((row, index) => {
      if (index === 0) {
        // Add header for Count
        return [...row, "Count"];
      } else if (row.length > 0) {
        // Calculate Count value by matching records
        let count = 0;
        
        // Count matching rows between the two sheets
        for (let i = 1; i < processedSalesData.length; i++) {
          const salesRow = processedSalesData[i];
          if (
            salesRow && 
            salesRow[salesDateColIndex] && 
            salesRow[salesCardBrandColIndex] &&
            row[dateColIndex] && 
            row[cardBrandColIndex] &&
            formatCompareValue(row[dateColIndex]) === formatCompareValue(salesRow[salesDateColIndex]) &&
            formatCompareValue(row[cardBrandColIndex]) === formatCompareValue(salesRow[salesCardBrandColIndex])
          ) {
            count++;
          }
        }
        
        return [...row, count];
      }
      return row;
    });

    // Filter to create the final data, keeping only rows with Count > 0
    const countColIndex = paymentsHubWithCount[0].length - 1;
    const filteredRows = paymentsHubWithCount.filter((row, index) => {
      return index === 0 || (row[countColIndex] && parseInt(row[countColIndex]) > 0);
    });

    // Select visible columns as in the VBA (date, customer name, total amount, discount, card brand, K-R)
    const finalData = filteredRows.map(row => {
      return [
        row[dateColIndex],
        row[customerNameColIndex],
        formatCurrency(row[totalAmountColIndex]),
        formatCurrency(row[discountingAmountColIndex]),
        row[cardBrandColIndex],
        formatCurrency(row[krColIndex]) // K-R
      ];
    });

    // Update header for the K-R column
    if (finalData.length > 0) {
      finalData[0][5] = "Total (-) Fee";
    }

    // Calculate card brand totals for reconciliation
    const paymentsHubTotals = calculateCardTotals(paymentsHubWithCount, cardBrandColIndex, krColIndex);
    const salesTotals = calculateCardTotals(processedSalesData, salesCardBrandColIndex, salesAmountColIndex);

    // Create reconciliation summary 
    const summaryData = [
      ["Hub Report", "Total", "", "Sales Report", "Total", "", "Difference"],
      [
        "Visa", 
        formatCurrency(paymentsHubTotals['visa'] || 0), 
        "", 
        "Visa", 
        formatCurrency(salesTotals['visa'] || 0), 
        "", 
        formatCurrency((paymentsHubTotals['visa'] || 0) - (salesTotals['visa'] || 0))
      ],
      [
        "Mastercard", 
        formatCurrency(paymentsHubTotals['mastercard'] || 0), 
        "", 
        "Mastercard", 
        formatCurrency(salesTotals['mastercard'] || 0), 
        "", 
        formatCurrency((paymentsHubTotals['mastercard'] || 0) - (salesTotals['mastercard'] || 0))
      ],
      [
        "American Express", 
        formatCurrency(paymentsHubTotals['american express'] || 0), 
        "", 
        "American Express", 
        formatCurrency(salesTotals['american express'] || 0), 
        "", 
        formatCurrency((paymentsHubTotals['american express'] || 0) - (salesTotals['american express'] || 0))
      ],
      [
        "Discover", 
        formatCurrency(paymentsHubTotals['discover'] || 0), 
        "", 
        "Discover", 
        formatCurrency(salesTotals['discover'] || 0), 
        "", 
        formatCurrency((paymentsHubTotals['discover'] || 0) - (salesTotals['discover'] || 0))
      ]
    ];

    // Combine finalData with summaryData into a single result dataset
    const maxRows = Math.max(finalData.length, summaryData.length);
    const resultData = [];
    
    for (let i = 0; i < maxRows; i++) {
      const finalRow = i < finalData.length ? 
        finalData[i] : Array(finalData[0] ? finalData[0].length : 6).fill("");
      const summaryRow = i < summaryData.length ? 
        summaryData[i] : Array(summaryData[0] ? summaryData[0].length : 7).fill("");

      resultData.push([...finalRow, "", ...summaryRow]);
    }

    return resultData;

  } catch (error) {
    console.error("Error processing data:", error);
    return [
      ["Error processing data: " + error.message]
    ];
  }
}

/**
 * Helper function to find the index of a column by name
 */
function findColumnIndex(headerRow, columnName) {
  if (!headerRow) return -1;
  
  return headerRow.findIndex(header =>
    header && header.toString().toLowerCase() === columnName.toLowerCase()
  );
}

/**
 * Helper function to calculate total by card brand
 */
function calculateCardTotals(data, cardBrandColIndex, amountColIndex) {
  const totals = {};
  
  if (cardBrandColIndex === -1 || amountColIndex === -1) {
    return totals;
  }
  
  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    if (row && row.length > Math.max(cardBrandColIndex, amountColIndex)) {
      if (row[cardBrandColIndex]) {
        const cardBrand = formatCompareValue(row[cardBrandColIndex]);
        const amount = parseFloat(row[amountColIndex]) || 0;
        
        if (!isNaN(amount)) {
          totals[cardBrand] = (totals[cardBrand] || 0) + amount;
        }
      }
    }
  }
  
  return totals;
}

/**
 * Helper function to format values for comparison
 */
function formatCompareValue(value) {
  if (value === null || value === undefined) return '';
  return value.toString().toLowerCase().trim();
}

/**
 * Helper function to format currency values
 */
function formatCurrency(value) {
  if (value === null || value === undefined || isNaN(parseFloat(value))) {
    return '';
  }
  
  const numValue = parseFloat(value);
  return numValue.toFixed(2);
}