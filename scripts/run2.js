/**
 * Excel Reconciliation Script - JavaScript implementation for comparing payment data
 * This function compares and reconciles data from two Excel files/sheets:
 * 1. Payments Hub Transaction
 * 2. Sales Totals
 * 
 * This JavaScript version replicates the VBA macro logic in the original "Sales Recon" Excel workbook
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

    // Find column indices in Sales Totals data
    const salesNameColIndex = findColumnIndex(salesTotalsData[0], "Name");
    const salesDateColIndex = findColumnIndex(salesTotalsData[0], "Date Closed");
    const salesAmountColIndex = findColumnIndex(salesTotalsData[0], "Amount");

    // Calculate K-R (Total Transaction Amount - Cash Discounting Amount)
    // This matches the VBA: "Range("AA1").FormulaR1C1 = "K-R" and Range("AA2").FormulaR1C1 = "=RC[-16]-RC[-9]""
    const paymentsHubWithKR = paymentsHubData.map((row, index) => {
      if (index === 0) {
        // Add header for K-R
        return [...row, "K-R"];
      } else if (row.length > 0) {
        // Parse currency values properly by removing $ signs and other non-numeric characters
        let totalAmount = 0;
        if (row[totalAmountColIndex]) {
          const totalStr = row[totalAmountColIndex].toString().replace(/[^\d.-]/g, "");
          totalAmount = parseFloat(totalStr) || 0;
        }
        
        let discountAmount = 0;
        if (row[discountingAmountColIndex]) {
          const discountStr = row[discountingAmountColIndex].toString().replace(/[^\d.-]/g, "");
          discountAmount = parseFloat(discountStr) || 0;
        }
        
        const krValue = totalAmount - discountAmount;
        return [...row, krValue];
      }
      return row;
    });

    // Get the K-R column index for later use
    const krColIndex = paymentsHubWithKR[0].length - 1;

    // Add Count column based on the COUNTIFS formula from the VBA:
    // "=COUNTIFS('Sales Totals'!C2,'Payments Hub Transaction'!RC1,'Sales Totals'!C1,'Payments Hub Transaction'!RC24,'Sales Totals'!C5,'Payments Hub Transaction'!RC27)"
    // This corresponds to matching the Sales Totals card type, date, and amount with Payments Hub data
    const paymentsHubWithCount = paymentsHubWithKR.map((row, index) => {
      if (index === 0) {
        // Add header for Count
        return [...row, "Count"];
      } else if (row.length > 0) {
        // Get the date, card brand, and KR value from the current row
        const hubDate = formatDate(row[dateColIndex]);
        const hubCardBrand = normalize(row[cardBrandColIndex]);
        const hubKR = row[krColIndex] !== undefined ? parseFloat(row[krColIndex]) : 0;
        
        // Count matching records in Sales Totals
        let count = 0;
        
        for (let i = 1; i < salesTotalsData.length; i++) {
          const salesRow = salesTotalsData[i];
          if (salesRow.length <= Math.max(salesNameColIndex, salesDateColIndex, salesAmountColIndex)) {
            continue;
          }
          
          const salesCardType = normalize(salesRow[salesNameColIndex]);
          const salesDate = formatDate(salesRow[salesDateColIndex]);
          let salesAmount = 0;
          
          if (salesRow[salesAmountColIndex]) {
            const amountStr = salesRow[salesAmountColIndex].toString().replace(/[^\d.-]/g, "");
            salesAmount = parseFloat(amountStr) || 0;
          }
          
          // Match on date, card type and amount (with small tolerance for floating point differences)
          if (hubDate === salesDate && 
              hubCardBrand === salesCardType && 
              Math.abs(hubKR - salesAmount) < 0.01) {
            count++;
          }
        }
        
        return [...row, count];
      }
      return row;
    });

    // Filter rows where Count = 0 (this matches the VBA filter operation)
    // "Columns("AB:AB").AutoFilter Field:=28, Criteria1:="0""
    const countColIndex = paymentsHubWithCount[0].length - 1;
    
    // Create the Final data by filtering rows where Count = 0
    const finalRows = [paymentsHubWithCount[0]]; // Always include header row
    
    for (let i = 1; i < paymentsHubWithCount.length; i++) {
      const row = paymentsHubWithCount[i];
      if (row.length > countColIndex) {
        const countValue = parseInt(row[countColIndex]) || 0;
        if (countValue === 0) {
          finalRows.push(row);
        }
      }
    }

    // Sort by Total Transaction Amount descending (similar to VBA sort)
    finalRows.sort((a, b) => {
      // Keep header row at the top
      if (a === finalRows[0]) return -1;
      if (b === finalRows[0]) return 1;
      
      // Parse amounts for sorting
      let amountA = 0;
      if (a[totalAmountColIndex]) {
        const amountStrA = a[totalAmountColIndex].toString().replace(/[^\d.-]/g, "");
        amountA = parseFloat(amountStrA) || 0;
      }
      
      let amountB = 0;
      if (b[totalAmountColIndex]) {
        const amountStrB = b[totalAmountColIndex].toString().replace(/[^\d.-]/g, "");
        amountB = parseFloat(amountStrB) || 0;
      }
      
      return amountB - amountA; // Descending order
    });

    // Create the final output data with selected columns
    const finalData = finalRows.map((row, index) => {
      if (index === 0) {
        // Use the exact header names from the VBA code
        return [
          "Date",
          "Customer Name",
          "Total Transaction Amount",
          "Cash Discounting Amount",
          "Card Brand",
          "Total (-) Fee" // This is the K-R column
        ];
      } else {
        // Format values for display
        return [
          formatDateForDisplay(row[dateColIndex]),
          row[customerNameColIndex] || "",
          formatCurrencyString(row[totalAmountColIndex]),
          formatCurrencyString(row[discountingAmountColIndex]),
          row[cardBrandColIndex] || "",
          formatCurrencyString(row[krColIndex])
        ];
      }
    });

    // Calculate card brand totals for the summary section
    // First for Payments Hub data
    const paymentsHubTotals = {
      visa: 0,
      mastercard: 0,
      "american express": 0
    };
    
    for (let i = 1; i < paymentsHubWithCount.length; i++) {
      const row = paymentsHubWithCount[i];
      if (row.length > Math.max(cardBrandColIndex, krColIndex)) {
        const cardBrand = normalize(row[cardBrandColIndex]);
        const amount = parseFloat(row[krColIndex]) || 0;
        
        if (cardBrand === "visa") {
          paymentsHubTotals.visa += amount;
        } else if (cardBrand === "mastercard") {
          paymentsHubTotals.mastercard += amount;
        } else if (cardBrand === "american express") {
          paymentsHubTotals["american express"] += amount;
        }
      }
    }
    
    // Then for Sales Totals data
    const salesTotals = {
      visa: 0,
      mastercard: 0,
      "american express": 0
    };
    
    for (let i = 1; i < salesTotalsData.length; i++) {
      const row = salesTotalsData[i];
      if (row.length > Math.max(salesNameColIndex, salesAmountColIndex)) {
        const cardType = normalize(row[salesNameColIndex]);
        let amount = 0;
        
        if (row[salesAmountColIndex]) {
          const amountStr = row[salesAmountColIndex].toString().replace(/[^\d.-]/g, "");
          amount = parseFloat(amountStr) || 0;
        }
        
        if (cardType === "visa") {
          salesTotals.visa += amount;
        } else if (cardType === "mastercard") {
          salesTotals.mastercard += amount;
        } else if (cardType === "american express") {
          salesTotals["american express"] += amount;
        }
      }
    }
    
    // Calculate differences
    const differences = {
      visa: paymentsHubTotals.visa - salesTotals.visa,
      mastercard: paymentsHubTotals.mastercard - salesTotals.mastercard,
      "american express": paymentsHubTotals["american express"] - salesTotals["american express"]
    };

    // Create summary section data - matches the VBA section that creates I1:O4
    const summaryData = [
      ["Hub Report", "Total", "", "Hub Report", "Total", "", "Difference"],
      [
        "Visa", 
        formatCurrencyString(paymentsHubTotals.visa), 
        "", 
        "Visa", 
        formatCurrencyString(salesTotals.visa), 
        "", 
        formatCurrencyString(differences.visa)
      ],
      [
        "Mastercard", 
        formatCurrencyString(paymentsHubTotals.mastercard), 
        "", 
        "Mastercard", 
        formatCurrencyString(salesTotals.mastercard), 
        "", 
        formatCurrencyString(differences.mastercard)
      ],
      [
        "American Express", 
        formatCurrencyString(paymentsHubTotals["american express"]), 
        "", 
        "American Express", 
        formatCurrencyString(salesTotals["american express"]), 
        "", 
        formatCurrencyString(differences["american express"])
      ]
    ];

    // Combine finalData with summaryData into a single result dataset
    // This matches the expected return format in the original code
    const resultData = [];
    
    // Determine the max number of rows needed
    const maxRows = Math.max(finalData.length, summaryData.length);
    
    // Combine the data horizontally with spacing in between
    for (let i = 0; i < maxRows; i++) {
      const finalRow = i < finalData.length ? 
        finalData[i] : Array(finalData[0] ? finalData[0].length : 6).fill("");
      const summaryRow = i < summaryData.length ? 
        summaryData[i] : Array(summaryData[0] ? summaryData[0].length : 7).fill("");

      resultData.push([...finalRow, "", ...summaryRow]);
    }

    // Return the combined data as a flat array
    return resultData;

  } catch (error) {
    console.error("Error processing data:", error);
    // Return error as a single row array to match expected format
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
    header && header.toString().toLowerCase().trim() === columnName.toLowerCase().trim()
  );
}

/**
 * Helper function to normalize text values for comparison
 */
function normalize(value) {
  if (!value) return "";
  return value.toString().toLowerCase().trim();
}

/**
 * Helper function to format date values for comparison
 * Converts various date formats to a standard MM/DD/YY format
 */
function formatDate(dateStr) {
  if (!dateStr) return "";
  
  let date;
  const dateString = dateStr.toString().trim();
  
  // Try to parse the date using various formats
  if (dateString.includes('/')) {
    // Handle format like MM/DD/YYYY or MM/DD/YY
    const parts = dateString.split('/');
    if (parts.length >= 3) {
      const month = parts[0].padStart(2, '0');
      const day = parts[1].padStart(2, '0');
      let year = parts[2];
      if (year.length > 2) {
        year = year.substring(year.length - 2);
      }
      return `${month}/${day}/${year}`;
    }
  }
  
  // Fallback to Date object parsing
  try {
    date = new Date(dateString);
    if (!isNaN(date.getTime())) {
      const month = String(date.getMonth() + 1).padStart(2, '0');
      const day = String(date.getDate()).padStart(2, '0');
      // Get last 2 digits of year
      const year = String(date.getFullYear()).slice(-2);
      return `${month}/${day}/${year}`;
    }
  } catch (e) {
    console.error("Error parsing date:", e);
  }
  
  return dateString; // Return original if parsing fails
}

/**
 * Format date for display in the final output
 * Adds "0:00" time component to match the VBA output format
 */
function formatDateForDisplay(dateStr) {
  if (!dateStr) return "";
  
  const formattedDate = formatDate(dateStr);
  // Add the time component (0:00) if it's not already there
  if (!formattedDate.includes(':')) {
    return `${formattedDate} 0:00`;
  }
  return formattedDate;
}

/**
 * Helper function to format currency values with $ sign and proper decimal places
 */
function formatCurrencyString(value) {
  if (value === undefined || value === null || value === "") return "";
  
  let numValue;
  if (typeof value === 'string') {
    // Remove any non-numeric characters except decimal point and negative sign
    const numStr = value.replace(/[^\d.-]/g, "");
    numValue = parseFloat(numStr);
  } else {
    numValue = parseFloat(value);
  }
  
  if (isNaN(numValue)) return "";
  
  // Format with 2 decimal places and $ sign
  return `$${numValue.toFixed(2)}`;
}