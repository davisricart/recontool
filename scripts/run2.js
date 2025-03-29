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

    // Process Sales Totals data
    // For Sales Totals, the structure is different - looking for Name (card type) and Amount columns
    const salesCardTypeIndex = findColumnIndex(salesTotalsData[0], "Name");
    const salesDateIndex = findColumnIndex(salesTotalsData[0], "Date Closed");
    const salesAmountIndex = findColumnIndex(salesTotalsData[0], "Amount");
    
    // Add Count column to track matching records
    const krColIndex = paymentsHubWithKR[0].length - 1;
    const paymentsHubWithCount = paymentsHubWithKR.map((row, index) => {
      if (index === 0) {
        // Add header for Count
        return [...row, "Count"];
      } else if (row.length > 0) {
        // Count matching transactions - based on the Card Brand and date
        let count = 0;
        
        // Compare hub date with sales date, considering different formats
        const hubDate = formatDate(row[dateColIndex]);
        const cardBrand = row[cardBrandColIndex] ? row[cardBrandColIndex].toString().toLowerCase().trim() : "";
        
        for (let i = 1; i < salesTotalsData.length; i++) {
          const salesRow = salesTotalsData[i];
          if (salesRow && salesRow.length > Math.max(salesCardTypeIndex, salesDateIndex)) {
            const salesDate = formatDate(salesRow[salesDateIndex]);
            const salesCardType = salesRow[salesCardTypeIndex] ? 
              salesRow[salesCardTypeIndex].toString().toLowerCase().trim() : "";
            
            // Match based on card type and date
            if (hubDate && salesDate && salesCardType && 
                salesCardType === cardBrand && 
                salesDate === hubDate) {
              count++;
            }
          }
        }
        
        return [...row, count];
      }
      return row;
    });

    // We won't filter based on count for now
    const filteredRows = paymentsHubWithCount;

    // Select visible columns for the report
    const finalData = filteredRows.map((row, index) => {
      if (index === 0) {
        // Explicitly set all headers
        return [
          "Date",
          "Customer Name",
          "Total Transaction Amount",
          "Cash Discounting Amount",
          "Card Brand",
          "Total (-) Fee"
        ];
      } else {
        return [
          row[dateColIndex] || "",
          row[customerNameColIndex] || "",
          formatCurrencyString(row[totalAmountColIndex]),
          formatCurrencyString(row[discountingAmountColIndex]),
          row[cardBrandColIndex] || "",
          formatCurrencyString(row[krColIndex]) // K-R
        ];
      }
    });

    // Calculate card brand totals from both data sources
    const paymentsHubTotals = calculateCardTotalsFromPaymentsHub(paymentsHubWithCount, cardBrandColIndex, krColIndex);
    const salesTotals = calculateCardTotalsFromSales(salesTotalsData, salesCardTypeIndex, salesAmountIndex);

    // Calculate differences
    const differences = {
      visa: (paymentsHubTotals.visa || 0) - (salesTotals.visa || 0),
      mastercard: (paymentsHubTotals.mastercard || 0) - (salesTotals.mastercard || 0),
      "american express": (paymentsHubTotals["american express"] || 0) - (salesTotals["american express"] || 0),
      discover: (paymentsHubTotals.discover || 0) - (salesTotals.discover || 0)
    };

    // Create reconciliation summary - using integer values for the totals to match expected format
    const summaryData = [
      ["Hub Report", "Total", "", "Sales Report", "Total", "", "Difference"],
      [
        "Visa", 
        Math.round(paymentsHubTotals.visa || 0).toString(), 
        "", 
        "Visa", 
        Math.round(salesTotals.visa || 0).toString(), 
        "", 
        Math.round(differences.visa || 0).toString()
      ],
      [
        "Mastercard", 
        Math.round(paymentsHubTotals.mastercard || 0).toString(), 
        "", 
        "Mastercard", 
        Math.round(salesTotals.mastercard || 0).toString(), 
        "", 
        Math.round(differences.mastercard || 0).toString()
      ],
      [
        "American Express", 
        Math.round(paymentsHubTotals["american express"] || 0).toString(), 
        "", 
        "American Express", 
        Math.round(salesTotals["american express"] || 0).toString(), 
        "", 
        Math.round(differences["american express"] || 0).toString()
      ],
      [
        "Discover", 
        Math.round(paymentsHubTotals.discover || 0).toString(), 
        "", 
        "Discover", 
        Math.round(salesTotals.discover || 0).toString(), 
        "", 
        Math.round(differences.discover || 0).toString()
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
    header && header.toString().toLowerCase().trim() === columnName.toLowerCase().trim()
  );
}

/**
 * Helper function to calculate card totals from Payments Hub data
 */
function calculateCardTotalsFromPaymentsHub(data, cardBrandColIndex, krColIndex) {
  const totals = {
    visa: 0,
    mastercard: 0,
    "american express": 0,
    discover: 0
  };
  
  if (cardBrandColIndex === -1 || krColIndex === -1) {
    return totals;
  }
  
  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    if (row && row.length > Math.max(cardBrandColIndex, krColIndex)) {
      if (row[cardBrandColIndex]) {
        const cardBrand = row[cardBrandColIndex].toString().toLowerCase().trim();
        const amount = parseFloat(row[krColIndex]) || 0;
        
        if (!isNaN(amount)) {
          // Add amounts to the appropriate card brand
          if (cardBrand === "visa") {
            totals.visa += amount;
          } else if (cardBrand === "mastercard") {
            totals.mastercard += amount;
          } else if (cardBrand === "american express") {
            totals["american express"] += amount;
          } else if (cardBrand === "discover") {
            totals.discover += amount;
          }
        }
      }
    }
  }
  
  return totals;
}

/**
 * Helper function to calculate card totals from Sales Totals data
 */
function calculateCardTotalsFromSales(data, cardTypeIndex, amountIndex) {
  const totals = {
    visa: 0,
    mastercard: 0,
    "american express": 0,
    discover: 0
  };
  
  if (cardTypeIndex === -1 || amountIndex === -1) {
    return totals;
  }
  
  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    if (row && row.length > Math.max(cardTypeIndex, amountIndex)) {
      if (row[cardTypeIndex]) {
        const cardType = row[cardTypeIndex].toString().toLowerCase().trim();
        
        // Parse the amount, removing currency symbols and spaces
        let amount = 0;
        if (row[amountIndex]) {
          const amountStr = row[amountIndex].toString().replace(/[^\d.-]/g, "");
          amount = parseFloat(amountStr) || 0;
        }
        
        if (!isNaN(amount)) {
          // Add amounts to the appropriate card type
          if (cardType === "visa") {
            totals.visa += amount;
          } else if (cardType === "mastercard") {
            totals.mastercard += amount;
          } else if (cardType === "american express") {
            totals["american express"] += amount;
          } else if (cardType === "discover") {
            totals.discover += amount;
          }
        }
      }
    }
  }
  
  return totals;
}

/**
 * Helper function to format date values for comparison
 * Handles different date formats
 */
function formatDate(dateStr) {
  if (!dateStr) return "";
  
  // Try to extract date components from various formats
  // First, clean up the string
  let cleanDateStr = dateStr.toString().trim();
  
  // Handle mm/dd/yyyy format
  const dateRegex1 = /(\d{1,2})\/(\d{1,2})\/(\d{4})/;
  const dateRegex2 = /(\d{1,2})\/(\d{1,2})\/(\d{2})/;
  // Handle yyyy-mm-dd format
  const dateRegex3 = /(\d{4})-(\d{1,2})-(\d{1,2})/;
  // Handle timestamp format like "3/13/25 19:58"
  const timestampRegex = /(\d{1,2})\/(\d{1,2})\/(\d{2})\s+\d{1,2}:\d{1,2}/;
  
  let month, day, year;
  
  if (timestampRegex.test(cleanDateStr)) {
    const match = cleanDateStr.match(timestampRegex);
    month = match[1].padStart(2, "0");
    day = match[2].padStart(2, "0");
    year = match[3].length === 2 ? (match[3] < "50" ? "20" + match[3] : "19" + match[3]) : match[3];
  } else if (dateRegex1.test(cleanDateStr)) {
    const match = cleanDateStr.match(dateRegex1);
    month = match[1].padStart(2, "0");
    day = match[2].padStart(2, "0");
    year = match[3];
  } else if (dateRegex2.test(cleanDateStr)) {
    const match = cleanDateStr.match(dateRegex2);
    month = match[1].padStart(2, "0");
    day = match[2].padStart(2, "0");
    year = match[3].length === 2 ? (match[3] < "50" ? "20" + match[3] : "19" + match[3]) : match[3];
  } else if (dateRegex3.test(cleanDateStr)) {
    const match = cleanDateStr.match(dateRegex3);
    year = match[1];
    month = match[2].padStart(2, "0");
    day = match[3].padStart(2, "0");
  } else {
    // Try to use JavaScript's Date parsing as a fallback
    try {
      const date = new Date(cleanDateStr);
      if (!isNaN(date.getTime())) {
        month = (date.getMonth() + 1).toString().padStart(2, "0");
        day = date.getDate().toString().padStart(2, "0");
        year = date.getFullYear().toString();
      } else {
        return cleanDateStr; // Return original if parsing fails
      }
    } catch (e) {
      return cleanDateStr; // Return original if parsing fails
    }
  }
  
  // Return normalized format for comparison: MM/DD/YYYY
  return month + "/" + day + "/" + year;
}

/**
 * Helper function to format currency string values
 * Preserves original currency format with $ sign for display
 */
function formatCurrencyString(value) {
  if (!value) return "";
  
  // Extract the numeric part from currency string
  const numStr = value.toString().replace(/[^\d.-]/g, "");
  const numValue = parseFloat(numStr);
  
  if (isNaN(numValue)) {
    return "";
  }
  
  // Return with dollar sign to match the expected format
  return "$" + numValue.toFixed(2) + " ";
}