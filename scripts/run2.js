/**
 * Helper function to normalize text values for comparison
 */
function normalize(value) {
  if (!value) return "";
  return value.toString().toLowerCase().trim();
}

/**
 * Helper function to format date values for comparison
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
  return `$${numValue.toFixed(2)} `; // Added space after the amount to match Excel format
}

/**
 * Better card brand comparison that handles variations
 */
function compareCardBrands(brand1, brand2) {
  if (!brand1 || !brand2) return false;
  
  const normBrand1 = brand1.toLowerCase().trim();
  const normBrand2 = brand2.toLowerCase().trim();
  
  // Direct match
  if (normBrand1 === normBrand2) return true;
  
  // Common variations
  const visaVariations = ['visa', 'vs', 'v'];
  const mcVariations = ['mastercard', 'master card', 'master', 'mc'];
  const amexVariations = ['american express', 'amex', 'ax', 'american exp'];
  const discoverVariations = ['discover', 'disc', 'dc'];
  
  // Check if both brands are variations of the same card type
  const isVisa = visaVariations.includes(normBrand1) && visaVariations.includes(normBrand2);
  const isMC = mcVariations.includes(normBrand1) && mcVariations.includes(normBrand2);
  const isAmex = amexVariations.includes(normBrand1) && amexVariations.includes(normBrand2);
  const isDiscover = discoverVariations.includes(normBrand1) && discoverVariations.includes(normBrand2);
  
  return isVisa || isMC || isAmex || isDiscover;
}

/**
 * Compare dates with better handling of formatting variations
 */
function compareDates(date1, date2) {
  if (!date1 || !date2) return false;
  
  // Direct string match
  if (date1 === date2) return true;
  
  // Try to normalize both dates to MM/DD/YY format
  let normalizedDate1 = date1;
  let normalizedDate2 = date2;
  
  // Handle MM/DD/YYYY vs MM/DD/YY
  if (date1.includes('/') && date2.includes('/')) {
    const parts1 = date1.split('/');
    const parts2 = date2.split('/');
    
    if (parts1.length >= 3 && parts2.length >= 3) {
      // Convert to MM/DD/YY format
      const month1 = parts1[0].padStart(2, '0');
      const day1 = parts1[1].padStart(2, '0');
      const year1 = parts1[2].length > 2 ? parts1[2].slice(-2) : parts1[2];
      
      const month2 = parts2[0].padStart(2, '0');
      const day2 = parts2[1].padStart(2, '0');
      const year2 = parts2[2].length > 2 ? parts2[2].slice(-2) : parts2[2];
      
      normalizedDate1 = `${month1}/${day1}/${year1}`;
      normalizedDate2 = `${month2}/${day2}/${year2}`;
      
      return normalizedDate1 === normalizedDate2;
    }
  }
  
  // Try parsing as Date objects and compare
  try {
    const dateObj1 = new Date(date1);
    const dateObj2 = new Date(date2);
    
    if (!isNaN(dateObj1.getTime()) && !isNaN(dateObj2.getTime())) {
      return dateObj1.getFullYear() === dateObj2.getFullYear() &&
             dateObj1.getMonth() === dateObj2.getMonth() &&
             dateObj1.getDate() === dateObj2.getDate();
    }
  } catch (e) {
    console.error("Error comparing dates:", e);
  }
  
  // Fallback
  return false;
}

/**
 * Calculate totals by card brand for summary
 */
function calculateTotalsByCardBrand(data, cardBrandColIndex, amountColIndex) {
  const totals = {
    visa: 0,
    mastercard: 0,
    "american express": 0,
    discover: 0
  };
  
  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    if (row.length <= Math.max(cardBrandColIndex, amountColIndex)) {
      continue;
    }
    
    let cardBrand = normalize(row[cardBrandColIndex]);
    let amount = 0;
    
    if (row[amountColIndex]) {
      const amountStr = typeof row[amountColIndex] === 'string' 
        ? row[amountColIndex].replace(/[^\d.-]/g, "") 
        : row[amountColIndex];
      amount = parseFloat(amountStr) || 0;
    }
    
    // Map various card brand formats to standard names
    if (cardBrand.includes("visa") || cardBrand === "vs" || cardBrand === "v") {
      totals.visa += amount;
    } else if (cardBrand.includes("master") || cardBrand === "mc") {
      totals.mastercard += amount;
    } else if (cardBrand.includes("american") || cardBrand.includes("amex") || cardBrand === "ax") {
      totals["american express"] += amount;
    } else if (cardBrand.includes("discover") || cardBrand === "disc" || cardBrand === "dc") {
      totals.discover += amount;
    }
  }
  
  return totals;
}

/**
 * More forgiving column finder that tries multiple potential column names
 */
function findColumnIndexFuzzy(headerRow, possibleNames) {
  if (!headerRow) return -1;
  
  // First try exact matches
  for (const name of possibleNames) {
    const exactIndex = headerRow.findIndex(header =>
      header && header.toString().toLowerCase().trim() === name.toLowerCase().trim()
    );
    if (exactIndex !== -1) return exactIndex;
  }
  
  // Then try contains matches
  for (const name of possibleNames) {
    const partialIndex = headerRow.findIndex(header =>
      header && header.toString().toLowerCase().trim().includes(name.toLowerCase().trim())
    );
    if (partialIndex !== -1) return partialIndex;
  }
  
  // Try even more fuzzy matches with common variations
  for (const name of possibleNames) {
    const fuzzyIndex = headerRow.findIndex(header => {
      if (!header) return false;
      const headerStr = header.toString().toLowerCase().trim();
      const nameStr = name.toLowerCase().trim();
      
      // Try removing spaces, underscores, etc.
      const cleanHeader = headerStr.replace(/[\s_-]/g, "");
      const cleanName = nameStr.replace(/[\s_-]/g, "");
      
      return cleanHeader.includes(cleanName) || cleanName.includes(cleanHeader);
    });
    if (fuzzyIndex !== -1) return fuzzyIndex;
  }
  
  return -1; // Not found
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
 * Find a sheet with a similar name to the target
 */
function findSheetWithSimilarName(sheetNames, targetName) {
  const targetLower = targetName.toLowerCase();
  
  // First try exact match
  const exactMatch = sheetNames.find(name => 
    name.toLowerCase() === targetLower
  );
  if (exactMatch) return exactMatch;
  
  // Then try contains match
  const containsMatch = sheetNames.find(name => 
    name.toLowerCase().includes(targetLower) || 
    targetLower.includes(name.toLowerCase())
  );
  
  return containsMatch || null;
}

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

    // Get sheets from the workbooks - we'll look for the sheets by name
    let paymentsHubSheet, salesTotalsSheet;
    
    // Try to find the exact sheet names as used in the VBA
    paymentsHubSheet = workbook1.Sheets["Payments Hub Transaction"] || null;
    salesTotalsSheet = workbook2.Sheets["Sales Totals"] || null;
    
    // If not found, try alternative names or first sheet as fallback
    if (!paymentsHubSheet) {
      // Look for sheets with similar names first
      const hubSheetName = findSheetWithSimilarName(workbook1.SheetNames, "Payments Hub");
      if (hubSheetName) {
        paymentsHubSheet = workbook1.Sheets[hubSheetName];
        console.log("Found Payments Hub sheet by similar name:", hubSheetName);
      } else {
        // Last resort: use the first sheet
        paymentsHubSheet = workbook1.Sheets[workbook1.SheetNames[0]];
        console.log("Using first sheet for Payments Hub:", workbook1.SheetNames[0]);
      }
    }
    
    if (!salesTotalsSheet) {
      const salesSheetName = findSheetWithSimilarName(workbook2.SheetNames, "Sales");
      if (salesSheetName) {
        salesTotalsSheet = workbook2.Sheets[salesSheetName];
        console.log("Found Sales Totals sheet by similar name:", salesSheetName);
      } else {
        // Last resort: use the first sheet
        salesTotalsSheet = workbook2.Sheets[workbook2.SheetNames[0]];
        console.log("Using first sheet for Sales Totals:", workbook2.SheetNames[0]);
      }
    }
    
    // Check if we have both sheets
    if (!paymentsHubSheet || !salesTotalsSheet) {
      throw new Error("Could not find required sheets in the Excel files");
    }

    // Convert sheets to JSON for easier processing - ensure we read dates as strings
    const paymentsHubData = XLSX.utils.sheet_to_json(paymentsHubSheet, {
      header: 1,
      defval: "",
      raw: false,
      dateNF: "MM/DD/YY"  // Match VBA date format
    });
    
    const salesTotalsData = XLSX.utils.sheet_to_json(salesTotalsSheet, {
      header: 1,
      defval: "",
      raw: false,
      dateNF: "MM/DD/YY"  // Match VBA date format
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

    // Find column indices in Payments Hub data - more robust column finding
    const dateColIndex = findColumnIndexFuzzy(paymentsHubData[0], ["Date", "Transaction Date"]);
    const customerNameColIndex = findColumnIndexFuzzy(paymentsHubData[0], ["Customer Name", "Customer", "Client Name"]);
    const totalAmountColIndex = findColumnIndexFuzzy(paymentsHubData[0], ["Total Transaction Amount", "Total Amount", "Transaction Amount"]);
    const discountingAmountColIndex = findColumnIndexFuzzy(paymentsHubData[0], ["Cash Discounting Amount", "Discount Amount", "Discounting"]);
    const cardBrandColIndex = findColumnIndexFuzzy(paymentsHubData[0], ["Card Brand", "Brand", "Card Type"]);

    // Find column indices in Sales Totals data - more robust column finding
    const salesNameColIndex = findColumnIndexFuzzy(salesTotalsData[0], ["Name", "Card Type", "Card Brand"]);
    const salesDateColIndex = findColumnIndexFuzzy(salesTotalsData[0], ["Date Closed", "Date", "Transaction Date"]);
    const salesAmountColIndex = findColumnIndexFuzzy(salesTotalsData[0], ["Amount", "Total", "Sale Amount"]);

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

    // Get the K-R column index
    const krColIndex = paymentsHubWithKR[0].length - 1;

    // Add Count column based on the specific COUNTIFS formula:
    // =COUNTIFS('Sales Totals'!$A:$A,'Payments Hub Transaction'!$X2,'Sales Totals'!$B:$B,'Payments Hub Transaction'!$AH2,'Sales Totals'!$E:$E,'Payments Hub Transaction'!$AA2,'Sales Totals'!$G:$G,'Payments Hub Transaction'!$AB2)
    const paymentsHubWithCount = paymentsHubWithKR.map((row, index) => {
      if (index === 0) {
        // Add header for Count
        return [...row, "Count"];
      } else if (row.length > 0) {
        // We need to access the exact columns used in the COUNTIFS formula
        // Note: JavaScript is 0-indexed, so we subtract 1 from Excel column indices
        
        // Payments Hub columns (using fixed positions from the formula, not header searching)
        // Convert from Excel column references to 0-based array indices
        const hubColX = 23;  // Excel column X (24) - 1
        const hubColAH = 33; // Excel column AH (34) - 1
        const hubColAA = 26; // Excel column AA (27) - 1
        const hubColAB = 27; // Excel column AB (28) - 1
        
        // Get the values from the appropriate columns
        // Use safe indexing to avoid errors if columns don't exist
        const hubValX = row.length > hubColX ? row[hubColX] : "";
        const hubValAH = row.length > hubColAH ? row[hubColAH] : "";
        const hubValAA = row.length > hubColAA ? row[hubColAA] : "";
        const hubValAB = row.length > hubColAB ? row[hubColAB] : "";
        
        // Sales Totals columns (from the formula)
        // Excel column references are 1-indexed, but our arrays are 0-indexed
        const salesColA = 0;  // Excel column A - 1
        const salesColB = 1;  // Excel column B - 1
        const salesColE = 4;  // Excel column E - 1
        const salesColG = 6;  // Excel column G - 1
        
        // Count matching records in Sales Totals
        let count = 0;
        
        for (let i = 1; i < salesTotalsData.length; i++) {
          const salesRow = salesTotalsData[i];
          
          // Skip if row doesn't have enough columns
          if (salesRow.length <= Math.max(salesColA, salesColB, salesColE, salesColG)) {
            continue;
          }
          
          // Get values from Sales Totals
          const salesValA = salesRow[salesColA] || "";
          const salesValB = salesRow[salesColB] || "";
          const salesValE = salesRow[salesColE] || "";
          const salesValG = salesRow[salesColG] || "";
          
          // Convert values to strings for comparison
          const normalizeForCompare = (val) => {
            if (val === null || val === undefined) return "";
            return val.toString().trim().toLowerCase();
          };
          
          // Parse numeric values for comparison with tolerance
          const parseAmount = (val) => {
            if (val === null || val === undefined) return 0;
            const str = val.toString().replace(/[^\d.-]/g, "");
            return parseFloat(str) || 0;
          };
          
          // For date values, we need to ensure consistent formatting
          const formatDateForCompare = (val) => {
            return formatDate(val);
          };
          
          // Compare values based on likely data types
          // Assuming A and X are dates, B and AH are card types, E and AA are amounts, G and AB are amounts
          const dateMatches = formatDateForCompare(salesValA) === formatDateForCompare(hubValX);
          const cardMatches = normalizeForCompare(salesValB) === normalizeForCompare(hubValAH);
          
          // For amounts, use tolerance to account for floating point differences
          const tolerance = 0.01;
          const amount1Matches = Math.abs(parseAmount(salesValE) - parseAmount(hubValAA)) < tolerance;
          const amount2Matches = Math.abs(parseAmount(salesValG) - parseAmount(hubValAB)) < tolerance;
          
          // Only increment count if ALL criteria match (this is how COUNTIFS works)
          if (dateMatches && cardMatches && amount1Matches && amount2Matches) {
            count++;
          }
        }
        
        return [...row, count];
      }
      return row;
    });

    // Get the Count column index
    const countColIndex = paymentsHubWithCount[0].length - 1;

    // Create the final output data with selected columns
    // IMPORTANT: Only include rows where Count = 0 (matches the requirement)
    const finalData = paymentsHubWithCount.filter((row, index) => {
      // Always include header row
      if (index === 0) return true;
      
      // Only include rows where Count = 0 (exact match to the requirement)
      const count = parseInt(row[countColIndex]) || 0;
      return count === 0;
    }).map((row, index) => {
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
    const paymentsHubTotals = calculateTotalsByCardBrand(paymentsHubWithCount, cardBrandColIndex, krColIndex);
    
    // Then for Sales Totals data
    const salesTotals = calculateTotalsByCardBrand(salesTotalsData, salesNameColIndex, salesAmountColIndex);
    
    // Calculate differences
    const differences = {
      visa: (paymentsHubTotals.visa || 0) - (salesTotals.visa || 0),
      mastercard: (paymentsHubTotals.mastercard || 0) - (salesTotals.mastercard || 0),
      "american express": (paymentsHubTotals["american express"] || 0) - (salesTotals["american express"] || 0),
      discover: (paymentsHubTotals.discover || 0) - (salesTotals.discover || 0)
    };

    // Create summary section data
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
      ],
      [
        "Discover", 
        formatCurrencyString(paymentsHubTotals.discover), 
        "", 
        "Discover", 
        formatCurrencyString(salesTotals.discover), 
        "", 
        formatCurrencyString(differences.discover)
      ]
    ];

    // Combine finalData with summaryData into a single result dataset
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

    // Return the combined data as a flat array with all columns
    return resultData;

  } catch (error) {
    console.error("Error processing data:", error);
    // Return error as a single row array to match expected format
    return [
      ["Error processing data: " + error.message]
    ];
  }
}
