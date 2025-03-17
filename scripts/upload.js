/**
 * Excel Reconciliation Script
 * 
 * This script compares "Payments Hub Transaction" and "Sales Totals" Excel files
 * and performs reconciliation calculations similar to the original VBA macro.
 * 
 * @param {Object} XLSX - The SheetJS library
 * @param {Object} workbook1 - "Payments Hub Transaction" workbook
 * @param {Object} workbook2 - "Sales Totals" workbook
 * @returns {Array} - Array of results for display
 */
function compareExcelFiles(XLSX, workbook1, workbook2) {
  try {
    // Step 1: Extract the data from workbooks
    const paymentsHub = extractData(XLSX, workbook1, 'Payments Hub Transaction');
    const salesTotals = extractData(XLSX, workbook2, 'Sales Totals');
    
    // Step 2: Process Payments Hub data (similar to VBA)
    processPaymentsHubData(paymentsHub);
    
    // Step 3: Process Sales Totals data (similar to VBA)
    processSalesTotalsData(salesTotals);
    
    // Step 4: Compare data between sheets
    const comparisonResults = compareData(paymentsHub, salesTotals);
    
    // Step 5: Create summary data
    const summaryData = createSummary(paymentsHub, salesTotals);
    
    // Step 6: Format final results
    return formatResults(comparisonResults, summaryData);
  } catch (error) {
    console.error("Error in compareExcelFiles:", error);
    // Return a default array with error message to prevent forEach errors
    return [["Error processing files", error.message]];
  }
}

/**
 * Extract data from workbook sheet
 */
function extractData(XLSX, workbook, expectedSheetName) {
  // Get the first sheet if it exists
  const firstSheetName = workbook.SheetNames[0];
  const sheet = workbook.Sheets[firstSheetName];
  
  // Convert sheet to JSON with headers
  const data = XLSX.utils.sheet_to_json(sheet, { header: 1, defval: '' });
  
  // Extract headers (first row)
  const headers = data[0].map(header => String(header).trim());
  
  // Map data to objects with header keys
  const rows = [];
  for (let i = 1; i < data.length; i++) {
    if (data[i].some(cell => cell !== '')) {  // Skip completely empty rows
      const row = {};
      for (let j = 0; j < headers.length; j++) {
        if (headers[j]) {  // Only add cells that have headers
          row[headers[j]] = data[i][j];
        }
      }
      rows.push(row);
    }
  }
  
  return {
    headers: headers,
    rows: rows
  };
}

/**
 * Process Payments Hub Transaction data (similar to VBA operations)
 */
function processPaymentsHubData(paymentsHub) {
  // In VBA: text to columns, calculations, etc.
  // We'll adapt these operations for JavaScript
  
  // Add K-R column calculation (as in VBA)
  paymentsHub.rows.forEach(row => {
    // Find column K and R by index or name
    const colK = findColumnByIndex(paymentsHub.headers, 10); // K is 11th column (index 10)
    const colR = findColumnByIndex(paymentsHub.headers, 17); // R is 18th column (index 17)
    
    if (colK && colR && row[colK] !== undefined && row[colR] !== undefined) {
      row['K-R'] = Number(row[colK]) - Number(row[colR]);
    } else {
      row['K-R'] = 0;
    }
  });
  
  // Add 'K-R' to headers if not already there
  if (!paymentsHub.headers.includes('K-R')) {
    paymentsHub.headers.push('K-R');
  }
}

/**
 * Process Sales Totals data (similar to VBA operations)
 */
function processSalesTotalsData(salesTotals) {
  // In VBA: adding Amount column, formatting, etc.
  
  // Calculate Amount column (Col G = Col E * 1 in VBA)
  salesTotals.rows.forEach(row => {
    const colE = findColumnByIndex(salesTotals.headers, 4); // E is 5th column (index 4)
    
    if (colE && row[colE] !== undefined) {
      row['Amount'] = Number(row[colE]);
    } else {
      row['Amount'] = 0;
    }
  });
  
  // Add 'Amount' to headers if not already there
  if (!salesTotals.headers.includes('Amount')) {
    salesTotals.headers.push('Amount');
  }
}

/**
 * Compare data between sheets (similar to COUNTIFS in VBA)
 */
function compareData(paymentsHub, salesTotals) {
  // We'll implement the comparison logic that was in the VBA
  // This is based on the COUNTIFS in the VBA code
  
  const results = [];
  
  // Extract the column indices we need based on VBA's logic
  const dateColumnPH = 'Date';
  const customerColumnPH = 'Customer Name';
  const amountColumnPH = 'Total Transaction Amount';
  const cardBrandColumnPH = 'Card Brand';
  const discountColumnPH = 'Cash Discounting Amount';
  
  // Add Count and Final Count columns to each row
  paymentsHub.rows.forEach(phRow => {
    let count = 0;
    let finalCount = 0;
    
    // Similar to COUNTIFS in VBA
    salesTotals.rows.forEach(stRow => {
      // Check match conditions 
      // This replicates the complex COUNTIFS from the VBA code
      const dateMatch = formatDate(phRow[dateColumnPH]) === formatDate(stRow['Date']);
      const amountMatch = Math.abs(Number(phRow[amountColumnPH]) - Number(stRow['Amount'])) < 0.01;
      
      if (dateMatch && amountMatch) {
        count++;
        
        // For Final Count, add additional matching criteria
        if (phRow[cardBrandColumnPH] === stRow['Card Type']) {
          finalCount++;
        }
      }
    });
    
    phRow['Count'] = count;
    phRow['Final Count'] = finalCount;
    
    // If final count is 0, add to results (these are unmatched records)
    if (finalCount === 0) {
      results.push({
        Date: phRow[dateColumnPH],
        'Customer Name': phRow[customerColumnPH],
        'Total Transaction Amount': phRow[amountColumnPH],
        'Cash Discounting Amount': phRow[discountColumnPH],
        'Card Brand': phRow[cardBrandColumnPH],
        'Total (-) Fee': phRow['K-R']
      });
    }
  });
  
  return results;
}

/**
 * Create summary data (similar to SumIfs in VBA)
 */
function createSummary(paymentsHub, salesTotals) {
  // This builds the summary table with card brands and totals
  const cardBrands = ['Visa', 'Mastercard', 'American Express', 'Discover'];
  const summary = [];
  
  cardBrands.forEach(brand => {
    // Hub Report Total
    const hubTotal = paymentsHub.rows
      .filter(row => row['Card Brand'] === brand)
      .reduce((sum, row) => sum + (Number(row['K-R']) || 0), 0);
    
    // Sales Report Total
    const salesTotal = salesTotals.rows
      .filter(row => row['Card Type'] === brand)
      .reduce((sum, row) => sum + (Number(row['Amount']) || 0), 0);
    
    // Calculate difference
    const difference = hubTotal - salesTotal;
    
    summary.push({
      'Brand': brand,
      'Hub Report': hubTotal.toFixed(2),
      'Sales Report': salesTotal.toFixed(2),
      'Difference': difference.toFixed(2),
      'Status': difference >= 0 ? 'positive' : 'negative'
    });
  });
  
  return summary;
}

/**
 * Format results for display
 */
function formatResults(comparisonResults, summaryData) {
  // Format as an array for display in table
  const results = [];
  
  // Add headers for the main results
  if (comparisonResults.length > 0) {
    const headers = ['Date', 'Customer Name', 'Total Transaction Amount', 'Cash Discounting Amount', 'Card Brand', 'Total (-) Fee'];
    results.push(headers);
    
    // Add data rows
    comparisonResults.forEach(row => {
      const dataRow = [
        formatDate(row.Date),
        row['Customer Name'],
        formatNumber(row['Total Transaction Amount']),
        formatNumber(row['Cash Discounting Amount']),
        row['Card Brand'],
        formatNumber(row['Total (-) Fee'])
      ];
      results.push(dataRow);
    });
  } else {
    // If no unmatched records, add "In Balance" message
    results.push(['Date', 'Customer Name', 'Total Transaction Amount', 'Cash Discounting Amount', 'Card Brand', 'Total (-) Fee']);
    results.push(['In Balance', '', '', '', '', '']);
  }
  
  // Add some space
  results.push(['', '', '', '', '', '']);
  results.push(['', '', '', '', '', '']);
  
  // Add Summary section headers
  results.push(['Hub Report', 'Total', '', 'Sales Report', 'Total', 'Difference']);
  
  // Add summary data
  summaryData.forEach(row => {
    results.push([
      row.Brand,
      row['Hub Report'],
      '',
      row.Brand,
      row['Sales Report'],
      row['Difference']
    ]);
  });
  
  return results;
}

// Utility Functions

/**
 * Find column by index, returns header name
 */
function findColumnByIndex(headers, index) {
  return headers[index] || null;
}

/**
 * Format date (handles various date inputs)
 */
function formatDate(dateValue) {
  if (!dateValue) return '';
  
  let date;
  if (dateValue instanceof Date) {
    date = dateValue;
  } else if (typeof dateValue === 'number') {
    // Handle Excel serial date
    date = new Date(Math.round((dateValue - 25569) * 86400 * 1000));
  } else {
    try {
      date = new Date(dateValue);
    } catch (e) {
      return String(dateValue);
    }
  }
  
  if (isNaN(date.getTime())) return String(dateValue);
  
  // Format as MM/DD/YYYY
  return (date.getMonth() + 1).toString().padStart(2, '0') + '/' + 
         date.getDate().toString().padStart(2, '0') + '/' + 
         date.getFullYear();
}

/**
 * Format number with 2 decimal places
 */
function formatNumber(value) {
  if (value === undefined || value === null || value === '') return '';
  
  const num = Number(value);
  return isNaN(num) ? value : num.toFixed(2);
}

// FIXED: Export the function for browser environments
// Remove Node.js module.exports and make it available globally
window.compareExcelFiles = compareExcelFiles;