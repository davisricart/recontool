// This function compares dates between two Excel files
// Shows only the differences between the dates in both files
// Structure adjusted to work with the Master V4 interface

function compareFiles(XLSX, workbook1, workbook2) {
  try {
    // Get the first sheet from each workbook
    const paymentsHubSheet = workbook1.Sheets[workbook1.SheetNames[0]];
    const salesTotalsSheet = workbook2.Sheets[workbook2.SheetNames[0]];
    
    // Convert the sheets to JSON
    const paymentsHubJson = XLSX.utils.sheet_to_json(paymentsHubSheet);
    const salesTotalsJson = XLSX.utils.sheet_to_json(salesTotalsSheet);
    
    // Normalize date formats for comparison
    // Extract dates from column A in Payments Hub
    const paymentsHubDates = paymentsHubJson.map(row => {
      const date = row.Date;
      // For handling string dates (like ISO format)
      if (typeof date === 'string') {
        const parsedDate = new Date(date);
        if (!isNaN(parsedDate.getTime())) {
          return `${parsedDate.getMonth() + 1}/${parsedDate.getDate()}/${parsedDate.getFullYear()}`;
        }
        return date;
      }
      // For handling Date objects
      else if (date instanceof Date) {
        return `${date.getMonth() + 1}/${date.getDate()}/${date.getFullYear()}`;
      }
      // If all else fails, return as string
      return String(date);
    });
    
    // Extract dates from column B (Date Closed) in Sales Totals
    const salesTotalsDates = salesTotalsJson.map(row => {
      const date = row['Date Closed'];
      // For handling string dates (already in MM/DD/YYYY format)
      if (typeof date === 'string') {
        // Check if we need to parse it (in case it's not MM/DD/YYYY)
        if (date.includes('-')) {
          const parsedDate = new Date(date);
          if (!isNaN(parsedDate.getTime())) {
            return `${parsedDate.getMonth() + 1}/${parsedDate.getDate()}/${parsedDate.getFullYear()}`;
          }
        }
        return date;
      }
      // For handling Date objects
      else if (date instanceof Date) {
        return `${date.getMonth() + 1}/${date.getDate()}/${date.getFullYear()}`;
      }
      // If all else fails, return as string
      return String(date);
    });
    
    // Remove duplicates
    const uniquePaymentsHubDates = [...new Set(paymentsHubDates)];
    const uniqueSalesTotalsDates = [...new Set(salesTotalsDates)];
    
    // Find dates in Payments Hub that are not in Sales Totals
    const inPaymentsNotInSales = uniquePaymentsHubDates.filter(date => 
      !uniqueSalesTotalsDates.includes(date)
    );
    
    // Find dates in Sales Totals that are not in Payments Hub
    const inSalesNotInPayments = uniqueSalesTotalsDates.filter(date => 
      !uniquePaymentsHubDates.includes(date)
    );
    
    // Create a table format that Master V4 expects
    // First row contains headers
    const tableData = [
      ['Comparison Type', 'Date', 'Comments']
    ];
    
    // Add dates from Payments Hub not in Sales
    inPaymentsNotInSales.forEach(date => {
      tableData.push(['In Payments Hub Only', date, '']);
    });
    
    // Add dates from Sales not in Payments Hub
    inSalesNotInPayments.forEach(date => {
      tableData.push(['In Sales Totals Only', date, '']);
    });
    
    // If no differences found, add a row indicating this
    if (inPaymentsNotInSales.length === 0 && inSalesNotInPayments.length === 0) {
      tableData.push(['No Differences', 'All dates match between files', '']);
    }
    
    return tableData;
  } catch (error) {
    // Return error in format that can be displayed in the table
    return [
      ['Error Type', 'Message', 'Details'],
      ['Comparison Error', error.message, '']
    ];
  }
}

// The Master V4 site will call this function
return compareFiles;