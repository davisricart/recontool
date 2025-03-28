/**
 * Excel Reconciliation Script - JavaScript conversion of VBA macro
 * This function compares and reconciles data from two Excel files/sheets
 * Similar to the original VBA MainRecon macro
 */
async function compareAndDisplayData(XLSX, file1Data, file2Data) {
  try {
    // Load workbooks
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
    
    // Get sheets (assuming similar structure to original VBA)
    // First file should have "Payments Hub Transaction" sheet
    const paymentsHubSheet = workbook1.Sheets[workbook1.SheetNames[0]];
    // Second file should have "Sales Totals" sheet
    const salesTotalsSheet = workbook2.Sheets[workbook2.SheetNames[0]];
    
    // Convert sheets to JSON for easier processing
    const paymentsHubData = XLSX.utils.sheet_to_json(paymentsHubSheet, {header: 1, defval: ""});
    const salesTotalsData = XLSX.utils.sheet_to_json(salesTotalsSheet, {header: 1, defval: ""});
    
    // Process data - Mirroring the VBA workflow
    // Step 1: Process Payments Hub Transaction data
    
    // Get column indices (assuming they match the original structure)
    const dateColIndex = findColumnIndex(paymentsHubData[0], "Date");
    const customerNameColIndex = findColumnIndex(paymentsHubData[0], "Customer Name");
    const totalAmountColIndex = findColumnIndex(paymentsHubData[0], "Total Transaction Amount");
    const discountingAmountColIndex = findColumnIndex(paymentsHubData[0], "Cash Discounting Amount");
    const cardBrandColIndex = findColumnIndex(paymentsHubData[0], "Card Brand");
    
    // Step 2: Calculate K-R (similar to AA column in VBA)
    // Original formula: =RC[-16]-RC[-9] (Total Transaction Amount - Cash Discounting Amount)
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
    
    // Step 3: Process Sales Totals data
    // Find relevant columns in Sales Totals
    const salesDateColIndex = findColumnIndex(salesTotalsData[0], "Date");
    const salesCardBrandColIndex = findColumnIndex(salesTotalsData[0], "Card Brand");
    const salesAmountColIndex = findColumnIndex(salesTotalsData[0], "Amount");
    
    // If amount column doesn't exist, calculate it from other columns
    let processedSalesData = salesTotalsData;
    if (salesAmountColIndex === -1) {
      const salesEColIndex = findColumnIndex(salesTotalsData[0], "E"); // Assuming E is the column with values
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
    
    // Step 4: Add Count column (similar to AB column in VBA)
    // Original formula: =COUNTIFS('Sales Totals'!C2,'Payments Hub Transaction'!RC1,'Sales Totals'!C1,'Payments Hub Transaction'!RC24,'Sales Totals'!C5,'Payments Hub Transaction'!RC27)
    const krColIndex = paymentsHubWithKR[0].length - 1;
    const paymentsHubWithCount = paymentsHubWithKR.map((row, index) => {
      if (index === 0) {
        // Add header for Count
        return [...row, "Count"];
      } else if (row.length > 0) {
        // Calculate Count value
        let count = 0;
        
        // Count matching rows between the two sheets
        for (let i = 1; i < processedSalesData.length; i++) {
          const salesRow = processedSalesData[i];
          if (
            row[dateColIndex] === salesRow[salesDateColIndex] &&
            row[cardBrandColIndex] === salesRow[salesCardBrandColIndex]
            // We could add more criteria here if needed
          ) {
            count++;
          }
        }
        
        return [...row, count];
      }
      return row;
    });
    
    // Step 5: Filter and create the final data (mirroring the "Final" sheet in VBA)
    const countColIndex = paymentsHubWithCount[0].length - 1;
    const filteredRows = paymentsHubWithCount.filter((row, index) => {
      return index === 0 || row[countColIndex] !== 0;
    });
    
    // Keep only the visible columns as in the VBA (date, customer name, total amount, discount, card brand, K-R)
    const finalData = filteredRows.map(row => {
      return [
        row[dateColIndex], 
        row[customerNameColIndex], 
        row[totalAmountColIndex], 
        row[discountingAmountColIndex], 
        row[cardBrandColIndex],
        row[krColIndex] // K-R
      ];
    });
    
    // Add header for the last column to match the VBA
    finalData[0][5] = "Total (-) Fee";
    
    // Step 6: Add reconciliation summary (mirroring the right side of "Final" sheet)
    // Calculate card brand totals
    const visaTotal = calculateTotalByCardBrand(paymentsHubData, cardBrandColIndex, totalAmountColIndex, "Visa");
    const mcTotal = calculateTotalByCardBrand(paymentsHubData, cardBrandColIndex, totalAmountColIndex, "Mastercard");
    const amexTotal = calculateTotalByCardBrand(paymentsHubData, cardBrandColIndex, totalAmountColIndex, "American Express");
    
    const visaTotalSales = calculateTotalByCardBrand(processedSalesData, salesCardBrandColIndex, salesAmountColIndex, "Visa");
    const mcTotalSales = calculateTotalByCardBrand(processedSalesData, salesCardBrandColIndex, salesAmountColIndex, "Mastercard");
    const amexTotalSales = calculateTotalByCardBrand(processedSalesData, salesCardBrandColIndex, salesAmountColIndex, "American Express");
    
    // Prepare summary data to display
    const summaryData = [
      ["Hub Report", "Total", "", "Hub Report", "Total", "", "Difference"],
      ["Visa", visaTotal, "", "Visa", visaTotalSales, "", visaTotal - visaTotalSales],
      ["Mastercard", mcTotal, "", "Mastercard", mcTotalSales, "", mcTotal - mcTotalSales],
      ["American Express", amexTotal, "", "American Express", amexTotalSales, "", amexTotal - amexTotalSales]
    ];
    
    // Combine finalData with summaryData
    const maxRows = Math.max(finalData.length, summaryData.length);
    const resultData = [];
    
    for (let i = 0; i < maxRows; i++) {
      const finalRow = i < finalData.length ? finalData[i] : Array(finalData[0].length).fill("");
      const summaryRow = i < summaryData.length ? summaryData[i] : Array(summaryData[0].length).fill("");
      
      resultData.push([...finalRow, "", ...summaryRow]);
    }
    
    // Format numbers for display
    for (let i = 1; i < resultData.length; i++) {
      for (let j = 0; j < resultData[i].length; j++) {
        if (typeof resultData[i][j] === 'number') {
          resultData[i][j] = resultData[i][j].toFixed(2);
        }
      }
    }
    
    return resultData;
    
  } catch (error) {
    console.error("Error processing data:", error);
    return [["Error processing data: " + error.message]];
  }
}

/**
 * Helper function to find the index of a column by name
 */
function findColumnIndex(headerRow, columnName) {
  return headerRow.findIndex(header => 
    header && header.toString().toLowerCase() === columnName.toLowerCase()
  );
}

/**
 * Helper function to calculate total by card brand
 */
function calculateTotalByCardBrand(data, cardBrandColIndex, amountColIndex, brandName) {
  let total = 0;
  
  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    if (
      row.length > Math.max(cardBrandColIndex, amountColIndex) && 
      row[cardBrandColIndex] && 
      row[cardBrandColIndex].toString().toLowerCase() === brandName.toLowerCase()
    ) {
      const amount = parseFloat(row[amountColIndex]) || 0;
      total += amount;
    }
  }
  
  return total;
}