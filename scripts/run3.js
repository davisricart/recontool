/**
 * Compare and process Excel files to analyze payment data.
 * 
 * @param {Object} XLSX - The SheetJS library object
 * @param {ArrayBuffer} file1 - The first uploaded Excel file data
 * @param {ArrayBuffer} file2 - The second uploaded Excel file data
 * @returns {Array} An array of arrays representing the processed data
 */
function compareAndDisplayData(XLSX, file1, file2) {
    // Helper function to format date for comparison
    function formatDateForComparison(date) {
        if (!(date instanceof Date) || isNaN(date)) {
            return '';
        }
        return `${date.getFullYear()}-${String(date.getMonth() + 1).padStart(2, '0')}-${String(date.getDate()).padStart(2, '0')}`;
    }
    
    // Step 1: Process First File
    // Parse the first Excel file
    const workbook1 = XLSX.read(file1, {
        cellDates: true,
        dateNF: 'yyyy-mm-dd'
    });
    
    // Get first sheet
    const sheetName1 = workbook1.SheetNames[0];
    const worksheet1 = workbook1.Sheets[sheetName1];
    
    // Convert worksheet to JSON
    const jsonData1 = XLSX.utils.sheet_to_json(worksheet1);
    
    // Step 2: Process Second File (if provided)
    let jsonData2 = [];
    let file2Headers = [];
    let dateClosedIndex = -1;
    let nameIndex = -1;
    let amountIndex = -1;
    
    if (file2) {
        // Parse the second Excel file
        const workbook2 = XLSX.read(file2, {
            cellDates: true,
            dateNF: 'yyyy-mm-dd'
        });
        
        // Get first sheet
        const sheetName2 = workbook2.SheetNames[0];
        const worksheet2 = workbook2.Sheets[sheetName2];
        
        // Convert worksheet to JSON with headers
        const data = XLSX.utils.sheet_to_json(worksheet2, { header: 1 });
        
        // Store headers and find required columns
        file2Headers = data[0] || [];
        
        dateClosedIndex = file2Headers.findIndex(header => 
            typeof header === "string" && header.trim().toLowerCase() === "date closed"
        );
        
        nameIndex = file2Headers.findIndex(header => 
            typeof header === "string" && header.trim().toLowerCase() === "name"
        );
        
        amountIndex = file2Headers.findIndex(header => 
            typeof header === "string" && header.trim().toLowerCase() === "amount"
        );
        
        // Get data rows (skip header)
        jsonData2 = data.slice(1);
        
        // Convert amount values to numbers
        if (amountIndex !== -1) {
            jsonData2.forEach(row => {
                if (amountIndex < row.length && row[amountIndex] !== undefined) {
                    let amount = row[amountIndex];
                    if (typeof amount === "string") {
                        amount = amount.replace(/[^0-9.-]+/g, "");
                    }
                    row[amountIndex] = parseFloat(amount) || 0;
                }
            });
        }
    }
    
    // Step 3: Filter First File and Add New Columns
    // Define columns to keep
    const columnsToKeep = ["Date", "Customer Name", "Total Transaction Amount", "Cash Discounting Amount", "Card Brand"];
    const newColumns = ["Total (-) Fee", "Count", "Final Count"];
    
    // Create result array starting with header
    const resultData = [columnsToKeep.concat(newColumns)];
    
    // Store processed first file data for comparisons
    const firstFileData = [];
    
    // Process each row of first file
    jsonData1.forEach(row => {
        const filteredRow = [];
        let firstFileDate = null;
        let cardBrand = "";
        let krValue = 0;
        
        // Filter columns
        columnsToKeep.forEach(column => {
            if (column === "Date") {
                if (row[column] instanceof Date) {
                    const date = row[column];
                    firstFileDate = new Date(date); // Clone date
                    firstFileDate.setHours(12, 0, 0, 0); // Normalize time component
                    
                    // Format as MM/DD/YYYY
                    const year = date.getFullYear();
                    const month = String(date.getMonth() + 1).padStart(2, '0');
                    const day = String(date.getDate()).padStart(2, '0');
                    filteredRow.push(`${month}/${day}/${year}`);
                } else {
                    // Handle string dates
                    filteredRow.push(row[column] !== undefined ? row[column] : "");
                    if (row[column]) {
                        try {
                            firstFileDate = new Date(row[column]);
                            firstFileDate.setHours(12, 0, 0, 0);
                        } catch (e) {
                            firstFileDate = null;
                        }
                    }
                }
            } else if (column === "Card Brand") {
                cardBrand = row[column] || "";
                filteredRow.push(cardBrand);
            } else {
                filteredRow.push(row[column] !== undefined ? row[column] : "");
            }
        });
        
        // Calculate K-R value
        const totalAmount = parseFloat(row["Total Transaction Amount"]) || 0;
        const discountAmount = parseFloat(row["Cash Discounting Amount"]) || 0;
        krValue = totalAmount - discountAmount;
        
        // Add K-R value (formatted to 2 decimal places)
        filteredRow.push(krValue.toFixed(2));
        
        // Calculate Count - matches in second file
        let countMatches = 0;
        
        if (dateClosedIndex !== -1 && nameIndex !== -1 && amountIndex !== -1 && firstFileDate) {
            jsonData2.forEach(secondRow => {
                if (secondRow.length > Math.max(dateClosedIndex, nameIndex, amountIndex)) {
                    // Get date from second file
                    let secondFileDate = null;
                    const dateValue = secondRow[dateClosedIndex];
                    
                    if (typeof dateValue === 'string') {
                        try {
                            secondFileDate = new Date(dateValue);
                            secondFileDate.setHours(12, 0, 0, 0);
                        } catch (e) {
                            secondFileDate = null;
                        }
                    } else if (dateValue instanceof Date) {
                        secondFileDate = new Date(dateValue);
                        secondFileDate.setHours(12, 0, 0, 0);
                    }
                    
                    // Get name and amount
                    const secondFileName = String(secondRow[nameIndex] || "").trim().toLowerCase();
                    const secondFileAmount = parseFloat(secondRow[amountIndex]) || 0;
                    
                    // Format dates for comparison
                    const firstFileDateStr = formatDateForComparison(firstFileDate);
                    const secondFileDateStr = formatDateForComparison(secondFileDate);
                    
                    // Compare values
                    const dateMatches = firstFileDateStr && secondFileDateStr && 
                        firstFileDateStr === secondFileDateStr;
                    
                    const nameMatches = secondFileName && cardBrand && (
                        cardBrand.trim().toLowerCase().includes(secondFileName) || 
                        secondFileName.includes(cardBrand.trim().toLowerCase())
                    );
                    
                    const amountMatches = Math.abs(krValue - secondFileAmount) < 0.01;
                    
                    if (dateMatches && nameMatches && amountMatches) {
                        countMatches++;
                    }
                }
            });
        }
        
        // Add count and empty final count
        filteredRow.push(countMatches.toString());
        filteredRow.push("");
        
        resultData.push(filteredRow);
        firstFileData.push(filteredRow);
    });
    
    // Step 4: Process Second File Data and Calculate Count2 Values
    if (file2 && file2Headers.length > 0) {
        const secondFileWithCount2 = [];
        
        // Process second file rows
        jsonData2.forEach(row => {
            const processedRow = [...row];
            let secondFileDate = null;
            let secondFileName = "";
            let secondFileAmount = 0;
            
            // Extract date
            if (dateClosedIndex !== -1 && dateClosedIndex < row.length) {
                const dateValue = row[dateClosedIndex];
                if (typeof dateValue === 'string') {
                    try {
                        secondFileDate = new Date(dateValue);
                        secondFileDate.setHours(12, 0, 0, 0);
                    } catch (e) {
                        secondFileDate = null;
                    }
                } else if (dateValue instanceof Date) {
                    secondFileDate = new Date(dateValue);
                    secondFileDate.setHours(12, 0, 0, 0);
                }
            }
            
            // Extract name and amount
            if (nameIndex !== -1 && nameIndex < row.length) {
                secondFileName = String(row[nameIndex] || "").trim().toLowerCase();
            }
            
            if (amountIndex !== -1 && amountIndex < row.length) {
                let amountValue = row[amountIndex];
                if (typeof amountValue === 'string') {
                    amountValue = amountValue.replace(/[^0-9.-]+/g, "");
                }
                secondFileAmount = parseFloat(amountValue) || 0;
            }
            
            // Format date for comparison
            const secondFileDateStr = formatDateForComparison(secondFileDate);
            
            // Count matches in first file
            let countMatches = 0;
            
            firstFileData.forEach(firstFileRow => {
                // Extract values from first file
                let firstFileDate = null;
                if (firstFileRow[0]) {
                    // Parse MM/DD/YYYY format
                    const parts = firstFileRow[0].split('/');
                    if (parts.length === 3) {
                        const month = parseInt(parts[0]) - 1;
                        const day = parseInt(parts[1]);
                        const year = parseInt(parts[2]);
                        firstFileDate = new Date(year, month, day);
                        firstFileDate.setHours(12, 0, 0, 0);
                    }
                }
                
                const firstFileDateStr = formatDateForComparison(firstFileDate);
                const firstFileCardBrand = String(firstFileRow[4] || "").trim().toLowerCase();
                const firstFileKR = parseFloat(firstFileRow[5] || 0);
                
                // Compare values
                const dateMatches = firstFileDateStr && secondFileDateStr && 
                    firstFileDateStr === secondFileDateStr;
                
                const nameMatches = secondFileName && firstFileCardBrand && (
                    firstFileCardBrand.includes(secondFileName) || 
                    secondFileName.includes(firstFileCardBrand)
                );
                
                const amountMatches = Math.abs(firstFileKR - secondFileAmount) < 0.01;
                
                if (dateMatches && nameMatches && amountMatches) {
                    countMatches++;
                }
            });
            
            // Add Count2 value
            processedRow.push(countMatches.toString());
            secondFileWithCount2.push(processedRow);
        });
        
        // Step 5: Calculate Final Count for First File Rows
        firstFileData.forEach((firstFileRow, index) => {
            // Extract values
            const date = firstFileRow[0]; 
            const cardBrand = String(firstFileRow[4] || "").trim().toLowerCase();
            const kr = parseFloat(firstFileRow[5] || 0);
            const count = parseInt(firstFileRow[6] || 0);
            
            // Parse date
            let firstFileDate = null;
            if (date) {
                const parts = date.split('/');
                if (parts.length === 3) {
                    const month = parseInt(parts[0]) - 1;
                    const day = parseInt(parts[1]);
                    const year = parseInt(parts[2]);
                    firstFileDate = new Date(year, month, day);
                    firstFileDate.setHours(12, 0, 0, 0);
                }
            }
            
            const firstFileDateStr = formatDateForComparison(firstFileDate);
            
            // Calculate Final Count
            let finalCount = 0;
            
            secondFileWithCount2.forEach(secondFileRow => {
                // Extract values from second file
                let secondFileDate = null;
                if (dateClosedIndex !== -1 && dateClosedIndex < secondFileRow.length) {
                    const dateValue = secondFileRow[dateClosedIndex];
                    if (typeof dateValue === 'string') {
                        try {
                            secondFileDate = new Date(dateValue);
                            secondFileDate.setHours(12, 0, 0, 0);
                        } catch (e) {
                            secondFileDate = null;
                        }
                    } else if (dateValue instanceof Date) {
                        secondFileDate = new Date(dateValue);
                        secondFileDate.setHours(12, 0, 0, 0);
                    }
                }
                
                const secondFileDateStr = formatDateForComparison(secondFileDate);
                
                const secondFileName = nameIndex !== -1 && nameIndex < secondFileRow.length ?
                    String(secondFileRow[nameIndex] || "").trim().toLowerCase() : "";
                    
                const secondFileAmount = amountIndex !== -1 && amountIndex < secondFileRow.length ?
                    parseFloat(secondFileRow[amountIndex]) || 0 : 0;
                    
                const secondFileCount2 = parseInt(secondFileRow[secondFileRow.length - 1] || 0);
                
                // Check all four criteria
                const dateMatches = firstFileDateStr && secondFileDateStr && 
                    firstFileDateStr === secondFileDateStr;
                
                const nameMatches = secondFileName && cardBrand && (
                    cardBrand.includes(secondFileName) || 
                    secondFileName.includes(cardBrand)
                );
                
                const amountMatches = Math.abs(kr - secondFileAmount) < 0.01;
                
                const countMatches = count === secondFileCount2;
                
                if (dateMatches && nameMatches && amountMatches && countMatches) {
                    finalCount++;
                }
            });
            
            // Update Final Count in result data (index + 1 because index 0 is header)
            resultData[index + 1][7] = finalCount.toString();
        });
    }
    
    // Step 6: Filter results to only show rows with Final Count = 0
    // and remove the Count and Final Count columns, and don't include second file data
    
    // Create a filtered result with only the columns we want to display
    const displayColumns = ["Date", "Customer Name", "Total Transaction Amount", "Cash Discounting Amount", "Card Brand", "Total (-) Fee"];
    
    // Create two separate data structures: one for the data and one for any formatting
    const filteredResults = [displayColumns]; // New header with renamed column and without Count/Final Count
    
    // Add first file rows that have Final Count = 0, but without the Count and Final Count columns
    for (let i = 1; i < resultData.length; i++) {
        const row = resultData[i];
        
        // Stop when we reach the separator (empty row)
        if (row.every(cell => cell === "")) {
            break;
        }
        
        // Check if Final Count is 0
        const finalCount = parseInt(row[7] || 0);
        if (finalCount === 0) {
            // Take just the first 6 columns
            const displayRow = row.slice(0, 6);
            
            // Convert numeric columns from strings to numbers for Excel formatting
            // Total Transaction Amount (index 2)
            if (displayRow[2] && !isNaN(parseFloat(displayRow[2]))) {
                displayRow[2] = parseFloat(displayRow[2]);
            }
            
            // Cash Discounting Amount (index 3)
            if (displayRow[3] && !isNaN(parseFloat(displayRow[3]))) {
                displayRow[3] = parseFloat(displayRow[3]);
            }
            
            // Total (-) Fee (index 5)
            if (displayRow[5] && !isNaN(parseFloat(displayRow[5]))) {
                displayRow[5] = parseFloat(displayRow[5]);
            }
            
            filteredResults.push(displayRow);
        }
    }
    
    // Calculate totals (but don't display them in the main table)
    let totalTransactionAmount = 0;
    let totalDiscountAmount = 0;
    let totalFee = 0;
    
    for (let i = 1; i < filteredResults.length; i++) {
        // Total Transaction Amount (index 2)
        totalTransactionAmount += parseFloat(filteredResults[i][2] || 0);
        
        // Cash Discounting Amount (index 3)
        totalDiscountAmount += parseFloat(filteredResults[i][3] || 0);
        
        // Total (-) Fee (index 5)
        totalFee += parseFloat(filteredResults[i][5] || 0);
    }
    
    // Add two blank rows as separators (without the totals row)
    filteredResults.push(["", "", "", "", "", ""]);
    filteredResults.push(["", "", "", "", "", ""]);
        
        // Create a more compact side-by-side headers for Hub Report and Sales Report, plus Difference
        filteredResults.push(["Card Brand", "Hub Report", "Sales Report", "Difference", "", ""]);
        
        // Collect unique card brands and totals from first file
        const cardBrandTotals = {};
        
        // Process all rows from the first file before filtering
        for (let i = 1; i < resultData.length; i++) {
            const row = resultData[i];
            
            // Stop at separator row
            if (row.every(cell => cell === "")) {
                break;
            }
            
            const cardBrand = row[4]; // Card Brand at index 4
            // Skip any row where the card brand contains "cash" (case-insensitive)
            if (cardBrand && cardBrand.toLowerCase().includes("cash")) {
                continue;
            }
            
            const netAmount = parseFloat(row[5] || 0); // Total (-) Fee
            
            if (cardBrand) {
                if (!cardBrandTotals[cardBrand]) {
                    cardBrandTotals[cardBrand] = 0;
                }
                
                cardBrandTotals[cardBrand] += netAmount;
            }
        }
        
        // Collect unique names and totals from second file
        const nameTotals = {};
        
        if (file2 && file2Headers.length > 0 && nameIndex !== -1 && amountIndex !== -1) {
            // Common names to match card brands (case-insensitive)
            const commonNames = {
                "visa": "Visa",
                "mastercard": "Mastercard",
                "master": "Mastercard",
                "american express": "American Express",
                "amex": "American Express",
                "discover": "Discover"
            };
            
            // Process all rows from the second file
            jsonData2.forEach(row => {
                if (row.length > Math.max(nameIndex, amountIndex)) {
                    const name = String(row[nameIndex] || "").trim();
                    // Skip if name contains "cash" (case-insensitive)
                    if (name.toLowerCase().includes("cash")) {
                        return;
                    }
                    
                    const amount = parseFloat(row[amountIndex]) || 0;
                    
                    if (name) {
                        let displayName = name;
                        
                        // Try to match with common card brand names
                        const lowerName = name.toLowerCase();
                        for (const [key, value] of Object.entries(commonNames)) {
                            if (lowerName.includes(key) || key.includes(lowerName)) {
                                displayName = value;
                                break;
                            }
                        }
                        
                        if (!nameTotals[displayName]) {
                            nameTotals[displayName] = 0;
                        }
                        
                        nameTotals[displayName] += amount;
                    }
                }
            });
        }
        
        // Add rows for each card brand side by side
        const commonCardBrands = ["Visa", "Mastercard", "American Express", "Discover"];
        
        // First add the common card brands in a specific order
        commonCardBrands.forEach(brand => {
            const leftValue = cardBrandTotals[brand] ? 
                parseFloat(cardBrandTotals[brand].toFixed(2)) : 0;
                
            const rightValue = nameTotals[brand] ? 
                parseFloat(nameTotals[brand].toFixed(2)) : 0;
                
            // Calculate difference (Hub Report - Sales Report)
            const difference = parseFloat((leftValue - rightValue).toFixed(2));
                
            filteredResults.push([
                brand,          // Card Brand
                leftValue,      // Hub Report
                rightValue,     // Sales Report
                difference,     // Difference
                "",
                ""
            ]);
            
            // Remove from objects so we don't process them again
            delete cardBrandTotals[brand];
            delete nameTotals[brand];
        });
        
        // Then add any other card brands that might be in the data (excluding the common ones)
        const otherBrands = new Set([
            ...Object.keys(cardBrandTotals).filter(b => !b.toLowerCase().includes("cash") && !commonCardBrands.includes(b)),
            ...Object.keys(nameTotals).filter(n => !n.toLowerCase().includes("cash") && !commonCardBrands.includes(n))
        ]);
        
        [...otherBrands].sort().forEach(brand => {
            const leftValue = cardBrandTotals[brand] ? 
                parseFloat(cardBrandTotals[brand].toFixed(2)) : 0;
                
            const rightValue = nameTotals[brand] ? 
                parseFloat(nameTotals[brand].toFixed(2)) : 0;
            
            // Calculate difference (Hub Report - Sales Report)
            const difference = parseFloat((leftValue - rightValue).toFixed(2));
                
            filteredResults.push([
                brand,          // Card Brand
                leftValue || 0, // Hub Report
                rightValue || 0,// Sales Report
                difference,     // Difference
                "",
                ""
            ]);
        });
    }
    
// If the cell colors don't work when using the website's built-in downloadResults function,
// We need to create a compatible version for the website code
function downloadResults(results) {
    // Create a workbook and worksheet
    const workbook = XLSX.utils.book_new();
    const worksheet = XLSX.utils.aoa_to_sheet(results);
    
    // Find the summary section
    let summaryRowStart = -1;
    for (let i = 0; i < results.length; i++) {
        if (results[i][0] === "Card Brand") {
            summaryRowStart = i;
            break;
        }
    }
    
    // If we found the summary section, add cell styling
    if (summaryRowStart > 0) {
        // Add cell styles for the difference column
        for (let i = summaryRowStart + 1; i < results.length; i++) {
            const row = results[i];
            if (row.length >= 4 && row[3] !== "" && !isNaN(parseFloat(row[3]))) {
                const difference = parseFloat(row[3]);
                const cellRef = XLSX.utils.encode_cell({r: i, c: 3}); // Column D (index 3)
                
                // Apply the style
                if (!worksheet[cellRef]) worksheet[cellRef] = {};
                worksheet[cellRef].s = {
                    fill: {
                        patternType: "solid",
                        fgColor: { rgb: difference < 0 ? "FFFF0000" : "FF92D050" } // Red if negative, green if positive/zero
                    },
                    font: {
                        color: { rgb: difference < 0 ? "FFFFFFFF" : "FF000000" } // White text on red, black on green
                    },
                    alignment: {
                        horizontal: "right"
                    }
                };
            }
        }
    }
    
    // Add the worksheet to the workbook
    XLSX.utils.book_append_sheet(workbook, worksheet, "Results");
    
    // Write the workbook and trigger download
    XLSX.writeFile(workbook, "Comparison_Results.xlsx");
}

// Return the filtered results data
return filteredResults;
}