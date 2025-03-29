// Check all four conditions
                const dateMatches = firstFileDateStr && secondFileDateStr && 
                    firstFileDateStr === secondFileDateStr;
                
                const nameMatches = secondFileName && cardBrand && (
                    cardBrand.toLowerCase().includes(secondFileName) || 
                    secondFileName.includes(cardBrand.toLowerCase())/**
 * This script processes two Excel files:
 * 1. For the first file, it keeps only the specified columns: "Date", "Customer Name", 
 *    "Total Transaction Amount", "Cash Discounting Amount", and "Card Brand".
 *    It also adds "K-R" (calculated as Total Amount - Cash Discounting) and "Count" columns.
 * 
 * 2. For the second file, it finds the "Amount" column, converts all values to number format,
 *    and adds a new "Count2" column.
 * 
 * @param {Object} XLSX - The SheetJS library object
 * @param {ArrayBuffer} file1 - The first uploaded Excel file data
 * @param {ArrayBuffer} file2 - The second uploaded Excel file data
 * @returns {Array} An array of arrays representing the processed data
 */
function compareAndDisplayData(XLSX, file1, file2) {
    // Parse both files first so we can do cross-file comparisons
    // ==========================================================
    
    // Parse the first Excel file
    const workbook1 = XLSX.read(file1, {
        cellDates: true,
        dateNF: 'yyyy-mm-dd'
    });
    
    // Get the first sheet from the first workbook
    const sheetName1 = workbook1.SheetNames[0];
    const worksheet1 = workbook1.Sheets[sheetName1];
    
    // Convert worksheet to JSON with headers
    const jsonData1 = XLSX.utils.sheet_to_json(worksheet1);
    
    // Parse the second Excel file if available
    let jsonData2 = [];
    let file2Headers = [];
    let dateClosedIndex = -1;
    let nameIndex = -1;
    let amountIndex = -1;
    
    if (file2) {
        try {
            // Parse the second Excel file
            const workbook2 = XLSX.read(file2, {
                cellDates: true
            });
            
            // Get the first sheet from the second workbook
            const sheetName2 = workbook2.SheetNames[0];
            const worksheet2 = workbook2.Sheets[sheetName2];
            
            // Convert to JSON array with headers
            const data = XLSX.utils.sheet_to_json(worksheet2, { header: 1 });
            
            // Store headers
            file2Headers = data[0] || [];
            
            // Find the index of required columns
            dateClosedIndex = file2Headers.findIndex(header => 
                typeof header === "string" && header.trim().toLowerCase() === "date closed"
            );
            
            nameIndex = file2Headers.findIndex(header => 
                typeof header === "string" && header.trim().toLowerCase() === "name"
            );
            
            amountIndex = file2Headers.findIndex(header => 
                typeof header === "string" && header.trim().toLowerCase() === "amount"
            );
            
            // Store data rows
            jsonData2 = data.slice(1); // Skip header row
            
            // Normalize amount values in the second file
            if (amountIndex !== -1) {
                jsonData2.forEach(row => {
                    if (amountIndex < row.length && row[amountIndex] !== undefined) {
                        // Convert the amount to number
                        let amount = row[amountIndex];
                        if (typeof amount === "string") {
                            amount = amount.replace(/[^0-9.-]+/g, "");
                        }
                        row[amountIndex] = parseFloat(amount) || 0;
                    }
                });
            }
        } catch (error) {
            console.error("Error processing second file:", error);
        }
    }
    
    // Process the first file
    // ======================
    
    // Define the columns to keep from first file
    const columnsToKeep = ["Date", "Customer Name", "Total Transaction Amount", "Cash Discounting Amount", "Card Brand"];
    
    // Define the new columns to add
    const newColumns = ["K-R", "Count", "Final Count"];
    
    // Create result array starting with the header row (including new columns)
    const resultData = [columnsToKeep.concat(newColumns)];
    
    // Process each row from the first file
    jsonData1.forEach((row, rowIndex) => {
        const filteredRow = [];
        
        // Extract values for comparison before filtering
        let firstFileDate = null;
        let cardBrand = "";
        let krValue = 0;
        
        // Keep only the specified columns
        columnsToKeep.forEach(column => {
            // For Date column, format as MM/DD/YYYY if it's a date object
            if (column === "Date") {
                if (row[column] instanceof Date) {
                    const date = row[column];
                    firstFileDate = date; // Store date for comparison
                    const year = date.getFullYear();
                    const month = String(date.getMonth() + 1).padStart(2, '0');
                    const day = String(date.getDate()).padStart(2, '0');
                    filteredRow.push(`${month}/${day}/${year}`);
                } else {
                    filteredRow.push(row[column] !== undefined ? row[column] : "");
                    if (row[column]) {
                        // Try to parse the date if it's a string
                        try {
                            firstFileDate = new Date(row[column]);
                        } catch (e) {
                            firstFileDate = null;
                        }
                    }
                }
            } else if (column === "Card Brand") {
                cardBrand = row[column] || "";
                filteredRow.push(cardBrand);
            } else {
                // For other columns, just add the value (or empty string if not present)
                filteredRow.push(row[column] !== undefined ? row[column] : "");
            }
        });
        
        // Calculate K-R value (Total Transaction Amount - Cash Discounting Amount)
        const totalAmount = parseFloat(row["Total Transaction Amount"]) || 0;
        const discountAmount = parseFloat(row["Cash Discounting Amount"]) || 0;
        krValue = totalAmount - discountAmount;
        
        // Add K-R value (formatted to 2 decimal places)
        filteredRow.push(krValue.toFixed(2));
        
        // Perform COUNTIF-like functionality for Count column
        let countMatches = 0;
        
        // Only compare if all necessary columns exist in the second file
        if (dateClosedIndex !== -1 && nameIndex !== -1 && amountIndex !== -1 && firstFileDate) {
            // Loop through second file rows to compare
            jsonData2.forEach(secondRow => {
                // Check if row has all the required fields
                if (secondRow.length > Math.max(dateClosedIndex, nameIndex, amountIndex)) {
                    // Get the date from second file
                    let secondFileDate = secondRow[dateClosedIndex];
                    if (typeof secondFileDate === 'string') {
                        // Try to parse string date
                        try {
                            secondFileDate = new Date(secondFileDate);
                        } catch (e) {
                            secondFileDate = null;
                        }
                    }
                    
                    // Get name and amount from second file
                    const secondFileName = String(secondRow[nameIndex] || "").trim().toLowerCase();
                    const secondFileAmount = parseFloat(secondRow[amountIndex]) || 0;
                    
                    // Format firstFileDate and secondFileDate to remove time component for comparison
                    const firstFileDateStr = firstFileDate instanceof Date ? 
                        `${firstFileDate.getFullYear()}-${String(firstFileDate.getMonth() + 1).padStart(2, '0')}-${String(firstFileDate.getDate()).padStart(2, '0')}` : '';
                    
                    const secondFileDateStr = secondFileDate instanceof Date ? 
                        `${secondFileDate.getFullYear()}-${String(secondFileDate.getMonth() + 1).padStart(2, '0')}-${String(secondFileDate.getDate()).padStart(2, '0')}` : '';
                    
                    // Compare values (case-insensitive for text)
                    const dateMatches = firstFileDateStr && secondFileDateStr && 
                        firstFileDateStr === secondFileDateStr;
                    
                    const nameMatches = secondFileName && cardBrand && (
                        cardBrand.trim().toLowerCase().includes(secondFileName) || 
                        secondFileName.includes(cardBrand.trim().toLowerCase())
                    );
                    
                    const amountMatches = Math.abs(krValue - secondFileAmount) < 0.01; // Allow small rounding differences
                    
                    // Increment count if all conditions match
                    if (dateMatches && nameMatches && amountMatches) {
                        countMatches++;
                    }
                }
            });
        }
        
        // Add count value
        filteredRow.push(countMatches.toString());
        
        // Add placeholder for Final Count (will be calculated after processing second file)
        filteredRow.push("");
        
        resultData.push(filteredRow);
        firstFileData.push(filteredRow);
    });
    
    // Process the second file for display (if provided)
    // ================================================
    if (file2 && file2Headers.length > 0) {
        // Create a separate array to store the Count2 values for each row in the second file
        const secondFileWithCount2 = [];
        
        // Process all data rows from second file to calculate Count2 values
        jsonData2.forEach((row, rowIndex) => {
            // Make a copy of the row
            const processedRow = [...row];
            
            // Get values for comparison
            let secondFileDate = null;
            let secondFileName = "";
            let secondFileAmount = 0;
            
            // Extract date from second file
            if (dateClosedIndex !== -1 && dateClosedIndex < row.length) {
                let dateValue = row[dateClosedIndex];
                if (typeof dateValue === 'string') {
                    try {
                        secondFileDate = new Date(dateValue);
                    } catch (e) {
                        secondFileDate = null;
                    }
                } else if (dateValue instanceof Date) {
                    secondFileDate = dateValue;
                }
            }
            
            // Extract name from second file
            if (nameIndex !== -1 && nameIndex < row.length) {
                secondFileName = String(row[nameIndex] || "").trim().toLowerCase();
            }
            
            // Extract amount from second file
            if (amountIndex !== -1 && amountIndex < row.length) {
                let amountValue = row[amountIndex];
                if (typeof amountValue === 'string') {
                    amountValue = amountValue.replace(/[^0-9.-]+/g, "");
                }
                secondFileAmount = parseFloat(amountValue) || 0;
            }
            
            // Format date for string comparison
            const secondFileDateStr = secondFileDate instanceof Date ? 
                `${secondFileDate.getFullYear()}-${String(secondFileDate.getMonth() + 1).padStart(2, '0')}-${String(secondFileDate.getDate()).padStart(2, '0')}` : '';
            
            // Compare with each row from first file
            let countMatches = 0;
            
            firstFileData.forEach(firstFileRow => {
                // Extract date, card brand, and K-R from first file row
                let firstFileDate = null;
                if (firstFileRow[0]) { // Date is at index 0
                    // Try to parse date from MM/DD/YYYY format
                    const parts = firstFileRow[0].split('/');
                    if (parts.length === 3) {
                        const month = parseInt(parts[0]) - 1;
                        const day = parseInt(parts[1]);
                        const year = parseInt(parts[2]);
                        firstFileDate = new Date(year, month, day);
                    }
                }
                
                const firstFileDateStr = firstFileDate instanceof Date ? 
                    `${firstFileDate.getFullYear()}-${String(firstFileDate.getMonth() + 1).padStart(2, '0')}-${String(firstFileDate.getDate()).padStart(2, '0')}` : '';
                
                const firstFileCardBrand = String(firstFileRow[4] || "").trim().toLowerCase(); // Card Brand is at index 4
                const firstFileKR = parseFloat(firstFileRow[5] || 0); // K-R is at index 5
                
                // Check if all conditions match
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
            
            // Add count value to Count2 column
            processedRow.push(countMatches.toString());
            
            // Store the row with Count2 value
            secondFileWithCount2.push(processedRow);
        });
        
        // Now calculate Final Count for first file rows
        firstFileData.forEach((firstFileRow, firstFileIndex) => {
            const date = firstFileRow[0]; // Date at index 0
            const cardBrand = String(firstFileRow[4] || "").trim().toLowerCase(); // Card Brand at index 4
            const kr = parseFloat(firstFileRow[5] || 0); // K-R at index 5
            const count = parseInt(firstFileRow[6] || 0); // Count at index 6
            
            // Parse date from MM/DD/YYYY format
            let firstFileDate = null;
            if (date) {
                const parts = date.split('/');
                if (parts.length === 3) {
                    const month = parseInt(parts[0]) - 1;
                    const day = parseInt(parts[1]);
                    const year = parseInt(parts[2]);
                    firstFileDate = new Date(year, month, day);
                }
            }
            
            const firstFileDateStr = firstFileDate instanceof Date ? 
                `${firstFileDate.getFullYear()}-${String(firstFileDate.getMonth() + 1).padStart(2, '0')}-${String(firstFileDate.getDate()).padStart(2, '0')}` : '';
            
            // Calculate Final Count based on all four criteria
            let finalCount = 0;
            
            secondFileWithCount2.forEach(secondFileRow => {
                // Extract values from second file row
                let secondFileDate = null;
                if (dateClosedIndex !== -1 && dateClosedIndex < secondFileRow.length) {
                    let dateValue = secondFileRow[dateClosedIndex];
                    if (typeof dateValue === 'string') {
                        try {
                            secondFileDate = new Date(dateValue);
                        } catch (e) {
                            secondFileDate = null;
                        }
                    } else if (dateValue instanceof Date) {
                        secondFileDate = dateValue;
                    }
                }
                
                const secondFileDateStr = secondFileDate instanceof Date ? 
                    `${secondFileDate.getFullYear()}-${String(secondFileDate.getMonth() + 1).padStart(2, '0')}-${String(secondFileDate.getDate()).padStart(2, '0')}` : '';
                
                const secondFileName = nameIndex !== -1 && nameIndex < secondFileRow.length ?
                    String(secondFileRow[nameIndex] || "").trim().toLowerCase() : "";
                    
                const secondFileAmount = amountIndex !== -1 && amountIndex < secondFileRow.length ?
                    parseFloat(secondFileRow[amountIndex]) || 0 : 0;
                    
                const secondFileCount2 = parseInt(secondFileRow[secondFileRow.length - 1] || 0);
                
                // Check all four conditions
                const dateMatches = firstFileDateStr && secondFileDateStr && 
                    firstFileDateStr === secondFileDateStr;
                
                const nameMatches = secondFileName && cardBrand && (
                    cardBrand.toLowerCase().includes(secondFileName) || 
                    secondFileName.includes(cardBrand.toLowerCase())
                );
                
                const amountMatches = Math.abs(kr - secondFileAmount) < 0.01;
                
                const countMatches = count === secondFileCount2;
                
                if (dateMatches && nameMatches && amountMatches && countMatches) {
                    finalCount++;
                }
            });
            
            // Update Final Count in the result data
            // (firstFileIndex + 1 because index 0 is the header row)
            resultData[firstFileIndex + 1][7] = finalCount.toString();
        });
        
        // Add a blank row separator
        resultData.push(Array(columnsToKeep.length + newColumns.length).fill(""));
        
        // Add header row for second file
        const secondFileHeadersWithCount2 = [...file2Headers, "Count2"];
        resultData.push(secondFileHeadersWithCount2);
        
        // Add the second file data rows
        secondFileWithCount2.forEach(row => {
            resultData.push(row);
        });
    }
    
    return resultData;
}
