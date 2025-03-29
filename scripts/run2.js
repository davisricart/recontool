/**
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
    const newColumns = ["K-R", "Count"];
    
    // Create result array starting with the header row (including new columns)
    const resultData = [columnsToKeep.concat(newColumns)];
    
    // Process each row from the first file
    jsonData1.forEach(row => {
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
        
        resultData.push(filteredRow);
    });
    
    // Process the second file for display (if provided)
    // ================================================
    if (file2 && file2Headers.length > 0) {
        // Add a blank row separator
        resultData.push(Array(columnsToKeep.length + newColumns.length).fill(""));
        
        // Add header row for second file
        const secondFileHeadersWithCount2 = [...file2Headers, "Count2"];
        resultData.push(secondFileHeadersWithCount2);
        
        // Process all data rows from second file
        jsonData2.forEach(row => {
            // Copy the row and add empty Count2 column
            const processedRow = [...row, ""];
            resultData.push(processedRow);
        });
    }
    
    return resultData;
}
