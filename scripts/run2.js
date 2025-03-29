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
    // Process the first file
    // ======================
    
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
    
    // Define the columns to keep
    const columnsToKeep = ["Date", "Customer Name", "Total Transaction Amount", "Cash Discounting Amount", "Card Brand"];
    
    // Define the new columns to add
    const newColumns = ["K-R", "Count"];
    
    // Create result array starting with the header row (including new columns)
    const resultData = [columnsToKeep.concat(newColumns)];
    
    // Process each row and keep only the specified columns
    jsonData1.forEach(row => {
        const filteredRow = [];
        
        columnsToKeep.forEach(column => {
            // For Date column, format as MM/DD/YYYY if it's a date object
            if (column === "Date" && row[column] instanceof Date) {
                const date = row[column];
                const year = date.getFullYear();
                const month = String(date.getMonth() + 1).padStart(2, '0');
                const day = String(date.getDate()).padStart(2, '0');
                filteredRow.push(`${month}/${day}/${year}`);
            } else {
                // For other columns, just add the value (or empty string if not present)
                filteredRow.push(row[column] !== undefined ? row[column] : "");
            }
        });
        
        // Calculate K-R value (Total Transaction Amount - Cash Discounting Amount)
        const totalAmount = parseFloat(row["Total Transaction Amount"]) || 0;
        const discountAmount = parseFloat(row["Cash Discounting Amount"]) || 0;
        const krValue = totalAmount - discountAmount;
        
        // Add K-R value (formatted to 2 decimal places)
        filteredRow.push(krValue.toFixed(2));
        
        // Add empty Count column
        filteredRow.push("");
        
        resultData.push(filteredRow);
    });
    
    // Process the second file (if provided)
    // =====================================
    if (file2) {
        try {
            // Parse the second Excel file
            const workbook2 = XLSX.read(file2, {
                cellDates: true
            });
            
            // Get the first sheet from the second workbook
            const sheetName2 = workbook2.SheetNames[0];
            const worksheet2 = workbook2.Sheets[sheetName2];
            
            // Convert to JSON to get headers
            const jsonData2 = XLSX.utils.sheet_to_json(worksheet2, { header: 1 });
            
            // Find the "Amount" column index
            const headers = jsonData2[0];
            const amountColumnIndex = headers.findIndex(header => 
                typeof header === "string" && header.trim().toLowerCase() === "amount"
            );
            
            // If "Amount" column exists, process it
            if (amountColumnIndex !== -1) {
                // Add a blank row separator
                resultData.push(Array(columnsToKeep.length + newColumns.length).fill(""));
                
                // Add header row for second file
                const secondFileHeaders = [...headers];
                secondFileHeaders.push("Count2"); // Add Count2 column
                resultData.push(secondFileHeaders);
                
                // Process all rows except header
                for (let i = 1; i < jsonData2.length; i++) {
                    const row = jsonData2[i];
                    
                    // Make a copy of the row
                    const processedRow = [...row];
                    
                    // Convert the Amount column to number if it exists in this row
                    if (amountColumnIndex < row.length && row[amountColumnIndex] !== undefined) {
                        // Try to convert to number
                        let amount = row[amountColumnIndex];
                        if (typeof amount === "string") {
                            // Remove any currency symbols, commas, etc.
                            amount = amount.replace(/[^0-9.-]+/g, "");
                        }
                        // Convert to number and format to 2 decimal places
                        processedRow[amountColumnIndex] = parseFloat(amount) || 0;
                    }
                    
                    // Add empty Count2 column
                    processedRow.push("");
                    
                    resultData.push(processedRow);
                }
            }
        } catch (error) {
            // If there's an error processing the second file, just continue with the first file results
            console.error("Error processing second file:", error);
        }
    }
    
    return resultData;
}
