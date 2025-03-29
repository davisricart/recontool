/**
 * This script processes two Excel files and returns a filtered version of the first file
 * keeping only the specified columns: "Date", "Customer Name", "Total Transaction Amount",
 * "Cash Discounting Amount", and "Card Brand".
 * 
 * @param {Object} XLSX - The SheetJS library object
 * @param {ArrayBuffer} file1 - The first uploaded Excel file data
 * @param {ArrayBuffer} file2 - The second uploaded Excel file data
 * @returns {Array} An array of arrays representing the filtered data
 */
function compareAndDisplayData(XLSX, file1, file2) {
    // Parse the first Excel file
    const workbook1 = XLSX.read(file1, {
        cellDates: true,
        dateNF: 'yyyy-mm-dd'
    });
    
    // Get the first sheet from the first workbook
    const sheetName1 = workbook1.SheetNames[0];
    const worksheet1 = workbook1.Sheets[sheetName1];
    
    // Convert worksheet to JSON with headers
    const jsonData = XLSX.utils.sheet_to_json(worksheet1);
    
    // Define the columns to keep
    const columnsToKeep = ["Date", "Customer Name", "Total Transaction Amount", "Cash Discounting Amount", "Card Brand"];
    
    // Define the new columns to add
    const newColumns = ["K-R", "Count"];
    
    // Create result array starting with the header row (including new columns)
    const resultData = [columnsToKeep.concat(newColumns)];
    
    // Process each row and keep only the specified columns
    jsonData.forEach(row => {
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
        
        // Add empty values for K-R and Count columns
        filteredRow.push(""); // K-R column (empty by default)
        filteredRow.push(""); // Count column (empty by default)
        
        resultData.push(filteredRow);
    });
    
    return resultData;
}
