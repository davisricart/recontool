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
    
    // Add an extra row at the end with the total in the K-R column
    const totalRow = ["", "", "", "", "TOTAL:"];
    
    // Calculate sum of all K-R values
    let krTotal = 0;
    for (let i = 1; i < resultData.length; i++) {
        // Get K-R value from each row (at 5th index after the 5 initial columns)
        const krValue = parseFloat(resultData[i][5]) || 0;
        krTotal += krValue;
    }
    
    // Add the total to the K-R column position
    totalRow.push(krTotal.toFixed(2));
    
    // Add empty Count cell
    totalRow.push("");
    
    // Add the total row to the result data
    resultData.push(totalRow);
    
    return resultData;
}
