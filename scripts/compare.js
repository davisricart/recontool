async function compareAndDisplayData(XLSX, file1, file2) {
    console.log("Starting comparison...");

    // Function to process the Payments Hub file
    function processPaymentsHub(content) {
        console.log("Processing Payments Hub file...");
        const workbook = XLSX.read(content, { type: 'array' });
        const sheetName = workbook.SheetNames[0];
        const sheet = workbook.Sheets[sheetName];
        let jsonData = XLSX.utils.sheet_to_json(sheet, { header: 1 });

        // Required headers for Payments Hub file
        const requiredHeaders = [
            "Date", "Customer Name", "Total Transaction Amount",
            "Cash Discounting Amount", "Card Brand"
        ];

        // Get header indexes
        const headerRow = jsonData[0].map(header => 
            typeof header === 'string' ? header.trim() : String(header)
        );
        const headerIndexes = requiredHeaders.map(header => 
            headerRow.findIndex(h => h.toLowerCase() === header.toLowerCase())
        );

        // Check if all required headers are present
        if (headerIndexes.includes(-1)) {
            throw new Error("Some required headers are missing in the Payments Hub file.");
        }

        // Process data rows
        let processedData = jsonData.slice(1).map(row => {
            // Ensure row is an array and has enough elements
            const safeRow = Array.isArray(row) ? row : [];
            
            let newRow = headerIndexes.map(index => 
                index >= 0 && index < safeRow.length ? 
                    (safeRow[index] !== null && safeRow[index] !== undefined ? 
                        String(safeRow[index]).trim() : "") 
                    : ""
            );

            // Format date (remove time if present)
            if (newRow[0]) {
                // Handle different date formats
                const dateStr = String(newRow[0]);
                const dateParts = dateStr.includes("/") ? 
                    dateStr.split("/") : 
                    dateStr.split(" ")[0].split("-");
                
                if (dateParts.length === 3) {
                    // Ensure consistent date format YYYY-MM-DD
                    const [part1, part2, part3] = dateParts;
                    const year = part3.length === 4 ? part3 : part1;
                    const month = part3.length === 4 ? part1 : part2;
                    const day = part3.length === 4 ? part2 : part3;
                    
                    newRow[0] = `${year}-${month.padStart(2, "0")}-${day.padStart(2, "0")}`;
                }
            }

            return newRow;
        });

        // Map processed data to objects
        return processedData.map(row => ({
            date: row[0] || "",
            customerName: row[1] || "",
            totalAmount: row[2] || "",
            cashDiscount: row[3] || "",
            cardBrand: row[4] || ""
        }));
    }

    // Function to process the Sales Totals file
    function processSalesTotals(content) {
        console.log("Processing Sales Totals file...");
        const workbook = XLSX.read(content, { type: 'array' });
        const sheetName = workbook.SheetNames[0];
        const sheet = workbook.Sheets[sheetName];
        let jsonData = XLSX.utils.sheet_to_json(sheet, { header: 1 });

        // Required headers for Sales Totals file
        const requiredHeaders = ["Name", "Date Closed", "Amount"];

        // Get header indexes
        const headerRow = jsonData[0].map(header => 
            typeof header === 'string' ? header.trim() : String(header)
        );
        const headerIndexes = requiredHeaders.map(header => 
            headerRow.findIndex(h => h.toLowerCase() === header.toLowerCase())
        );

        // Check if all required headers are present
        if (headerIndexes.includes(-1)) {
            throw new Error("Some required headers are missing in the Sales Totals file.");
        }

        // Process data rows
        let processedData = jsonData.slice(1).map(row => {
            // Ensure row is an array and has enough elements
            const safeRow = Array.isArray(row) ? row : [];
            
            let newRow = headerIndexes.map(index => 
                index >= 0 && index < safeRow.length ? 
                    (safeRow[index] !== null && safeRow[index] !== undefined ? 
                        String(safeRow[index]).trim() : "") 
                    : ""
            );

            // Format date (remove time if present)
            if (newRow[1]) {
                // Handle different date formats
                const dateStr = String(newRow[1]);
                const dateParts = dateStr.includes("/") ? 
                    dateStr.split("/") : 
                    dateStr.split(" ")[0].split("-");
                
                if (dateParts.length === 3) {
                    // Ensure consistent date format YYYY-MM-DD
                    const [part1, part2, part3] = dateParts;
                    const year = part3.length === 4 ? part3 : part1;
                    const month = part3.length === 4 ? part1 : part2;
                    const day = part3.length === 4 ? part2 : part3;
                    
                    newRow[1] = `${year}-${month.padStart(2, "0")}-${day.padStart(2, "0")}`;
                }
            }

            return newRow;
        });

        // Map processed data to objects
        return processedData.map(row => ({
            name: row[0] || "",
            dateClosed: row[1] || "",
            amount: row[2] || ""
        }));
    }

    try {
        console.log("Reading and processing files...");
        // Determine which file is which based on their content
        const workbook1 = XLSX.read(file1, { type: 'array' });
        const workbook2 = XLSX.read(file2, { type: 'array' });
        
        const sheet1 = workbook1.Sheets[workbook1.SheetNames[0]];
        const sheet2 = workbook2.Sheets[workbook2.SheetNames[0]];
        
        const headers1 = XLSX.utils.sheet_to_json(sheet1, { header: 1 })[0].map(h => 
            typeof h === 'string' ? h.trim().toLowerCase() : String(h).toLowerCase()
        );
        const headers2 = XLSX.utils.sheet_to_json(sheet2, { header: 1 })[0].map(h => 
            typeof h === 'string' ? h.trim().toLowerCase() : String(h).toLowerCase()
        );

        let paymentsData, salesData;
        
        // Determine which file is Payments Hub and which is Sales Totals
        if (headers1.includes("total transaction amount") && headers1.includes("customer name")) {
            paymentsData = processPaymentsHub(file1);
            salesData = processSalesTotals(file2);
        } else {
            paymentsData = processPaymentsHub(file2);
            salesData = processSalesTotals(file1);
        }

        console.log("Files processed successfully. Matching records...");

        // Initialize matched records with headers
        let matchedRecords = [
            ["Date", "Customer Name", "Total Transaction Amount", "Cash Discounting Amount", "Card Brand", "Sales Amount"]
        ];

        // Match records from Payments Hub and Sales Totals
        paymentsData.forEach(payment => {
            let match = salesData.find(sale =>
                sale.dateClosed === payment.date &&
                sale.name.toLowerCase().trim() === payment.customerName.toLowerCase().trim()
            );

            // If a match is found, add it to the results
            if (match) {
                matchedRecords.push([
                    payment.date,
                    payment.customerName,
                    payment.totalAmount,
                    payment.cashDiscount,
                    payment.cardBrand,
                    match.amount
                ]);
            }
        });

        console.log("Comparison complete. Returning results...");
        return matchedRecords;
    } catch (error) {
        console.error("Error during comparison:", error);
        // Return an array with the error message
        return [["Error", error.message]];
    }
}