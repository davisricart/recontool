async function compareAndDisplayData(XLSX, file1, file2) {
    console.log("Starting comparison...");

    // Function to process the Payments Hub file
    async function processPaymentsHub(content) {
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
        const headerRow = jsonData[0];
        const headerIndexes = requiredHeaders.map(header => headerRow.indexOf(header));

        // Check if all required headers are present
        if (headerIndexes.includes(-1)) {
            throw new Error("Some required headers are missing in the Payments Hub file.");
        }

        // Process data rows
        let processedData = jsonData.slice(1).map(row => {
            let newRow = headerIndexes.map(index => row[index] || "");
            // Format date (remove time if present)
            if (newRow[0]) newRow[0] = newRow[0].split(" ")[0] || "";
            return newRow;
        });

        // Map processed data to objects
        return processedData.map(row => ({
            date: row[0],
            customerName: row[1],
            totalAmount: row[2],
            cashDiscount: row[3],
            cardBrand: row[4]
        }));
    }

    // Function to process the Sales Totals file
    async function processSalesTotals(content) {
        console.log("Processing Sales Totals file...");
        const workbook = XLSX.read(content, { type: 'array' });
        const sheetName = workbook.SheetNames[0];
        const sheet = workbook.Sheets[sheetName];
        let jsonData = XLSX.utils.sheet_to_json(sheet, { header: 1 });

        // Required headers for Sales Totals file
        const requiredHeaders = ["Name", "Date Closed", "Amount"];

        // Get header indexes
        const headerRow = jsonData[0];
        const headerIndexes = requiredHeaders.map(header => headerRow.indexOf(header));

        // Check if all required headers are present
        if (headerIndexes.includes(-1)) {
            throw new Error("Some required headers are missing in the Sales Totals file.");
        }

        // Process data rows
        let processedData = jsonData.slice(1).map(row => {
            let newRow = headerIndexes.map(index => row[index] || "");
            // Format date (remove time if present)
            if (newRow[1]) newRow[1] = newRow[1].split(" ")[0] || "";
            return newRow;
        });

        // Map processed data to objects
        return processedData.map(row => ({
            name: row[0],
            dateClosed: row[1],
            amount: row[2]
        }));
    }

    try {
        console.log("Reading and processing files...");
        const paymentsData = await processPaymentsHub(file1);
        const salesData = await processSalesTotals(file2);

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