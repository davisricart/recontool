async function processPaymentsHub(file) {
    return new Promise((resolve, reject) => {
        const reader = new FileReader();
        reader.onload = (event) => {
            const data = new Uint8Array(event.target.result);
            const workbook = XLSX.read(data, { type: 'array' });

            const sheetName = workbook.SheetNames[0];
            const sheet = workbook.Sheets[sheetName];

            let jsonData = XLSX.utils.sheet_to_json(sheet, { header: 1 });

            const requiredHeaders = [
                "Date", "Customer Name", "Total Transaction Amount",
                "Cash Discounting Amount", "Card Brand"
            ];

            const headerRow = jsonData[0];
            const headerIndexes = requiredHeaders.map(header => headerRow.indexOf(header));

            if (headerIndexes.includes(-1)) {
                reject("Some required headers are missing in the Payments Hub file.");
                return;
            }

            let processedData = jsonData.slice(1).map(row => {
                let newRow = headerIndexes.map(index => row[index] || "");

                if (newRow[0]) {
                    newRow[0] = newRow[0].split(" ")[0] || "";
                }

                return newRow;
            });

            let paymentsData = processedData.map(row => ({
                date: row[0],
                customerName: row[1],
                totalAmount: row[2],
                cashDiscount: row[3],
                cardBrand: row[4]
            }));

            resolve(paymentsData);
        };

        reader.onerror = (error) => reject(error);
        reader.readAsArrayBuffer(file);
    });
}

async function processSalesTotals(file) {
    return new Promise((resolve, reject) => {
        const reader = new FileReader();
        reader.onload = (event) => {
            const data = new Uint8Array(event.target.result);
            const workbook = XLSX.read(data, { type: 'array' });

            const sheetName = workbook.SheetNames[0];
            const sheet = workbook.Sheets[sheetName];

            let jsonData = XLSX.utils.sheet_to_json(sheet, { header: 1 });

            const requiredHeaders = ["Name", "Date Closed", "Amount"];

            const headerRow = jsonData[0];
            const headerIndexes = requiredHeaders.map(header => headerRow.indexOf(header));

            if (headerIndexes.includes(-1)) {
                reject("Some required headers are missing in the Sales Totals file.");
                return;
            }

            let processedData = jsonData.slice(1).map(row => {
                let newRow = headerIndexes.map(index => row[index] || "");

                if (newRow[1]) {
                    newRow[1] = newRow[1].split(" ")[0] || "";
                }

                return newRow;
            });

            let salesData = processedData.map(row => ({
                name: row[0],
                dateClosed: row[1],
                amount: row[2]
            }));

            resolve(salesData);
        };

        reader.onerror = (error) => reject(error);
        reader.readAsArrayBuffer(file);
    });
}

async function compareAndDisplayData() {
    const file1 = document.getElementById('file1').files[0];
    const file2 = document.getElementById('file2').files[0];

    if (!file1 || !file2) {
        alert("Please upload both files.");
        return;
    }

    try {
        const paymentsData = await processPaymentsHub(file1);
        const salesData = await processSalesTotals(file2);

        let matchedRecords = [];

        paymentsData.forEach(payment => {
            let match = salesData.find(sale =>
                sale.dateClosed === payment.date &&
                sale.name.toLowerCase().trim() === payment.customerName.toLowerCase().trim()
            );

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

        if (matchedRecords.length === 0) {
            alert("No matches found.");
            return;
        }

        displayResults(matchedRecords);

    } catch (error) {
        console.error("Error processing files:", error);
    }
}

function displayResults(data) {
    const previewTable = document.getElementById('previewTable');
    const previewArea = document.getElementById('previewArea');
    const downloadButton = document.getElementById('downloadButton');

    previewTable.innerHTML = "";
    previewArea.style.display = "block";

    let headerRow = document.createElement("tr");
    ["Date", "Customer Name", "Total Transaction Amount", "Cash Discounting Amount", "Card Brand", "Sales Amount"].forEach(headerText => {
        let th = document.createElement("th");
        th.textContent = headerText;
        headerRow.appendChild(th);
    });
    previewTable.appendChild(headerRow);

    data.forEach(rowData => {
        let row = document.createElement("tr");
        rowData.forEach(cellData => {
            let td = document.createElement("td");
            td.textContent = cellData;
            row.appendChild(td);
        });
        previewTable.appendChild(row);
    });

    downloadButton.style.display = "inline-block";
    downloadButton.onclick = () => downloadResults(data);
}

function downloadResults(results) {
    const workbook = XLSX.utils.book_new();
    const worksheet = XLSX.utils.aoa_to_sheet([["Date", "Customer Name", "Total Transaction Amount", "Cash Discounting Amount", "Card Brand", "Sales Amount"], ...results]);
    XLSX.utils.book_append_sheet(workbook, worksheet, "Results");
    XLSX.writeFile(workbook, "Comparison_Results.xlsx");
}

document.getElementById('compareButton').addEventListener('click', compareAndDisplayData);
