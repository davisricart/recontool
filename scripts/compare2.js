const XLSX = require('xlsx');

// Function to compare Column A from two Excel files
function compareColumnA(file1Path, file2Path) {
    // Read both files
    const workbook1 = XLSX.readFile(file1Path);
    const workbook2 = XLSX.readFile(file2Path);

    // Access the first sheet from each workbook
    const sheet1 = workbook1.Sheets[workbook1.SheetNames[0]];
    const sheet2 = workbook2.Sheets[workbook2.SheetNames[0]];

    // Extract Column A data from both sheets
    const columnA1 = XLSX.utils.sheet_to_json(sheet1, { header: 1 }).map(row => row[0]);
    const columnA2 = XLSX.utils.sheet_to_json(sheet2, { header: 1 }).map(row => row[0]);

    // Ensure both columns have the same number of rows
    if (columnA1.length !== columnA2.length) {
        return 'No';
    }

    // Compare each cell in Column A
    for (let i = 0; i < columnA1.length; i++) {
        if (columnA1[i] !== columnA2[i]) {
            return 'No';
        }
    }

    return 'Yes';
}

// Example usage
const file1Path = './file1.xlsx'; // Path to the first Excel file
const file2Path = './file2.xlsx'; // Path to the second Excel file

const result = compareColumnA(file1Path, file2Path);
console.log(`Do all cells in Column A match? ${result}`);
