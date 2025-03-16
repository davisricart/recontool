const XLSX = require('xlsx');

// Function to compare Column A from two Excel files
function compareColumnA(file1Path, file2Path) {
    const workbook1 = XLSX.readFile(file1Path);
    const workbook2 = XLSX.readFile(file2Path);

    const sheet1 = workbook1.Sheets[workbook1.SheetNames[0]];
    const sheet2 = workbook2.Sheets[workbook2.SheetNames[0]];

    const columnAFile1 = XLSX.utils.sheet_to_json(sheet1, { header: 1 }).map(row => row[0]);
    const columnAFile2 = XLSX.utils.sheet_to_json(sheet2, { header: 1 }).map(row => row[0]);

    if (columnAFile1.length !== columnAFile2.length) {
        return 'No';
    }

    for (let i = 0; i < columnAFile1.length; i++) {
        if (columnAFile1[i] !== columnAFile2[i]) {
            return 'No';
        }
    }

    return 'Yes';
}

// Export the function
module.exports = compareColumnA;
