<!DOCTYPE html>
<html lang="en">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>Compare Excel Files</title>
    <script src="https://cdnjs.cloudflare.com/ajax/libs/xlsx/0.18.5/xlsx.full.min.js"></script>
    <style>
        :root {
            --primary: #4361ee;
            --primary-hover: #3a56d4;
            --secondary: #e6f0ff;
            --text: #333;
            --light-text: #6c757d;
            --border: #dee2e6;
            --success: #38b000;
            --danger: #d00000;
            --background: #f8f9fa;
        }
        
        body { 
            font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif; 
            margin: 0;
            padding: 20px;
            background-color: var(--background);
            color: var(--text);
            line-height: 1.6;
        }
        
        h1 { 
            color: var(--primary); 
            margin-bottom: 1.5rem;
            border-bottom: 2px solid var(--secondary);
            padding-bottom: 0.5rem;
        }
        
        button { 
            padding: 10px 16px; 
            background: var(--primary); 
            color: white; 
            border: none; 
            border-radius: 4px;
            cursor: pointer; 
            margin-right: 10px;
            font-weight: 500;
            transition: background-color 0.2s ease;
            box-shadow: 0 2px 4px rgba(0, 0, 0, 0.1);
        }
        
        button:hover { 
            background: var(--primary-hover); 
            box-shadow: 0 4px 8px rgba(0, 0, 0, 0.1);
        }
        
        input { 
            margin: 10px 0; 
            padding: 10px; 
            width: 100%; 
            max-width: 300px;
            border: 1px solid var(--border);
            border-radius: 4px;
        }
        
        input:focus {
            outline: none;
            border-color: var(--primary);
            box-shadow: 0 0 0 3px rgba(67, 97, 238, 0.15);
        }
        
        #downloadArea { 
            margin-top: 20px; 
            display: none; 
        }
        
        #previewArea { 
            margin-top: 20px; 
            display: none;
            background: white;
            border-radius: 8px;
            padding: 15px;
            box-shadow: 0 2px 10px rgba(0, 0, 0, 0.05);
        }
        
        table { 
            border-collapse: collapse; 
            width: 100%; 
            margin-top: 15px;
            box-shadow: 0 1px 3px rgba(0, 0, 0, 0.1);
            border-radius: 4px;
            overflow: hidden;
        }
        
        th, td { 
            border: 1px solid var(--border); 
            padding: 12px; 
            text-align: left; 
        }
        
        th { 
            background-color: var(--secondary); 
            color: var(--primary);
            font-weight: 600;
        }
        
        tr:nth-child(even) { 
            background-color: #f2f6ff; 
        }
        
        .drop-area {
            border: 2px dashed var(--border);
            border-radius: 8px;
            padding: 25px;
            text-align: center;
            margin: 15px 0;
            max-width: 300px;
            transition: all 0.3s ease;
            background-color: white;
        }
        
        .drop-area.highlight {
            border-color: var(--primary);
            background-color: var(--secondary);
        }
        
        .file-info {
            font-size: 0.9em;
            margin-top: 8px;
            color: var(--light-text);
        }
        
        #status { 
            color: var(--danger); 
            font-weight: 500;
            padding: 10px 0;
            min-height: 20px;
        }
        
        label {
            display: block;
            margin-bottom: 5px;
            font-weight: 500;
            color: var(--text);
        }
        
        #uploadForm {
            display: flex;
            flex-direction: column;
            gap: 20px;
            max-width: 800px;
        }
        
        .file-upload-section {
            background: white;
            padding: 20px;
            border-radius: 8px;
            box-shadow: 0 2px 10px rgba(0, 0, 0, 0.05);
        }
        
        .input-group {
            margin-bottom: 15px;
        }

        input[type="file"] {
            padding: 8px;
            border: 1px solid var(--border);
            border-radius: 4px;
            width: 100%;
            max-width: 300px;
        }

        h3 {
            color: var(--primary);
            margin-top: 0;
            margin-bottom: 15px;
        }

        #downloadButton {
            background-color: var(--success);
        }

        #downloadButton:hover {
            background-color: #2d9200;
        }
    </style>
</head>
<body>
    <h1>Compare Excel Files</h1>
    <form id="uploadForm">
        <div class="file-upload-section">
            <div class="input-group">
                <label for="file1">Upload First File:</label>
                <input type="file" id="file1" accept=".xlsx">
                <div class="drop-area" id="dropArea1">
                    <p>Drag & drop your first Excel file here</p>
                </div>
            </div>
            
            <div class="input-group">
                <label for="file2">Upload Second File:</label>
                <input type="file" id="file2" accept=".xlsx">
                <div class="drop-area" id="dropArea2">
                    <p>Drag & drop your second Excel file here</p>
                </div>
            </div>
        </div>

        <div class="file-upload-section">
            <div class="input-group">
                <label for="scriptName">Enter Code:</label>
                <input type="text" id="scriptName" placeholder="Enter script code name" required>
            </div>

            <div>
                <button type="button" id="compareButton">Run Code</button>
                <button type="button" id="downloadButton" style="display: none;">Download Results</button>
            </div>
            <p id="status"></p>
        </div>
    </form>

    <div id="previewArea">
        <h3>Comparison Results</h3>
        <table id="previewTable"></table>
    </div>

    <script>
        const REPO_OWNER = 'davisricart';
        const REPO_NAME = 'recontool';
        const SCRIPT_PATH = 'scripts';

        // Drag and drop functionality
        function setupDragDrop(dropArea, fileInput) {
            // Prevent default drag behaviors
            ['dragenter', 'dragover', 'dragleave', 'drop'].forEach(eventName => {
                dropArea.addEventListener(eventName, preventDefaults, false);
                document.body.addEventListener(eventName, preventDefaults, false);
            });

            // Highlight drop area when item is dragged over
            ['dragenter', 'dragover'].forEach(eventName => {
                dropArea.addEventListener(eventName, highlight, false);
            });

            ['dragleave', 'drop'].forEach(eventName => {
                dropArea.addEventListener(eventName, unhighlight, false);
            });

            // Handle dropped files
            dropArea.addEventListener('drop', handleDrop, false);

            function preventDefaults(e) {
                e.preventDefault();
                e.stopPropagation();
            }

            function highlight() {
                dropArea.classList.add('highlight');
            }

            function unhighlight() {
                dropArea.classList.remove('highlight');
            }

            function handleDrop(e) {
                const dt = e.dataTransfer;
                const files = dt.files;

                // Set the files to the file input
                fileInput.files = files;
                
                // Trigger change event on file input
                const event = new Event('change');
                fileInput.dispatchEvent(event);
            }
        }

        // Setup drag and drop for both file inputs
        setupDragDrop(
            document.getElementById('dropArea1'), 
            document.getElementById('file1')
        );
        setupDragDrop(
            document.getElementById('dropArea2'), 
            document.getElementById('file2')
        );

        async function fetchComparisonScript(scriptName) {
            const formattedScriptName = `${scriptName}.js`;
            const url = `https://api.github.com/repos/${REPO_OWNER}/${REPO_NAME}/contents/${SCRIPT_PATH}/${formattedScriptName}`;

            try {
                const response = await fetch(url);
                if (!response.ok) {
                    throw new Error('Code not found, please try again. If the error continues please contact us');
                }
                const data = await response.json();
                const decodedContent = atob(data.content);
                return decodedContent;
            } catch (error) {
                throw error;
            }
        }

        function readExcelFile(file) {
            return new Promise((resolve, reject) => {
                const reader = new FileReader();
                reader.onload = (e) => resolve(e.target.result);
                reader.onerror = (e) => reject(e);
                reader.readAsArrayBuffer(file);
            });
        }

        async function compareFiles() {
            const file1 = document.getElementById('file1').files[0];
            const file2 = document.getElementById('file2').files[0];
            const scriptName = document.getElementById('scriptName').value.trim();
            const status = document.getElementById('status');
            const previewTable = document.getElementById('previewTable');
            const previewArea = document.getElementById('previewArea');
            const downloadButton = document.getElementById('downloadButton');

            status.textContent = '';
            previewTable.innerHTML = '';
            previewArea.style.display = 'none';
            downloadButton.style.display = 'none';

            if (!file1 || !file2) {
                status.textContent = 'Please upload both files.';
                return;
            }
            if (!scriptName) {
                status.textContent = 'Please enter a script code.';
                return;
            }

            try {
                status.textContent = 'Fetching script...';
                const scriptContent = await fetchComparisonScript(scriptName);

                // Create an async function that uses the compareAndDisplayData from the script
                const CompareFunction = new Function('XLSX', 'file1', 'file2', 
                    `return (async () => {
                        ${scriptContent}
                        return compareAndDisplayData(XLSX, file1, file2);
                    })();`
                );

                status.textContent = 'Reading files...';
                const content1 = await readExcelFile(file1);
                const content2 = await readExcelFile(file2);

                status.textContent = 'Comparing files...';
                const result = await CompareFunction(XLSX, content1, content2);

                result.forEach((row, rowIndex) => {
                    const tr = document.createElement('tr');
                    row.forEach(cell => {
                        const td = document.createElement(rowIndex === 0 ? 'th' : 'td');
                        td.textContent = cell || '';
                        tr.appendChild(td);
                    });
                    previewTable.appendChild(tr);
                });

                previewArea.style.display = 'block';
                downloadButton.style.display = 'inline-block';
                downloadButton.onclick = () => downloadResults(result);

                status.textContent = 'Comparison complete!';
                status.style.color = 'var(--success)';
            } catch (error) {
                status.textContent = error.message;
                status.style.color = 'var(--danger)';
                console.error(error);
            }
        }

        function downloadResults(results) {
            const workbook = XLSX.utils.book_new();
            const worksheet = XLSX.utils.aoa_to_sheet(results);
            XLSX.utils.book_append_sheet(workbook, worksheet, 'Results');
            XLSX.writeFile(workbook, 'Comparison_Results.xlsx');
        }

        document.getElementById('compareButton').addEventListener('click', compareFiles);
    </script>
</body>
</html>