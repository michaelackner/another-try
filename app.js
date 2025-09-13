class ExcelProcessor {
    constructor() {
        this.workbook = null;
        this.processedData = null;
        this.summaryMetrics = {};
        this.setupEventListeners();
    }

    setupEventListeners() {
        const fileInput = document.getElementById('fileInput');
        const clearFile = document.getElementById('clearFile');
        const processButton = document.getElementById('processButton');
        const downloadButton = document.getElementById('downloadButton');

        fileInput.addEventListener('change', this.handleFileSelect.bind(this));
        clearFile.addEventListener('click', this.clearFile.bind(this));
        processButton.addEventListener('click', this.processFile.bind(this));
        downloadButton.addEventListener('click', this.downloadFile.bind(this));
    }

    handleFileSelect(event) {
        const file = event.target.files[0];
        if (!file) return;

        const fileInfo = document.getElementById('fileInfo');
        const fileName = document.getElementById('fileName');
        const processButton = document.getElementById('processButton');

        fileName.textContent = file.name;
        fileInfo.style.display = 'flex';
        processButton.disabled = false;

        this.hideError();
        this.hideResults();
    }

    clearFile() {
        const fileInput = document.getElementById('fileInput');
        const fileInfo = document.getElementById('fileInfo');
        const processButton = document.getElementById('processButton');

        fileInput.value = '';
        fileInfo.style.display = 'none';
        processButton.disabled = true;

        this.hideError();
        this.hideResults();
    }

    async processFile() {
        this.showLoading();
        this.hideError();
        this.hideResults();

        try {
            const fileInput = document.getElementById('fileInput');
            const file = fileInput.files[0];

            if (!file) {
                throw new Error('No file selected');
            }

            // Validate file type
            if (!file.name.match(/\.(xlsx|xls)$/i)) {
                throw new Error('Please select a valid Excel file (.xlsx or .xls)');
            }

            // Read the Excel file
            const arrayBuffer = await file.arrayBuffer();
            this.workbook = XLSX.read(arrayBuffer, { type: 'array', cellDates: true });

            // Validate file structure
            this.validateFile();

            // Process the data (use setTimeout to allow UI updates)
            await new Promise(resolve => setTimeout(resolve, 100));
            await this.processData();

            // Show results
            this.showResults();

        } catch (error) {
            console.error('Processing error:', error);
            this.showError(error.message);
        } finally {
            this.hideLoading();
        }
    }

    validateFile() {
        const settings = this.getSettings();

        // Check if required sheets exist
        const sheetNames = this.workbook.SheetNames;

        const rawSheet1 = settings.rawSheet1Name || sheetNames[0];
        const rawSheet2 = settings.rawSheet2Name || sheetNames[1];
        const rawSheet3 = settings.rawSheet3Name || sheetNames[2];

        if (!this.workbook.Sheets[rawSheet1]) {
            throw new Error(`Raw Sheet 1 "${rawSheet1}" not found`);
        }
        if (!this.workbook.Sheets[rawSheet2]) {
            throw new Error(`Raw Sheet 2 "${rawSheet2}" not found`);
        }
        if (!this.workbook.Sheets[rawSheet3]) {
            throw new Error(`Raw Sheet 3 "${rawSheet3}" not found`);
        }

        // Validate essential columns in each sheet
        this.validateSheetColumns(rawSheet1, ['B', 'AA', 'M', 'L', 'Q', 'AB', 'AD', 'AL', 'X', 'BZ']);
        this.validateSheetColumns(rawSheet2, ['N', 'AQ', 'AV']);
        this.validateSheetColumns(rawSheet3, ['M', 'BR', 'CN']);

        // Check if Raw Sheet 1 has data rows beyond header
        const sheet1Data = XLSX.utils.sheet_to_json(this.workbook.Sheets[rawSheet1], { header: 1 });
        if (sheet1Data.length < 2) {
            throw new Error('Raw Sheet 1 must have at least one data row');
        }
    }

    validateSheetColumns(sheetName, requiredColumns) {
        const sheet = this.workbook.Sheets[sheetName];
        const range = XLSX.utils.decode_range(sheet['!ref']);

        for (const col of requiredColumns) {
            const colIndex = XLSX.utils.decode_col(col);
            if (colIndex > range.e.c) {
                throw new Error(`Required column ${col} not found in sheet ${sheetName}`);
            }
        }
    }

    async processData() {
        const settings = this.getSettings();

        // Step 1: Build formatted report
        const newWorkbook = this.buildStep1Report(settings);

        // Step 2: Enrich with data from other sheets
        this.enrichStep2Data(newWorkbook, settings);

        this.processedData = newWorkbook;
    }

    buildStep1Report(settings) {
        const newWorkbook = XLSX.utils.book_new();
        const outputSheetName = settings.outputSheetName;

        // Create headers A-V
        const headers = [
            'Varo deal', 'VSA deal', 'VESSEL', 'VMAG %', 'L/C costs',
            'Load insp', 'Discharge inspection', 'Superintendent', 'CIN insurance',
            'CLI insurance', 'Provisional charge', 'TOTAL USD', 'VARO comments',
            'Product', 'Hedge', 'Qty BBL', 'Inco', 'Contractual Location',
            'Risk', 'Date', 'VSA comments', 'Additional information'
        ];

        // Get raw data from Sheet 1
        const rawSheet1Name = settings.rawSheet1Name || this.workbook.SheetNames[0];
        const rawSheet1 = this.workbook.Sheets[rawSheet1Name];
        const rawData = XLSX.utils.sheet_to_json(rawSheet1, { header: 1 });

        // Skip first row (titles) and process data
        const dataRows = rawData.slice(1).filter(row => row && row.length > 0);

        // Map raw data to new format
        const mappedData = dataRows.map(row => {
            const newRow = new Array(22).fill(''); // A-V = 22 columns

            // Map specific columns: B←B, C←AA, N←M, O←L, P←Q, Q←AB, R←AD, S←AL, T←X
            newRow[1] = row[1] || ''; // B ← B
            newRow[2] = row[26] || ''; // C ← AA (26 is column AA)
            newRow[13] = row[12] || ''; // N ← M
            newRow[14] = row[11] || ''; // O ← L
            newRow[15] = row[16] || ''; // P ← Q
            newRow[16] = row[27] || ''; // Q ← AB
            newRow[17] = row[29] || ''; // R ← AD
            newRow[18] = row[37] || ''; // S ← AL
            newRow[19] = this.parseDate(row[23]); // T ← X (23 is column X)

            return newRow;
        });

        // Sort by date (column T, index 19)
        mappedData.sort((a, b) => {
            const dateA = this.parseDate(a[19]);
            const dateB = this.parseDate(b[19]);

            if (!dateA && !dateB) return 0;
            if (!dateA) return 1; // NaT at end
            if (!dateB) return -1;

            return dateA - dateB;
        });

        // Group by month and insert spacers
        const finalData = [headers]; // Start with headers
        let currentMonth = null;
        let currentYear = null;

        mappedData.forEach(row => {
            const date = this.parseDate(row[19]);
            if (date) {
                const month = date.getMonth();
                const year = date.getFullYear();

                if (currentMonth !== month || currentYear !== year) {
                    // Insert 3 blank rows (except for first month)
                    if (currentMonth !== null) {
                        finalData.push(new Array(22).fill(''));
                        finalData.push(new Array(22).fill(''));
                        finalData.push(new Array(22).fill(''));
                    }

                    currentMonth = month;
                    currentYear = year;

                    // Mark this row to get month label in column A
                    row._isFirstOfMonth = true;
                    row._monthDate = date;
                }
            }

            // Format date as DD/MM/YYYY for display
            if (row[19] && typeof row[19] === 'object') {
                row[19] = this.formatDateDDMMYYYY(row[19]);
            }

            finalData.push(row);
        });

        // Create worksheet
        const worksheet = XLSX.utils.aoa_to_sheet(finalData);

        // Apply formatting and styling
        this.applyStep1Formatting(worksheet, finalData.length);

        XLSX.utils.book_append_sheet(newWorkbook, worksheet, outputSheetName);

        return newWorkbook;
    }

    enrichStep2Data(workbook, settings) {
        const outputSheetName = settings.outputSheetName;
        const worksheet = workbook.Sheets[outputSheetName];

        // Build lookup tables
        const lookupTables = this.buildLookupTables(settings);

        // Initialize lock system
        const locks = new Set();

        // Get worksheet data
        const data = XLSX.utils.sheet_to_json(worksheet, { header: 1, defval: '' });

        // Process each data row (skip header and month spacer rows)
        for (let i = 1; i < data.length; i++) {
            const row = data[i];

            // Skip empty rows and month spacer rows
            if (!row[1] && !row[13] && !row[14]) continue; // No deal, product, or hedge

            const deal = this.normalize(row[1]); // VSA deal (B)
            const product = this.normalize(row[13]); // Product (N)
            const hedge = this.normalize(row[14]); // Hedge (O)

            // Rule 1: MIDLANDS product
            if (product === 'MIDLANDS') {
                for (let col = 4; col <= 10; col++) { // E-K
                    row[col] = 0;
                    locks.add(`${i},${col}`);
                }
            }

            // Rule 2: WHB+CIF deals
            if (lookupTables.whbCifDeals.has(deal)) {
                if (!locks.has(`${i},8`)) { // I not locked
                    row[8] = 0;
                    locks.add(`${i},8`);
                }
                if (!locks.has(`${i},9`)) { // J not locked
                    row[9] = 0;
                    locks.add(`${i},9`);
                }
            }

            // Rule 3: LC Costs (E) - Don't override zeros
            if (!locks.has(`${i},4`) && row[4] !== 0) {
                const botCost = lookupTables.costsMap.get(`${deal},BOT`) || 0;
                const blcCost = lookupTables.costsMap.get(`${deal},BLC`) || 0;
                const totalCost = botCost + blcCost;
                if (totalCost !== 0) {
                    row[4] = totalCost;
                }
            }

            // Rule 4: CIN insurance (I) - Don't override zeros
            if (!locks.has(`${i},8`) && row[8] !== 0) {
                const cinCost = lookupTables.costsMap.get(`${deal},CIN`) || 0;
                if (cinCost !== 0) {
                    row[8] = cinCost;
                }
            }

            // Rule 5: CLI insurance (J) - Don't override zeros
            if (!locks.has(`${i},9`) && row[9] !== 0) {
                const cliCost = lookupTables.costsMap.get(`${deal},CLI`) || 0;
                if (cliCost !== 0) {
                    row[9] = cliCost;
                }
            }

            // Rule 6: TOTAL calculation (L) - Calculate as numeric sum
            let total = 0;
            for (let col = 4; col <= 10; col++) { // E-K columns (indices 4-10)
                const val = row[col];
                if (val && typeof val === 'number') {
                    total += val;
                }
            }
            row[11] = total;

            // Rule 7: Hedge to VSA comments (U)
            if (hedge && lookupTables.hedgeToBR.has(hedge)) {
                row[20] = lookupTables.hedgeToBR.get(hedge);
            }

            // Rule 8: Hedge to Additional information (V)
            if (hedge && lookupTables.hedgeToCN.has(hedge)) {
                row[21] = lookupTables.hedgeToCN.get(hedge);
            }
        }

        // Update worksheet with modified data
        const newWorksheet = XLSX.utils.aoa_to_sheet(data);
        this.applyStep1Formatting(newWorksheet, data.length);

        workbook.Sheets[outputSheetName] = newWorksheet;

        // Calculate summary metrics
        this.calculateSummaryMetrics(data, lookupTables, locks);
    }

    buildLookupTables(settings) {
        const rawSheet1Name = settings.rawSheet1Name || this.workbook.SheetNames[0];
        const rawSheet2Name = settings.rawSheet2Name || this.workbook.SheetNames[1];
        const rawSheet3Name = settings.rawSheet3Name || this.workbook.SheetNames[2];
        const dealColumn = settings.dealColumnName || 'N';

        // WHB+CIF deals from Sheet 1
        const whbCifDeals = new Set();
        const sheet1Data = XLSX.utils.sheet_to_json(this.workbook.Sheets[rawSheet1Name], { header: 1 });
        sheet1Data.slice(1).forEach(row => {
            if (row && row[27] === 'CIF' && row[77] === 'WHB') { // AB=CIF, BZ=WHB
                const deal = this.normalize(row[1]); // Use column B (index 1) for deal number
                if (deal) whbCifDeals.add(deal);
            }
        });

        // Costs map from Sheet 2
        const costsMap = new Map();
        const sheet2Data = XLSX.utils.sheet_to_json(this.workbook.Sheets[rawSheet2Name], { header: 1 });

        sheet2Data.slice(1).forEach(row => {
            if (row) {
                const deal = this.normalize(row[13]); // N
                const type = this.normalize(row[42]); // AQ
                const amount = parseFloat(row[47]) || 0; // AV

                if (deal && type) {
                    const key = `${deal},${type}`;
                    costsMap.set(key, (costsMap.get(key) || 0) + amount);
                }
            }
        });

        // Hedge maps from Sheet 3
        const hedgeToBR = new Map();
        const hedgeToCN = new Map();
        const sheet3Data = XLSX.utils.sheet_to_json(this.workbook.Sheets[rawSheet3Name], { header: 1 });

        sheet3Data.slice(1).forEach(row => {
            if (row) {
                const hedge = this.normalize(row[12]); // M (column 13, index 12)
                const brValue = row[69]; // BR (column 70, index 69)
                const cnValue = row[91]; // CN (column 92, index 91)

                if (hedge) {
                    if (brValue && !hedgeToBR.has(hedge)) {
                        hedgeToBR.set(hedge, brValue);
                    }
                    if (cnValue && !hedgeToCN.has(hedge)) {
                        hedgeToCN.set(hedge, cnValue);
                    }
                }
            }
        });

        return { whbCifDeals, costsMap, hedgeToBR, hedgeToCN };
    }

    applyStep1Formatting(worksheet, numRows) {
        // Set column widths exactly as Python script
        const colWidths = [
            14, 16, 18, 10,      // A, B, C, D
            12, 14, 20, 16, 14, 14, 18,  // E, F, G, H, I, J, K (cost columns)
            14, 26,              // L, M
            16, 12, 12, 12, 22, 10, 12, 26, 26  // N, O, P, Q, R, S, T, U, V
        ];

        worksheet['!cols'] = colWidths.map(width => ({ width }));

        // Freeze panes at A2
        worksheet['!freeze'] = { xSplit: 1, ySplit: 1 };

        // Set row heights
        worksheet['!rows'] = [];
        worksheet['!rows'][0] = { hpt: 45 }; // Header row height = 45 (triple height)

        // Note: Advanced Excel styling (double borders, thick borders, fills)
        // would require xlsx-style library or server-side processing
        // The web version focuses on data transformation and basic formatting
    }

    calculateSummaryMetrics(data, lookupTables, locks) {
        let rowsProcessed = 0;
        let midlandsRows = 0;
        let whbCifDeals = 0;
        let costUpdates = { E: 0, I: 0, J: 0 };
        let hedgeMatches = { U: 0, V: 0 };

        const processedDeals = new Set();

        for (let i = 1; i < data.length; i++) {
            const row = data[i];

            // Skip empty rows and month spacer rows
            if (!row[1] && !row[13] && !row[14]) continue;

            rowsProcessed++;

            const deal = this.normalize(row[1]);
            const product = this.normalize(row[13]);

            // Count MIDLANDS rows
            if (product === 'MIDLANDS') {
                midlandsRows++;
            }

            // Count WHB+CIF deals
            if (lookupTables.whbCifDeals.has(deal)) {
                processedDeals.add(deal);
            }

            // Count cost updates
            if (row[4] && !locks.has(`${i},4`)) costUpdates.E++;
            if (row[8] && !locks.has(`${i},8`)) costUpdates.I++;
            if (row[9] && !locks.has(`${i},9`)) costUpdates.J++;

            // Count hedge matches
            if (row[20]) hedgeMatches.U++;
            if (row[21]) hedgeMatches.V++;
        }

        whbCifDeals = processedDeals.size;

        this.summaryMetrics = {
            rowsProcessed,
            midlandsRows,
            whbCifDeals,
            costUpdates: costUpdates.E + costUpdates.I + costUpdates.J,
            hedgeMatches: hedgeMatches.U + hedgeMatches.V
        };
    }

    normalize(value) {
        return value ? String(value).trim().toUpperCase() : '';
    }

    parseDate(value) {
        if (!value) return null;

        // Try parsing as Excel date number
        if (typeof value === 'number') {
            return new Date((value - 25569) * 86400 * 1000);
        }

        // Try parsing as date string
        const date = new Date(value);
        return isNaN(date.getTime()) ? null : date;
    }

    formatDateDDMMYYYY(date) {
        if (!date) return '';
        const day = String(date.getDate()).padStart(2, '0');
        const month = String(date.getMonth() + 1).padStart(2, '0');
        const year = date.getFullYear();
        return `${day}/${month}/${year}`;
    }

    getSettings() {
        return {
            outputSheetName: document.getElementById('outputSheetName').value,
            rawSheet1Name: document.getElementById('rawSheet1Name').value,
            rawSheet2Name: document.getElementById('rawSheet2Name').value,
            rawSheet3Name: document.getElementById('rawSheet3Name').value,
            dealColumnName: document.getElementById('dealColumnName').value
        };
    }

    showLoading() {
        document.getElementById('processButton').style.display = 'none';
        document.getElementById('loadingSpinner').style.display = 'flex';
    }

    hideLoading() {
        document.getElementById('processButton').style.display = 'inline-block';
        document.getElementById('loadingSpinner').style.display = 'none';
    }

    showError(message) {
        const errorSection = document.getElementById('errorSection');
        const errorMessage = document.getElementById('errorMessage');
        errorMessage.textContent = message;
        errorSection.style.display = 'block';
    }

    hideError() {
        document.getElementById('errorSection').style.display = 'none';
    }

    showResults() {
        const resultsSection = document.getElementById('resultsSection');

        // Show summary metrics
        this.displayMetrics();

        // Show preview
        this.displayPreview();

        resultsSection.style.display = 'block';
    }

    hideResults() {
        document.getElementById('resultsSection').style.display = 'none';
    }

    displayMetrics() {
        const metricsGrid = document.getElementById('metricsGrid');
        const metrics = this.summaryMetrics;

        metricsGrid.innerHTML = `
            <div class="metric-card">
                <h4>Rows Processed</h4>
                <div class="value">${metrics.rowsProcessed || 0}</div>
            </div>
            <div class="metric-card">
                <h4>MIDLANDS Rows</h4>
                <div class="value">${metrics.midlandsRows || 0}</div>
            </div>
            <div class="metric-card">
                <h4>WHB+CIF Deals</h4>
                <div class="value">${metrics.whbCifDeals || 0}</div>
            </div>
            <div class="metric-card">
                <h4>Cost Updates</h4>
                <div class="value">${metrics.costUpdates || 0}</div>
            </div>
            <div class="metric-card">
                <h4>Hedge Matches</h4>
                <div class="value">${metrics.hedgeMatches || 0}</div>
            </div>
        `;
    }

    displayPreview() {
        if (!this.processedData) return;

        const outputSheetName = this.getSettings().outputSheetName;
        const worksheet = this.processedData.Sheets[outputSheetName];
        const data = XLSX.utils.sheet_to_json(worksheet, { header: 1, defval: '' });

        const previewTable = document.getElementById('previewTable');

        // Show first 20 rows
        const previewData = data.slice(0, 21); // Header + 20 data rows

        let tableHTML = '';

        previewData.forEach((row, index) => {
            const isHeader = index === 0;
            const tag = isHeader ? 'th' : 'td';
            const rowClass = isHeader ? '' : ' class="data-row"';

            tableHTML += `<tr${rowClass}>`;
            for (let i = 0; i < 22; i++) { // A-V columns
                let cellValue = row[i] || '';

                // Handle formula cells
                if (cellValue && typeof cellValue === 'object' && cellValue.f) {
                    cellValue = `=${cellValue.f}`;
                }

                tableHTML += `<${tag}>${cellValue}</${tag}>`;
            }
            tableHTML += '</tr>';
        });

        previewTable.innerHTML = tableHTML;
    }

    downloadFile() {
        if (!this.processedData) return;

        // Convert workbook to array buffer
        const wbout = XLSX.write(this.processedData, { type: 'array', bookType: 'xlsx' });

        // Create blob and download
        const blob = new Blob([wbout], { type: 'application/octet-stream' });
        const url = URL.createObjectURL(blob);

        const a = document.createElement('a');
        a.href = url;
        a.download = 'formatted_output.xlsx';
        document.body.appendChild(a);
        a.click();
        document.body.removeChild(a);
        URL.revokeObjectURL(url);
    }
}

// Initialize the application when DOM is loaded
document.addEventListener('DOMContentLoaded', () => {
    new ExcelProcessor();
});