const XLSX = require('xlsx');

// Create test Excel file with sample data
function createTestExcel() {
    const workbook = XLSX.utils.book_new();

    // Sheet 1: WHB/CIF & base row data
    const sheet1Data = [
        // Headers (row 0)
        ['A', 'B', 'C', 'D', 'E', 'F', 'G', 'H', 'I', 'J', 'K', 'L', 'M', 'N', 'O', 'P', 'Q', 'R', 'S', 'T', 'U', 'V', 'W', 'X', 'Y', 'Z', 'AA', 'AB', 'AC', 'AD', 'AE', 'AF', 'AG', 'AH', 'AI', 'AJ', 'AK', 'AL', 'AM', 'AN', 'AO', 'AP', 'AQ', 'AR', 'AS', 'AT', 'AU', 'AV', 'AW', 'AX', 'AY', 'AZ', 'BA', 'BB', 'BC', 'BD', 'BE', 'BF', 'BG', 'BH', 'BI', 'BJ', 'BK', 'BL', 'BM', 'BN', 'BO', 'BP', 'BQ', 'BR', 'BS', 'BT', 'BU', 'BV', 'BW', 'BX', 'BY', 'BZ'],

        // Data rows - B←B, T←X mapping
        ['', 'DEAL001', '', '', '', '', '', '', '', '', '', 'HEDGE001', 'PROD001', '', '', '', '100', '', '', '', '', '', '', '2024-01-15', '', '', 'VESSEL1', 'FOB', '', 'SINGAPORE', '', '', '', '', '', '', '', 'OIL', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', 'WHB'],
        ['', 'DEAL002', '', '', '', '', '', '', '', '', '', 'HEDGE002', 'MIDLANDS', '', '', '', '200', '', '', '', '', '', '', '2024-01-20', '', '', 'VESSEL2', 'CIF', '', 'ROTTERDAM', '', '', '', '', '', '', '', 'GAS', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', 'WHB'],
        ['', 'DEAL003', '', '', '', '', '', '', '', '', '', 'HEDGE003', 'PROD003', '', '', '', '150', '', '', '', '', '', '', '2024-02-10', '', '', 'VESSEL3', 'FOB', '', 'HOUSTON', '', '', '', '', '', '', '', 'CRUDE', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '']
    ];

    // Sheet 2: Costs data
    const sheet2Data = [
        // Headers
        ['A', 'B', 'C', 'D', 'E', 'F', 'G', 'H', 'I', 'J', 'K', 'L', 'M', 'N', 'O', 'P', 'Q', 'R', 'S', 'T', 'U', 'V', 'W', 'X', 'Y', 'Z', 'AA', 'AB', 'AC', 'AD', 'AE', 'AF', 'AG', 'AH', 'AI', 'AJ', 'AK', 'AL', 'AM', 'AN', 'AO', 'AP', 'AQ', 'AR', 'AS', 'AT', 'AU', 'AV', 'AW'],

        // Cost data
        ['', '', '', '', '', '', '', '', '', '', '', '', '', 'DEAL001', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', 'BOT', '', '', '', '', '', 1500],
        ['', '', '', '', '', '', '', '', '', '', '', '', '', 'DEAL001', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', 'BLC', '', '', '', '', '', 800],
        ['', '', '', '', '', '', '', '', '', '', '', '', '', 'DEAL001', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', 'CIN', '', '', '', '', '', 300],
        ['', '', '', '', '', '', '', '', '', '', '', '', '', 'DEAL002', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', 'CLI', '', '', '', '', '', 400],
        ['', '', '', '', '', '', '', '', '', '', '', '', '', 'DEAL003', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', 'BOT', '', '', '', '', '', 2000],
        ['', '', '', '', '', '', '', '', '', '', '', '', '', 'DEAL003', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', 'CIN', '', '', '', '', '', 250]
    ];

    // Sheet 3: Hedge data
    const sheet3Data = [
        // Headers (simplified - only showing relevant columns)
        new Array(100).fill('').map((_, i) => String.fromCharCode(65 + Math.floor(i / 26)) + String.fromCharCode(65 + (i % 26))),

        // Hedge data
        (() => {
            const row1 = new Array(100).fill('');
            row1[12] = 'HEDGE001'; // M column
            row1[69] = 'VSA Comment 1'; // BR column
            row1[91] = 'Additional Info 1'; // CN column
            return row1;
        })(),

        (() => {
            const row2 = new Array(100).fill('');
            row2[12] = 'HEDGE002'; // M column
            row2[69] = 'VSA Comment 2'; // BR column
            row2[91] = 'Additional Info 2'; // CN column
            return row2;
        })(),

        (() => {
            const row3 = new Array(100).fill('');
            row3[12] = 'HEDGE003'; // M column
            row3[69] = 'VSA Comment 3'; // BR column
            row3[91] = 'Additional Info 3'; // CN column
            return row3;
        })()
    ];

    // Create worksheets
    const worksheet1 = XLSX.utils.aoa_to_sheet(sheet1Data);
    const worksheet2 = XLSX.utils.aoa_to_sheet(sheet2Data);
    const worksheet3 = XLSX.utils.aoa_to_sheet(sheet3Data);

    XLSX.utils.book_append_sheet(workbook, worksheet1, "Sheet1");
    XLSX.utils.book_append_sheet(workbook, worksheet2, "Sheet2");
    XLSX.utils.book_append_sheet(workbook, worksheet3, "Sheet3");

    // Write to file
    XLSX.writeFile(workbook, 'test_data.xlsx');
    console.log('Test Excel file created: test_data.xlsx');
}

createTestExcel();