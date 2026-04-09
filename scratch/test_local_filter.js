const ExcelJS = require('exceljs');
const path = require('path');
const fs = require('fs');

// Mock checkIsRed
function checkIsRed(font) {
    if (!font || !font.color) return false;
    if (font.color.argb) {
        const argb = font.color.argb.toUpperCase();
        if (argb.includes('FF0000') || argb === 'FFFF0000' || argb === 'FFC00000') return true;
    }
    if (font.color.indexed === 10 || font.color.indexed === 2) return true;
    return false;
}

async function processExcelFile(filePath, originalName, mode) {
    const workbook = new ExcelJS.Workbook();
    await workbook.xlsx.readFile(filePath);

    let finalVisibleRed = 0;
    let finalHidden     = 0;
    const isSomeMode    = mode === 'some';
    const isNA          = /\bUSA\b|\bCANADA\b/i.test(originalName);

    if (isSomeMode) {
        const sheet2 = workbook.worksheets[1];
        if (sheet2) {
            let stats          = {};
            let highestSomeNum = -1;
            let latestSomeKey  = "Some";
            let foundAnySome   = false;

            sheet2.eachRow((row) => {
                let hasSomeInRow = false;
                let rowSomeKey   = null;
                let rowSomeNum   = -1;

                row.eachCell((cell) => {
                    let cellText = '';
                    let isRed    = false;
                    if (cell.value && cell.value.richText) {
                        cell.value.richText.forEach(rt => { cellText += rt.text; if (checkIsRed(rt.font)) isRed = true; });
                    } else {
                        cellText = cell.value ? cell.value.toString() : '';
                        if (checkIsRed(cell.font)) isRed = true;
                    }
                    const match = cellText.trim().match(/^some(\d*)$/i);
                    if (match && isRed) {
                        hasSomeInRow = true; foundAnySome = true;
                        const num = match[1] === "" ? 0 : parseInt(match[1], 10);
                        if (num > rowSomeNum) { rowSomeNum = num; rowSomeKey = num === 0 ? "Some" : `Some${num}`; }
                    }
                });

                if (hasSomeInRow) {
                    if (rowSomeNum > highestSomeNum) { highestSomeNum = rowSomeNum; latestSomeKey = rowSomeKey; }
                    if (!stats[rowSomeKey]) stats[rowSomeKey] = { shown: 0, hidden: 0, shownLocal: 0, hiddenLocal: 0 };
                    let col2Value = row.getCell(2).value;
                    if (col2Value && typeof col2Value === 'object' && col2Value.result !== undefined) col2Value = col2Value.result;
                    const col2Str = col2Value !== null && col2Value !== undefined ? col2Value.toString().trim() : '';
                    
                    const isLocal = col2Str.toLowerCase().includes('local');
                    if (/\d/.test(col2Str)) {
                        stats[rowSomeKey].shown++;
                        if (isLocal) stats[rowSomeKey].shownLocal++;
                    } else {
                        stats[rowSomeKey].hidden++;
                        if (isLocal) stats[rowSomeKey].hiddenLocal++;
                    }
                }
            });

            if (foundAnySome) {
                finalVisibleRed = stats[latestSomeKey].shown;
                finalHidden     = stats[latestSomeKey].hidden;

                if (!isNA && (finalVisibleRed + finalHidden) >= 200) {
                    finalVisibleRed = Math.max(0, finalVisibleRed - stats[latestSomeKey].shownLocal);
                    finalHidden     = Math.max(0, finalHidden - stats[latestSomeKey].hiddenLocal);
                }
            }
        }
    } else {
        const sheet = workbook.getWorksheet('Sheet1');
        if (sheet) {
            let stats         = {};
            let highestNewNum = -1;
            let latestNewKey  = "New";
            let foundAnyNew   = false;
            let globalHidden  = 0;
            let globalVisible = 0;
            let globalHiddenLocal  = 0;
            let globalVisibleLocal = 0;

            sheet.eachRow((row) => {
                const isHidden  = row.hidden;
                let hasNewInRow = false;
                let rowNewKey   = null;
                let rowNewNum   = -1;
                let hasRedFont  = false;
                const col1      = row.getCell(1).value;
                const hasData   = col1 !== null && col1 !== undefined && col1.toString().trim() !== '';

                // Count "local" entries in Column 2
                let col2Value = row.getCell(2).value;
                if (col2Value && typeof col2Value === 'object' && col2Value.result !== undefined) col2Value = col2Value.result;
                const col2Str = col2Value !== null && col2Value !== undefined ? col2Value.toString().trim() : '';
                const isLocal = col2Str.toLowerCase().includes('local');

                row.eachCell((cell) => {
                    let cellText = '';
                    let isRed    = false;
                    if (cell.value && cell.value.richText) {
                        cell.value.richText.forEach(rt => { cellText += rt.text; if (checkIsRed(rt.font)) isRed = true; });
                    } else {
                        cellText = cell.value ? cell.value.toString() : '';
                        if (checkIsRed(cell.font)) isRed = true;
                    }
                    const match = cellText.trim().match(/^new(\d*)$/i);
                    if (match) {
                        hasNewInRow = true; foundAnyNew = true;
                        const num = match[1] === "" ? 0 : parseInt(match[1], 10);
                        if (num > rowNewNum) { rowNewNum = num; rowNewKey = num === 0 ? "New" : `New${num}`; }
                    }
                    if (isRed) hasRedFont = true;
                });

                if (isHidden) {
                    globalHidden++;
                    if (isLocal) globalHiddenLocal++;
                } else if (hasData) {
                    globalVisible++;
                    if (isLocal) globalVisibleLocal++;
                }

                if (hasNewInRow) {
                    if (rowNewNum > highestNewNum) { highestNewNum = rowNewNum; latestNewKey = rowNewKey; }
                    if (!stats[rowNewKey]) stats[rowNewKey] = { visibleRed: 0, hidden: 0, visibleRedLocal: 0, hiddenLocal: 0 };
                    if (isHidden) {
                        stats[rowNewKey].hidden++;
                        if (isLocal) stats[rowNewKey].hiddenLocal++;
                    } else if (hasRedFont) {
                        stats[rowNewKey].visibleRed++;
                        if (isLocal) stats[rowNewKey].visibleRedLocal++;
                    }
                }
            });

            if (foundAnyNew) {
                finalVisibleRed = stats[latestNewKey].visibleRed;
                finalHidden     = stats[latestNewKey].hidden;

                if (!isNA && (finalVisibleRed + finalHidden) >= 200) {
                    finalVisibleRed = Math.max(0, finalVisibleRed - stats[latestNewKey].visibleRedLocal);
                    finalHidden     = Math.max(0, finalHidden - stats[latestNewKey].hiddenLocal);
                }
            } else {
                finalVisibleRed = globalVisible;
                finalHidden     = globalHidden;

                if (!isNA && (finalVisibleRed + finalHidden) >= 200) {
                    finalVisibleRed = Math.max(0, finalVisibleRed - globalVisibleLocal);
                    finalHidden     = Math.max(0, finalHidden - globalHiddenLocal);
                }
            }
        }
    }

    const ext           = path.extname(originalName);
    const baseName      = path.basename(originalName, ext);
    const finalFileName = isSomeMode ? `Some ${baseName}${ext}` : originalName;

    return { finalFileName, shown: finalVisibleRed, hidden: finalHidden };
}

async function runTest() {
    console.log('--- Test 1: User Case (232 total: 53 shown, 179 hidden local) ---');
    const workbook = new ExcelJS.Workbook();
    const sheet = workbook.addWorksheet('Sheet1');
    
    // Header
    sheet.addRow(['Col1', 'Col2', 'Col3']);
    
    // 53 visible red rows
    for (let i = 0; i < 53; i++) {
        const row = sheet.addRow(['Data', 'Standard', 'New']);
        row.getCell(3).font = { color: { argb: 'FFFF0000' } }; // Red
    }
    // 179 hidden local rows
    for (let i = 0; i < 179; i++) {
        const row = sheet.addRow(['Data', 'Local', 'New']);
        row.hidden = true;
    }
    
    const testFile = 'user_case_test.xlsx';
    await workbook.xlsx.writeFile(testFile);
    
    const result1 = await processExcelFile(testFile, 'International_Show_GERMANY.xlsx', 'new');
    console.log('Result:', result1);
    // Expected: total 232. shown 53 (0 local). hidden 179 (all local).
    // Final stats: shown 53, hidden 0.

    fs.unlinkSync(testFile);
}

runTest().catch(console.error);
