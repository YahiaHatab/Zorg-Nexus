const express = require('express');
const fs = require('fs');
const path = require('path');
const ExcelJS = require('exceljs');
const multer = require('multer');

const app = express();
const PORT = 3005;

// --- CONFIGURATION ---
const configPath = path.join(__dirname, 'config.json');
let config = JSON.parse(fs.readFileSync(configPath));
const usersPath = path.join(__dirname, 'users.json');
const historyPath = path.join(__dirname, 'history.json');
const monthNames = ["Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"];

if (!fs.existsSync(config.tempZone)) fs.mkdirSync(config.tempZone, { recursive: true });
if (!fs.existsSync(historyPath)) fs.writeFileSync(historyPath, JSON.stringify({ files: [], logs: [] }));

const upload = multer({ dest: config.tempZone });

function saveHistory(newFile, newLog) {
    const history = JSON.parse(fs.readFileSync(historyPath));
    history.files.unshift(newFile);
    history.logs.unshift(newLog);
    if (history.files.length > 500) history.files.pop();
    if (history.logs.length > 500) history.logs.pop();
    fs.writeFileSync(historyPath, JSON.stringify(history, null, 2));
}

function checkIsRed(font) {
    if (!font || !font.color) return false;
    if (font.color.argb) {
        const argb = font.color.argb.toUpperCase();
        if (argb.includes('FF0000') || argb === 'FFFF0000' || argb === 'FFC00000') return true;
    }
    if (font.color.indexed === 10 || font.color.indexed === 2) return true;
    return false;
}

// --- EXCEL LOGIC ---
async function processExcelFile(filePath, originalName, username, mode) {
    const workbook = new ExcelJS.Workbook();
    await workbook.xlsx.readFile(filePath);
    
    let finalVisibleRed = 0;
    let finalHidden = 0;
    const isSomeMode = mode === 'some';

    if (isSomeMode) {
        const sheet2 = workbook.worksheets[1]; 
        if (sheet2) {
            let stats = {};
            let highestSomeNum = -1;
            let latestSomeKey = "Some";
            let foundAnySome = false;

            sheet2.eachRow((row) => {
                let hasSomeInRow = false;
                let rowSomeKey = null;
                let rowSomeNum = -1;

                row.eachCell((cell) => {
                    let cellText = '';
                    let isRed = false;
                    if (cell.value && cell.value.richText) {
                        cell.value.richText.forEach(rt => {
                            cellText += rt.text;
                            if (checkIsRed(rt.font)) isRed = true;
                        });
                    } else {
                        cellText = cell.value ? cell.value.toString() : '';
                        if (checkIsRed(cell.font)) isRed = true;
                    }
                    const match = cellText.trim().match(/^some(\d*)$/i);
                    if (match && isRed) {
                        hasSomeInRow = true;
                        foundAnySome = true;
                        const numStr = match[1];
                        const num = numStr === "" ? 0 : parseInt(numStr, 10);
                        if (num > rowSomeNum) {
                            rowSomeNum = num;
                            rowSomeKey = num === 0 ? "Some" : `Some${num}`;
                        }
                    }
                });

                if (hasSomeInRow) {
                    if (rowSomeNum > highestSomeNum) { highestSomeNum = rowSomeNum; latestSomeKey = rowSomeKey; }
                    if (!stats[rowSomeKey]) stats[rowSomeKey] = { shown: 0, hidden: 0 };
                    
                    let col2Value = row.getCell(2).value;
                    if (col2Value && typeof col2Value === 'object' && col2Value.result !== undefined) col2Value = col2Value.result;
                    const col2Str = col2Value !== null && col2Value !== undefined ? col2Value.toString().trim() : '';
                    if (/\d/.test(col2Str)) stats[rowSomeKey].shown++;
                    else stats[rowSomeKey].hidden++;
                }
            });

            if (foundAnySome) {
                finalVisibleRed = stats[latestSomeKey].shown;
                finalHidden = stats[latestSomeKey].hidden;
            }
        }
    } else {
        const sheet = workbook.getWorksheet('Sheet1');
        if (sheet) {
            let stats = {};
            let highestNewNum = -1;
            let latestNewKey = "New"; 
            let foundAnyNew = false;
            let globalHidden = 0;
            let globalVisible = 0; 

            sheet.eachRow((row) => {
                const isHidden = row.hidden;
                let hasNewInRow = false;
                let rowNewKey = null;
                let rowNewNum = -1;
                let hasRedFont = false;

                const col1 = row.getCell(1).value;
                const hasData = col1 !== null && col1 !== undefined && col1.toString().trim() !== '';

                row.eachCell((cell) => {
                    let cellText = '';
                    let isRed = false;
                    if (cell.value && cell.value.richText) {
                        cell.value.richText.forEach(rt => {
                            cellText += rt.text;
                            if (checkIsRed(rt.font)) isRed = true;
                        });
                    } else {
                        cellText = cell.value ? cell.value.toString() : '';
                        if (checkIsRed(cell.font)) isRed = true;
                    }
                    const match = cellText.trim().match(/^new(\d*)$/i);
                    if (match) {
                        hasNewInRow = true;
                        foundAnyNew = true;
                        const numStr = match[1];
                        const num = numStr === "" ? 0 : parseInt(numStr, 10);
                        if (num > rowNewNum) {
                            rowNewNum = num;
                            rowNewKey = num === 0 ? "New" : `New${num}`;
                        }
                    }
                    if (isRed) hasRedFont = true;
                });

                if (isHidden) globalHidden++;
                else if (hasData) globalVisible++;

                if (hasNewInRow) {
                    if (rowNewNum > highestNewNum) { highestNewNum = rowNewNum; latestNewKey = rowNewKey; }
                    if (!stats[rowNewKey]) stats[rowNewKey] = { visibleRed: 0, hidden: 0 };
                    if (isHidden) stats[rowNewKey].hidden++;
                    else if (hasRedFont) stats[rowNewKey].visibleRed++;
                }
            });

            if (foundAnyNew) {
                finalVisibleRed = stats[latestNewKey].visibleRed;
                finalHidden = stats[latestNewKey].hidden;
            } else {
                finalVisibleRed = globalVisible;
                finalHidden = globalHidden;
            }
        }
    }

    const ext = path.extname(originalName);
    const baseName = path.basename(originalName, ext);
    let finalFileName = originalName;
    if (isSomeMode) finalFileName = `Some ${baseName}${ext}`;

    const reportPath = path.join(__dirname, 'Z-Report.xlsx');
    const reportWb = new ExcelJS.Workbook();
    let reportSheet;

    const columnsDef = [
        { header: 'Agent', key: 'agent', width: 15 },
        { header: 'Total (Hidden + Shown)', key: 'total', width: 28 },
        { header: 'Shown Count', key: 'shown', width: 18 },
        { header: 'File Name', key: 'filename', width: 45 },
        { header: 'Date', key: 'date', width: 15 },
        { header: 'Mode', key: 'mode', width: 15 }
    ];

    if (fs.existsSync(reportPath)) {
        await reportWb.xlsx.readFile(reportPath);
        reportSheet = reportWb.getWorksheet(1);
        reportSheet.columns = columnsDef; 
    } else {
        reportSheet = reportWb.addWorksheet('Report');
        reportSheet.columns = columnsDef;
        const headerRow = reportSheet.getRow(1);
        headerRow.eachCell((cell) => {
            cell.font = { name: 'Arial', size: 11, bold: true, color: { argb: 'FFFFFFFF' } }; 
            cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF4F81BD' } }; 
            cell.alignment = { vertical: 'middle', horizontal: 'center' };
            cell.border = { top: { style: 'thin' }, left: { style: 'thin' }, bottom: { style: 'thin' }, right: { style: 'thin' } };
        });
    }

    const todayDate = new Date().toISOString().split('T')[0];
    const totalCount = finalVisibleRed + finalHidden;
    
    const newRow = reportSheet.addRow({
        agent: username,
        total: totalCount,
        shown: finalVisibleRed,
        filename: finalFileName, 
        date: todayDate,
        mode: mode.toUpperCase()
    });

    newRow.eachCell((cell, colNumber) => {
        cell.font = { name: 'Arial', size: 11, bold: true };
        cell.border = { top: { style: 'thin' }, left: { style: 'thin' }, bottom: { style: 'thin' }, right: { style: 'thin' } };
        cell.alignment = { vertical: 'middle', horizontal: colNumber === 4 ? 'left' : 'center' };
    });

    await reportWb.xlsx.writeFile(reportPath);

    return { finalFileName, shown: finalVisibleRed, hidden: finalHidden };
}

// --- API ENDPOINTS ---
app.use(express.static('public'));
app.use(express.json());

app.post('/api/login', (req, res) => {
    const { username, password } = req.body;
    const users = JSON.parse(fs.readFileSync(usersPath));
    const user = users.find(u => u.username === username && u.password === password);
    if (user) res.json({ success: true, username: user.username });
    else res.json({ success: false, message: 'Invalid credentials' });
});

// PRIVATE DATA ROUTE (Filters by User)
app.get('/api/user/data', (req, res) => {
    const username = req.query.username;
    const history = JSON.parse(fs.readFileSync(historyPath));
    
    const userFiles = history.files.filter(f => f.agent === username);
    const userLogs = history.logs.filter(l => l.agent === username);
    
    res.json({ success: true, files: userFiles, logs: userLogs });
});

// --- NEW: Personal Z-Report Endpoint ---
app.get('/api/user/report', async (req, res) => {
    const username = req.query.username;
    const reportPath = path.join(__dirname, 'Z-Report.xlsx');
    
    if (!fs.existsSync(reportPath)) {
        return res.json({ success: true, data: [] });
    }
    
    try {
        const workbook = new ExcelJS.Workbook();
        await workbook.xlsx.readFile(reportPath);
        const sheet = workbook.getWorksheet(1);
        
        let reportData = [];
        
        sheet.eachRow((row, rowNumber) => {
            if (rowNumber === 1) return; // Skip the header row
            const rowAgent = row.getCell(1).value ? row.getCell(1).value.toString() : '';
            
            // Only grab data if the agent name matches the logged-in user
            if (rowAgent === username) {
                reportData.push({
                    total: parseInt(row.getCell(2).value) || 0,
                    shown: parseInt(row.getCell(3).value) || 0,
                    filename: row.getCell(4).value ? row.getCell(4).value.toString() : '',
                    date: row.getCell(5).value ? row.getCell(5).value.toString() : '',
                    mode: row.getCell(6).value ? row.getCell(6).value.toString() : ''
                });
            }
        });
        
        // Reverse it so the newest files are at the top!
        res.json({ success: true, data: reportData.reverse() });
    } catch (error) {
        res.status(500).json({ success: false, error: error.message });
    }
});

app.post('/api/upload', upload.single('file'), async (req, res) => {
    if (!req.file) return res.status(400).json({ error: 'No file uploaded.' });
    
    const { username, mode } = req.body;
    const originalName = req.file.originalname;
    const tempPath = req.file.path;
    
    try {
        const users = JSON.parse(fs.readFileSync(usersPath));
        const activeUser = users.find(u => u.username === username);
        if (!activeUser || !activeUser.archivePath) throw new Error(`Archive path missing for ${username}`);
        const userArchiveBase = activeUser.archivePath;

        const stats = await processExcelFile(tempPath, originalName, username, mode);
        
        const date = new Date();
        const folderName = `${date.getDate()}-${date.getMonth() + 1}`;
        const month = monthNames[date.getMonth()];
        const year = date.getFullYear().toString();
        const isNorthAmerica = /\bUSA\b|\bCANADA\b/i.test(originalName);
        const region = isNorthAmerica ? 'USA' : 'UK';
        
        const personalDir = path.join(userArchiveBase, year, month, folderName, region);
        if (!fs.existsSync(personalDir)) fs.mkdirSync(personalDir, { recursive: true });
        const finalPersonalPath = path.join(personalDir, stats.finalFileName);

        if (mode === 'some') {
            fs.copyFileSync(tempPath, finalPersonalPath);
        } else {
            const regionBase = isNorthAmerica ? config.usBase : config.ukBase;
            const regionDir = path.join(regionBase, year, month, folderName);
            if (!fs.existsSync(regionDir)) fs.mkdirSync(regionDir, { recursive: true });
            const finalRegionPath = path.join(regionDir, stats.finalFileName);
            fs.copyFileSync(tempPath, finalPersonalPath);
            fs.copyFileSync(tempPath, finalRegionPath);
        }

        fs.unlinkSync(tempPath);

        // Update History for UI
        const timeStr = `${String(date.getHours()).padStart(2,'0')}:${String(date.getMinutes()).padStart(2,'0')}:${String(date.getSeconds()).padStart(2,'0')}`;
        const fileRecord = {
            agent: username, name: stats.finalFileName, size: req.file.size,
            mtime: date.toISOString(), region: region, destPath: `${year}/${month}/${folderName}/${region}`, status: 'sorted'
        };
        const logRecord = {
            agent: username, ts: timeStr, msg: `Sorted [${mode.toUpperCase()}]: ${stats.finalFileName} (Shown: ${stats.shown}, Hidden: ${stats.hidden})`, type: 'success'
        };
        saveHistory(fileRecord, logRecord);

        res.json({ success: true, stats });

    } catch (error) {
        if (fs.existsSync(tempPath)) fs.unlinkSync(tempPath);
        const timeStr = new Date().toLocaleTimeString();
        saveHistory(
            { agent: username, name: originalName, size: req.file.size, mtime: new Date().toISOString(), region: 'UNK', destPath: 'ERROR', status: 'error' },
            { agent: username, ts: timeStr, msg: `Error: ${error.message}`, type: 'error' }
        );
        res.status(500).json({ success: false, error: error.message });
    }
});

// --- ADMIN DASHBOARD DATA ROUTE ---
app.get('/api/admin/report', async (req, res) => {
    const reportPath = path.join(__dirname, 'Z-Report.xlsx');
    
    if (!fs.existsSync(reportPath)) {
        return res.json({ success: true, data: [] });
    }
    
    try {
        const workbook = new ExcelJS.Workbook();
        await workbook.xlsx.readFile(reportPath);
        const sheet = workbook.getWorksheet(1);
        
        let reportData = [];
        
        sheet.eachRow((row, rowNumber) => {
            if (rowNumber === 1) return; // Skip header
            reportData.push({
                agent: row.getCell(1).value ? row.getCell(1).value.toString() : 'Unknown',
                total: parseInt(row.getCell(2).value) || 0,
                shown: parseInt(row.getCell(3).value) || 0,
                filename: row.getCell(4).value ? row.getCell(4).value.toString() : '',
                date: row.getCell(5).value ? row.getCell(5).value.toString() : '',
                mode: row.getCell(6).value ? row.getCell(6).value.toString() : ''
            });
        });
        
        res.json({ success: true, data: reportData.reverse() });
    } catch (error) {
        res.status(500).json({ success: false, error: error.message });
    }
});

app.listen(PORT, () => {
    console.log(`\n> ZORG-NEXUS ACTIVE (Classic UI Mode)`);
    console.log(`> http://localhost:${PORT}\n`);
});