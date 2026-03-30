const express = require('express');
const http    = require('http');
const { Server } = require('socket.io');
const fs      = require('fs');
const path    = require('path');
const ExcelJS = require('exceljs');
const multer  = require('multer');

const app    = express();
const server = http.createServer(app);
const io     = new Server(server);
const PORT   = 3005;

// ─────────────────────────────────────────────
//  BOOTSTRAP — ensure all required files exist
// ─────────────────────────────────────────────
const configPath    = path.join(__dirname, 'config.json');
const usersPath     = path.join(__dirname, 'users.json');
const historyPath   = path.join(__dirname, 'history.json');
const analyticsPath = path.join(__dirname, 'analytics.json');

const DEFAULT_CONFIG = {
    tempZone: path.join(__dirname, 'temp'),
    usBase:   path.join(__dirname, 'output', 'US'),
    ukBase:   path.join(__dirname, 'output', 'UK'),
};

const DEFAULT_USERS = [
    { username: 'Admin', password: 'admin', archivePath: path.join(__dirname, 'archive') }
];

if (!fs.existsSync(configPath)) {
    fs.writeFileSync(configPath, JSON.stringify(DEFAULT_CONFIG, null, 2));
    console.log('> Created default config.json');
}
if (!fs.existsSync(usersPath)) {
    fs.writeFileSync(usersPath, JSON.stringify(DEFAULT_USERS, null, 2));
    console.log('> Created default users.json');
}
if (!fs.existsSync(historyPath)) {
    fs.writeFileSync(historyPath, JSON.stringify({ files: [], logs: [] }, null, 2));
}
if (!fs.existsSync(analyticsPath)) {
    fs.writeFileSync(analyticsPath, JSON.stringify({}, null, 2));
    console.log('> Created analytics.json');
}

// Load config AFTER ensuring it exists
let config = JSON.parse(fs.readFileSync(configPath));

// Ensure all critical directories exist
[config.tempZone, config.usBase, config.ukBase].forEach(dir => {
    if (dir && !fs.existsSync(dir)) fs.mkdirSync(dir, { recursive: true });
});

const monthNames = ["Jan","Feb","Mar","Apr","May","Jun","Jul","Aug","Sep","Oct","Nov","Dec"];
const upload     = multer({ dest: config.tempZone });

// ─────────────────────────────────────────────
//  UNDO REGISTRY (5-Minute Window)
// ─────────────────────────────────────────────
const undoRegistry = new Map();

setInterval(() => {
    const now = Date.now();
    for (const [txnId, data] of undoRegistry.entries()) {
        if (now - data.timestamp > 5 * 60 * 1000) undoRegistry.delete(txnId);
    }
}, 60 * 1000);

// ─────────────────────────────────────────────
//  HELPERS
// ─────────────────────────────────────────────
function reloadConfig() {
    config = JSON.parse(fs.readFileSync(configPath));
}

function saveHistory(newFile, newLog) {
    const history = JSON.parse(fs.readFileSync(historyPath));
    if (newFile) history.files.unshift(newFile);
    if (newLog)  history.logs.unshift(newLog);
    if (history.files.length > 500) history.files.pop();
    if (history.logs.length  > 500) history.logs.pop();
    fs.writeFileSync(historyPath, JSON.stringify(history, null, 2));
}

// ── Analytics ledger helpers ──────────────────
function loadAnalytics() {
    return JSON.parse(fs.readFileSync(analyticsPath));
}

function saveAnalytics(data) {
    fs.writeFileSync(analyticsPath, JSON.stringify(data, null, 2));
}

/**
 * Append one upload record to analytics.json.
 *
 * Structure:
 *   analytics[dateKey] = {
 *     summary: {
 *       totalFiles: N,
 *       totalLeads: N,     // shown + hidden combined
 *       totalShown: N,
 *       byAgent: {
 *         [username]: { files: N, leads: N, shown: N }
 *       }
 *     },
 *     records: [
 *       { transactionId, agent, filename, mode, shown, hidden, total, time }
 *     ]
 *   }
 */
function analyticsAddRecord(dateKey, username, record) {
    const data = loadAnalytics();

    if (!data[dateKey]) {
        data[dateKey] = {
            summary: { totalFiles: 0, totalLeads: 0, totalShown: 0, byAgent: {} },
            records: []
        };
    }

    const day = data[dateKey];

    // Global summary
    day.summary.totalFiles++;
    day.summary.totalLeads += record.total;
    day.summary.totalShown += record.shown;

    // Per-agent summary
    if (!day.summary.byAgent[username]) {
        day.summary.byAgent[username] = { files: 0, leads: 0, shown: 0 };
    }
    day.summary.byAgent[username].files++;
    day.summary.byAgent[username].leads += record.total;
    day.summary.byAgent[username].shown += record.shown;

    // Individual record
    day.records.push(record);

    saveAnalytics(data);
}

/**
 * Remove one record from analytics.json by transactionId and deduct its tallies.
 */
function analyticsRemoveRecord(dateKey, transactionId) {
    const data = loadAnalytics();
    if (!data[dateKey]) return;

    const day    = data[dateKey];
    const recIdx = day.records.findIndex(r => r.transactionId === transactionId);
    if (recIdx === -1) return;

    const rec = day.records[recIdx];

    // Deduct global summary
    day.summary.totalFiles = Math.max(0, day.summary.totalFiles - 1);
    day.summary.totalLeads = Math.max(0, day.summary.totalLeads - rec.total);
    day.summary.totalShown = Math.max(0, day.summary.totalShown - rec.shown);

    // Deduct per-agent summary
    const agentSummary = day.summary.byAgent[rec.agent];
    if (agentSummary) {
        agentSummary.files = Math.max(0, agentSummary.files - 1);
        agentSummary.leads = Math.max(0, agentSummary.leads - rec.total);
        agentSummary.shown = Math.max(0, agentSummary.shown - rec.shown);
        if (agentSummary.files === 0) delete day.summary.byAgent[rec.agent];
    }

    // Remove record
    day.records.splice(recIdx, 1);

    // Remove day key if empty
    if (day.records.length === 0) delete data[dateKey];

    saveAnalytics(data);
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

// ─────────────────────────────────────────────
//  DAILY RESET
//  No longer touches any Excel files.
//  Clears history.json and signals all frontends.
// ─────────────────────────────────────────────
async function performDailyReset() {
    console.log(`\n> [RESET] Starting daily reset at ${new Date().toISOString()}`);

    fs.writeFileSync(historyPath, JSON.stringify({ files: [], logs: [] }, null, 2));
    console.log('> [RESET] history.json cleared.');

    io.emit('daily_reset');
    console.log('> [RESET] daily_reset event broadcast.');
    console.log('> [RESET] Complete.\n');
}

// ─────────────────────────────────────────────
//  MIDNIGHT CRON
// ─────────────────────────────────────────────
function scheduleMidnightReset() {
    function msUntilMidnight() {
        const now  = new Date();
        const next = new Date(now);
        next.setHours(24, 0, 0, 0);
        return next - now;
    }

    function arm() {
        const delay = msUntilMidnight();
        console.log(`> [CRON] Next reset in ${Math.round(delay / 1000 / 60)} minutes.`);
        setTimeout(async () => {
            await performDailyReset();
            arm();
        }, delay);
    }

    arm();
}

scheduleMidnightReset();

// ─────────────────────────────────────────────
//  EXCEL PROCESSING
//  Pure extraction — returns { finalFileName, shown, hidden }.
//  No longer writes any report file.
// ─────────────────────────────────────────────
async function processExcelFile(filePath, originalName, mode) {
    const workbook = new ExcelJS.Workbook();
    await workbook.xlsx.readFile(filePath);

    let finalVisibleRed = 0;
    let finalHidden     = 0;
    const isSomeMode    = mode === 'some';

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
                finalHidden     = stats[latestSomeKey].hidden;
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

            sheet.eachRow((row) => {
                const isHidden  = row.hidden;
                let hasNewInRow = false;
                let rowNewKey   = null;
                let rowNewNum   = -1;
                let hasRedFont  = false;
                const col1      = row.getCell(1).value;
                const hasData   = col1 !== null && col1 !== undefined && col1.toString().trim() !== '';

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
                finalHidden     = stats[latestNewKey].hidden;
            } else {
                finalVisibleRed = globalVisible;
                finalHidden     = globalHidden;
            }
        }
    }

    const ext           = path.extname(originalName);
    const baseName      = path.basename(originalName, ext);
    const finalFileName = isSomeMode ? `Some ${baseName}${ext}` : originalName;

    return { finalFileName, shown: finalVisibleRed, hidden: finalHidden };
}

// ─────────────────────────────────────────────
//  MIDDLEWARE
// ─────────────────────────────────────────────
app.use(express.static('public'));
app.use(express.json());

// ─────────────────────────────────────────────
//  AUTH
// ─────────────────────────────────────────────
app.post('/api/login', (req, res) => {
    const { username, password } = req.body;
    const users = JSON.parse(fs.readFileSync(usersPath));
    const user  = users.find(u => u.username === username && u.password === password);
    if (user) res.json({ success: true, username: user.username });
    else       res.json({ success: false, message: 'Invalid credentials' });
});

// ─────────────────────────────────────────────
//  USER DATA  (today's activity from history.json)
// ─────────────────────────────────────────────
app.get('/api/user/data', (req, res) => {
    const username = req.query.username;
    const history  = JSON.parse(fs.readFileSync(historyPath));
    res.json({
        success: true,
        files: history.files.filter(f => f.agent === username),
        logs:  history.logs.filter(l  => l.agent === username)
    });
});

// ─────────────────────────────────────────────
//  USER REPORT  (all-time from analytics.json)
// ─────────────────────────────────────────────
app.get('/api/user/report', (req, res) => {
    try {
        const username = req.query.username;
        const data     = loadAnalytics();
        const result   = [];

        // Newest date first
        for (const dateKey of Object.keys(data).sort((a, b) => b.localeCompare(a))) {
            const day = data[dateKey];
            // Newest record within the day first
            for (const rec of [...day.records].reverse()) {
                if (rec.agent === username) {
                    result.push({
                        date:     dateKey,
                        mode:     rec.mode,
                        filename: rec.filename,
                        shown:    rec.shown,
                        total:    rec.total
                    });
                }
            }
        }

        res.json({ success: true, data: result });
    } catch (e) {
        res.status(500).json({ success: false, error: e.message });
    }
});

// ─────────────────────────────────────────────
//  UPLOAD
// ─────────────────────────────────────────────
app.post('/api/upload', upload.single('file'), async (req, res) => {
    if (!req.file) return res.status(400).json({ error: 'No file uploaded.' });

    const { username, mode } = req.body;
    const originalName       = req.file.originalname;
    const tempPath           = req.file.path;

    try {
        reloadConfig();
        const users      = JSON.parse(fs.readFileSync(usersPath));
        const activeUser = users.find(u => u.username === username);
        if (!activeUser || !activeUser.archivePath) throw new Error(`Archive path missing for ${username}`);

        // Pure extraction — no report file written
        const stats      = await processExcelFile(tempPath, originalName, mode);
        const date       = new Date();
        const dateKey    = date.toISOString().split('T')[0];
        const folderName = `${date.getDate()}-${date.getMonth() + 1}`;
        const month      = monthNames[date.getMonth()];
        const year       = date.getFullYear().toString();
        const isNA       = /\bUSA\b|\bCANADA\b/i.test(originalName);
        const region     = isNA ? 'USA' : 'UK';
        const timeStr    = `${String(date.getHours()).padStart(2,'0')}:${String(date.getMinutes()).padStart(2,'0')}:${String(date.getSeconds()).padStart(2,'0')}`;
        const totalCount = stats.shown + stats.hidden;

        // ── Copy files ──
        const personalDir = path.join(activeUser.archivePath, year, month, folderName, region);
        if (!fs.existsSync(personalDir)) fs.mkdirSync(personalDir, { recursive: true });
        const finalPersonalPath = path.join(personalDir, stats.finalFileName);
        const savedPaths        = [finalPersonalPath];

        if (mode === 'some') {
            fs.copyFileSync(tempPath, finalPersonalPath);
        } else {
            const regionBase = isNA ? config.usBase : config.ukBase;
            const regionDir  = path.join(regionBase, year, month, folderName);
            if (!fs.existsSync(regionDir)) fs.mkdirSync(regionDir, { recursive: true });
            const regionDest = path.join(regionDir, stats.finalFileName);
            fs.copyFileSync(tempPath, finalPersonalPath);
            fs.copyFileSync(tempPath, regionDest);
            savedPaths.push(regionDest);
        }

        fs.unlinkSync(tempPath);

        // ── Transaction ID ──
        const transactionId = Date.now().toString(36) + Math.random().toString(36).substr(2, 9);

        // ── Write to analytics.json ──
        analyticsAddRecord(dateKey, username, {
            transactionId,
            agent:    username,
            filename: stats.finalFileName,
            mode:     mode.toUpperCase(),
            shown:    stats.shown,
            hidden:   stats.hidden,
            total:    totalCount,
            time:     timeStr
        });

        // ── Write to history.json ──
        saveHistory(
            { agent: username, name: stats.finalFileName, size: req.file.size, mtime: date.toISOString(), region, destPath: `${year}/${month}/${folderName}/${region}`, status: 'sorted', transactionId },
            { agent: username, ts: timeStr, msg: `Sorted [${mode.toUpperCase()}]: ${stats.finalFileName} (Shown: ${stats.shown}, Hidden: ${stats.hidden})`, type: 'success' }
        );

        // ── Undo registry ──
        undoRegistry.set(transactionId, {
            timestamp: Date.now(),
            username,
            filename:  stats.finalFileName,
            dateKey,
            savedPaths
        });

        // ── Broadcast to admin dashboard ──
        io.emit('new_upload', {
            agent:    username,
            total:    totalCount,
            shown:    stats.shown,
            hidden:   stats.hidden,
            filename: stats.finalFileName,
            date:     dateKey,
            mode:     mode.toUpperCase()
        });

        res.json({ success: true, stats, transactionId });

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

// ─────────────────────────────────────────────
//  UNDO
// ─────────────────────────────────────────────
app.post('/api/undo', async (req, res) => {
    const { transactionId, username } = req.body;
    const record = undoRegistry.get(transactionId);

    if (!record) return res.status(400).json({ success: false, error: 'Undo expired or invalid.' });
    if (record.username !== username) return res.status(403).json({ success: false, error: 'Unauthorized undo attempt.' });

    try {
        // 1. Delete physical files
        record.savedPaths.forEach(p => {
            if (fs.existsSync(p)) fs.unlinkSync(p);
        });

        // 2. Remove from analytics.json and deduct tallies
        analyticsRemoveRecord(record.dateKey, transactionId);

        // 3. Remove from history.json
        const history = JSON.parse(fs.readFileSync(historyPath));
        const fileIdx = history.files.findIndex(f => f.transactionId === transactionId);
        if (fileIdx > -1) history.files.splice(fileIdx, 1);

        const now     = new Date();
        const timeStr = `${String(now.getHours()).padStart(2,'0')}:${String(now.getMinutes()).padStart(2,'0')}:${String(now.getSeconds()).padStart(2,'0')}`;
        history.logs.unshift({ agent: username, ts: timeStr, msg: `[UNDO] Reverted: ${record.filename}`, type: 'warning' });
        fs.writeFileSync(historyPath, JSON.stringify(history, null, 2));

        // 4. Clear from registry
        undoRegistry.delete(transactionId);

        res.json({ success: true });
    } catch (err) {
        console.error('Undo Error:', err);
        res.status(500).json({ success: false, error: 'Failed to process undo.' });
    }
});

// ─────────────────────────────────────────────
//  ANALYTICS — full ledger
// ─────────────────────────────────────────────
app.get('/api/analytics', (req, res) => {
    try {
        res.json({ success: true, data: loadAnalytics() });
    } catch (e) {
        res.status(500).json({ success: false, error: e.message });
    }
});

// ─────────────────────────────────────────────
//  ADMIN — DELETE SINGLE RECORD
//  Body: { dateKey: "YYYY-MM-DD", transactionId: "..." }
// ─────────────────────────────────────────────
app.delete('/api/admin/record', (req, res) => {
    try {
        const { dateKey, transactionId } = req.body;
        if (!dateKey || !transactionId) {
            return res.status(400).json({ success: false, error: 'dateKey and transactionId are required.' });
        }
        const data = loadAnalytics();
        if (!data[dateKey]) {
            return res.status(404).json({ success: false, error: `No data for date ${dateKey}.` });
        }
        const recIdx = data[dateKey].records.findIndex(r => r.transactionId === transactionId);
        if (recIdx === -1) {
            return res.status(404).json({ success: false, error: 'Record not found.' });
        }
        // Reuse existing helper — deducts summaries and removes the record
        analyticsRemoveRecord(dateKey, transactionId);
        console.log(`> [ADMIN] Deleted record ${transactionId} from ${dateKey}`);
        res.json({ success: true });
    } catch (e) {
        console.error('Admin delete record error:', e);
        res.status(500).json({ success: false, error: e.message });
    }
});

// ─────────────────────────────────────────────
//  EXPORT — on-demand Excel download for a date
//  GET /api/export/:date   e.g. /api/export/2026-03-30
// ─────────────────────────────────────────────
app.get('/api/export/:date', async (req, res) => {
    try {
        const dateKey = req.params.date;
        const data    = loadAnalytics();
        const day     = data[dateKey];

        if (!day || !day.records || day.records.length === 0) {
            return res.status(404).json({ success: false, error: `No data found for ${dateKey}` });
        }

        const wb    = new ExcelJS.Workbook();
        const sheet = wb.addWorksheet('Report');

        // Column definitions — identical layout to the old Z-Report
        sheet.columns = [
            { header: 'Agent',                  key: 'agent',    width: 15 },
            { header: 'Total (Hidden + Shown)',  key: 'total',    width: 28 },
            { header: 'Shown Count',             key: 'shown',    width: 18 },
            { header: 'File Name',               key: 'filename', width: 50 },
            { header: 'Date',                    key: 'date',     width: 15 },
            { header: 'Mode',                    key: 'mode',     width: 15 },
            { header: 'Time',                    key: 'time',     width: 12 },
        ];

        // Style header row
        const headerRow = sheet.getRow(1);
        headerRow.eachCell((cell) => {
            cell.font      = { name: 'Arial', size: 11, bold: true, color: { argb: 'FFFFFFFF' } };
            cell.fill      = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF4F81BD' } };
            cell.alignment = { vertical: 'middle', horizontal: 'center' };
            cell.border    = { top: { style: 'thin' }, left: { style: 'thin' }, bottom: { style: 'thin' }, right: { style: 'thin' } };
        });

        // Data rows
        for (const rec of day.records) {
            const row = sheet.addRow({
                agent:    rec.agent,
                total:    rec.total,
                shown:    rec.shown,
                filename: rec.filename,
                date:     dateKey,
                mode:     rec.mode,
                time:     rec.time || '',
            });
            row.eachCell((cell, colNumber) => {
                cell.font      = { name: 'Arial', size: 11 };
                cell.alignment = { vertical: 'middle', horizontal: colNumber === 4 ? 'left' : 'center' };
                cell.border    = { top: { style: 'thin' }, left: { style: 'thin' }, bottom: { style: 'thin' }, right: { style: 'thin' } };
            });
        }

        // Summary row at the bottom
        const summaryRow = sheet.addRow({
            agent:    'TOTAL',
            total:    day.summary.totalLeads,
            shown:    day.summary.totalShown,
            filename: `${day.summary.totalFiles} file(s) processed`,
            date:     dateKey,
            mode:     '—',
            time:     '—',
        });
        summaryRow.eachCell((cell) => {
            cell.font      = { name: 'Arial', size: 11, bold: true };
            cell.fill      = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFD9E1F2' } };
            cell.alignment = { vertical: 'middle', horizontal: 'center' };
            cell.border    = { top: { style: 'thin' }, left: { style: 'thin' }, bottom: { style: 'thin' }, right: { style: 'thin' } };
        });

        // Stream directly to browser as a download
        const filename = `Report ${dateKey}.xlsx`;
        res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
        res.setHeader('Content-Disposition', `attachment; filename="${filename}"`);
        await wb.xlsx.write(res);
        res.end();

    } catch (e) {
        console.error('Export error:', e);
        res.status(500).json({ success: false, error: e.message });
    }
});

// ─────────────────────────────────────────────
//  ADMIN — REPORT DATA  (from analytics.json)
// ─────────────────────────────────────────────
app.get('/api/admin/report', (req, res) => {
    try {
        const data    = loadAnalytics();
        const records = [];

        for (const dateKey of Object.keys(data).sort((a, b) => b.localeCompare(a))) {
            for (const rec of [...data[dateKey].records].reverse()) {
                records.push({
                    agent:    rec.agent,
                    total:    rec.total,
                    shown:    rec.shown,
                    filename: rec.filename,
                    date:     dateKey,
                    mode:     rec.mode
                });
            }
        }

        res.json({ success: true, data: records });
    } catch (e) {
        res.status(500).json({ success: false, error: e.message });
    }
});

// ─────────────────────────────────────────────
//  ADMIN — CONFIG EDITOR
// ─────────────────────────────────────────────
app.get('/api/admin/config', (req, res) => {
    try {
        res.json({ success: true, config: JSON.parse(fs.readFileSync(configPath)) });
    } catch (e) {
        res.status(500).json({ success: false, error: e.message });
    }
});

app.post('/api/admin/config', (req, res) => {
    try {
        const newConfig = req.body;
        if (!newConfig || typeof newConfig !== 'object') throw new Error('Invalid config payload');
        ['usBase', 'ukBase', 'tempZone'].forEach(key => {
            if (newConfig[key]) fs.mkdirSync(newConfig[key], { recursive: true });
        });
        fs.writeFileSync(configPath, JSON.stringify(newConfig, null, 2));
        reloadConfig();
        res.json({ success: true });
    } catch (e) {
        res.status(500).json({ success: false, error: e.message });
    }
});

// ─────────────────────────────────────────────
//  ADMIN — USERS MANAGER
// ─────────────────────────────────────────────
app.get('/api/admin/users', (req, res) => {
    try {
        res.json({ success: true, users: JSON.parse(fs.readFileSync(usersPath)) });
    } catch (e) {
        res.status(500).json({ success: false, error: e.message });
    }
});

app.post('/api/admin/users', (req, res) => {
    try {
        const users = req.body;
        if (!Array.isArray(users)) throw new Error('Payload must be an array of users');
        if (!users.some(u => u.username === 'Admin')) throw new Error('Cannot remove the Admin user');
        fs.writeFileSync(usersPath, JSON.stringify(users, null, 2));
        res.json({ success: true });
    } catch (e) {
        res.status(500).json({ success: false, error: e.message });
    }
});

// ─────────────────────────────────────────────
//  ADMIN — MANUAL RESET
// ─────────────────────────────────────────────
app.post('/api/admin/reset', async (req, res) => {
    try {
        await performDailyReset();
        res.json({ success: true, message: 'History cleared and frontends notified.' });
    } catch (e) {
        res.status(500).json({ success: false, error: e.message });
    }
});

// ─────────────────────────────────────────────
//  START
// ─────────────────────────────────────────────
server.listen(PORT, () => {
    console.log(`\n> ZORG-NEXUS ACTIVE`);
    console.log(`> http://localhost:${PORT}\n`);
});
