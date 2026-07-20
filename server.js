const express = require('express');
const http = require('http');
const { Server } = require('socket.io');
const fs = require('fs');
const path = require('path');
const ExcelJS = require('exceljs');
const multer = require('multer');

const app = express();
const server = http.createServer(app);
const io = new Server(server);
const PORT = 3005;
const activeFloor = {}; // Tracks agent status and timers
let isAutoDispatch = true; // Global toggle for Automated vs Manual show assignment

// ─────────────────────────────────────────────
//  BOOTSTRAP — ensure all required files exist
// ─────────────────────────────────────────────
const configPath = path.join(__dirname, 'config.json');
const usersPath = path.join(__dirname, 'users.json');
const historyPath = path.join(__dirname, 'history.json');
const analyticsPath = path.join(__dirname, 'analytics.json');

const DEFAULT_CONFIG = {
    tempZone: path.join(__dirname, 'temp'),
    usBase: path.join(__dirname, 'output', 'US'),
    ukBase: path.join(__dirname, 'output', 'UK'),
    cxlTags: ['Pricing', 'Duplicates', 'SameList'],
    useManualDestFolder: false,
    manualDestFolder: ""
};

const DEFAULT_USERS = [
    { username: 'Admin', password: 'admin', role: 'Admin', archivePath: path.join(__dirname, 'archive') }
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

const showsPath = path.join(__dirname, 'shows.json');
if (!fs.existsSync(showsPath)) {
    fs.writeFileSync(showsPath, JSON.stringify([], null, 2));
    console.log('> Created default shows.json');
}

// Load config AFTER ensuring it exists
let config = { ...DEFAULT_CONFIG, ...JSON.parse(fs.readFileSync(configPath)) };

// Ensure all critical directories exist
[config.tempZone, config.usBase, config.ukBase].forEach(dir => {
    if (dir && !fs.existsSync(dir)) fs.mkdirSync(dir, { recursive: true });
});

const monthNames = ["Jan", "Feb", "March", "April", "May", "June", "July", "August", "Sep", "Oct", "Nov", "Dec"];
const upload = multer({ dest: config.tempZone });

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
    config = { ...DEFAULT_CONFIG, ...JSON.parse(fs.readFileSync(configPath)) };
}

function saveHistory(newFile, newLog) {
    const history = JSON.parse(fs.readFileSync(historyPath));
    if (newFile) history.files.unshift(newFile);
    if (newLog) history.logs.unshift(newLog);
    if (history.files.length > 500) history.files.pop();
    if (history.logs.length > 500) history.logs.pop();
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
            summary: { totalFiles: 0, totalLeads: 0, totalShown: 0, reasonBreakdown: {}, byAgent: {} },
            records: []
        };
    }

    const day = data[dateKey];

    // Global summary
    day.summary.totalFiles++;
    day.summary.totalLeads += record.total;
    day.summary.totalShown += record.shown;

    // Accumulate reason breakdown into day summary
    if (!day.summary.reasonBreakdown) day.summary.reasonBreakdown = {};
    if (record.reasonBreakdown) {
        for (const [reason, count] of Object.entries(record.reasonBreakdown)) {
            day.summary.reasonBreakdown[reason] = (day.summary.reasonBreakdown[reason] || 0) + count;
        }
    }

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

    const day = data[dateKey];
    const recIdx = day.records.findIndex(r => r.transactionId === transactionId);
    if (recIdx === -1) return;

    const rec = day.records[recIdx];

    // Deduct global summary
    day.summary.totalFiles = Math.max(0, day.summary.totalFiles - 1);
    day.summary.totalLeads = Math.max(0, day.summary.totalLeads - rec.total);
    day.summary.totalShown = Math.max(0, day.summary.totalShown - rec.shown);

    // Deduct reason breakdown from day summary
    if (rec.reasonBreakdown && day.summary.reasonBreakdown) {
        for (const [reason, count] of Object.entries(rec.reasonBreakdown)) {
            if (day.summary.reasonBreakdown[reason] !== undefined) {
                day.summary.reasonBreakdown[reason] = Math.max(0, day.summary.reasonBreakdown[reason] - count);
                if (day.summary.reasonBreakdown[reason] === 0) delete day.summary.reasonBreakdown[reason];
            }
        }
    }

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
        // Match standard red (FFFF0000), dark red (FFC00000), or missing alpha red (FF0000)
        // Explicitly avoiding .includes('FF0000') which matches solid black (FF000000)
        if (argb === 'FFFF0000' || argb === 'FFC00000' || argb === 'FF0000') return true;
    }
    if (font.color.indexed === 10 || font.color.indexed === 2) return true;
    return false;
}

// ─────────────────────────────────────────────
//  DAILY RESET
//  No longer touches any Excel files.
//  Clears history.json and signals all frontends.
// ─────────────────────────────────────────────
// ─────────────────────────────────────────────
//  DAILY RESET
//  Clears history.json and wipes today's data from analytics.
// ─────────────────────────────────────────────
async function performDailyReset() {
    console.log(`\n> [RESET] Starting daily reset at ${new Date().toISOString()}`);

    fs.writeFileSync(historyPath, JSON.stringify({ files: [], logs: [] }, null, 2));
    console.log('> [RESET] history.json cleared.');

    // Wipe today's data from analytics.json so it doesn't reload on refresh
    try {
        const dateKey = new Date().toISOString().split('T')[0];
        const analytics = JSON.parse(fs.readFileSync(analyticsPath));
        if (analytics[dateKey]) {
            delete analytics[dateKey];
            fs.writeFileSync(analyticsPath, JSON.stringify(analytics, null, 2));
            console.log(`> [RESET] Removed today's (${dateKey}) data from analytics.json.`);
        }
    } catch (e) {
        console.error("> [RESET] Error modifying analytics.json:", e);
    }

    io.emit('daily_reset');
    console.log('> [RESET] daily_reset event broadcast.');
    console.log('> [RESET] Complete.\n');
}

// ─────────────────────────────────────────────
//  MIDNIGHT CRON
// ─────────────────────────────────────────────
function scheduleMidnightReset() {
    function msUntilMidnight() {
        const now = new Date();
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
async function processExcelFile(filePath, originalName) {
    const workbook = new ExcelJS.Workbook();
    await workbook.xlsx.readFile(filePath);

    let finalVisibleRed = 0;
    let finalHidden = 0;
    let newShown = 0;
    let newHidden = 0;
    let someShown = 0;
    let someHidden = 0;
    const isNA = /\bUSA\b|\bCANADA\b/i.test(originalName);

    // ── Hidden-row reason breakdown ──────────────
    // For each hidden row: read Col B for the reason keyword.
    // Exception: if Col B is not a known reason keyword AND Col C contains "Local", label it "Local".
    // Known keywords are matched case-insensitively as exact strings.
    const reasonBreakdown = {};
    const KNOWN_REASONS = ['Local', 'NF', 'Removed', 'Repeated', 'No Num'];

    function getCellStr(row, colIndex) {
        let val = row.getCell(colIndex).value;
        if (val && typeof val === 'object' && val.result !== undefined) val = val.result;
        if (val && val.richText) val = val.richText.map(r => r.text).join('');
        return val !== null && val !== undefined ? val.toString().trim() : '';
    }

    function bumpReason(reason) {
        reasonBreakdown[reason] = (reasonBreakdown[reason] || 0) + 1;
    }

    const sheet = workbook.getWorksheet('Sheet1') || workbook.worksheets[0];
    if (sheet) {
        let stats = {};

        let highestNewNum = -1;
        let latestNewKey = "New";
        let foundAnyNew = false;

        let highestSomeNum = -1;
        let latestSomeKey = "Some";
        let foundAnySome = false;

        let globalHidden = 0;
        let globalVisible = 0;
        let globalHiddenLocal = 0;
        let globalVisibleLocal = 0;

        sheet.eachRow((row) => {
            const isHidden = row.hidden;
            let hasRedFont = false;

            let hasNewInRow = false;
            let rowNewKey = null;
            let rowNewNum = -1;

            let hasSomeInRow = false;
            let rowSomeKey = null;
            let rowSomeNum = -1;

            const col1 = row.getCell(1).value;
            const hasData = col1 !== null && col1 !== undefined && col1.toString().trim() !== '';

            // Count "local" entries in Column 2
            let col2Value = row.getCell(2).value;
            if (col2Value && typeof col2Value === 'object' && col2Value.result !== undefined) col2Value = col2Value.result;
            const col2Str = col2Value !== null && col2Value !== undefined ? col2Value.toString().trim() : '';
            const isLocal = col2Str.toLowerCase().includes('local');

            row.eachCell((cell) => {
                let cellText = '';
                let isRed = false;
                if (cell.value && cell.value.richText) {
                    cell.value.richText.forEach(rt => { cellText += rt.text; if (checkIsRed(rt.font)) isRed = true; });
                } else {
                    cellText = cell.value ? cell.value.toString() : '';
                    if (checkIsRed(cell.font)) isRed = true;
                }

                const matchNew = cellText.trim().match(/^new(\d*)$/i);
                if (matchNew) {
                    hasNewInRow = true; foundAnyNew = true;
                    const num = matchNew[1] === "" ? 0 : parseInt(matchNew[1], 10);
                    if (num > rowNewNum) { rowNewNum = num; rowNewKey = num === 0 ? "New" : `New${num}`; }
                    if (isRed) hasRedFont = true;
                }

                const matchSome = cellText.trim().match(/^some(\d*)$/i);
                if (matchSome) {
                    hasSomeInRow = true; foundAnySome = true;
                    const num = matchSome[1] === "" ? 0 : parseInt(matchSome[1], 10);
                    if (num > rowSomeNum) { rowSomeNum = num; rowSomeKey = num === 0 ? "Some" : `Some${num}`; }
                    if (isRed) hasRedFont = true;
                }
            });

            if (isHidden) {
                globalHidden++;
                if (isLocal) globalHiddenLocal++;

                // ── Classify hidden row reason ──
                const colBStr = getCellStr(row, 2);
                const colCStr = getCellStr(row, 3);

                // Check Col B for a known reason (case-insensitive exact match)
                const colBReason = KNOWN_REASONS.find(r => r.toLowerCase() === colBStr.toLowerCase());

                if (colBReason) {
                    bumpReason(colBReason);
                } else if (colCStr.toLowerCase() === 'local') {
                    // Col B has a number/ID, Col C says "Local"
                    bumpReason('Local');
                } else {
                    // No recognisable reason
                    bumpReason('Other');
                }
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

            if (hasSomeInRow) {
                if (rowSomeNum > highestSomeNum) { highestSomeNum = rowSomeNum; latestSomeKey = rowSomeKey; }
                if (!stats[rowSomeKey]) stats[rowSomeKey] = { visibleRed: 0, hidden: 0, visibleRedLocal: 0, hiddenLocal: 0 };
                if (isHidden) {
                    stats[rowSomeKey].hidden++;
                    if (isLocal) stats[rowSomeKey].hiddenLocal++;
                } else if (hasRedFont) {
                    stats[rowSomeKey].visibleRed++;
                    if (isLocal) stats[rowSomeKey].visibleRedLocal++;
                }
            }
        });

        let newShownLocal = 0;
        let newHiddenLocal = 0;
        let someShownLocal = 0;
        let someHiddenLocal = 0;

        if (foundAnyNew || foundAnySome) {
            if (foundAnyNew) {
                newShown = stats[latestNewKey].visibleRed;
                newHidden = stats[latestNewKey].hidden;
                newShownLocal = stats[latestNewKey].visibleRedLocal;
                newHiddenLocal = stats[latestNewKey].hiddenLocal;
            }
            if (foundAnySome) {
                someShown = stats[latestSomeKey].visibleRed;
                someHidden = stats[latestSomeKey].hidden;
                someShownLocal = stats[latestSomeKey].visibleRedLocal;
                someHiddenLocal = stats[latestSomeKey].hiddenLocal;

                if (someShown === 0) {
                    someHidden = 0;
                    someHiddenLocal = 0;
                }
            }

            if (!isNA && (newShown + newHidden + someShown + someHidden) >= 200) {
                newShown = Math.max(0, newShown - newShownLocal);
                newHidden = Math.max(0, newHidden - newHiddenLocal);
                someShown = Math.max(0, someShown - someShownLocal);
                someHidden = Math.max(0, someHidden - someHiddenLocal);
            }

            if (someShown === 0) {
                someHidden = 0;
            }
        } else {
            newShown = globalVisible;
            newHidden = globalHidden;

            if (!isNA && (newShown + newHidden) >= 200) {
                newShown = Math.max(0, newShown - globalVisibleLocal);
                newHidden = Math.max(0, newHidden - globalHiddenLocal);
            }
        }

        finalVisibleRed = newShown + someShown;
        finalHidden = newHidden + someHidden;
    }

    return {
        finalFileName: originalName,
        shown: finalVisibleRed,
        hidden: finalHidden,
        newShown,
        newHidden,
        someShown,
        someHidden,
        reasonBreakdown
    };
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
    const user = users.find(u => u.username === username && u.password === password);
    // Return the role, fallback to Agent if they don't have one yet
    if (user) res.json({ success: true, username: user.username, role: user.role || 'Agent' });
    else res.json({ success: false, message: 'Invalid credentials' });
});

// ─────────────────────────────────────────────
//  USER DATA  (today's activity from history.json)
// ─────────────────────────────────────────────
app.get('/api/user/data', (req, res) => {
    const username = req.query.username;
    const history = JSON.parse(fs.readFileSync(historyPath));

    let analytics = {};
    try {
        analytics = JSON.parse(fs.readFileSync(analyticsPath));
    } catch (e) {
        console.error('Failed to read analytics inside user data:', e);
    }

    const enrichedFiles = history.files.filter(f => f.agent === username).map(f => {
        let note = '';
        if (f.transactionId) {
            for (const dateKey of Object.keys(analytics)) {
                const rec = analytics[dateKey].records && analytics[dateKey].records.find(r => r.transactionId === f.transactionId);
                if (rec) {
                    note = rec.note || '';
                    break;
                }
            }
        }
        return { ...f, note };
    });

    res.json({
        success: true,
        files: enrichedFiles,
        logs: history.logs.filter(l => l.agent === username)
    });
});

// ─────────────────────────────────────────────
//  USER REPORT  (all-time from analytics.json)
// ─────────────────────────────────────────────
app.get('/api/user/report', (req, res) => {
    try {
        const username = req.query.username;
        const data = loadAnalytics();
        const result = [];

        // Newest date first
        for (const dateKey of Object.keys(data).sort((a, b) => b.localeCompare(a))) {
            const day = data[dateKey];
            // Newest record within the day first
            for (const rec of [...day.records].reverse()) {
                if (rec.agent === username) {
                    result.push({
                        date: dateKey,
                        mode: rec.mode,
                        filename: rec.filename,
                        shown: rec.shown,
                        total: rec.total,
                        transactionId: rec.transactionId,
                        note: rec.note || ''
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
//  RECORD NOTE EDITOR
// ─────────────────────────────────────────────
app.post('/api/record/note', (req, res) => {
    try {
        const { date, transactionId, note, username } = req.body;
        if (!date || !transactionId) {
            return res.status(400).json({ success: false, error: 'Missing date or transactionId' });
        }

        const data = loadAnalytics();
        const day = data[date];
        if (!day || !day.records) {
            return res.status(404).json({ success: false, error: 'Day records not found' });
        }

        const rec = day.records.find(r => r.transactionId === transactionId);
        if (!rec) {
            return res.status(404).json({ success: false, error: 'Record not found' });
        }

        // Authorization check: only Admin, or the agent who processed this show, can edit the note
        if (rec.agent !== username) {
            // Check if user is Admin in users.json to bypass
            const users = JSON.parse(fs.readFileSync(usersPath));
            const user = users.find(u => u.username === username);
            if (!user || user.role !== 'Admin') {
                return res.status(403).json({ success: false, error: 'Unauthorized to edit this show note' });
            }
        }

        rec.note = note || '';
        saveAnalytics(data);

        // Broadcast to clients in real-time
        io.emit('record_note_updated', { date, transactionId, note: rec.note });

        res.json({ success: true, note: rec.note });
    } catch (e) {
        res.status(500).json({ success: false, error: e.message });
    }
});

// ─────────────────────────────────────────────
//  SHOW HOPPER
// ─────────────────────────────────────────────
app.get('/api/shows', (req, res) => {
    try {
        const shows = JSON.parse(fs.readFileSync(showsPath));
        res.json({ success: true, shows: shows });
    } catch (e) {
        res.status(500).json({ success: false, error: e.message });
    }
});

app.post('/api/shows/upload', upload.single('file'), async (req, res) => {
    if (!req.file) return res.status(400).json({ error: 'No file uploaded.' });
    try {
        const workbook = new ExcelJS.Workbook();
        await workbook.xlsx.readFile(req.file.path);
        const sheet = workbook.worksheets[0];
        const newShows = [];

        let lastShow = null;
        sheet.eachRow((row, rowNumber) => {
            if (rowNumber === 1) return; // skip header

            const getVal = (col) => {
                let cell = row.getCell(col).value;
                if (!cell) return '';
                if (typeof cell === 'object' && cell.text) return cell.text;
                if (typeof cell === 'object' && cell.hyperlink) return cell.hyperlink;
                return cell.toString();
            };

            const showName = getVal(1).trim();
            const link = getVal(2).trim();
            const status = getVal(3).trim();

            if (showName) {
                if (!status) {
                    lastShow = {
                        id: Date.now().toString(36) + Math.random().toString(36).substr(2, 9),
                        showName: showName,
                        link: link ? [link] : [],
                        agentName: getVal(4),
                        ld: getVal(5),
                        lists: getVal(6),
                        comment: getVal(7),
                        date: getVal(8),
                        status: 'Pending',
                        pinnedTo: null // for assignment logic
                    };
                    newShows.push(lastShow);
                } else {
                    lastShow = null;
                }
            } else if (!showName && link && lastShow) {
                // Continuation row for the last show
                lastShow.link.push(link);
                const ld = getVal(5), lists = getVal(6), comment = getVal(7);
                if (ld) lastShow.ld += (lastShow.ld ? '\n' : '') + ld;
                if (lists) lastShow.lists += (lastShow.lists ? '\n' : '') + lists;
                if (comment) lastShow.comment += (lastShow.comment ? '\n' : '') + comment;
            }
        });

        fs.unlinkSync(req.file.path);
        const currentShows = JSON.parse(fs.readFileSync(showsPath));
        const updatedShows = [...currentShows, ...newShows];
        fs.writeFileSync(showsPath, JSON.stringify(updatedShows, null, 2));

        res.json({ success: true, added: newShows.length, shows: updatedShows });
    } catch (error) {
        if (fs.existsSync(req.file.path)) fs.unlinkSync(req.file.path);
        res.status(500).json({ success: false, error: error.message });
    }
});

app.post('/api/shows/update', (req, res) => {
    // Save reorder and pinning
    fs.writeFileSync(showsPath, JSON.stringify(req.body.shows, null, 2));
    res.json({ success: true });
});

app.post('/api/shows/reorder', (req, res) => {
    const { reorderedIds } = req.body;
    let currentShows = JSON.parse(fs.readFileSync(showsPath));

    const pendingShows = currentShows.filter(s => s.status === 'Pending' || !s.status);
    const otherShows = currentShows.filter(s => s.status !== 'Pending' && s.status);

    pendingShows.sort((a, b) => {
        const idxA = reorderedIds.indexOf(a.id);
        const idxB = reorderedIds.indexOf(b.id);
        if (idxA === -1) return 1;
        if (idxB === -1) return -1;
        return idxA - idxB;
    });

    const newShows = [...pendingShows, ...otherShows];
    fs.writeFileSync(showsPath, JSON.stringify(newShows, null, 2));
    io.emit('hopper_updated');
    res.json({ success: true });
});

app.get('/api/shows/next', (req, res) => {
    const { username } = req.query;
    const currentShows = JSON.parse(fs.readFileSync(showsPath));

    // 1. Check if the agent ALREADY has an "In Progress" show
    const existingInProgressShow = currentShows.find(s => s.agentName === username && s.status === 'In Progress');
    if (existingInProgressShow) {
        return res.json({ success: true, show: existingInProgressShow });
    }

    // 2. Look for a Pending show specifically Pinned/Assigned to this user
    // We allow fetching of pinned shows even if Manual Mode is active.
    const pinnedShow = currentShows.find(s => s.agentName === username && s.status === 'Pending');

    if (pinnedShow) {
        pinnedShow.status = 'In Progress';
        fs.writeFileSync(showsPath, JSON.stringify(currentShows, null, 2));
        io.emit('hopper_updated');
        return res.json({ success: true, show: pinnedShow });
    }

    // 3. If no pinned show, check the Global Dispatch Mode
    if (!isAutoDispatch) {
        return res.json({ success: false, message: 'MANUAL_MODE', error: 'System is currently in Manual Dispatch mode. Please wait for an assignment.' });
    }

    // 4. Automated assignment: Look for the first UNASSIGNED Pending show
    const unassignedShow = currentShows.find(s => (!s.agentName || s.agentName.trim() === '') && (s.status === 'Pending' || !s.status));

    if (unassignedShow) {
        unassignedShow.status = 'In Progress';
        unassignedShow.agentName = username;
        fs.writeFileSync(showsPath, JSON.stringify(currentShows, null, 2));
        io.emit('hopper_updated');
        return res.json({ success: true, show: unassignedShow });
    }

    // 5. If nothing was found, let the agent know the queue is empty
    return res.json({ success: false, message: 'No shows available' });
});

app.get('/api/shows/active', (req, res) => {
    const { username } = req.query;
    const currentShows = JSON.parse(fs.readFileSync(showsPath));
    const activeShow = currentShows.find(s => s.agentName === username && s.status === 'In Progress');

    if (activeShow) {
        res.json({ success: true, show: activeShow });
    } else {
        res.json({ success: false });
    }
});

app.post('/api/shows/complete', (req, res) => {
    const { id } = req.body;
    const currentShows = JSON.parse(fs.readFileSync(showsPath));
    const show = currentShows.find(s => s.id === id);
    if (show) {
        show.status = 'Done';
        fs.writeFileSync(showsPath, JSON.stringify(currentShows, null, 2));
        io.emit('hopper_updated');
        res.json({ success: true });
    } else {
        res.status(404).json({ success: false, error: 'Show not found' });
    }
});

app.post('/api/shows/clear', (req, res) => {
    fs.writeFileSync(showsPath, JSON.stringify([], null, 2));
    io.emit('hopper_updated');
    res.json({ success: true });
});

app.get('/api/shows/export', async (req, res) => {
    try {
        const currentShows = JSON.parse(fs.readFileSync(showsPath));
        const wb = new ExcelJS.Workbook();
        const sheet = wb.addWorksheet('Shows Hopper');

        sheet.columns = [
            { header: 'Show Name', key: 'showName', width: 30 },
            { header: 'Link', key: 'link', width: 40 },
            { header: 'Status', key: 'status', width: 15 },
            { header: 'Agent Name', key: 'agentName', width: 20 },
            { header: 'L/D', key: 'ld', width: 15 },
            { header: 'Name of Lists', key: 'lists', width: 25 },
            { header: 'Comment', key: 'comment', width: 30 },
            { header: 'Date', key: 'date', width: 15 },
        ];

        sheet.getRow(1).font = { bold: true };

        currentShows.forEach(s => {
            const links = Array.isArray(s.link) ? s.link.join('\n') : (s.link || '');
            sheet.addRow({
                showName: s.showName,
                link: links,
                status: s.status || 'Pending',
                agentName: s.agentName || '',
                ld: s.ld || '',
                lists: s.lists || '',
                comment: s.comment || '',
                date: s.date || ''
            });
        });

        res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
        const today = new Date().toISOString().split('T')[0];
        res.setHeader('Content-Disposition', `attachment; filename="Shows_Log_${today}.xlsx"`);
        await wb.xlsx.write(res);
        res.end();
    } catch (e) {
        res.status(500).json({ success: false, error: e.message });
    }
});

app.post('/api/shows/cancel', (req, res) => {
    const { id, reason } = req.body;
    const currentShows = JSON.parse(fs.readFileSync(showsPath));
    const show = currentShows.find(s => s.id === id);

    if (show) {
        show.status = `CXL/${reason}`;

        // --- Automated File Scrubbing ---
        try {
            const config = JSON.parse(fs.readFileSync(configPath));
            const users = JSON.parse(fs.readFileSync(usersPath));
            const dirsToScrub = [config.usBase, config.ukBase];

            // Add the specific agent's archive path if assigned
            if (show.agentName) {
                const agent = users.find(u => u.username === show.agentName);
                if (agent && agent.archivePath) dirsToScrub.push(agent.archivePath);
            }

            // Scan and delete matching files
            dirsToScrub.forEach(dir => {
                if (dir && fs.existsSync(dir)) {
                    const files = fs.readdirSync(dir);
                    files.forEach(file => {
                        // Match files containing the show name
                        if (file.includes(show.showName)) {
                            try {
                                fs.unlinkSync(path.join(dir, file));
                            } catch (e) { console.error('Scrub failed for:', file); }
                        }
                    });
                }
            });
        } catch (e) {
            console.error("Error reading config for scrubbing:", e);
        }

        // --- Save CXL to Analytics Ledger ---
        try {
            const dateKey = new Date().toISOString().split('T')[0];
            const timeStr = new Date().toLocaleTimeString('en-US');
            const analytics = JSON.parse(fs.readFileSync(analyticsPath));

            if (!analytics[dateKey]) {
                analytics[dateKey] = { summary: { totalFiles: 0, totalLeads: 0, totalShown: 0, byAgent: {} }, records: [] };
            }

            analytics[dateKey].records.push({
                transactionId: Date.now().toString(36),
                time: timeStr,
                agent: show.agentName || 'Unassigned',
                mode: `CXL`,
                filename: show.showName,
                total: 0, shown: 0, hidden: 0,
                reason: reason
            });
            fs.writeFileSync(analyticsPath, JSON.stringify(analytics, null, 2));
        } catch (e) { console.error("Error logging CXL to analytics:", e); }

        fs.writeFileSync(showsPath, JSON.stringify(currentShows, null, 2));

        // Broadcast the cancellation to disconnect the agent
        io.emit('show_cancelled', {
            id: show.id,
            showName: show.showName,
            reason: reason,
            agentName: show.agentName
        });

        io.emit('hopper_updated');
        res.json({ success: true });
    } else {
        res.status(404).json({ success: false, error: 'Show not found' });
    }
});

// ─────────────────────────────────────────────
//  UPLOAD
// ─────────────────────────────────────────────
app.post('/api/upload', upload.single('file'), async (req, res) => {
    if (!req.file) return res.status(400).json({ error: 'No file uploaded.' });

    const username = req.body.username;
    const mode = 'standard';
    const originalName = Buffer.from(req.file.originalname, 'latin1').toString('utf8');
    const tempPath = req.file.path;

    try {
        reloadConfig();
        const users = JSON.parse(fs.readFileSync(usersPath));
        const activeUser = users.find(u => u.username === username);
        if (!activeUser || !activeUser.archivePath) throw new Error(`Archive path missing for ${username}`);

        // Pure extraction — no report file written
        const stats = await processExcelFile(tempPath, originalName);
        let date = new Date();
        if (req.body.uploadDate) {
            const [y, m, d] = req.body.uploadDate.split('-').map(Number);
            date = new Date(y, m - 1, d, 12, 0, 0);
        }
        const dateKey = date.toISOString().split('T')[0];
        const dayPadded = String(date.getDate()).padStart(2, '0');
        const monthNum = String(date.getMonth() + 1);
        const folderName = `${dayPadded}-${monthNum}`;
        const month = monthNames[date.getMonth()];
        const year = date.getFullYear().toString();
        const isNA = /\bUSA\b|\bCANADA\b/i.test(originalName);
        const region = isNA ? 'USA' : 'UK';
        const h12upload = date.getHours() % 12 || 12;
        const amPmUpload = date.getHours() < 12 ? 'AM' : 'PM';
        const timeStr = `${String(h12upload).padStart(2, '0')}:${String(date.getMinutes()).padStart(2, '0')}:${String(date.getSeconds()).padStart(2, '0')} ${amPmUpload}`;
        const totalCount = stats.shown + stats.hidden;

        // ── Copy files ──
        let savedPaths = [];
        if (config.useManualDestFolder === true || config.useManualDestFolder === 'true') {
            const baseDestDir = config.manualDestFolder;
            if (!baseDestDir) throw new Error("Manual destination folder path is not configured.");
            const destDir = path.join(baseDestDir, month, folderName);
            if (!fs.existsSync(destDir)) fs.mkdirSync(destDir, { recursive: true });
            const destPath = path.join(destDir, stats.finalFileName);
            fs.copyFileSync(tempPath, destPath);
            savedPaths.push(destPath);
        } else {
            const personalDir = path.join(activeUser.archivePath, year, month, folderName);
            if (!fs.existsSync(personalDir)) fs.mkdirSync(personalDir, { recursive: true });
            const finalPersonalPath = path.join(personalDir, stats.finalFileName);
            fs.copyFileSync(tempPath, finalPersonalPath);
            savedPaths.push(finalPersonalPath);

            if (activeUser.enableSharedOutput === true || activeUser.enableSharedOutput === 'true') {
                const outputBase = config.sharedBasePath || path.dirname(config.usBase);
                const regionDir = path.join(outputBase, year, month, folderName);
                if (!fs.existsSync(regionDir)) fs.mkdirSync(regionDir, { recursive: true });
                const regionDest = path.join(regionDir, stats.finalFileName);
                fs.copyFileSync(tempPath, regionDest);
                savedPaths.push(regionDest);
            }
        }

        fs.unlinkSync(tempPath);

        // ── Transaction ID ──
        const transactionId = Date.now().toString(36) + Math.random().toString(36).substr(2, 9);

        // ── Write to analytics.json ──
        analyticsAddRecord(dateKey, username, {
            transactionId,
            agent: username,
            filename: stats.finalFileName,
            mode: mode.toUpperCase(),
            shown: stats.shown,
            hidden: stats.hidden,
            total: totalCount,
            time: timeStr,
            newShown: stats.newShown || 0,
            newHidden: stats.newHidden || 0,
            someShown: stats.someShown || 0,
            someHidden: stats.someHidden || 0,
            reasonBreakdown: stats.reasonBreakdown || {}
        });

        // ── Write to history.json ──
        saveHistory(
            { agent: username, name: stats.finalFileName, size: req.file.size, mtime: date.toISOString(), uploadTime: Date.now(), region, destPath: `${year}/${month}/${folderName}`, status: 'sorted', transactionId },
            { agent: username, ts: timeStr, msg: `Sorted [${mode.toUpperCase()}]: ${stats.finalFileName} (Shown: ${stats.shown}, Hidden: ${stats.hidden})`, type: 'success' }
        );

        // ── Undo registry ──
        undoRegistry.set(transactionId, {
            timestamp: Date.now(),
            username,
            filename: stats.finalFileName,
            dateKey,
            savedPaths
        });

        // ── Broadcast to admin dashboard ──
        io.emit('new_upload', {
            agent: username,
            total: totalCount,
            shown: stats.shown,
            hidden: stats.hidden,
            filename: stats.finalFileName,
            date: dateKey,
            mode: mode.toUpperCase(),
            newShown: stats.newShown || 0,
            newHidden: stats.newHidden || 0,
            someShown: stats.someShown || 0,
            someHidden: stats.someHidden || 0,
            reasonBreakdown: stats.reasonBreakdown || {}
        });

        res.json({ success: true, stats, transactionId });

    } catch (error) {
        if (fs.existsSync(tempPath)) fs.unlinkSync(tempPath);
        const errDate = new Date();
        const h12err = errDate.getHours() % 12 || 12;
        const amPmErr = errDate.getHours() < 12 ? 'AM' : 'PM';
        const timeStr = `${String(h12err).padStart(2, '0')}:${String(errDate.getMinutes()).padStart(2, '0')}:${String(errDate.getSeconds()).padStart(2, '0')} ${amPmErr}`;
        saveHistory(
            { agent: username, name: originalName, size: req.file.size, mtime: new Date().toISOString(), region: 'UNK', destPath: 'ERROR', status: 'error' },
            { agent: username, ts: timeStr, msg: `Error: ${error.message}`, type: 'error' }
        );
        res.status(500).json({ success: false, error: error.message });
    }
});

// ─────────────────────────────────────────────
//  BULK UPLOAD
// ─────────────────────────────────────────────
app.post('/api/upload-bulk', upload.array('files', 20), async (req, res) => {
    if (!req.files || req.files.length === 0) return res.status(400).json({ error: 'No files uploaded.' });

    const username = req.body.username;
    const mode = 'standard';

    try {
        reloadConfig();
        const users = JSON.parse(fs.readFileSync(usersPath));
        const activeUser = users.find(u => u.username === username);
        if (!activeUser || !activeUser.archivePath) throw new Error(`Archive path missing for ${username}`);

        const results = [];

        let date = new Date();
        if (req.body.uploadDate) {
            const [y, m, d] = req.body.uploadDate.split('-').map(Number);
            date = new Date(y, m - 1, d, 12, 0, 0);
        }

        for (const file of req.files) {
            const originalName = Buffer.from(file.originalname, 'latin1').toString('utf8');
            const tempPath = file.path;

            try {
                // Pure extraction — no report file written
                const stats = await processExcelFile(tempPath, originalName);
                const dateKey = date.toISOString().split('T')[0];
                const dayPadded = String(date.getDate()).padStart(2, '0');
                const monthNum = String(date.getMonth() + 1);
                const folderName = `${dayPadded}-${monthNum}`;
                const month = monthNames[date.getMonth()];
                const year = date.getFullYear().toString();
                const isNA = /\bUSA\b|\bCANADA\b/i.test(originalName);
                const region = isNA ? 'USA' : 'UK';
                const h12upload = date.getHours() % 12 || 12;
                const amPmUpload = date.getHours() < 12 ? 'AM' : 'PM';
                const timeStr = `${String(h12upload).padStart(2, '0')}:${String(date.getMinutes()).padStart(2, '0')}:${String(date.getSeconds()).padStart(2, '0')} ${amPmUpload}`;
                const totalCount = stats.shown + stats.hidden;

                // ── Copy files ──
                let savedPaths = [];
                if (config.useManualDestFolder === true || config.useManualDestFolder === 'true') {
                    const baseDestDir = config.manualDestFolder;
                    if (!baseDestDir) throw new Error("Manual destination folder path is not configured.");
                    const destDir = path.join(baseDestDir, month, folderName);
                    if (!fs.existsSync(destDir)) fs.mkdirSync(destDir, { recursive: true });
                    const destPath = path.join(destDir, stats.finalFileName);
                    fs.copyFileSync(tempPath, destPath);
                    savedPaths.push(destPath);
                } else {
                    const personalDir = path.join(activeUser.archivePath, year, month, folderName);
                    if (!fs.existsSync(personalDir)) fs.mkdirSync(personalDir, { recursive: true });
                    const finalPersonalPath = path.join(personalDir, stats.finalFileName);
                    fs.copyFileSync(tempPath, finalPersonalPath);
                    savedPaths.push(finalPersonalPath);

                    if (activeUser.enableSharedOutput === true || activeUser.enableSharedOutput === 'true') {
                        const outputBase = config.sharedBasePath || path.dirname(config.usBase);
                        const regionDir = path.join(outputBase, year, month, folderName);
                        if (!fs.existsSync(regionDir)) fs.mkdirSync(regionDir, { recursive: true });
                        const regionDest = path.join(regionDir, stats.finalFileName);
                        fs.copyFileSync(tempPath, regionDest);
                        savedPaths.push(regionDest);
                    }
                }

                fs.unlinkSync(tempPath);

                // ── Transaction ID ──
                const transactionId = Date.now().toString(36) + Math.random().toString(36).substr(2, 9);

                // ── Write to analytics.json ──
                analyticsAddRecord(dateKey, username, {
                    transactionId,
                    agent: username,
                    filename: stats.finalFileName,
                    mode: mode.toUpperCase(),
                    shown: stats.shown,
                    hidden: stats.hidden,
                    total: totalCount,
                    time: timeStr,
                    newShown: stats.newShown || 0,
                    newHidden: stats.newHidden || 0,
                    someShown: stats.someShown || 0,
                    someHidden: stats.someHidden || 0,
                    reasonBreakdown: stats.reasonBreakdown || {}
                });

                // ── Write to history.json ──
                saveHistory(
                    { agent: username, name: stats.finalFileName, size: file.size, mtime: date.toISOString(), uploadTime: Date.now(), region, destPath: `${year}/${month}/${folderName}`, status: 'sorted', transactionId },
                    { agent: username, ts: timeStr, msg: `Sorted [${mode.toUpperCase()}]: ${stats.finalFileName} (Shown: ${stats.shown}, Hidden: ${stats.hidden})`, type: 'success' }
                );

                // ── Undo registry ──
                undoRegistry.set(transactionId, {
                    timestamp: Date.now(),
                    username,
                    filename: stats.finalFileName,
                    dateKey,
                    savedPaths
                });

                // ── Broadcast to admin dashboard ──
                io.emit('new_upload', {
                    agent: username,
                    total: totalCount,
                    shown: stats.shown,
                    hidden: stats.hidden,
                    filename: stats.finalFileName,
                    date: dateKey,
                    mode: mode.toUpperCase(),
                    newShown: stats.newShown || 0,
                    newHidden: stats.newHidden || 0,
                    someShown: stats.someShown || 0,
                    someHidden: stats.someHidden || 0,
                    reasonBreakdown: stats.reasonBreakdown || {}
                });

                results.push({ success: true, filename: stats.finalFileName, stats, transactionId });
            } catch (fileError) {
                if (fs.existsSync(tempPath)) fs.unlinkSync(tempPath);
                const errDate = new Date();
                const h12err = errDate.getHours() % 12 || 12;
                const amPmErr = errDate.getHours() < 12 ? 'AM' : 'PM';
                const timeStr = `${String(h12err).padStart(2, '0')}:${String(errDate.getMinutes()).padStart(2, '0')}:${String(errDate.getSeconds()).padStart(2, '0')} ${amPmErr}`;

                saveHistory(
                    { agent: username, name: originalName, size: file.size, mtime: new Date().toISOString(), region: 'UNK', destPath: 'ERROR', status: 'error' },
                    { agent: username, ts: timeStr, msg: `Error in bulk upload for ${originalName}: ${fileError.message}`, type: 'error' }
                );

                results.push({ success: false, filename: originalName, error: fileError.message });
            }
        }

        res.json({ success: true, results });
    } catch (globalError) {
        if (req.files) {
            req.files.forEach(file => {
                if (fs.existsSync(file.path)) fs.unlinkSync(file.path);
            });
        }
        res.status(500).json({ success: false, error: globalError.message });
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

        const now = new Date();
        const h12undo = now.getHours() % 12 || 12;
        const amPmUndo = now.getHours() < 12 ? 'AM' : 'PM';
        const timeStr = `${String(h12undo).padStart(2, '0')}:${String(now.getMinutes()).padStart(2, '0')}:${String(now.getSeconds()).padStart(2, '0')} ${amPmUndo}`;
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
//  LEADERBOARD — Today's top agents
// ─────────────────────────────────────────────
app.get('/api/leaderboard', (req, res) => {
    try {
        const data = loadAnalytics();
        const todayKey = new Date().toISOString().split('T')[0];
        const day = data[todayKey];
        if (!day || !day.summary || !day.summary.byAgent) {
            return res.json({ success: true, leaderboard: [] });
        }

        const leaderboard = Object.entries(day.summary.byAgent).map(([name, stats]) => ({
            name,
            shown: stats.shown,
            files: stats.files,
            leads: stats.leads
        })).sort((a, b) => b.shown - a.shown);

        res.json({ success: true, leaderboard });
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
        const rec = data[dateKey].records.find(r => r.transactionId === transactionId);
        if (!rec) {
            return res.status(404).json({ success: false, error: 'Record not found.' });
        }
        const agent = rec.agent;

        // Reuse existing helper — deducts summaries and removes the record from analytics
        analyticsRemoveRecord(dateKey, transactionId);

        // Remove from history.json so it is deleted from the agent's main page uploads
        const history = JSON.parse(fs.readFileSync(historyPath));
        const fileIdx = history.files.findIndex(f => f.transactionId === transactionId);
        if (fileIdx > -1) {
            history.files.splice(fileIdx, 1);
            fs.writeFileSync(historyPath, JSON.stringify(history, null, 2));
        }

        // Broadcast to both agent and admin frontends for real-time synchronization
        io.emit('record_deleted', { transactionId, dateKey, agent });

        console.log(`> [ADMIN] Deleted record ${transactionId} from ${dateKey}`);
        res.json({ success: true });
    } catch (e) {
        console.error('Admin delete record error:', e);
        res.status(500).json({ success: false, error: e.message });
    }
});

// Helper to rebuild daily metrics from scratch based on current records
function rebuildDaySummary(day) {
    day.summary = { totalFiles: 0, totalLeads: 0, totalShown: 0, byAgent: {} };
    for (const rec of day.records) {
        if (rec.mode === 'CXL') continue; // Skip cancelled records in daily statistics

        day.summary.totalFiles++;
        day.summary.totalLeads += (Number(rec.total) || 0);
        day.summary.totalShown += (Number(rec.shown) || 0);

        const agent = rec.agent || 'Unassigned';
        if (!day.summary.byAgent[agent]) {
            day.summary.byAgent[agent] = { files: 0, leads: 0, shown: 0 };
        }
        day.summary.byAgent[agent].files++;
        day.summary.byAgent[agent].leads += (Number(rec.total) || 0);
        day.summary.byAgent[agent].shown += (Number(rec.shown) || 0);
    }
}

// ─────────────────────────────────────────────
//  ADMIN — UPDATE SINGLE RECORD
//  Body: { dateKey, transactionId, agent, filename, mode, newShown, newHidden, someShown, someHidden, note }
// ─────────────────────────────────────────────
app.put('/api/admin/record', (req, res) => {
    try {
        const { dateKey, transactionId, agent, filename, mode, newShown, newHidden, someShown, someHidden, note } = req.body;
        if (!dateKey || !transactionId) {
            return res.status(400).json({ success: false, error: 'dateKey and transactionId are required.' });
        }

        const data = loadAnalytics();
        if (!data[dateKey]) {
            return res.status(404).json({ success: false, error: `No data for date ${dateKey}.` });
        }

        const rec = data[dateKey].records.find(r => r.transactionId === transactionId);
        if (!rec) {
            return res.status(404).json({ success: false, error: 'Record not found.' });
        }

        const oldAgent = rec.agent;

        if (agent !== undefined) rec.agent = agent;
        if (filename !== undefined) rec.filename = filename;
        if (mode !== undefined) rec.mode = mode.toUpperCase();
        if (note !== undefined) rec.note = note;

        if (rec.mode === 'CXL') {
            rec.newShown = 0;
            rec.newHidden = 0;
            rec.someShown = 0;
            rec.someHidden = 0;
            rec.shown = 0;
            rec.hidden = 0;
            rec.total = 0;
        } else {
            if (newShown !== undefined) rec.newShown = Number(newShown) || 0;
            if (newHidden !== undefined) rec.newHidden = Number(newHidden) || 0;
            if (someShown !== undefined) rec.someShown = Number(someShown) || 0;
            if (someHidden !== undefined) rec.someHidden = Number(someHidden) || 0;

            rec.shown = (rec.newShown || 0) + (rec.someShown || 0);
            rec.hidden = (rec.newHidden || 0) + (rec.someHidden || 0);
            rec.total = rec.shown + rec.hidden;
        }

        rebuildDaySummary(data[dateKey]);
        saveAnalytics(data);

        // Sync with history.json files list
        const history = JSON.parse(fs.readFileSync(historyPath));
        const fileIdx = history.files.findIndex(f => f.transactionId === transactionId);
        if (fileIdx > -1) {
            if (agent !== undefined) history.files[fileIdx].agent = agent;
            if (filename !== undefined) history.files[fileIdx].name = filename;
            fs.writeFileSync(historyPath, JSON.stringify(history, null, 2));
        }

        // Broadcast updated event
        io.emit('record_updated', {
            transactionId,
            dateKey,
            record: rec,
            oldAgent
        });

        console.log(`> [ADMIN] Updated record ${transactionId} for date ${dateKey}`);
        res.json({ success: true, record: rec });
    } catch (e) {
        console.error('Admin update record error:', e);
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
        const data = loadAnalytics();
        const day = data[dateKey];

        if (!day || !day.records || day.records.length === 0) {
            return res.status(404).json({ success: false, error: `No data found for ${dateKey}` });
        }

        const wb = new ExcelJS.Workbook();
        const sheet = wb.addWorksheet('Report');

        // Column definitions
        sheet.columns = [
            { header: 'Agent', key: 'agent', width: 15 },
            { header: 'Total Gross', key: 'total', width: 18 },
            { header: 'Total Net', key: 'shown', width: 18 },
            { header: 'New Net', key: 'newShown', width: 15 },
            { header: 'New Hidden', key: 'newHidden', width: 15 },
            { header: 'Some Net', key: 'someShown', width: 15 },
            { header: 'Some Hidden', key: 'someHidden', width: 15 },
            { header: 'File Name', key: 'filename', width: 50 },
            { header: 'Date', key: 'date', width: 15 },
            { header: 'Mode', key: 'mode', width: 15 },
            { header: 'Time', key: 'time', width: 12 },
            { header: 'Notes', key: 'note', width: 30 }
        ];

        // Style header row
        const headerRow = sheet.getRow(1);
        headerRow.eachCell((cell) => {
            cell.font = { name: 'Arial', size: 11, bold: true, color: { argb: 'FFFFFFFF' } };
            cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF4F81BD' } };
            cell.alignment = { vertical: 'middle', horizontal: 'center' };
            cell.border = { top: { style: 'thin' }, left: { style: 'thin' }, bottom: { style: 'thin' }, right: { style: 'thin' } };
        });

        // Data rows
        for (const rec of day.records) {
            const row = sheet.addRow({
                agent: rec.agent,
                total: rec.total,
                shown: rec.shown,
                newShown: rec.newShown || 0,
                newHidden: rec.newHidden || 0,
                someShown: rec.someShown || 0,
                someHidden: rec.someHidden || 0,
                filename: rec.filename,
                date: dateKey,
                mode: rec.mode,
                time: rec.time || '',
                note: rec.note || '',
            });
            row.eachCell((cell, colNumber) => {
                cell.font = { name: 'Arial', size: 11 };
                cell.alignment = { vertical: 'middle', horizontal: colNumber === 8 ? 'left' : 'center' };
                cell.border = { top: { style: 'thin' }, left: { style: 'thin' }, bottom: { style: 'thin' }, right: { style: 'thin' } };
            });
        }

        // Summary row at the bottom
        const summaryRow = sheet.addRow({
            agent: 'TOTAL',
            total: day.summary.totalLeads,
            shown: day.summary.totalShown,
            newShown: day.records.reduce((sum, r) => sum + (r.newShown || 0), 0),
            newHidden: day.records.reduce((sum, r) => sum + (r.newHidden || 0), 0),
            someShown: day.records.reduce((sum, r) => sum + (r.someShown || 0), 0),
            someHidden: day.records.reduce((sum, r) => sum + (r.someHidden || 0), 0),
            filename: `${day.summary.totalFiles} file(s) processed`,
            date: dateKey,
            mode: '—',
            time: '—',
            note: '—',
        });
        summaryRow.eachCell((cell) => {
            cell.font = { name: 'Arial', size: 11, bold: true };
            cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFD9E1F2' } };
            cell.alignment = { vertical: 'middle', horizontal: 'center' };
            cell.border = { top: { style: 'thin' }, left: { style: 'thin' }, bottom: { style: 'thin' }, right: { style: 'thin' } };
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
        const data = loadAnalytics();
        const records = [];

        for (const dateKey of Object.keys(data).sort((a, b) => b.localeCompare(a))) {
            for (const rec of [...data[dateKey].records].reverse()) {
                records.push({
                    agent: rec.agent,
                    total: rec.total,
                    shown: rec.shown,
                    filename: rec.filename,
                    date: dateKey,
                    mode: rec.mode
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
        const incomingConfig = req.body;
        if (!incomingConfig || typeof incomingConfig !== 'object') throw new Error('Invalid config payload');
        const existingConfig = JSON.parse(fs.readFileSync(configPath));
        const newConfig = { ...existingConfig, ...incomingConfig };
        ['sharedBasePath', 'usBase', 'ukBase', 'tempZone'].forEach(key => {
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

        // Enforce that at least ONE user has the Admin role
        const adminCount = users.filter(u => u.role === 'Admin').length;
        if (adminCount === 0) throw new Error('System must have at least one Admin account.');

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
//  ADMIN — CLOSE SERVER
// ─────────────────────────────────────────────
app.post('/api/admin/close-server', (req, res) => {
    try {
        const { username } = req.body;
        const users = JSON.parse(fs.readFileSync(usersPath));
        const user = users.find(u => u.username === username);
        if (!user || user.role !== 'Admin') {
            return res.status(403).json({ success: false, error: 'Unauthorized. Only admins can close the server.' });
        }
        res.json({ success: true, message: 'Server is closing down...' });
        console.log(`> Server shutdown initiated by admin: ${username}`);
        setTimeout(() => {
            process.exit(0);
        }, 1000);
    } catch (e) {
        res.status(500).json({ success: false, error: e.message });
    }
});

// ─────────────────────────────────────────────
//  LIVE FLOOR TRACKING (SOCKET.IO)
// ─────────────────────────────────────────────
io.on('connection', (socket) => {
    // Sync current dispatch mode on connection
    socket.emit('dispatchModeChanged', isAutoDispatch);

    socket.on('agent_status', (data) => {
        // Track the specific socket connection ID with the username
        socket.username = data.agent;
        activeFloor[data.agent] = {
            status: data.status,
            show: data.show,
            startTime: data.status === 'Working' ? Date.now() : null
        };
        io.emit('floor_update', activeFloor);
    });

    socket.on('toggleDispatchMode', (status) => {
        isAutoDispatch = status;
        console.log(`> [DISPATCH] Mode changed to: ${isAutoDispatch ? 'AUTOMATED' : 'MANUAL'}`);
        io.emit('dispatchModeChanged', isAutoDispatch);
    });

    socket.on('requestManualShow', (agentData) => {
        console.log(`> [DISPATCH] Manual show requested by: ${agentData.name} (${agentData.id})`);
        io.emit('manual_show_requested', agentData);
    });

    socket.on('bulkAssignShows', (data) => {
        const { agentId, showIds } = data; // agentId is the username string
        if (!agentId || !showIds || !Array.isArray(showIds)) return;

        console.log(`> [DISPATCH] Bulk assignment for ${agentId}: ${showIds.length} shows`);

        try {
            let currentShows = JSON.parse(fs.readFileSync(showsPath));

            currentShows.forEach(s => {
                if (showIds.includes(s.id)) {
                    s.agentName = agentId;
                    s.status = 'Pending';
                }
            });

            fs.writeFileSync(showsPath, JSON.stringify(currentShows, null, 2));

            // 1. Broadcast to all admins that the table needs refresh
            io.emit('hopper_updated');

            // 2. Notify the specific agent so their Briefing Room updates
            for (const [id, s] of io.of("/").sockets) {
                if (s.username === agentId) {
                    s.emit('showsAssigned');
                    console.log(`> [DISPATCH] Notified targeted agent: ${agentId}`);
                }
            }
        } catch (err) {
            console.error('[DISPATCH] Error during bulk assignment:', err);
        }
    });

    socket.on('disconnect', () => {
        if (socket.username && activeFloor[socket.username]) {
            activeFloor[socket.username].status = 'Offline';
            io.emit('floor_update', activeFloor);
        }
    });
});

// ─────────────────────────────────────────────
//  START
// ─────────────────────────────────────────────
server.listen(PORT, () => {
    console.log(`\n> ZORG-NEXUS ACTIVE`);
    console.log(`> http://localhost:${PORT}\n`);
});
