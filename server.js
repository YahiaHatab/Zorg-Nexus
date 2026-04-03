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
const PORT   = 3017;
const activeFloor = {}; // Tracks agent status and timers

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
    cxlTags:  ['Pricing', 'Duplicates', 'SameList']
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

const showsPath = path.join(__dirname, 'shows.json');
if (!fs.existsSync(showsPath)) {
    fs.writeFileSync(showsPath, JSON.stringify([], null, 2));
    console.log('> Created default shows.json');
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
        if(idxA === -1) return 1;
        if(idxB === -1) return -1;
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
    
    // 2. Look for a Pending show specifically Pinned/Assigned to this user FIRST
    const pinnedShow = currentShows.find(s => s.agentName === username && s.status === 'Pending');
    
    let showToAssign = null;
    
    if (pinnedShow) {
        showToAssign = pinnedShow;
    } else {
        // 3. If no pinned show, look for the first UNASSIGNED Pending show
        // (Ensures we don't accidentally give away a show pinned to a different agent)
        const unassignedShow = currentShows.find(s => (!s.agentName || s.agentName.trim() === '') && s.status === 'Pending');
        
        if (unassignedShow) {
            showToAssign = unassignedShow;
        }
    }
    
    // 4. If nothing was found, let the agent know the queue is empty
    if (!showToAssign) {
        return res.json({ success: false, message: 'No shows available' });
    }
    
    // 5. Mark the chosen show as In Progress and officially assign it
    showToAssign.status = 'In Progress';
    showToAssign.agentName = username;
    
    fs.writeFileSync(showsPath, JSON.stringify(currentShows, null, 2));
    io.emit('hopper_updated'); // Notify admin dashboard
    
    res.json({ success: true, show: showToAssign });
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
                            } catch(e) { console.error('Scrub failed for:', file); }
                        }
                    });
                }
            });
        } catch(e) {
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
        } catch(e) { console.error("Error logging CXL to analytics:", e); }

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
        const h12upload  = date.getHours() % 12 || 12;
        const amPmUpload = date.getHours() < 12 ? 'AM' : 'PM';
        const timeStr    = `${String(h12upload).padStart(2,'0')}:${String(date.getMinutes()).padStart(2,'0')}:${String(date.getSeconds()).padStart(2,'0')} ${amPmUpload}`;
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
        const errDate = new Date();
        const h12err  = errDate.getHours() % 12 || 12;
        const amPmErr = errDate.getHours() < 12 ? 'AM' : 'PM';
        const timeStr = `${String(h12err).padStart(2,'0')}:${String(errDate.getMinutes()).padStart(2,'0')}:${String(errDate.getSeconds()).padStart(2,'0')} ${amPmErr}`;
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
        const h12undo  = now.getHours() % 12 || 12;
        const amPmUndo = now.getHours() < 12 ? 'AM' : 'PM';
        const timeStr  = `${String(h12undo).padStart(2,'0')}:${String(now.getMinutes()).padStart(2,'0')}:${String(now.getSeconds()).padStart(2,'0')} ${amPmUndo}`;
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
//  LIVE FLOOR TRACKING (SOCKET.IO)
// ─────────────────────────────────────────────
io.on('connection', (socket) => {
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
