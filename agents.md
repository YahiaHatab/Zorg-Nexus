# AI Agent Developer Guide (`agents.md`)

This document is the authoritative instruction manual and operational guideline for AI coding assistants (e.g. Antigravity, Gemini, Claude, Copilot) maintaining, debugging, or extending the **Zorg Nexus** codebase.

---

## 1. Project Identity & Architecture Overview

Zorg Nexus is a monolithic, zero-build Node.js application managing trade show exhibitor lead research, real-time agent floor monitoring, Excel auditing, and regional file dispatching.

### Technology Stack
- **Server**: Node.js (CommonJS), Express 5.2.1, HTTP, Socket.IO 4.8.3.
- **Spreadsheet Engine**: ExcelJS 4.4.0.
- **File Uploads**: Multer 2.1.1 (staged to local `temp/`).
- **Desktop Companion**: Electron 43.3.0 floating HUD widget.
- **Frontend**: Vanilla ES6 JavaScript, HTML5, CSS3 with CSS variables & Glassmorphism, Chart.js 4.4.2. No build step (no Webpack, Vite, React, or TypeScript).
- **Persistence**: Plain JSON file-based storage in the project root.
- **Host OS**: Windows 10/11 environment with VBScript background runners (`ZORG.vbs`, `WIDGET.vbs`, `Run.vbs`).

---

## 2. Repository Layout & File Responsibilities

```
Zorg-Nexus/
├── .gitignore                   # Excludes runtime data, node_modules, excel files, temp zones
├── package.json                 # Project dependencies & npm scripts (start, widget, start:all)
├── server.js                    # Core monolith (REST APIs, WebSockets, Excel processing, storage)
├── ZORG.vbs                     # Background launcher for Node server + browser
├── WIDGET.vbs                   # Background launcher for Electron widget
├── Run.vbs                      # General runner shortcut
├── spec.md                      # Complete system specification
├── agents.md                    # This AI developer guide
│
├── public/                      # Static web assets served by Express
│   ├── index.html               # Agent Portal (Workspace, Briefing Room, Live Leaderboard)
│   ├── admin.html               # Admin Dashboard (Floor Monitor, Hopper, Ledger, Config, Users)
│   └── widget.html              # Electron HUD widget interface
│
├── widget/                      # Electron desktop client
│   ├── main.js                  # Electron main process (BrowserWindow, IPC, auto-resize)
│   └── preload.js               # Context bridge for widget
│
├── [Runtime / Ignored Files]    # DO NOT COMMIT TO GIT
│   ├── config.json              # Runtime configuration & destination folders
│   ├── users.json               # User credentials, roles, and themes
│   ├── shows.json               # Show Hopper queue dataset
│   ├── history.json             # Recent upload records and audit logs
│   ├── analytics.json           # Daily/monthly metrics ledger
│   ├── agent_shows_history.json # Completed/cancelled shows history
│   ├── output/                  # Formatted and sorted Excel lead files (US/UK)
│   ├── archive/                 # Agent backup copies of uploaded files
│   ├── temp/                    # Multer upload staging area
│   └── scratch/                 # Diagnostic/test scripts
```

---

## 3. Strict Coding Conventions & Rules

### 3.1 Backend Architecture (`server.js`)
1. **Single-File Monolith**:
   - `server.js` houses all business logic (~2,250+ lines).
   - When modifying `server.js`, make surgical edits. Do NOT rewrite entire route blocks unnecessarily.
   - Always verify syntax after editing:
     ```powershell
     node -c server.js
     ```
2. **Synchronous File Persistence**:
   - Data is stored in root-level JSON files. Read via `fs.readFileSync` and write via `fs.writeFileSync(..., JSON.stringify(data, null, 2))`.
   - Maintain defensive parsing: always wrap `JSON.parse` in `try/catch` or ensure fallback defaults exist.
   - Respect the startup bootstrap block (lines 16–60 of `server.js`), which auto-creates default `config.json`, `users.json`, `history.json`, `analytics.json`, and `shows.json` if missing.
3. **Username Propagation**:
   - When a username is renamed or updated, always call `propagateUsernameChange(oldName, newName)`. It propagates changes across `shows.json`, `history.json`, and `analytics.json`.
4. **Undo Registry**:
   - The server maintains a rolling 5-minute `undoRegistry` (Map). Any new file upload must register its `transactionId`, file paths, and database state so `/api/undo` can roll it back cleanly.

### 3.2 Excel Parsing & Reason Tag Normalization
1. **Row Auditing**:
   - ExcelJS row indices are **1-based**.
   - Check `row.hidden === true` or `row.height === 0` to determine whether a lead is hidden.
   - Do not assume cell values are plain strings; cell values in ExcelJS can be objects (`{ text: ... }`, `{ formula: ... }`, `{ hyperlink: ... }`). Use helper unwrap methods.
2. **Reason Tag Mapping**:
   - Hidden row disqualification tags MUST pass through `normalizeReason()`.
   - Canonical reasons include: `Int'l`, `Publisher`, `CXL`, `NF`, `No Num`, `Repeated`, `Local`, `Blank`, `Invalid Reason`.
   - Check `config.cxlTags` (default: `['Pricing', 'Duplicates', 'SameList']`) for cancellation tag logic.

### 3.3 Frontend Conventions (`public/`)
1. **Zero-Build Architecture**:
   - Do NOT introduce external frameworks (React, Vue, Svelte), module bundlers (Vite, Webpack), or TypeScript compilation steps.
   - All code in `index.html` and `admin.html` must remain standard ES6 JavaScript and native DOM APIs.
2. **Styling & Theming**:
   - Follow the established design system: modern CSS variables (`--bg-main`, `--bg-panel`, `--accent-blue`, `--accent-cyan`, `--success`, `--danger`, etc.).
   - Support both Light and Dark themes via `body.dark-mode`.
   - Preserve glassmorphism panel styles (`.glass-panel`, `--glass-bg`, `--glass-border`).
3. **Socket.IO Real-Time Sync**:
   - Use the native client loaded from `/socket.io/socket.io.js`.
   - Listen for `floorUpdate`, `queueUpdate`, and `showAssigned` to keep UI state synchronized without polling.

### 3.4 Windows & Electron Environment
1. **Paths**:
   - Always use `path.join(__dirname, ...)` for file operations to ensure Windows backslash compatibility.
2. **Process Management**:
   - Background tasks on Windows are initiated via VBScript (`ZORG.vbs`, `WIDGET.vbs`). Do not kill processes abruptly without checking port 3005.

---

## 4. Git & Security Guardrails

> [!CAUTION]
> **Strict Git Hygiene**:
> 1. NEVER stage or commit `node_modules/` to git.
> 2. NEVER stage `.xlsx`, `.xls`, or `.csv` files containing trade show lead data.
> 3. NEVER commit runtime database state: `users.json`, `config.json`, `shows.json`, `history.json`, `analytics.json`, or `agent_shows_history.json`.
> 4. Ensure `scratch/`, `temp/`, `output/`, and `archive/` directories remain untracked.
> 5. If adding new data files or output zones, update [.gitignore](file:///c:/Users/Yehya/Desktop/Zorg-Nexus/.gitignore) immediately.

---

## 5. Standard Verification Checklist

Before completing any task or making commits in this repo, run through this checklist:

1. **Syntax Check**:
   ```powershell
   node -c server.js
   node -c widget/main.js
   node -c widget/preload.js
   ```
2. **Git Status Inspection**:
   ```powershell
   git status
   ```
   Ensure no runtime data (`.json`), spreadsheets (`.xlsx`), or `node_modules/` are marked as untracked or staged.
3. **JSON Schema Integrity**:
   Verify that any updates to JSON structures preserve existing keys and backward compatibility with previous analytics or history records.
4. **WebSocket Event Verification**:
   If modifying API endpoints that alter show or floor state, ensure corresponding Socket.IO events (`io.emit('floorUpdate', ...)` or `io.emit('queueUpdate', ...)`) are triggered.

