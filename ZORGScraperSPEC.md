# ZORG-Ω Technical Specification & Module Reference (`SPEC.md`)
**Version:** 6.1.0  
**Status:** Living Technical Standard  
**Runtime:** Node.js (ES6+ CommonJS)  
**Core Process:** `app.js` (Express 5 + Socket.io 4)  
**Port:** `3000` (`http://localhost:3000`)

---

## 1. System Overview & Architecture

ZORG-Ω is a modular B2B trade show exhibitor extraction hub, telemetry monitor, autonomous engine validator, and multi-agent operations hub. It provides an operational dashboard and headless execution pipeline where human operators and automated AI agents execute scraping protocols against diverse trade show floorplans, directories, and exhibition management platforms.

```
                               ┌──────────────────────────────────────────────────────────┐
                               │                   Browser Web Client                     │
                               │       (public/index.html, client.js, Tailwind CDN)       │
                               └─────────────────────────────┬────────────────────────────┘
                                                             │ HTTP POST /run (Single URL or Legacy Mode)
                                                             │ HTTP GET/POST Canary & Importer APIs
                                                             │ WebSocket (Live Telemetry, HUD, Pulse, Logs)
                                                             ▼
┌──────────────────────────────────────────────────────────────────────────────────────────────────────────────────────────┐
│                                                  ZORG-Ω Core (app.js)                                                    │
├──────────────────────────────────────────────────────────────────────────────────────────────────────────────────────────┤
│ • IP Bouncer & Auth Firewall (bans.json, users.json)                                                                     │
│ • Run Lock & Session Tracker (activeEngines, activeUsers, activeScrapes)                                                 │
│ • AsyncLocalStorage Context & Global HTTP Cancellation Interceptors (Axios / Fetch)                                      │
│ • Server Resource Pulse (/api/health -> CPU, RAM, active scrapes)                                                        │
│ • Universal Telemetry Interceptor (Array.prototype.push -> Scrape Rate, ETA, Exhibitor Cards, Live HUD Preview)           │
│ • Dynamic Engine Loader & Require Cache Invalidation                                                                     │
├──────────────────────────────────────────────────────────────────────────────────────────────────────────────────────────┤
│                                            Modular Subsystems (lib/)                                                     │
│ ┌───────────────────────────┐ ┌───────────────────────────┐ ┌───────────────────────────┐ ┌────────────────────────────┐ │
│ │   lib/orchestrator.js     │ │     lib/detector.js       │ │      lib/adapter.js       │ │    lib/gemini-recon.js     │ │
│ │ Pipeline Orchestrator     │ │ Tier 1 Regex Fast-Path    │ │ Modern Engine Object /    │ │ Tier 3 Zero-Day Recon &    │ │
│ │ Ingestion & Self-Healing  │ │ Tier 2 DOM/Script Sniffer │ │ Legacy Function Bridge    │ │ Zero-Record Recovery       │ │
│ └─────────────┬─────────────┘ └─────────────┬─────────────┘ └─────────────┬─────────────┘ └─────────────┬──────────────┘ │
│               │                             │                             │                             │                │
│ ┌─────────────┴─────────────┐ ┌─────────────┴─────────────┐               └─────────────────────────────┘                │
│ │      lib/canary.js        │ │  lib/engine-importer.js   │                                                              │
│ │ Autonomous Health Monitor │ │ Deterministic Validator   │                                                              │
│ │ Benchmarking & Latency    │ │ Hot-Reload Code Injection │                                                              │
│ └───────────────────────────┘ └───────────────────────────┘                                                              │
├──────────────────────────────────────────────────────────────────────────────────────────────────────────────────────────┤
│                                                      Engines/                                                            │
│ ┌──────────────────────┐  ┌──────────────────────┐  ┌──────────────────────┐  ┌────────────────────────────────────────┐ │
│ │     Engines/a2z.js   │  │    Engines/Eshow.js  │  │ Engines/cadmiumv2.js │  │     Engines/<custom-engine>.js         │ │
│ │   (A2Z Floorplan)    │  │    (eShow Platform)  │  │ (Conference Harvester│  │  (26+ Dynamic & Platform Scrapers)    │ │
│ └──────────┬───────────┘  └──────────┬───────────┘  └──────────┬───────────┘  └───────────────────┬────────────────────┘ │
│            │                         │                         │                                  │                      │
│            └─────────────────────────┼─────────────────────────┴──────────────────────────────────┘                      │
│                                      ▼ Returns rawRecords Array                                                          │
│ ┌──────────────────────────────────────────────────────────────────────────────────────────────────────────────────────┐ │
│ │                                            ZORG-Ω Standardizer Engine                                                │ │
│ │ • Record Sanitization (clean() transforms null/undefined/empty/'N/A'/'NULL' -> 'N/A')                                 │ │
│ │ • Case-Insensitive Deduplication (uniqueKey = cleanName.toLowerCase())                                               │ │
│ │ • Multi-Booth Consolidation (Appends new booths: "A-101, B-205")                                                     │ │
│ │ • Normalized 8-Column Transformation (Exhibitors Sheet)                                                              │ │
│ │ • Dynamic Export Name Derivation: deriveUniqueExportName({ customFileName, eventName, engineName, rawUrl, mode })      │ │
│ │ • XLSX Binary Buffer Compilation (XLSX.write(wb, { type: 'buffer', bookType: 'xlsx' }))                              │ │
│ │ • File Archival (public/exports/<timestamp>-<name>.xlsx) & HTTP Binary Stream Delivery                                │ │
│ └──────────────────────────────────────────────────────────────────────────────────────────────────────────────────────┘ │
└──────────────────────────────────────────────────────────────────────────────────────────────────────────────────────────┘
```

### 1.1 Core Execution Paradigms

1. **Modular Engine Isolation:**
   - Every scraping engine resides in `Engines/` as an independent CommonJS module.
   - Engines share zero state with each other and never import other engines.
   - Core server routes and orchestrator dynamically import engines using `require(path)` after invalidating `require.cache[require.resolve(path)]`, allowing hot-reloading without server restarts.

2. **Dual Engine Classification:**
   - **General Engines:** Built to handle any event hosted on a platform family (e.g., `Engines/a2z.js`, `Engines/expofp-general.js`, `Engines/map-dynamics-general.js`, `Engines/cadmiumv2.js`, `Engines/ungerboeck.js`, `Engines/mapyourshow.js`). They accept dynamic input parameters (show slugs, event keys, session cookies, or pasted cURL commands) from the UI palette or resolve them autonomously from target URLs.
   - **Specific Engines:** Built for a single, one-off trade show (e.g., `Engines/engine_fespa2026.js`, `Engines/uniti-expo-2026.js`). Target endpoints and parameters are hardcoded; frontend parameters are ignored.

3. **In-Flight Cancellation via `AsyncLocalStorage`:**
   - When a client aborts an extraction or disconnects from `POST /run`, `app.js` sets `runState.aborted = true`.
   - Node's `AsyncLocalStorage` context propagates `runState` down the asynchronous execution tree.
   - Global Axios and Fetch interceptors immediately reject outbound requests with `OPERATION_ABORTED_BY_SYSTEM`, halting socket traffic instantly without lingering background processes.

4. **Universal Telemetry, Live HUD & Real-Time Record Interception:**
   - In addition to the standard `emitLog(message)` callback streaming to console and Socket.io (`new-log`), `app.js` globally intercepts `Array.prototype.push` during scrape execution.
   - Any engine pushing record objects containing `"Company Name"` or `"company"` automatically updates real-time operational telemetry:
     - Scrape rate (records/sec) and dynamic completion ETA calculation.
     - Live exhibitor preview cards streaming directly to the operator HUD.
     - Live extraction progress (`current`, `total`, `percent`, `firstItem` sample preview).
   - This provides immediate visual confirmation that the engine is capturing valid records before full completion.

5. **Test Mode Protocol (`testMode: true`):**
   - The runner integrates a dry-run interceptor embedded in the `Array.prototype.push` proxy.
   - When 5 valid records containing `"Company Name"` are captured in test mode, the engine execution is short-circuited via a `TEST_MODE_LIMIT_MET` exception, verified, and sent to the Standardizer without scraping entire catalogs.

---

### 1.2 Autonomous Single-URL Ingestion Architecture

ZORG-Ω transitions from manual engine-selection and manual-parameter entry to a zero-touch, single-input URL ingestion architecture. An operator supplies solely a trade show directory URL; ZORG-Ω autonomously identifies the platform, resolves hidden event keys/tokens/endpoints, and dispatches the scraper.

```
                  ┌─────────────────────────────────────────┐
                  │    Operator Input: Single Target URL    │
                  └────────────────────┬────────────────────┘
                                       │
                                       ▼
                  ┌─────────────────────────────────────────┐
                  │      Tier 1: Fast Path Regex Match      │
                  │        (engine.urlPatterns)             │
                  └───────┬─────────────────────────┬───────┘
                          │ Match Found             │ No Match
                          ▼                         ▼
                  ┌───────────────┐     ┌───────────────────────────────────┐
                  │ Candidate     │     │ Tier 2: Deep Path DOM Sniffer     │
                  │ Engine        │     │ (Fetch Root HTML -> fingerprint)  │
                  └───────┬───────┘     └───────────┬───────────────────┬───┘
                          │                         │ Match Found       │ No Match
                          │                         ▼                   ▼
                          │             ┌───────────────────┐ ┌─────────────────────────┐
                          │             │ Candidate Engine  │ │ Tier 3: Gemini Fallback │
                          │             └─────────┬─────────┘ │ Recon API (JSON Schema) │
                          │                       │           └────────────┬────────────┘
                          │                       │                        │
                          ▼                       ▼                        ▼
                  ┌─────────────────────────────────────────────────────────────┐
                  │      Autonomous Resolution: engine.resolveConfig()          │
                  │  (Extracts hidden tokens, show keys, & API endpoints)       │
                  └──────────────────────────────┬──────────────────────────────┘
                                                 │
                                                 ▼
                  ┌─────────────────────────────────────────────────────────────┐
                  │       Autonomous Dispatch: engine.run(config)               │
                  │       (Paginates, streams telemetry, returns records)       │
                  └──────────────────────────────┬──────────────────────────────┘
                                                 │
                                                 ▼
                                     ┌───────────────────────┐
                                     │ Records Captured > 0? │
                                     └───────┬───────────────┘
                                             │
                             ┌───────────────┴───────────────┐
                             │ Yes                           │ No (Zero Records)
                             ▼                               ▼
                 ┌───────────────────────┐       ┌──────────────────────────────────────┐
                 │ Standardizer Engine   │       │ Step 4: Zero-Record Self-Healing     │
                 │ Deduplication & XLSX  │       │ (healZeroRecordPage via Gemini Recon)│
                 └───────────────────────┘       └──────────────────┬───────────────────┘
                                                                    │ Recovered Correct URL?
                                                                    ▼
                                                         [Re-Dispatch Pipeline]
```

#### 1.2.1 Two-Tier Detection Pipeline (`lib/detector.js`)
1. **Tier 1: Fast Path URL Regex Matching (`urlPatterns`):**
   - Every modern engine declares an array of canonical regular expressions (`urlPatterns: [ /.../i ]`).
   - Incoming directory URLs are matched against registered patterns synchronously.
   - If a distinct match is identified, candidate engine selection completes with zero network overhead.

2. **Tier 2: Deep Path DOM & Script Sniffer (`fingerprint(html, url)`):**
   - When Tier 1 yields no match (due to custom domains, white-labeled exhibitor portals, or dynamic redirects), ZORG-Ω executes a lightweight HTTP GET request to fetch the directory's initial HTML document using realistic browser headers.
   - The sniffer invokes each candidate engine's synchronous `fingerprint(html, url)` predicate.
   - Fingerprint routines inspect embedded `<script>` tag sources, CDN asset bundles, global window variables (e.g., `window.__INITIAL_STATE__`, `window.a2z`), unique DOM containers (e.g., `#event-harvester-app`, `div[data-booth]`), or platform-specific meta tags.

#### 1.2.2 Tier 3: Autonomous Fallback Recon Engine (`lib/gemini-recon.js`)
- **Activation Scope:** Activated strictly as a fallback when deterministic URL regex (Tier 1) and DOM fingerprints (Tier 2) match zero known engines.
- **Structured Reconnaissance:** Transmits structured initial HTML metadata, scripts, and iframe sources to the Gemini Recon API with a constrained JSON schema. The model identifies the platform type, extracts hidden REST/GraphQL endpoints, and flags the required event identifier.
- **Strict Anti-Slop Restriction:** The Gemini Recon API must **NEVER** be invoked inside pagination or record-extraction loops. It operates exclusively during the initial pre-flight reconnaissance phase on root page markup.

#### 1.2.3 Step 4: Zero-Record Self-Healing Recovery (`healZeroRecordPage`)
- When a matching engine successfully runs but extracts 0 records (due to navigation redirection, embedded iframes, or floorplan tab switching), the orchestrator triggers autonomous self-healing via `healZeroRecordPage()`.
- The recon agent analyzes the root page structure to locate the actual exhibitor list URL or iframe source. If a corrected list URL is discovered, the orchestrator updates configuration and re-dispatches extraction automatically.

#### 1.2.4 Autonomous Parameter Resolution (`resolveConfig`)
- After engine identification, ZORG-Ω invokes `await engine.resolveConfig(url, rawHtml, emitLog)`.
- The engine autonomously extracts event slugs, tokens, exhibitor portal IDs, and API base URLs from the URL path/query or the initial DOM/script content.
- Human operators are never required to manually copy and paste internal show keys or cURL strings if the parameters are discoverable within the initial web assets.

### 1.3 Frontend UI Architecture & Immutable DOM Contracts (`public/`)

ZORG-Ω enforces a strict **Dark-Only Industrial Console** aesthetic (Zinc monochrome palette `#09090b`–`#18181d`, `1px solid #27272a` hairline borders, `6px`–`8px` radii, tabular monospace telemetry, zero decorative background orbs or glassmorphic blur).

1. **Separation of Concerns:**
   - `public/style.css` owns all design tokens (`:root`), component primitives (`.card`, `.inp`, `.btn-primary`, `.btn-ghost`, `.badge`), and layout rules.
   - `public/index.html` defines semantic structure and immutable DOM hooks.
   - `public/client.js` manages WebSocket listeners (`new-log`, `update-logs`), HUD telemetry bindings, and REST dispatches (`POST /run`, `/api/canary-*`, `/api/engine/*`).

2. **Immutable DOM Hook Registry:**
   Any UI modification MUST preserve the exact `id` attributes and event bindings for:
   - **Navigation & Views:** `#mainGrid`, `#homePage`, `#libraryView`, `#scraperView`, `#libraryNavBtn`, `#engineCountBadge`, `#canaryFleetBtn`, `#canaryFleetDot`, `#canaryFleetStatus`, `#userProfile`, `#notificationWidget`, `#notifBadge`, `#notificationPanel`, `#notificationList`, `#userCardWrapper`, `#userCardBtn`, `#profileInitial`, `#welcome-user`, `#scrapeCount`, `#userDropdownMenu`, `#dropdownUserName`, `#dropdownAdminSection`, `#adminUpdateCategory`, `#adminUpdateText`, `#dropdownAdminActions`, `#themeToggleBtn` (hidden stub for JS compatibility).
   - **Left & Right Sidebars + Resizers:** `#noticeBoard`, `#sidebarRecentEngines`, `#noticeBoardToggle`, `#bulletinAddBtn`, `#bulletinInputContainer`, `#bulletinTextInput`, `#bulletinFilePreview`, `#bulletinFileName`, `#bulletinFileInput`, `#bulletinSubmitBtn`, `#bulletinList`, `#updatesList`, `#recentFilesPanel`, `#recentFilesToggleBtn`, `#recentFilesCount`, `#recentFilesToggleChevron`, `#recentFilesDrawer`, `#recentFilesList`, `#recentFilesEmpty`, `#onlineList`, `#engineActivityList`, `#pulseStatusText`, `#pulseBar`, `#pulseLoadText`, `#activeScrapesText`, `#resizeHandleLeft`, `#resizeHandleRight`, `#telemetryResizeHandle`.
   - **Universal Ingest & Scraper HUD:** `#universalIngestSection`, `#universalUrl`, `#universalRunBtn`, `#detectionStatusBadge`, `#detectionStepIcon`, `#detectionStepText`, `#detectionStepSub`, `#manualToggleBtn`, `#manualToggleText`, `#manualPaletteContent`, `#instructionArrow`, `#instructionContentWrapper`, `#instructionContent`, `#mode`, `#pinnedEnginesBar`, `#paletteBtn`, `#dropdownBtnSearchIcon`, `#dropdownBtnFavicon`, `#dropdownBtnText`, `#engineListContainer`, `#standardBox`, `#url`, `#curlBox`, `#curl`, `#eshowBox`, `#token`, `#cadBox`, `#cadEventId`, `#cadClientId`, `#cadEventKey`, `#marketBox`, `#dynamicShowId`, `#cookie`, `#customInputBox`, `#file`, `#btn`, `#stopBtn`, `#clearInputsBtn`, `#counterArea`, `#loader`, `#completeText`, `#recordStats`, `#rawCountValue`, `#cleanCountValue`, `#errorText`, `#extractionProgressHUD`, `#hudPulseDot`, `#hudStatusTitle`, `#hudEngineBadge`, `#hudCount`, `#hudTotalSlash`, `#hudTotal`, `#hudRateBox`, `#hudRate`, `#hudEtaBox`, `#hudEta`, `#hudPercent`, `#hudProgressBar`, `#hudTickerSub`, `#hudLiveTicker`, `#hudTickerPlaceholder`, `#telemetryBox`, `#pauseTelemetryBtn`, `#pauseIcon`, `#pauseText`, `#telemetry`, `#adminTestMode`, `#testModeToggle`, `#testModeSwitch`.

---

## 2. Directory & File Structure

```
ZORG/
├── .git/                                 # Git version control metadata
├── .gitignore                            # Comprehensive ignore rules (databases, exports, uploads, sandboxes, env)
├── Docs/                                 # Technical documentation & architecture matrices
│   ├── ENGINE_OVERHAUL_MATRIX.md         # Complete engine migration audit matrix & regex/fingerprint catalog
│   ├── Skill.md                          # ZORG-Ω Engine Architect skill & Network Spy recon schema
│   └── ZORG_CONTEXT.md                   # Legacy architectural notes & critical AI directives
├── Engines/                              # Modular scraping protocols & platform engines
│   ├── a2z.js                            # A2Z / SmallWorld floorplan general engine
│   ├── Algolia.js                        # NürnbergMesse / Algolia search vector scraper
│   ├── cadmiumv2.js                      # Cadmium / Conference Harvester V5 engine (supersedes legacy Cadmium.js)
│   ├── cphichina.js                      # CPHI China API scraper
│   ├── Dusseldorf.js                     # Messe Düsseldorf directory scraper
│   ├── electramining.js                  # Electra Mining scraper
│   ├── engine_fespa2026.js               # FESPA 2026 specific engine
│   ├── Eshow.js                          # eShow concurrent API engine
│   ├── eurosatory.js                     # Eurosatory catalog scraper
│   ├── expofp-general.js                 # ExpoFP general engine (data.js parser)
│   ├── expogenie.js                      # ExpoGenie platform scraper
│   ├── goeshowlist.js                    # goEshowList request URL parser
│   ├── Informa.js                        # Informa Markets cURL-based engine
│   ├── lasvegasmarket.js                 # Las Vegas Market cURL scraper
│   ├── map-dynamics-general.js           # Map-Dynamics marketplace general engine (supersedes legacy MapDynamics.js)
│   ├── mapyourshow.js                    # MapYourShow (MYS) trade show directory engine
│   ├── ncs2026.js                        # National Convenience Show 2026 scraper
│   ├── pheedloop-general.js              # PheedLoop event code scraper
│   ├── salone2026.js                     # Salone del Mobile 2026 scraper
│   ├── sensor-test-alpha.js              # Sensor+Test crawler
│   ├── smarterev5.js                     # Smartere platform scraper
│   ├── southeast.js                      # SouthEast Green scraper
│   ├── tristate.js                       # Expotracker floorplan scraper
│   ├── ungerboeck.js                     # Ungerboeck / Momentus Virtual Floor Plan (VFP) engine
│   ├── uniti-expo-2026.js                # UNITI expo 2026 scraper
│   ├── windeurope.js                     # WindEurope platform scraper
│   ├── zendesk-ekr-vitafoods.js          # Vitafoods exhibitor portal scraper
│   └── _sandbox_*.js                     # [EPHEMERAL] Scratch/recon files (MUST NOT BE COMMITTED)
├── Icons/                                # Application branding, icons & favicons
├── lib/                                  # Core autonomous architecture subsystems
│   ├── adapter.js                        # Bidirectional modern engine contract & legacy caller bridge
│   ├── canary.js                         # Canary health diagnostic engine & latency profiler
│   ├── detector.js                       # Two-tier platform detector (urlPatterns regex & DOM sniffer)
│   ├── engine-importer.js                # Deterministic syntax validator, code sanitizer & registry manager
│   ├── gemini-recon.js                   # Tier 3 AI zero-day recon & zero-record self-healing fallback
│   └── orchestrator.js                   # Central autonomous single-URL pipeline orchestrator
├── public/                               # Static frontend assets served by Express
│   ├── exports/                          # Generated XLSX files (<timestamp>-<filename>.xlsx)
│   ├── uploads/                          # Ephemeral uploads for bulletin board attachments
│   ├── client.js                         # Browser client UI controller, Canary dashboard & Socket.io listeners
│   ├── index.html                        # Industrial dark operator dashboard UI (Dark Mode Only)
│   └── style.css                         # Zinc design tokens, component primitives & layout
├── scripts/                              # Developer and operational CLI utilities
│   └── import-engine.js                  # CLI tool for validating, previewing, and injecting engines
├── utils/                                # Shared server-side helper utilities
│   └── helpers.js                        # Super-greedy recursive booth finder (findBoothDeeply)
├── app.js                                # Primary server, execution manager, standardizer & socket hub
├── bans.json                             # Flat-file database of banned client IP addresses
├── bulletin.json                         # Flat-file database of notice board pins & broadcasts
├── canary-health.json                    # Flat-file database storing engine diagnostic health snapshots
├── deleted_engines.json                  # Flat-file tracking disabled default engine IDs
├── engines.json                          # Dynamic custom engine registry & parameter definitions
├── logs.json                             # Flat-file audit log of administrative updates & engine lifecycle
├── package.json                          # Project dependencies & Node runtime metadata
├── package-lock.json                     # Deterministic dependency tree lockfile
├── README.md                             # Architectural overview & operator quickstart guide
├── readme.txt                            # Windows operator quickstart manual
├── Skill.md                              # ZORG-Ω Engine Architect skill & Network Spy recon schema
└── users.json                            # Flat-file database mapping client IPs to agent names & roles
```

### 2.1 File Invariants & Ephemeral Policies

- **Ephemeral Exports (`public/exports/`):** Output spreadsheets are saved here upon successful extraction. The directory is cleared via `/admin/clear-exports`. Files here must never be committed to Git.
- **Temporary Uploads (`public/uploads/`):** Files uploaded via `/upload-bulletin-file` reside here. Must never be committed to Git.
- **Scratch Files (`Engines/_sandbox_*.js`):** Used strictly for temporary recon and dry-run debugging. Must follow the `_sandbox_<timestamp>_<name>.js` naming pattern and must be deleted prior to marking tasks complete.
- **Protected State Files:** `bans.json`, `bulletin.json`, `canary-health.json`, `deleted_engines.json`, `logs.json`, and `users.json` are flat-file stores updated at runtime. `users.json`, `canary-health.json`, and `bans.json` are explicitly ignored by `.gitignore`.

---

## 3. Engine Interface Contract

To support autonomous single-URL ingestion while preserving platform stability, every scraper residing in `Engines/*.js` must export a standardized Engine Object Contract conforming to the ZORG-Ω specification.

### 3.1 Unified Engine Object Export Schema

```javascript
/**
 * ZORG-Ω Autonomous Scraper Engine Contract
 */
module.exports = {
    // Unique URL-safe identifier matching filename stem
    id: 'engine-id',

    // Human-readable platform label for telemetry and reporting
    name: 'Platform Label',

    // Fast-path URL matching expressions
    urlPatterns: [
        /https?:\/\/(?:www\.)?example\.com\/directory\/([a-zA-Z0-9_-]+)/i
    ],

    // Deep-path DOM and markup fingerprinting
    fingerprint: (html, url) => {
        return html.includes('data-platform="example"') || /example\.com/i.test(url);
    },

    // Autonomous parameter & token resolution (Zero Manual Input)
    async resolveConfig(url, rawHtml, emitLog) {
        // Extracts tokens, show keys, API URLs autonomously from URL and DOM
        return {
            eventSlug: 'sample-slug',
            apiEndpoint: 'https://api.example.com/v1'
        };
    },

    // Core extraction loop returning standardized raw records
    async run(config, emitLog, runState) {
        // Core extraction loop returning rawRecords array
        return rawRecords;
    }
};
```

### 3.2 Property & Method Specifications

| Property / Method | Type | Required | Description |
| :--- | :--- | :--- | :--- |
| `id` | `string` | Yes | Unique, lowercase, URL-safe alphanumeric identifier exactly matching the file basename in `Engines/<id>.js`. |
| `name` | `string` | Yes | Human-readable platform name displayed in logs, socket telemetry, and exported report manifests. |
| `urlPatterns` | `Array<RegExp>` | Yes | Array of compiled regular expressions matching canonical directory URLs, subdomains, and exhibition paths. |
| `fingerprint` | `Function` | Yes | Synchronous predicate `(html: string, url: string) => boolean`. Inspects initial HTML markup, script tags, or DOM structures for platform signatures. |
| `resolveConfig` | `Function` | Yes | Asynchronous resolver `async (url: string, rawHtml: string, emitLog: Function) => Promise<Object>`. Autonomously derives all necessary event keys, API endpoints, and tokens. |
| `run` | `Function` | Yes | Asynchronous runner `async (config: Object, emitLog: Function, runState: Object) => Promise<Array<Object>>`. Executes extraction and returns the `rawRecords` array. |

#### Method Argument Signatures:
1. **`resolveConfig(url, rawHtml, emitLog)`:**
   - `url` (`string`): The raw operator-provided directory URL.
   - `rawHtml` (`string`): The initial HTML response fetched during Tier 2 fingerprinting (or fetched on demand if matched via Tier 1).
   - `emitLog` (`Function`): Telemetry logger: `emitLog(msg: string)`.
   - **Returns:** Resolves to a clean `config` object containing all required operational variables for `run()`.

2. **`run(config, emitLog, runState)`:**
   - `config` (`Object`): Operational configuration produced by `resolveConfig()` (or forwarded via legacy parameters).
   - `emitLog` (`Function`): Real-time logger: `emitLog("Fetching page 3...")`. Automatically stamps time, streams to terminal console, and broadcasts to frontend telemetry via Socket.io.
   - `runState` (`Object`): Operational state container containing `runState.aborted` (`boolean`) and optional `runState.updateProgress(current, total)` to drive the UI progress bar.
   - **Returns:** Resolves to an array of raw exhibitor record objects (`rawRecords`).

### 3.3 Lifecycle & Execution Rules

1. **Zero Disk I/O:** Engines must NEVER generate Excel, CSV, or JSON files on disk. The engine's sole duty is to extract, map, and return an in-memory array of raw record objects. The core Standardizer in `app.js` handles file generation.
2. **Abortion Compliance:** Every iteration loop (pagination, batch, or item-by-item processing) MUST evaluate `runState.aborted`:
   ```javascript
   if (runState && runState.aborted) {
       emitLog("Extraction aborted by operator. Terminating gracefully.");
       break;
   }
   ```
3. **Graceful Degradation:** Network failures on individual profile detail pages must NOT crash the entire scrape. Individual item errors should be logged via `emitLog` and return default `'N/A'` fields. Only systemic failures (e.g., initial master list fetch failure, invalid authentication token) should throw a fatal exception.
4. **Exception Handling:** Unrecoverable errors must be logged via `emitLog("FATAL ERROR: ...")` and rethrown so that `app.js` can clean up locks and return an HTTP 500 status to the client.

### 3.4 Backward Compatibility & Adapter Layer (`lib/adapter.js`)

ZORG-Ω maintains 100% backward compatibility with legacy single-function engines (`module.exports = async function(params, emitLog, runState)`):

1. **Bidirectional Adapter (`lib/adapter.js`):**
   - Modern engine objects are wrapped into callable functions via `adapt(engine)` to allow legacy callers to invoke them directly.
   - Legacy functions are normalized into standard engine objects via `normalizeEngine(engineId, module)` so that the autonomous platform detector and orchestrator can treat all engines uniformly.
2. **Unified Dispatch Endpoint (`POST /run`):**
   - Both autonomous single-URL extraction and legacy manual parameter execution are handled seamlessly by a single endpoint: `POST /run`.
   - If `req.body.url` is supplied, or `mode === 'auto'`, the server activates the autonomous pipeline via `orchestrator.execute(url, options)`.
   - If manual parameters (e.g., cURL command, show ID, auth token) are supplied for a specific engine ID, the server dispatches directly to the selected engine module.

---

## 4. Autonomous Canary Health Monitoring System (`lib/canary.js`)

To guarantee scraping reliability across dynamic exhibitor portals, ZORG-Ω integrates an automated canary health monitor.

### 4.1 System Purpose & Mechanics
- Runs lightweight dry-run tests (`testMode: true` with a 5-record threshold) against curated target URLs.
- Measures network latency in milliseconds.
- Validates data integrity (ensuring captured records contain non-empty `"Company Name"`).
- Automatically updates [`canary-health.json`](file:///c:/Users/Yehya/Desktop/ZORG/canary-health.json).

### 4.2 Health Classifications
| Status | Meaning | Criteria |
| :--- | :--- | :--- |
| `HEALTHY` (Green) | Production Ready | Test completed successfully within timeout; extracted >= 5 valid records. |
| `DEGRADED` (Yellow) | Slow or Partial | Test succeeded but latency exceeded 12,000ms, or partial record schema detected. |
| `FAILED` (Red) | Broken / Action Needed | Request timed out, returned 0 records, or encountered an HTTP/parsing error. |
| `UNTESTED` (Gray) | Unverified | Engine is registered but has not yet undergone an automated canary test. |

### 4.3 Canary Health REST API
- `GET /api/canary-health`: Returns the full JSON mapping of engine health states, latencies, and timestamps.
- `POST /api/canary-test/:engineId`: Executes an on-demand canary check for a single engine and returns the diagnostic report.
- `POST /api/canary-test-all`: Iterates across the entire engine registry sequentially, updating health records.

---

## 5. Autonomous Engine Importer & Dynamic Script Validation (`lib/engine-importer.js`)

ZORG-Ω supports dynamic, zero-restart scraper injection and validation via both CLI and UI.

### 5.1 Architecture & Pipeline
1. **Code Sanitization:** Strips markdown code fences (```javascript ... ```) and extracts raw executable code.
2. **Deterministic Syntax Validation:** Evaluates code in a secure Node.js `vm.Script` sandbox. If parsing yields syntax errors, the import fails immediately with exact line numbers.
3. **Contract Adherence Check:** Verifies presence of required properties:
   - `id`: Non-empty alphanumeric string.
   - `name`: Human-readable label.
   - `urlPatterns`: Non-empty regex array.
   - `resolveConfig`: Autonomous configuration function.
   - `run`: Primary extraction function.
   - `runState.aborted`: Abort check evaluator.
4. **Hot-Reload Injection:** Writes valid code to `Engines/<id>.js`, invalidates require cache, and registers or updates metadata in `engines.json`.

### 5.2 CLI Tooling (`scripts/import-engine.js`)
Operators and CI/CD agents can validate and import engines via the terminal:
```powershell
# Validate syntax and SPEC contract without modifying disk
node scripts/import-engine.js --validate Engines/my-engine.js

# Inject engine from file into ZORG
node scripts/import-engine.js Engines/my-engine.js

# Inject engine code piped from standard input
Get-Content script.js | node scripts/import-engine.js --stdin
```

### 5.3 REST Endpoints for Dynamic Engine Management
- `POST /api/engine/preview-script`: Runs `validateEngineCode(code)` and returns validation status and warnings without writing to disk.
- `POST /api/engine/import-script`: Executes deterministic validation, writes to `Engines/<id>.js`, updates `engines.json`, and returns the registered engine metadata.
- `GET /engine/:id`: Retrieves engine source code and registration metadata.
- `DELETE /delete-engine/:id`: Safely disables an engine and logs the administrative action.

---

## 6. Normalized Exhibitor Data Schema

The ZORG-Ω platform enforces a two-tier data schema: the **Raw Engine Output** and the **Standardized SheetJS Export**.

### 6.1 Raw Engine Output Schema (Pushed by Engines)

Every record object in the array returned by an engine must provide key-value pairs matching the following recognized properties:

| Property Name | Required? | Type | Fallback Value | Description |
| :--- | :--- | :--- | :--- | :--- |
| `"Company Name"` | **CRITICAL** | `string` | — | The legal or trade name of the exhibiting organization. If this field evaluates to `'N/A'`, empty string, `null`, or `undefined`, the Standardizer permanently discards the record. |
| `"Booth"` | Recommended | `string` | `"N/A"` | Booth number, stand code, or pavilion identifier (e.g., `"A-101"`, `"Hall 3 / Stand 402"`). Multiple booths will be concatenated. |
| `"Phone"` | Optional | `string` | `"N/A"` | Telephone or mobile contact number. |
| `"Country"` | Optional | `string` | `"N/A"` | Exhibitor country of origin or registration. |
| `"Contact Name"` | Optional | `string` | `"N/A"` | Primary contact representative. The standardizer also checks `"Contact Person"` as an alias. |
| `"Email"` | Optional | `string` | `"N/A"` | Primary or corporate inquiry email address. |
| `"Address"` | Optional | `string` | `"N/A"` | Physical street address, mailing address, or headquarters location. |
| `"City"` | Optional | `string` | `"N/A"` | Municipality or metropolitan area. |
| `"Website"` | Optional | `string` | `"N/A"` | Company official homepage URL or direct exhibitor directory portal link. |

*Note: For nested API responses where booth keys vary wildly, engines can import `findBoothDeeply` from `../utils/helpers`:*
```javascript
const { findBoothDeeply } = require('../utils/helpers');
const booth = findBoothDeeply(item); // Super-greedy recursive search across common booth keys
```

### 6.2 Standardizer Transformation & Cleaning Logic

When `rawRecords` are returned to `app.js`, the internal Standardizer applies deterministic normalization:

```javascript
const clean = (val) => {
    if (val === undefined || val === null) return 'N/A';
    const str = String(val).trim();
    if (str === '' || str.toUpperCase() === 'N/A' || str.toUpperCase() === 'NULL' || str.toUpperCase() === 'UNDEFINED') return 'N/A';
    return str;
};
```

1. **Filtering:** Any item where `clean(item["Company Name"]) === 'N/A'` is dropped.
2. **Deduplication:** Merged by `cleanName.toLowerCase()`.
3. **Booth Aggregation:** If a company appears across multiple booths, new booth strings are appended:
   ```javascript
   existingEntry["Booth"] = existingEntry["Booth"] !== 'N/A'
       ? `${existingEntry["Booth"]}, ${newBooth}`
       : newBooth;
   ```

### 6.3 Normalized 8-Column Workbook Output (`Exhibitors` Sheet)

The final generated workbook contains an `"Exhibitors"` worksheet with the exact following 8 standardized column headers:

| Index | Column Header | Source Mapping from Raw Engine Record |
| :---: | :--- | :--- |
| **1** | `Exhibitor Name` | `clean(item["Company Name"])` |
| **2** | `Phone Number` | `clean(item["Phone"])` |
| **3** | `Country` | `clean(item["Country"])` |
| **4** | `Contact Name` | `clean(item["Contact Name"] \|\| item["Contact Person"])` |
| **5** | `Email` | `clean(item["Email"])` |
| **6** | `Mails` | `clean(item["Address"] \|\| item["City"])` *(Note: Standardized column name is verbatim `"Mails"`)* |
| **7** | `Booth` | `clean(item["Booth"])` *(Consolidated list)* |
| **8** | `Website` | `clean(item["Website"])` |

### 6.4 Dynamic Export Name Generation (`deriveUniqueExportName`)

Spreadsheet export files are named intelligently by `deriveUniqueExportName()` in `app.js`:
- Incorporates user-provided custom file names, discovered `eventName`, URL show keys, or domain slugs.
- Automatically suffixes the active engine name and timestamp: `<timestamp>-<DerivedName>.xlsx`.
- Stored temporarily in `public/exports/` and made available for instant download.

---

## 7. Networking, Rate-Limiting & Error Boundaries

### 7.1 Permitted HTTP Libraries & Dependencies

- **Axios:** (`const axios = require('axios');`) — Primary HTTP client.
- **Native Fetch:** Supported across Node 18+.
- **Cheerio:** (`const cheerio = require('cheerio');`) — Fast server-side DOM parsing.
- **p-limit:** (`const pLimit = require('p-limit');`) — Concurrency throttling for outbound detail requests.
- **XLSX / ExcelJS:** Standardized binary workbook generation.
- **Multer:** Ephemeral upload management for engine imports and bulletin attachments.
- **STRICT PROHIBITION:** Puppeteer, Playwright, Selenium, and all headless browser automation frameworks are strictly forbidden in ZORG-Ω engines. All scraping must be conducted via direct HTTP requests, API endpoint replication, or static script evaluation.

### 7.2 Concurrency & Rate Limiting

To avoid HTTP 429 (Too Many Requests), IP bans, or socket exhaustion, engines must throttle concurrent outbound requests using either `p-limit` or chunked batch loops.

#### Pattern A: `p-limit` Concurrency Pool
```javascript
const axios = require('axios');
const pLimit = require('p-limit');
const limit = (typeof pLimit === 'function') ? pLimit(10) : pLimit.default(10);

const tasks = items.map(item => limit(async () => {
    if (runState && runState.aborted) return null;
    await new Promise(r => setTimeout(r, 200)); // Staggering delay
    const res = await axios.get(item.url, { headers, timeout: 15000 });
    return parseItem(res.data);
}));

const results = (await Promise.all(tasks)).filter(Boolean);
```

#### Pattern B: Chunked Batch Processing with Stagger Delay
```javascript
const BATCH_SIZE = 5;
const DELAY_MS = 800;

for (let i = 0; i < totalItems; i += BATCH_SIZE) {
    if (runState && runState.aborted) break;

    const chunk = items.slice(i, i + BATCH_SIZE);
    const chunkResults = await Promise.all(chunk.map(async (item) => {
        try {
            const resp = await axios.get(item.detailUrl, { timeout: 10000 });
            return extractDetails(resp.data);
        } catch (e) {
            emitLog(`Non-fatal error on ${item.name}: ${e.message}`);
            return null;
        }
    }));

    chunkResults.forEach(r => { if (r) rawRecords.push(r); });
    emitLog(`Progress: ${Math.min(i + BATCH_SIZE, totalItems)} / ${totalItems}`);
    await new Promise(r => setTimeout(r, DELAY_MS));
}
```

### 7.3 Exponential Backoff & 429 Handling

```javascript
for (let attempt = 0; attempt < 3; attempt++) {
    try {
        const res = await axios.get(targetUrl, { headers, timeout: 10000 });
        return res.data;
    } catch (err) {
        if (err.response && err.response.status === 429) {
            const delay = 1500 * (attempt + 1);
            emitLog(`HTTP 429 Rate limit detected. Backing off for ${delay}ms...`);
            await new Promise(r => setTimeout(r, delay));
        } else if (attempt === 2) {
            throw err;
        } else {
            await new Promise(r => setTimeout(r, 500));
        }
    }
}
```

---

## 8. Adding a New Engine (Step-by-Step Contract)

Follow this precise workflow whenever creating or integrating a new scraper in ZORG-Ω:

### Step 1: Network Reconnaissance & Parameter Mapping
1. Inspect the target exhibition site in DevTools (F12 Network and Elements Tabs) or use the **ZORG-Ω Network Spy** Chrome extension.
2. Identify URL routing rules and formulate compiled regexes for `urlPatterns`.
3. Locate distinct DOM markers, script tags, or global variables for `fingerprint(html, url)`.
4. Determine whether the site serves:
   - A single JSON catalog payload (e.g., Algolia index, ExpoFP `data.js`).
   - Paginated REST/GraphQL endpoints (e.g., CPHI, Finderr Cloud).
   - An initial master list followed by secondary profile popups (e.g., Cadmium, eShow, A2Z).
5. Map configuration variables (event slugs, tenant tokens, API base paths) and implement autonomous extraction inside `resolveConfig()`:
   - **Zero Manual Input Mandate:** If an event key or token is present in the URL, initial HTML document, or embedded script tags, the engine must extract it automatically. Engines requiring human copying of discoverable parameters will fail verification.

### Step 2: Implement Scraper Module in `Engines/<id>.js`
1. Choose a unique, URL-safe alphanumeric identifier (e.g., `cadmiumv2`, `a2z`, `map-dynamics-general`, `ungerboeck`).
2. Create the file: `Engines/<id>.js`.
3. Export the standard Engine Object Contract conforming to Section 3.1.

### Step 3: Deterministic Validation
Execute deterministic syntax and interface contract checks:
```powershell
# 1. Deterministic syntax validation via Node
node --check Engines/<id>.js

# 2. Complete SPEC interface contract and metadata validation
node scripts/import-engine.js --validate Engines/<id>.js
```

### Step 4: Register Engine in `engines.json`
Append the engine metadata object to `engines.json`:
```json
{
  "id": "my-target-platform",
  "name": "My Target Platform",
  "instruction": "Paste the trade show directory URL.",
  "inputType": "Directory URL",
  "category": "General",
  "targetUrl": "https://myplatform.com/"
}
```

### Step 5: Verification & Canary Health Run
1. Run a test extraction via the UI or `POST /run` with `testMode: true`.
2. Run a canary health verification check:
   ```powershell
   # Trigger canary health check via API
   Invoke-RestMethod -Uri http://localhost:3000/api/canary-test/my-target-platform -Method POST
   ```
3. Confirm status registers as `HEALTHY` in `canary-health.json`.
