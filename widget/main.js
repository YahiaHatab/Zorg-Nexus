const { app, BrowserWindow, Tray, Menu, ipcMain, shell } = require('electron');
const path = require('path');

let mainWindow = null;
let tray = null;
let isAlwaysOnTop = true;

// Prevent duplicate widget instances
const gotTheLock = app.requestSingleInstanceLock();
if (!gotTheLock) {
    app.quit();
} else {
    app.on('second-instance', () => {
        if (mainWindow) {
            if (mainWindow.isMinimized()) mainWindow.restore();
            mainWindow.show();
            mainWindow.focus();
        }
    });
}

function createWindow() {
    mainWindow = new BrowserWindow({
        width: 320,
        height: 380,
        frame: false,
        transparent: true,
        alwaysOnTop: isAlwaysOnTop,
        resizable: false,
        skipTaskbar: false,
        hasShadow: true,
        webPreferences: {
            nodeIntegration: false,
            contextIsolation: true,
            preload: path.join(__dirname, 'preload.js')
        }
    });

    const portalUrl = 'http://localhost:3005/widget.html';
    mainWindow.loadURL(portalUrl).catch(() => {
        console.log('> Waiting for Zorg-Nexus server on http://localhost:3005...');
        const retryInterval = setInterval(() => {
            mainWindow.loadURL(portalUrl).then(() => {
                clearInterval(retryInterval);
            }).catch(() => {});
        }, 1000);
    });

    mainWindow.on('closed', () => {
        mainWindow = null;
    });

    // Create System Tray
    createTray();
}

function createTray() {
    if (tray) return;

    tray = new Tray(createTrayIcon());
    tray.setToolTip('ZORG NEXUS | Desktop Widget');

    const updateContextMenu = () => {
        const contextMenu = Menu.buildFromTemplate([
            {
                label: 'ZORG NEXUS Desktop Widget',
                enabled: false
            },
            { type: 'separator' },
            {
                label: 'Show / Hide Widget',
                click: () => {
                    if (mainWindow.isVisible()) {
                        mainWindow.hide();
                    } else {
                        mainWindow.show();
                        mainWindow.focus();
                    }
                }
            },
            {
                label: 'Always On Top',
                type: 'checkbox',
                checked: isAlwaysOnTop,
                click: (item) => {
                    isAlwaysOnTop = item.checked;
                    if (mainWindow) mainWindow.setAlwaysOnTop(isAlwaysOnTop);
                }
            },
            { type: 'separator' },
            {
                label: '🌐 Open Web Portal',
                click: () => {
                    shell.openExternal('http://localhost:3005');
                }
            },
            { type: 'separator' },
            {
                label: '❌ Quit Widget',
                click: () => {
                    app.isQuitting = true;
                    app.quit();
                }
            }
        ]);
        tray.setContextMenu(contextMenu);
    };

    updateContextMenu();

    tray.on('double-click', () => {
        if (mainWindow) {
            mainWindow.show();
            mainWindow.focus();
        }
    });
}

function createTrayIcon() {
    const { nativeImage } = require('electron');
    const svg = `
    <svg xmlns="http://www.w3.org/2000/svg" width="16" height="16" viewBox="0 0 16 16">
        <rect width="16" height="16" rx="4" fill="#1e293b"/>
        <path d="M4 8L8 4L12 8M8 4V12" stroke="#3b82f6" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"/>
    </svg>`;
    return nativeImage.createFromBuffer(Buffer.from(svg));
}

// IPC Handlers
ipcMain.handle('widget-close', () => {
    if (mainWindow) mainWindow.hide();
});

ipcMain.handle('widget-minimize', () => {
    if (mainWindow) mainWindow.minimize();
});

ipcMain.handle('widget-open-portal', () => {
    shell.openExternal('http://localhost:3005');
});

ipcMain.handle('widget-toggle-ontop', () => {
    isAlwaysOnTop = !isAlwaysOnTop;
    if (mainWindow) mainWindow.setAlwaysOnTop(isAlwaysOnTop);
    return isAlwaysOnTop;
});

app.whenReady().then(createWindow);

app.on('window-all-closed', () => {
    if (process.platform !== 'darwin') {
        app.quit();
    }
});

app.on('activate', () => {
    if (mainWindow === null) {
        createWindow();
    }
});
