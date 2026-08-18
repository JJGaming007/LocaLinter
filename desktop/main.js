'use strict';

/**
 * LocaLinter desktop — Electron main process.
 *
 * The web build was two things: a website, and a local agent the website
 * called over HTTP because a page cannot open a USB cable. Here they are one
 * process. The agent's server still runs, but on loopback inside this app, and
 * it also serves the UI — so the window loads an ordinary http:// origin
 * rather than file://.
 *
 * That one decision buys a lot: every /api call is same-origin (no CORS, no
 * private-network preflight, no agent-URL box, no Reconnect), and Google
 * sign-in still sees an origin it is willing to work with.
 */

const path = require('path');
const http = require('http');
const { app, BrowserWindow, shell, dialog, ipcMain } = require('electron');

const AGENT_PORT = Number(process.env.PORT || 8790);
const AGENT_ORIGIN = `http://127.0.0.1:${AGENT_PORT}`;

// Where the UI lives. Packaged, electron-builder puts the repo alongside the
// asar as unpacked resources; from a checkout it is simply the parent folder.
const UI_DIR = app.isPackaged
  ? path.join(process.resourcesPath, 'app')
  : path.join(__dirname, '..');

// Packaged, this file lives inside the asar while the agent is unpacked beside
// it, so a relative require would not resolve. Both paths are derived from
// UI_DIR instead.
const AGENT_ENTRY = path.join(UI_DIR, 'agent', 'server.js');

/**
 * The window icon, and — via setAppUserModelId below — the taskbar one.
 *
 * Windows takes the taskbar icon from the window's own icon, but only groups it
 * under this app (rather than under "Electron") once the App User Model ID is
 * set, so the two have to be done together.
 */
const APP_ID = 'com.supergaming.localinter';
const ICON_PATH = path.join(UI_DIR, 'assets', 'icon.ico');

/**
 * Config, route maps and run output must land somewhere writable. Inside a
 * packaged app the install folder is not, so point the agent's existing
 * override at the per-user data directory before it is loaded.
 */
process.env.LOCALINTER_DATA_DIR = path.join(app.getPath('userData'), 'agent');
process.env.LOCALINTER_UI_DIR = UI_DIR;

let mainWindow = null;

/** Poll rather than reach into the agent: it owns its own listen(). */
function waitForAgent(timeoutMs = 15000) {
  const started = Date.now();
  return new Promise((resolve, reject) => {
    const attempt = () => {
      const req = http.get(`${AGENT_ORIGIN}/api/health`, (res) => {
        res.resume();
        if (res.statusCode === 200) return resolve();
        retry();
      });
      req.on('error', retry);
      req.setTimeout(1000, () => { req.destroy(); });
    };
    const retry = () => {
      if (Date.now() - started > timeoutMs) {
        return reject(new Error(`The agent did not come up on ${AGENT_ORIGIN} within ${timeoutMs / 1000}s.`));
      }
      setTimeout(attempt, 200);
    };
    attempt();
  });
}

function createWindow() {
  mainWindow = new BrowserWindow({
    width: 1100,
    height: 720,
    minWidth: 860,
    minHeight: 580,
    backgroundColor: '#f6f5f1',        // matches --bg-base, so no flash on open
    show: false,
    autoHideMenuBar: true,
    title: 'LocaLinter',
    icon: ICON_PATH,
    // Frameless with the system's own buttons painted over our title bar: the
    // app owns the whole surface, but minimise/maximise/close still behave and
    // look exactly like every other Windows app.
    //
    // Deliberately no `frame: false` alongside this. On Windows the overlay
    // *is* the frameless mode, and asking for both suppresses the overlay —
    // taking the window buttons with it, and leaving a window that never
    // reaches ready-to-show, so it stays hidden and the app looks dead.
    titleBarStyle: 'hidden',
    // Light is the default theme, so the strip starts light; the renderer
    // re-sends the real colours as soon as it knows which theme is stored.
    titleBarOverlay: process.platform === 'win32' ? {
      color: '#ffffff',                // --toolbar-bg
      symbolColor: '#6b6a64',          // --text-muted
      height: 44,
    } : false,
    webPreferences: {
      contextIsolation: true,
      nodeIntegration: false,          // the UI is still just a web app
      sandbox: true,
      preload: path.join(__dirname, 'preload.js'),
    },
  });

  // show:false avoids a white flash, but it means the window is invisible
  // until something says otherwise — so take whichever signal arrives first
  // rather than trusting a single event to fire.
  const reveal = () => {
    if (mainWindow && !mainWindow.isVisible()) mainWindow.show();
  };
  mainWindow.once('ready-to-show', reveal);
  mainWindow.webContents.once('did-finish-load', reveal);
  mainWindow.on('closed', () => { mainWindow = null; });

  // A window created with show:false that never reaches ready-to-show is
  // invisible *and* silent. Say why, and show it anyway rather than leaving
  // the app running with no way to see it.
  mainWindow.webContents.on('did-fail-load', (_e, code, desc, url) => {
    console.error(`[window] failed to load ${url}: ${desc} (${code})`);
  });
  mainWindow.webContents.on('render-process-gone', (_e, details) => {
    console.error(`[window] renderer gone: ${details.reason}`);
  });
  mainWindow.webContents.on('console-message', (_e, level, message, line, sourceId) => {
    if (level >= 2) console.error(`[renderer] ${message} (${sourceId}:${line})`);
  });
  // Last resort. If both signals somehow fail, an invisible window with no
  // explanation is the worst outcome, so force it and say what happened.
  setTimeout(() => {
    if (mainWindow && !mainWindow.isVisible()) {
      const b = mainWindow.getBounds();
      console.error(`[window] neither ready-to-show nor did-finish-load fired; forcing it visible at ${b.x},${b.y} ${b.width}x${b.height}`);
      mainWindow.show();
    }
  }, 6000);

  // Anything that is not our own origin belongs in the real browser — Google
  // sign-in especially, which refuses to run inside an embedded webview.
  mainWindow.webContents.setWindowOpenHandler(({ url }) => {
    shell.openExternal(url);
    return { action: 'deny' };
  });
  mainWindow.webContents.on('will-navigate', (e, url) => {
    if (!url.startsWith(AGENT_ORIGIN)) {
      e.preventDefault();
      shell.openExternal(url);
    }
  });

  mainWindow.loadURL(AGENT_ORIGIN + '/');
}

// One device, one agent port, one window.
if (!app.requestSingleInstanceLock()) {
  app.quit();
} else {
  app.on('second-instance', () => {
    if (mainWindow) {
      if (mainWindow.isMinimized()) mainWindow.restore();
      mainWindow.focus();
    }
  });

  // Must be set before the first window is created, or Windows keeps the
  // default "Electron" identity and shows its icon on the taskbar.
  if (process.platform === 'win32') app.setAppUserModelId(APP_ID);

  app.whenReady().then(async () => {
    try {
      require(AGENT_ENTRY);                // starts listening on AGENT_PORT
      await waitForAgent();
      createWindow();
    } catch (e) {
      dialog.showErrorBox('LocaLinter could not start', `${e.message}\n\nIs something else already using port ${AGENT_PORT}?`);
      app.quit();
    }
  });

  // The renderer owns the theme; the title-bar strip belongs to the window.
  ipcMain.on('titlebar-theme', (_e, colors) => {
    if (!mainWindow || process.platform !== 'win32' || !colors) return;
    try {
      mainWindow.setTitleBarOverlay({
        color: String(colors.color),
        symbolColor: String(colors.symbolColor),
        height: 44,
      });
    } catch { /* the window may be closing */ }
  });

  app.on('window-all-closed', () => app.quit());
}
