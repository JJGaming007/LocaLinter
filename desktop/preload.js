'use strict';

/**
 * The one thing the renderer cannot do for itself.
 *
 * The window is frameless with Windows' own buttons painted over our title bar,
 * and that strip's colours are owned by the main process. With a light and a
 * dark theme the strip has to follow the theme, so the renderer needs a way to
 * say which one is active. Nothing else crosses this bridge.
 */
const { contextBridge, ipcRenderer } = require('electron');

contextBridge.exposeInMainWorld('localinterShell', {
  setTitleBarTheme: (colors) => ipcRenderer.send('titlebar-theme', colors),
});
