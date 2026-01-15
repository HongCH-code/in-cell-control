const { app, BrowserWindow, ipcMain } = require('electron');
const path = require('path');
const { SerialPort } = require('serialport');

let mainWindow;
let serialPort = null;

function createWindow() {
  mainWindow = new BrowserWindow({
    width: 1024,
    height: 800,
    minWidth: 800,
    minHeight: 600,
    title: 'In Cell Parameter Setting',
    webPreferences: {
      nodeIntegration: false,
      contextIsolation: true,
      preload: path.join(__dirname, 'preload.js')
    }
  });

  // Load the index.html file
  mainWindow.loadFile('index.html');

  // Hide menu bar (optional - remove this line if you want the menu)
  mainWindow.setMenuBarVisibility(false);

  mainWindow.on('closed', () => {
    // Close serial port when window closes
    if (serialPort && serialPort.isOpen) {
      serialPort.close();
    }
    mainWindow = null;
  });
}

// Serial Port IPC Handlers

// List available serial ports
ipcMain.handle('serial:list', async () => {
  try {
    const ports = await SerialPort.list();
    return { success: true, ports };
  } catch (error) {
    return { success: false, error: error.message };
  }
});

// Connect to serial port
ipcMain.handle('serial:connect', async (event, portPath, baudRate) => {
  try {
    // Close existing connection if any
    if (serialPort && serialPort.isOpen) {
      await new Promise((resolve) => serialPort.close(resolve));
    }

    serialPort = new SerialPort({
      path: portPath,
      baudRate: parseInt(baudRate),
      autoOpen: false
    });

    // Open the port
    await new Promise((resolve, reject) => {
      serialPort.open((err) => {
        if (err) reject(err);
        else resolve();
      });
    });

    // Listen for data
    serialPort.on('data', (data) => {
      if (mainWindow) {
        mainWindow.webContents.send('serial:data', data.toString());
      }
    });

    // Listen for errors
    serialPort.on('error', (err) => {
      if (mainWindow) {
        mainWindow.webContents.send('serial:error', err.message);
      }
    });

    // Listen for close
    serialPort.on('close', () => {
      if (mainWindow) {
        mainWindow.webContents.send('serial:status', { connected: false });
      }
    });

    mainWindow.webContents.send('serial:status', { connected: true, port: portPath, baudRate });
    return { success: true };
  } catch (error) {
    return { success: false, error: error.message };
  }
});

// Disconnect from serial port
ipcMain.handle('serial:disconnect', async () => {
  try {
    if (serialPort && serialPort.isOpen) {
      await new Promise((resolve) => serialPort.close(resolve));
      serialPort = null;
    }
    return { success: true };
  } catch (error) {
    return { success: false, error: error.message };
  }
});

// Send data through serial port
ipcMain.handle('serial:send', async (event, data) => {
  try {
    if (!serialPort || !serialPort.isOpen) {
      return { success: false, error: 'Serial port is not connected' };
    }

    await new Promise((resolve, reject) => {
      serialPort.write(data, (err) => {
        if (err) reject(err);
        else resolve();
      });
    });

    // Drain to ensure all data is sent
    await new Promise((resolve) => serialPort.drain(resolve));

    return { success: true };
  } catch (error) {
    return { success: false, error: error.message };
  }
});

// When Electron is ready, create window
app.whenReady().then(createWindow);

// Quit when all windows are closed (except on macOS)
app.on('window-all-closed', () => {
  if (process.platform !== 'darwin') {
    app.quit();
  }
});

// On macOS, re-create window when dock icon is clicked
app.on('activate', () => {
  if (BrowserWindow.getAllWindows().length === 0) {
    createWindow();
  }
});

// Clean up on app quit
app.on('before-quit', () => {
  if (serialPort && serialPort.isOpen) {
    serialPort.close();
  }
});
