const { app, BrowserWindow, ipcMain, dialog } = require('electron');
const path = require('path');
const fs = require('fs');
const { SerialPort } = require('serialport');
const XLSX = require('xlsx');

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
ipcMain.handle('serial:connect', async (event, portPath, baudRate, dataBits = 8, parity = 'none', stopBits = 1) => {
  try {
    // Close existing connection if any
    if (serialPort && serialPort.isOpen) {
      await new Promise((resolve) => serialPort.close(resolve));
    }

    serialPort = new SerialPort({
      path: portPath,
      baudRate: parseInt(baudRate),
      dataBits: parseInt(dataBits),
      parity: parity,
      stopBits: parseInt(stopBits),
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

// ============================================
// File Import/Export IPC Handlers
// ============================================

// Import from Excel file (.xlsx, .xls)
ipcMain.handle('file:import', async () => {
  try {
    const result = await dialog.showOpenDialog(mainWindow, {
      title: 'Import Settings from Excel',
      filters: [
        { name: 'Excel Files', extensions: ['xlsx', 'xls'] },
        { name: 'All Files', extensions: ['*'] }
      ],
      properties: ['openFile']
    });

    if (result.canceled || result.filePaths.length === 0) {
      return { success: false, canceled: true };
    }

    const filePath = result.filePaths[0];
    const workbook = XLSX.readFile(filePath);
    const sheet = workbook.Sheets[workbook.SheetNames[0]];
    const data = XLSX.utils.sheet_to_json(sheet, { header: 1 });

    // Parse data: B column (index 1) = parameter name, C column (index 2) = value
    const parameters = {};
    for (const row of data) {
      if (row && row[1]) {
        const paramName = String(row[1]).trim();
        const value = row[2] !== undefined && row[2] !== null ? row[2] : '';
        if (paramName && paramName !== 'Parameter') { // Skip header row
          parameters[paramName] = value;
        }
      }
    }

    return { success: true, parameters, filePath };
  } catch (error) {
    return { success: false, error: error.message };
  }
});

// Export to Excel file (.xlsx) with same structure as import template
ipcMain.handle('file:export', async (event, data) => {
  try {
    const { parameters, modelName } = data;

    // Generate filename with timestamp
    const now = new Date();
    const timestamp = now.getFullYear().toString() +
      String(now.getMonth() + 1).padStart(2, '0') +
      String(now.getDate()).padStart(2, '0') + '_' +
      String(now.getHours()).padStart(2, '0') +
      String(now.getMinutes()).padStart(2, '0') +
      String(now.getSeconds()).padStart(2, '0');
    const defaultName = `${modelName || 'Settings'}_${timestamp}.xlsx`;

    const result = await dialog.showSaveDialog(mainWindow, {
      title: 'Export Settings to Excel',
      defaultPath: defaultName,
      filters: [
        { name: 'Excel Files', extensions: ['xlsx'] },
        { name: 'All Files', extensions: ['*'] }
      ]
    });

    if (result.canceled || !result.filePath) {
      return { success: false, canceled: true };
    }

    // Define parameter structure with categories and ranges
    const parameterStructure = [
      { category: '', param: 'Model', range: '' },
      { category: 'AC Timing', param: 'T0', range: '0~334' },
      { category: '', param: 'T1', range: '0~334' },
      { category: '', param: 'T2', range: '0~334' },
      { category: '', param: 'T3', range: '0~334' },
      { category: '', param: 'T4', range: '0~334' },
      { category: '', param: 'T5', range: '0~334' },
      { category: '', param: 'TSW1b', range: '0~334' },
      { category: '', param: 'TSW2b', range: '0~334' },
      { category: '', param: 'TSW3b', range: '0~334' },
      { category: '', param: 'TSW3a', range: '0~334' },
      { category: '', param: 'TSW2a', range: '0~334' },
      { category: '', param: 'TSW1a', range: '0~334' },
      { category: 'DC Electrial', param: 'VD', range: '7~30, step 0.5' },
      { category: '', param: 'VG', range: '-18~-6, step 0.25' },
      { category: 'Programble Gain Setting', param: 'CS0_Wiper0', range: '0~255' },
      { category: '', param: 'CS0_Wiper1', range: '0~255' },
      { category: '', param: 'CS0_Wiper2', range: '0~255' },
      { category: '', param: 'CS0_Wiper3', range: '0~255' },
      { category: '', param: 'CS1_Wiper0', range: '0~255' },
      { category: '', param: 'CS1_Wiper1', range: '0~255' },
      { category: '', param: 'CS1_Wiper2', range: '0~255' },
      { category: '', param: 'CS1_Wiper3', range: '0~255' },
      { category: '', param: 'CS2_Wiper0', range: '0~255' },
      { category: '', param: 'CS2_Wiper1', range: '0~255' },
      { category: '', param: 'CS2_Wiper2', range: '0~255' },
      { category: '', param: 'CS2_Wiper3', range: '0~255' },
      { category: '', param: 'CS3_Wiper0', range: '0~255' },
      { category: '', param: 'CS3_Wiper1', range: '0~255' },
      { category: '', param: 'CS3_Wiper2', range: '0~255' },
      { category: '', param: 'CS3_Wiper3', range: '0~255' },
      { category: '', param: 'CS4_Wiper0', range: '0~255' },
      { category: '', param: 'CS4_Wiper1', range: '0~255' },
      { category: '', param: 'CS4_Wiper2', range: '0~255' },
      { category: '', param: 'CS4_Wiper3', range: '0~255' },
      { category: '', param: 'CS5_Wiper0', range: '0~255' },
      { category: '', param: 'CS5_Wiper1', range: '0~255' },
      { category: '', param: 'CS5_Wiper2', range: '0~255' },
      { category: '', param: 'CS5_Wiper3', range: '0~255' }
    ];

    // Build worksheet data: [Category, Parameter, Value, Range]
    const wsData = [
      ['Parameter', null, 'Value', 'Range'] // Header row
    ];

    for (const item of parameterStructure) {
      const value = parameters[item.param] !== undefined ? parameters[item.param] : '';
      wsData.push([
        item.category || null,
        item.param,
        value === '' ? null : value,
        item.range
      ]);
    }

    // Create workbook and worksheet
    const ws = XLSX.utils.aoa_to_sheet(wsData);
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, '工作表1');

    // Write file
    XLSX.writeFile(wb, result.filePath);
    return { success: true, filePath: result.filePath };
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
