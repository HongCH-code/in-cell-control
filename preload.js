const { contextBridge, ipcRenderer } = require('electron');

// Expose serial port APIs to renderer process
contextBridge.exposeInMainWorld('serialAPI', {
  // Get list of available serial ports
  listPorts: () => ipcRenderer.invoke('serial:list'),

  // Connect to a serial port
  connect: (path, baudRate, dataBits = 8, parity = 'none', stopBits = 1) =>
    ipcRenderer.invoke('serial:connect', path, baudRate, dataBits, parity, stopBits),

  // Disconnect from serial port
  disconnect: () => ipcRenderer.invoke('serial:disconnect'),

  // Send data through serial port
  send: (data) => ipcRenderer.invoke('serial:send', data),

  // Listen for connection status changes
  onStatusChange: (callback) => {
    ipcRenderer.on('serial:status', (event, status) => callback(status));
  },

  // Listen for received data
  onData: (callback) => {
    ipcRenderer.on('serial:data', (event, data) => callback(data));
  },

  // Listen for errors
  onError: (callback) => {
    ipcRenderer.on('serial:error', (event, error) => callback(error));
  }
});
