global.Office = {
  onReady: jest.fn(),
  HostType: { Outlook: 'Outlook' },
  Context: {
    mailbox: {
      item: null,
      addHandlerAsync: jest.fn()
    },
    officeTheme: null
  },
  AsyncResultStatus: { Succeeded: 'succeeded' },
  CoercionType: { Text: 'text' },
  EventType: { ItemChanged: 'itemChanged' }
};

global.WebSocket = {
  OPEN: 1,
  CONNECTING: 0,
  CLOSED: 3
};

global.crypto = {
  randomUUID: jest.fn(() => Math.random().toString(36).substring(2))
};

const localStorageData = {};
const mockLocalStorage = {
  getItem: jest.fn((key) => localStorageData[key] || null),
  setItem: jest.fn((key, value) => { localStorageData[key] = value; }),
  removeItem: jest.fn((key) => { delete localStorageData[key]; }),
  clear: jest.fn(() => { for (let k in localStorageData) delete localStorageData[k]; }),
  get data() { return localStorageData; },
  set data(v) { for (let k in v) localStorageData[k] = v[k]; }
};
Object.defineProperty(global, 'localStorage', {
  value: mockLocalStorage,
  writable: true,
  configurable: true
});
