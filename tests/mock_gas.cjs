'use strict';
class MockSheet {
  constructor(name, rows) {
    this.name = name;
    this.rows = rows.map((r) => r.slice());
  }
  getName() { return this.name; }
  getLastRow() { return this.rows.length; }
  getLastColumn() { return this.rows.length ? this.rows[0].length : 0; }
  getDataRange() {
    const self = this;
    return { getValues: () => self.rows.map((r) => r.slice()) };
  }
  getRange(row, col, numRows, numCols) {
    const self = this;
    if (numRows === undefined) numRows = 1;
    if (numCols === undefined) numCols = 1;
    return {
      getValues() {
        const out = [];
        for (let i = 0; i < numRows; i++) {
          const r = self.rows[row - 1 + i] || [];
          out.push(r.slice(col - 1, col - 1 + numCols));
        }
        return out;
      },
      setValues(vals) {
        for (let i = 0; i < numRows; i++) {
          const rIdx = row - 1 + i;
          while (self.rows.length <= rIdx) self.rows.push([]);
          for (let j = 0; j < numCols; j++) {
            self.rows[rIdx][col - 1 + j] = vals[i][j];
          }
        }
      },
      setValue(val) {
        const rIdx = row - 1;
        while (self.rows.length <= rIdx) self.rows.push([]);
        self.rows[rIdx][col - 1] = val;
      },
      setFontWeight() { return this; },
      setBackground() { return this; },
    };
  }
  setFrozenRows() {}
  appendRow(arr) { this.rows.push(arr.slice()); }
}

class MockSpreadsheet {
  constructor(id) { this.id = id; this.sheets = {}; }
  getSheetByName(name) { return this.sheets[name] || null; }
  insertSheet(name) { const sh = new MockSheet(name, []); this.sheets[name] = sh; return sh; }
  _addSheet(name, rows) { this.sheets[name] = new MockSheet(name, rows); return this.sheets[name]; }
}

function buildMockEnv() {
  const ownSS = new MockSpreadsheet('OWN');
  const extSS = {};
  global.SpreadsheetApp = {
    getActiveSpreadsheet: () => ownSS,
    openById: (id) => { if (!extSS[id]) extSS[id] = new MockSpreadsheet(id); return extSS[id]; },
  };

  const folders = {};
  class MockFolder {
    constructor(id, name) { this.id = id; this.name = name; this.files = []; }
    getName() { return this.name; }
    getUrl() { return `https://drive.google.com/drive/folders/${this.id}`; }
    getId() { return this.id; }
    getFiles() {
      let i = 0; const files = this.files;
      return { hasNext: () => i < files.length, next: () => files[i++] };
    }
    createFolder(name) {
      const id = 'folder_' + Math.random().toString(36).slice(2);
      const f = new MockFolder(id, name); folders[id] = f; return f;
    }
    createFile(blob) {
      const id = 'file_' + Math.random().toString(36).slice(2);
      const file = { id, name: blob.name, bytes: blob.bytes, mime: blob.mime, getUrl: () => `https://drive.google.com/file/d/${id}/view`, getName: () => blob.name, getId: () => id, setSharing: () => {} };
      this.files.push(file); return file;
    }
  }
  folders['ROOT_MAU_WORD'] = new MockFolder('ROOT_MAU_WORD', 'Mau_Word');
  folders['ROOT_MAU_WORD'].files.push(
    { getName: () => 'BM-FSC-01.docx', getUrl: () => 'https://drive/x1', getId: () => 'x1' },
    { getName: () => 'BM-FSC-02.docx', getUrl: () => 'https://drive/x2', getId: () => 'x2' },
  );

  global.DriveApp = {
    getFolderById: (id) => { if (!folders[id]) throw new Error('Không tìm thấy thư mục Drive id=' + id); return folders[id]; },
    Access: { ANYONE_WITH_LINK: 'ANYONE_WITH_LINK' },
    Permission: { VIEW: 'VIEW' },
  };

  global.Utilities = {
    formatDate: (date, tz, fmt) => {
      const d = new Date(date);
      const pad = (n) => String(n).padStart(2, '0');
      return `${d.getFullYear()}-${pad(d.getMonth() + 1)}-${pad(d.getDate())} ${pad(d.getHours())}:${pad(d.getMinutes())}`;
    },
    base64Decode: (str) => {
      if (!/^[A-Za-z0-9+/]*={0,2}$/.test(str)) {
        throw new Error('Đối số không hợp lệ: ký tự ngoài bảng mã base64');
      }
      return Buffer.from(str, 'base64');
    },
    newBlob: (bytes, mime, name) => ({ bytes, mime, name }),
  };

  global.Session = { getScriptTimeZone: () => 'Asia/Ho_Chi_Minh' };

  const scriptProps = {};
  global.PropertiesService = {
    getScriptProperties: () => ({
      getProperty: (k) => (k in scriptProps ? scriptProps[k] : null),
      setProperty: (k, v) => { scriptProps[k] = v; },
    }),
  };

  global.HtmlService = {
    createHtmlOutputFromFile: () => ({ setTitle() { return this; }, addMetaTag() { return this; } }),
  };

  global.Logger = { log: () => {} };

  global.LockService = {
    getScriptLock: () => ({ waitLock: () => {}, releaseLock: () => {} }),
  };

  const cacheStore = {};
  global.CacheService = {
    getScriptCache: () => ({
      get: (k) => (k in cacheStore ? cacheStore[k] : null),
      put: (k, v) => { cacheStore[k] = v; },
      removeAll: (keys) => { keys.forEach((k) => { delete cacheStore[k]; }); },
    }),
  };

  return { ownSS, extSS, MockFolder, cacheStore };
}

module.exports = { buildMockEnv, MockSheet, MockSpreadsheet };
