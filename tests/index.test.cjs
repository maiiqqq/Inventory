const test = require('node:test');
const assert = require('node:assert/strict');
const fs = require('node:fs');
const path = require('node:path');
const vm = require('node:vm');

const root = path.resolve(__dirname, '..');

function fixture(ledgerRows, displayRows, cacheRows = []) {
  const sheets = new Map();
  class Sheet {
    constructor(name, rows, display) {
      this.name = name;
      this.rows = rows;
      this.display = display;
    }
    getName() { return this.name; }
    getLastRow() { return this.rows.length; }
    getMaxColumns() { return 26; }
    getDataRange() {
      return {
        getValues: () => this.rows.map(row => [...row]),
        getDisplayValues: () => (this.display || this.rows).map(row => [...row])
      };
    }
    getRange(row, col, height, width) {
      return {
        setValues: values => {
          assert.equal(values.length, height);
          values.forEach((v, i) => {
            assert.equal(v.length, width);
            const r = this.rows[row + i - 1] || [];
            v.forEach((cell, j) => { r[col + j - 1] = cell; });
            this.rows[row + i - 1] = r;
          });
          return { setFontWeight() {} };
        }
      };
    }
    appendRow(row) { this.rows.push([...row]); }
    deleteRow(row) { this.rows.splice(row - 1, 1); }
    clearContents() { this.rows = []; }
    showColumns() {}
    hideColumns() {}
    setFrozenRows() {}
  }
  const add = (name, rows, display) => {
    const sheet = new Sheet(name, rows, display);
    sheets.set(name, sheet);
    return sheet;
  };
  const ledger = add('บัญชีคุม', [
    ['สิ่งอุปกรณ์', 'หมายเลข', 'ปีงบประมาณ', 'โครงการ', 'หน่วย'],
    ...ledgerRows
  ], [
    ['สิ่งอุปกรณ์', 'หมายเลข', 'ปีงบประมาณ', 'โครงการ', 'หน่วย'],
    ...displayRows
  ]);
  add('สถานะโครงการ', [
    ['หน่วย', 'โครงการ', 'สถานะ', 'วงเงิน'],
    ['หน่วย ก', 'โครงการ A', 'อนุมัติ', 1000]
  ]);
  const cache = add('_Cache', [
    ['fileId', 'modifiedTime', 'fileName', 'url', 'folder', 'json'],
    ...cacheRows
  ]);
  const ss = {
    getSheetByName: name => sheets.get(name) || null,
    getSheets: () => [...sheets.values()],
    insertSheet: name => add(name, []),
    getName: () => 'fixture'
  };
  const ctx = vm.createContext({
    SpreadsheetApp: { openById: () => ss },
    PropertiesService: { getScriptProperties: () => ({ getProperty: () => 'fixture-key' }) },
    DriveApp: { getFolderById: () => ({
      getName: () => 'fixture',
      getFilesByType: () => ({ hasNext: () => false }),
      getFolders: () => ({ hasNext: () => false })
    }) },
    MimeType: { PDF: 'application/pdf' },
    Logger: { log() {} }
  });
  for (const file of ['Config.gs', 'Code.gs', 'Index.gs']) {
    vm.runInContext(fs.readFileSync(path.join(root, file), 'utf8'), ctx, { filename: file });
  }
  return { ctx, ledger, cache, sheets };
}

function item(name, number, overrides = {}) {
  return {
    number, name, unit: 'ชุด', owner_unit: 'หน่วย ก',
    project: 'โครงการ A', category: 'กล้อง', subtags: ['CCTV'],
    summary: 'กล้องสำหรับทดสอบ', rawText: name,
    ...overrides
  };
}

function entry(fileName, items, id = fileName) {
  return {
    [id]: {
      fileName, url: 'https://example.invalid/' + fileName, folder: 'fixture',
      items
    }
  };
}

test('file number range wins over a conflicting OCR number', () => {
  const { ctx } = fixture(
    [
      ['กล้องวงจรปิด', 1, 69, 'โครงการ A', 'หน่วย ก'],
      ['กล้องวงจรปิด', 2, 69, 'โครงการ A', 'หน่วย ก']
    ],
    [
      ['กล้องวงจรปิด', 'ส-001', '69', 'โครงการ A', 'หน่วย ก'],
      ['กล้องวงจรปิด', 'ส-002', '69', 'โครงการ A', 'หน่วย ก']
    ]
  );
  ctx.rebuildIndexSheet_(entry('2.pdf', [item('กล้องวงจรปิด', 1)]));
  const result = ctx.apiSearch({});
  assert.equal(result.total, 1);
  assert.equal(result.items[0].code, 'ส-002/69');
  assert.equal(result.items[0].number, 2);
  assert.equal(result.items[0].docNumber, 1);
  assert.equal(ctx.apiSearch({ number: 1 }).total, 0);
  assert.equal(ctx.apiSearch({ number: 2 }).total, 1);
});

test('unknown file range and ambiguous identical names leave the number unverified', () => {
  const data = [
    ['กล้องวงจรปิด', 1, 69, 'โครงการ A', 'หน่วย ก'],
    ['กล้องวงจรปิด', 2, 69, 'โครงการ A', 'หน่วย ก']
  ];
  const display = data.map((r, i) => [r[0], 'ส-00' + (i + 1), '69', r[3], r[4]]);
  const { ctx } = fixture(data, display);
  ctx.rebuildIndexSheet_(entry('99.pdf', [item('กล้องวงจรปิด', 99)]));
  assert.equal(ctx.apiSearch({}).items[0].code, '');
  assert.equal(ctx.apiSearch({ number: 99 }).total, 0);
  ctx.rebuildIndexSheet_(entry('unknown.pdf', [item('กล้องวงจรปิด', null)]));
  assert.equal(ctx.apiSearch({}).items[0].code, '');
});

test('a loose name match without a file range does not claim an official number', () => {
  const { ctx } = fixture(
    [['กล้องวงจรปิด', 1, 69, 'โครงการ A', 'หน่วย ก']],
    [['กล้องวงจรปิด', 'ส-001', '69', 'โครงการ A', 'หน่วย ก']]
  );
  ctx.rebuildIndexSheet_(entry('unknown.pdf', [item('กล้อง', 1)]));
  assert.equal(ctx.apiSearch({}).items[0].code, '');
});

test('duplicate PDF extracts collapse to one official item', () => {
  const { ctx } = fixture(
    [['กล้องวงจรปิด', 1, 69, 'โครงการ A', 'หน่วย ก']],
    [['กล้องวงจรปิด', 'ส-001', '69', 'โครงการ A', 'หน่วย ก']]
  );
  ctx.rebuildIndexSheet_({
    ...entry('1.pdf', [item('กล้องวงจรปิด', 1), item('กล้องวงจรปิด', 1)]),
    ...entry('1-2.pdf', [item('กล้องวงจรปิด', 1, { summary: 'รายละเอียดเพิ่มเติม' })])
  });
  assert.equal(ctx.apiSearch({}).total, 1);
  assert.equal(ctx.apiStats().total, 1);
});

test('formatted text numbers parse and do not repeat the year suffix', () => {
  const { ctx } = fixture(
    [['กล้องวงจรปิด', 'ส-002/69', 69, 'โครงการ A', 'หน่วย ก']],
    [['กล้องวงจรปิด', 'ส-002/69', '69', 'โครงการ A', 'หน่วย ก']]
  );
  ctx.rebuildIndexSheet_(entry('2.pdf', [item('กล้องวงจรปิด', 2)]));
  assert.equal(ctx.apiSearch({}).items[0].code, 'ส-002/69');
  assert.equal(ctx.apiSearch({ number: 2 }).total, 1);
});

test('files removed from Drive are removed from cache and index on rebuild', () => {
  const { ctx, cache } = fixture(
    [['กล้องวงจรปิด', 1, 69, 'โครงการ A', 'หน่วย ก']],
    [['กล้องวงจรปิด', 'ส-001', '69', 'โครงการ A', 'หน่วย ก']],
    [['old-file', '2026-01-01', '1.pdf', 'https://example.invalid/1.pdf', 'old', JSON.stringify([item('กล้องวงจรปิด', 1)])]]
  );
  ctx.buildIndex();
  assert.equal(cache.getLastRow(), 1);
  assert.equal(ctx.apiSearch({}).total, 0);
});

test('older duplicate cache rows are removed before reading the latest file data', () => {
  const { ctx, cache } = fixture(
    [['กล้องวงจรปิด', 1, 69, 'โครงการ A', 'หน่วย ก']],
    [['กล้องวงจรปิด', 'ส-001', '69', 'โครงการ A', 'หน่วย ก']],
    [
      ['same-file', 'old', '1.pdf', 'https://example.invalid/1.pdf', 'old', JSON.stringify([item('กล้องวงจรปิด', 1)])],
      ['same-file', 'new', '1.pdf', 'https://example.invalid/1.pdf', 'new', JSON.stringify([item('กล้องวงจรปิด', 1, { summary: 'ข้อมูลใหม่' })])]
    ]
  );
  const loaded = ctx.loadCache_();
  assert.equal(cache.getLastRow(), 2);
  assert.equal(loaded.map['same-file'].modifiedTime, 'new');
  assert.equal(loaded.map['same-file'].rowIndex, 2);
});

test('sorting happens before the 300-result limit', () => {
  const { ctx, ledger, sheets } = fixture([], []);
  const index = new ledger.constructor('Index_อุปกรณ์', [['code', 'ocr', 'name']]);
  for (let n = 1; n <= 301; n++) {
    const row = Array(16).fill('');
    row[0] = String(n);
    row[2] = n === 301 ? 'A first' : 'Z ' + String(n).padStart(3, '0');
    index.rows.push(row);
  }
  sheets.set('Index_อุปกรณ์', index);
  const result = ctx.apiSearch({ sort: 'name' });
  assert.equal(result.total, 301);
  assert.equal(result.items.length, 300);
  assert.equal(result.items[0].name, 'A first');
});

test('the same owner unit alone cannot assign another project status', () => {
  const { ctx } = fixture([], []);
  const control = ctx.readControlSheet_();
  assert.equal(control.matchProject('โครงการ B', 'หน่วย ก').row, undefined);
  assert.equal(control.matchProject('โครงการ A', 'หน่วย ก').status, 'อนุมัติ');
});
