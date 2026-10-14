/**
 * Shared test helpers: an in-memory stand-in for the parts of the Google
 * Sheets API the analyzer uses, plus source-extraction and check helpers.
 * Not a test itself, so it lives under tests/lib/.
 */
const fs = require('fs');
const path = require('path');

const root = path.join(__dirname, '..', '..');
const read = f => fs.readFileSync(path.join(root, f), 'utf8').replace(/^﻿/, '');

function extractFunction(source, name) {
  const startMatch = source.match(new RegExp('function\\s+' + name + '\\s*\\('));
  if (!startMatch) throw new Error('function ' + name + ' not found');
  const i = startMatch.index;
  const braceStart = source.indexOf('{', i);
  let depth = 0, end = -1;
  for (let j = braceStart; j < source.length; j++) {
    if (source[j] === '{') depth++;
    else if (source[j] === '}') { depth--; if (depth === 0) { end = j + 1; break; } }
  }
  if (end === -1) throw new Error('could not find end of function ' + name);
  return source.slice(i, end);
}

// A sheet backed by an array of rows. Every read is logged so tests can assert
// which columns a function touches (e.g. that it never pulls the HTML column).
class FakeSheet {
  constructor(headers, rows) {
    this.grid = [headers.slice()].concat((rows || []).map(r => r.slice()));
    this.reads = [];
  }
  getLastRow() { return this.grid.length; }
  getLastColumn() { return Math.max.apply(null, this.grid.map(r => r.length)); }
  // Google Sheets parses what appendRow writes as if the user had typed it:
  //   - one leading apostrophe is the "treat as text" marker and is NOT stored
  //   - a string starting with = + - @ is a formula; a chunk of HTML that happens to
  //     start that way is not one, so it comes back as an error value, not the text
  // Modelling both is what lets a test see cache chunks being corrupted.
  appendRow(arr) {
    this.grid.push(arr.map(v => {
      if (typeof v !== 'string') return v;
      if (v.charAt(0) === "'") return v.slice(1);
      if (/^[=+\-@]/.test(v)) return '#ERROR!';
      return v;
    }));
  }
  deleteRow(n) { this.grid.splice(n - 1, 1); }
  clearContents() { this.grid = []; }
  autoResizeColumns() {}
  setFrozenRows() {}
  // How many cells of one column (1-based) were read since the last reset.
  cellsReadInColumn(col, fromRow) {
    return this.reads.reduce((n, r) =>
      (col >= r.col && col < r.col + r.cols && r.row + r.rows - 1 >= (fromRow || 1))
        ? n + r.rows : n, 0);
  }
  resetReads() { this.reads = []; }

  _finder(colStart, firstRow, numRows, text) {
    const self = this;
    let entire = false;
    const finder = {
      matchEntireCell(v) { entire = v; return finder; },
      findAll() {
        const hits = [];
        for (let r = 0; r < numRows; r++) {
          const row = self.grid[firstRow - 1 + r];
          if (!row) continue;
          const cell = (row[colStart - 1] === undefined ? '' : row[colStart - 1]).toString();
          const match = entire ? cell.toLowerCase() === String(text).toLowerCase()
                               : cell.toLowerCase().indexOf(String(text).toLowerCase()) !== -1;
          if (match) hits.push({ getRow: () => firstRow + r });
        }
        return hits;
      },
      findNext() { const all = finder.findAll(); return all.length ? all[0] : null; },
    };
    return finder;
  }

  getRange(a, b, c, d) {
    const self = this;
    const chain = { setFontWeight() { return chain; }, setBackground() { return chain; }, setFontColor() { return chain; },
      setNote(n) { (self.notes = self.notes || {})[a + ',' + b] = n; return chain; },
      setNumberFormat(f) { (self.formats = self.formats || {})[a + ',' + b] = f; return chain; } };
    if (typeof a === 'string') {                       // whole column, e.g. 'A:A'
      const col = a.charCodeAt(0) - 64;
      return Object.assign(chain, {
        createTextFinder: text => self._finder(col, 1, self.grid.length, text),
      });
    }
    const nr = c || 1, nc = d || 1;
    return Object.assign(chain, {
      createTextFinder: text => self._finder(b, a, nr, text),
      getValues() {
        self.reads.push({ row: a, col: b, rows: nr, cols: nc });
        const out = [];
        for (let r = 0; r < nr; r++) {
          const row = self.grid[a - 1 + r] || [];
          const o = [];
          for (let k = 0; k < nc; k++) o.push(row[b - 1 + k] === undefined ? '' : row[b - 1 + k]);
          out.push(o);
        }
        return out;
      },
      getValue() { return this.getValues()[0][0]; },
      setValues(vals) {
        for (let r = 0; r < vals.length; r++) {
          while (self.grid.length < a + r) self.grid.push([]);
          const row = self.grid[a - 1 + r];
          for (let k = 0; k < vals[r].length; k++) { while (row.length < b + k) row.push(''); row[b - 1 + k] = vals[r][k]; }
        }
        return chain;
      },
      setValue(v) {
        while (self.grid.length < a) self.grid.push([]);
        const row = self.grid[a - 1];
        while (row.length < b) row.push('');
        row[b - 1] = v;
        return chain;
      },
    });
  }
}

function makeCheck() {
  const state = { failures: 0 };
  function check(name, cond, detail) {
    console.log('  ' + (cond ? 'PASS' : 'FAIL') + '  ' + name);
    if (!cond) { state.failures++; if (detail) console.log('          ' + detail); }
  }
  check.state = state;
  return check;
}

module.exports = { read, extractFunction, FakeSheet, makeCheck };
