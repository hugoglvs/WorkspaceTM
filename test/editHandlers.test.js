const test = require('node:test');
const assert = require('node:assert');

const { createMockSpreadsheet, setupGasGlobals } = require('./mockGas');
Object.assign(global, require('../src/Tables'));
const schema = require('../src/Schema');
global.SCHEMA_SPEC = schema.SCHEMA_SPEC;
global.schemaColumnLabel = schema.schemaColumnLabel;
global.schemaHeaderRow = schema.schemaHeaderRow;
global.FONT_FAMILY = schema.FONT_FAMILY;
global.FONT_SIZE = schema.FONT_SIZE;
global.CURRENT_SCHEMA_VERSION = schema.CURRENT_SCHEMA_VERSION;
global.DATE_FORMAT = schema.DATE_FORMAT;
Object.assign(global, require('../src/Typography'));
Object.assign(global, require('../src/Locations'));
const { freshInit } = require('../src/Init');
const edits = require('../src/EditHandlers');

let ss;
test.beforeEach(() => {
  ss = createMockSpreadsheet();
  setupGasGlobals(ss);
  freshInit(ss);
});

function editOn(tab, row, col, value) {
  var cell = ss.getSheetByName(tab).getRange(row, col);
  cell.setValue(value);
  return { range: cell, value: value, source: ss };
}

test('every table key with handlers exists in the table registry', () => {
  Object.keys(edits.EDIT_HANDLERS).forEach(function(key) {
    assert.ok(TABLE_REGISTRY[key], key + ' is a registered table');
  });
});

test('a view tab runs only its own handler, a data table runs before → common checks → after', () => {
  var calls = [];
  var saved = { weeklyMatrix: edits.EDIT_HANDLERS.weeklyMatrix, members: edits.EDIT_HANDLERS.members };
  edits.EDIT_HANDLERS.weeklyMatrix = { only: function matrix() { calls.push('matrix'); } };
  edits.EDIT_HANDLERS.members = { before: [function b1() { calls.push('before'); }], after: [function a1() { calls.push('after'); }] };
  try {
    if (!ss.getSheetByName('Weekly Presence')) ss.insertSheet('Weekly Presence');
    var r1 = edits.dispatchSheetEdit(editOn('Weekly Presence', 3, 3, 'Remote'));
    assert.deepStrictEqual(calls, ['matrix']);
    assert.strictEqual(r1.ran.length, 1, 'no common checks on a view');

    calls = [];
    var r2 = edits.dispatchSheetEdit(editOn('Members', 2, 3, '0102'));
    assert.deepStrictEqual(calls, ['before', 'after']);
    assert.deepStrictEqual(r2.ran.map(r => r.split(' ')[0]), ['b1', 'editCheckRequired', 'editCheckProtected', 'a1']);
  } finally {
    Object.assign(edits.EDIT_HANDLERS, saved);
  }
});

test('a handler error is logged and the next handlers still run; "stop" ends the edit', () => {
  var calls = [];
  var saved = edits.EDIT_HANDLERS.projects;
  edits.EDIT_HANDLERS.projects = { before: [function boom() { throw new Error('x'); }, function stopper() { calls.push('stop'); return 'stop'; }],
    after: [function never() { calls.push('never'); }] };
  try {
    var res = edits.dispatchSheetEdit(editOn('Projects', 2, 2, 'P'));
    assert.deepStrictEqual(calls, ['stop']);
    assert.strictEqual(res.stopped, true);
    assert.ok(Logger.getLogs().some(l => /boom \(projects\) note: Error: x/.test(l)));
  } finally {
    edits.EDIT_HANDLERS.projects = saved;
  }
});

test('an untracked tab runs nothing', () => {
  var other = ss.insertSheet('Scratch');
  var cell = other.getRange(1, 1);
  var res = edits.dispatchSheetEdit({ range: cell, value: 'x', source: ss });
  assert.strictEqual(res.tableKey, null);
  assert.deepStrictEqual(res.ran.map(r => r.split(' ')[0]), ['editCheckRequired', 'editCheckProtected']);
});
