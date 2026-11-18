// Run with: node --test tests/core.test.js
'use strict';
const test = require('node:test');
const assert = require('node:assert/strict');
const Core = require('../docs/word/core.js');

const ZW = '‌';

test('expandToWords stops at spaces and keeps the half-space inside the word', () => {
  const t = 'او می' + ZW + 'رود و نمی' + ZW + 'آید';
  const s = t.indexOf('رود');
  assert.deepEqual(Core.expandToWords(t, s, s + 1), [3, 3 + ('می' + ZW + 'رود').length]);
});

test('a space before a comma is fixed together with the word before it', () => {
  const t = 'کتاب ، دفتر';
  const edits = Core.planEdits(t, [{ s: 4, e: 6, rep: '،' }]);
  assert.equal(edits.length, 1);
  assert.equal(edits[0].old, 'کتاب ،');
  assert.equal(edits[0].neu, 'کتاب،');
  assert.equal(Core.applyEdits(t, edits), 'کتاب، دفتر');
});

test('two fixes in one word become one edit', () => {
  const t = 'اين كتاب خوب است';
  const edits = Core.planEdits(t, [{ s: 1, e: 2, rep: 'ی' }, { s: 4, e: 5, rep: 'ک' }]);
  assert.equal(edits.length, 2);
  const t2 = 'كتابهاي من';
  const e2 = Core.planEdits(t2, [{ s: 0, e: 1, rep: 'ک' }, { s: 6, e: 7, rep: 'ی' }]);
  assert.equal(e2.length, 1);
  assert.equal(e2[0].neu, 'کتابهای');
});

test('a repeated word is made unique by taking a neighbour in', () => {
  const t = 'می رود و باز می رود و می ماند';
  const second = t.indexOf('می رود', 5);
  const edits = Core.planEdits(t, [{ s: second + 2, e: second + 3, rep: ZW }]);
  assert.equal(edits.length, 1);
  assert.equal(edits[0].n, 1, 'context should be unique: ' + edits[0].old);
  assert.equal(Core.occurrences(t, edits[0].old).length, 1);
  assert.equal(Core.applyEdits(t, edits), 'می رود و باز می' + ZW + 'رود و می ماند');
});

test('all occurrences fixed at once stay separate and exact', () => {
  const t = 'می رود و می رود';
  const edits = Core.planEdits(t, [{ s: 2, e: 3, rep: ZW }, { s: 11, e: 12, rep: ZW }]);
  assert.equal(Core.applyEdits(t, edits), 'می' + ZW + 'رود و می' + ZW + 'رود');
  edits.forEach((ed) => assert.ok(ed.k >= 0 && ed.k < ed.n));
});

test('overlapping fixes: the first one wins', () => {
  const t = 'مورد بررسی قرار داد';
  const edits = Core.planEdits(t, [{ s: 0, e: 15, rep: 'بررسی' }, { s: 5, e: 10, rep: 'x' }]);
  assert.equal(Core.applyEdits(t, edits), 'بررسی داد');
});

test('search variants escape the caret and offer caret codes for the half-space', () => {
  assert.deepEqual(Core.searchVariants('a^b'), ['a^^b']);
  assert.deepEqual(Core.searchVariants('می' + ZW + 'رود'), ['می' + ZW + 'رود', 'می^xرود']);
});

test('token ranges map back to offsets only when they add up', () => {
  const p = 'یک دو سه';
  const offs = Core.mapTokenRanges(['یک ', 'دو ', 'سه'], p);
  assert.deepEqual(offs, [[0, 3], [3, 6], [6, 8]]);
  assert.deepEqual(Core.tokensCovering(offs, 3, 5), [1, 1]);
  assert.deepEqual(Core.tokensCovering(offs, 1, 7), [0, 2]);
  assert.equal(Core.mapTokenRanges(['یک', 'دو'], p), null);
});

test('edits longer than Word can search are marked', () => {
  const long = 'x'.repeat(300);
  const edits = Core.planEdits(long + ' a', [{ s: 0, e: 1, rep: 'y' }]);
  assert.equal(edits[0].tooLong, true);
});

test('snippet cuts at word boundaries and Persian digits render', () => {
  const t = 'این یک جملهٔ نسبتاً بلند است که در آن واژهٔ نادرست آمده است و ادامه دارد';
  const s = t.indexOf('نادرست');
  const sn = Core.snippet(t, s, s + 6, 12);
  assert.equal(sn.mid, 'نادرست');
  assert.ok(sn.before.startsWith('…'));
  assert.equal(Core.faNum(120), '۱۲۰');
});

test('nonOverlapping keeps the earlier and longer finding', () => {
  const out = Core.nonOverlapping([{ s: 5, e: 9 }, { s: 0, e: 6 }, { s: 6, e: 8 }]);
  assert.deepEqual(out, [{ s: 0, e: 6 }, { s: 6, e: 8 }]);
});
