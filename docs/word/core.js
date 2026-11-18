/*!
 * Virastar core — pure text logic shared by the Word task pane and the tests.
 * No Office.js, no DOM. Offsets are UTF-16 code units, as in JavaScript strings.
 */
(function (root, factory) {
  if (typeof module === 'object' && module.exports) {
    module.exports = factory();
  } else {
    root.VirastarCore = factory();
  }
}(typeof self !== 'undefined' ? self : this, function () {
  'use strict';

  var MAX_SEARCH = 255;          // Word refuses longer search strings
  var BOUNDARY = /[\s￼\u0000-\u001f]/; // spaces, tabs, line breaks, object placeholders; never the ZWNJ

  function isBoundary(ch) {
    return ch === undefined || ch === '' || BOUNDARY.test(ch);
  }

  /** Grow [s, e) until it starts and ends at a whitespace boundary. */
  function expandToWords(text, s, e) {
    var a = s, b = e;
    while (a > 0 && !isBoundary(text.charAt(a - 1))) { a--; }
    while (b < text.length && !isBoundary(text.charAt(b))) { b++; }
    return [a, b];
  }

  /** Start offsets of every occurrence of needle, overlapping ones included. */
  function occurrences(text, needle) {
    var out = [];
    if (!needle) { return out; }
    var i = text.indexOf(needle);
    while (i !== -1) {
      out.push(i);
      i = text.indexOf(needle, i + 1);
    }
    return out;
  }

  /** One more whole word to the left (dir < 0) or right (dir > 0) of [a, b). */
  function growByWord(text, a, b, dir) {
    if (dir < 0) {
      var i = a;
      while (i > 0 && isBoundary(text.charAt(i - 1))) { i--; }
      while (i > 0 && !isBoundary(text.charAt(i - 1))) { i--; }
      return [i, b];
    }
    var j = b;
    while (j < text.length && isBoundary(text.charAt(j))) { j++; }
    while (j < text.length && !isBoundary(text.charAt(j))) { j++; }
    return [a, j];
  }

  function mergeGroups(groups) {
    groups.sort(function (x, y) { return x.a - y.a || x.b - y.b; });
    var out = [];
    groups.forEach(function (g) {
      var last = out[out.length - 1];
      if (last && g.a < last.b) {
        last.b = Math.max(last.b, g.b);
        last.fix = last.fix.concat(g.fix);
      } else {
        out.push({ a: g.a, b: g.b, fix: g.fix.slice() });
      }
    });
    return out;
  }

  /**
   * Turn replacements inside one paragraph into the smallest set of safe edits.
   *
   * fixes: [{s, e, rep}] with rep a string. Overlapping fixes are dropped, first one wins.
   * Each edit covers whole words, because a whole word can be found again in Word by
   * searching, while a bare character offset cannot. An edit carries:
   *   a, b   span in the paragraph text      old, neu  text before and after
   *   k, n   it is occurrence k of n         tooLong   Word cannot search for it
   *   fix    the fixes it applies
   */
  function planEdits(text, fixes, opts) {
    opts = opts || {};
    var sorted = fixes.slice().sort(function (x, y) { return x.s - y.s || y.e - x.e; });
    var kept = [], last = -1;
    sorted.forEach(function (f) {
      if (f.s < last || f.s < 0 || f.e > text.length || f.e < f.s) { return; }
      kept.push(f);
      last = f.e;
    });
    var groups = kept.map(function (f) {
      var ab = expandToWords(text, f.s, f.e);
      return { a: ab[0], b: ab[1], fix: [f] };
    });
    groups = mergeGroups(groups);

    // Make every edit unique in the paragraph by taking in neighbouring words.
    for (var round = 0; round < 6; round++) {
      var changed = false;
      groups.forEach(function (g) {
        var old = text.slice(g.a, g.b);
        if (occurrences(text, old).length <= 1) { return; }
        var next = growByWord(text, g.a, g.b, round % 2 === 0 ? 1 : -1);
        if (next[0] === g.a && next[1] === g.b) {
          next = growByWord(text, g.a, g.b, round % 2 === 0 ? -1 : 1);
        }
        if ((next[0] !== g.a || next[1] !== g.b) && next[1] - next[0] <= MAX_SEARCH) {
          g.a = next[0];
          g.b = next[1];
          changed = true;
        }
      });
      if (!changed) { break; }
      groups = mergeGroups(groups);
    }

    var edits = [];
    groups.forEach(function (g) {
      var old = text.slice(g.a, g.b);
      var neu = old;
      g.fix.slice().sort(function (x, y) { return y.s - x.s; }).forEach(function (f) {
        neu = neu.slice(0, f.s - g.a) + f.rep + neu.slice(f.e - g.a);
      });
      if (neu === old && !opts.keepNoop) { return; }
      var occ = occurrences(text, old);
      edits.push({
        a: g.a, b: g.b, old: old, neu: neu,
        k: occ.indexOf(g.a), n: occ.length,
        tooLong: old.length > MAX_SEARCH || old.length === 0,
        fix: g.fix
      });
    });
    return edits;
  }

  /** The paragraph text that should result once these edits are in. */
  function applyEdits(text, edits) {
    var out = text;
    edits.slice().sort(function (x, y) { return y.a - x.a; }).forEach(function (ed) {
      out = out.slice(0, ed.a) + ed.neu + out.slice(ed.b);
    });
    return out;
  }

  /**
   * The ways to ask Word for a string. First the string itself; then, if it holds
   * characters Word's Find writes with caret codes, the caret-coded form.
   * A literal caret is always doubled.
   */
  function searchVariants(s) {
    var literal = s.replace(/\^/g, '^^');
    var coded = literal.replace(/‌/g, '^x').replace(/‍/g, '^z').replace(/ /g, '^s');
    return coded === literal ? [literal] : [literal, coded];
  }

  /**
   * Offsets of consecutive ranges whose texts should add up to the paragraph.
   * Returns [[start, end], …] or null when they do not add up.
   */
  function mapTokenRanges(texts, paragraph) {
    var joined = texts.join('');
    if (joined !== paragraph && joined !== paragraph + '\r' && joined !== paragraph + '\n') { return null; }
    var out = [], pos = 0;
    texts.forEach(function (t) {
      out.push([pos, Math.min(pos + t.length, paragraph.length)]);
      pos += t.length;
    });
    return out;
  }

  /** Indices of the first and last token that together cover [a, b). */
  function tokensCovering(offsets, a, b) {
    var i = -1, j = -1;
    for (var x = 0; x < offsets.length; x++) {
      if (i === -1 && offsets[x][1] > a) { i = x; }
      if (offsets[x][0] < b) { j = x; }
    }
    if (i === -1 || j === -1 || j < i) { return null; }
    return [i, j];
  }

  /** Text around a finding for display, cut at word boundaries. */
  function snippet(text, s, e, radius) {
    radius = radius || 28;
    var a = Math.max(0, s - radius), b = Math.min(text.length, e + radius);
    if (a > 0) {
      var sp = text.indexOf(' ', a);
      a = (sp !== -1 && sp < s) ? sp + 1 : a;
    }
    if (b < text.length) {
      var sp2 = text.lastIndexOf(' ', b);
      b = (sp2 !== -1 && sp2 > e) ? sp2 : b;
    }
    return {
      before: (a > 0 ? '…' : '') + text.slice(a, s),
      mid: text.slice(s, e),
      after: text.slice(e, b) + (b < text.length ? '…' : '')
    };
  }

  /** A key that follows a finding while the words around it stay the same. */
  function keyOf(code, text, s, e) {
    return code + '|' + text.slice(Math.max(0, s - 12), Math.min(text.length, e + 12)) + '|' + (s - Math.max(0, s - 12));
  }

  /** Drop findings that overlap an earlier one; the input order decides who stays. */
  function nonOverlapping(findings) {
    var out = [], last = -1;
    findings.slice().sort(function (x, y) { return x.s - y.s || y.e - x.e; }).forEach(function (f) {
      if (f.s >= last && f.e > f.s) {
        out.push(f);
        last = f.e;
      }
    });
    return out;
  }

  var FA_DIGITS = '۰۱۲۳۴۵۶۷۸۹';
  function faNum(n) {
    return String(n).replace(/[0-9]/g, function (d) { return FA_DIGITS.charAt(+d); });
  }

  return {
    MAX_SEARCH: MAX_SEARCH,
    expandToWords: expandToWords,
    occurrences: occurrences,
    planEdits: planEdits,
    applyEdits: applyEdits,
    searchVariants: searchVariants,
    mapTokenRanges: mapTokenRanges,
    tokensCovering: tokensCovering,
    snippet: snippet,
    keyOf: keyOf,
    nonOverlapping: nonOverlapping,
    faNum: faNum
  };
}));
