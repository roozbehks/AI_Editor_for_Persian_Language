/*
 * A small stand-in for Office.js, used only by the add-in tests.
 *
 * It keeps a document as a list of paragraphs and imitates the parts of the Word API the
 * task pane uses, including the batch model: nothing is read or changed before
 * context.sync(), and reading a property that was not loaded throws. It can also imitate
 * the ways Word's Find may treat the half-space, so each fallback in the task pane is
 * exercised. It says nothing about how the real Word behaves.
 *
 * Configure with window.__MOCK__ before this script loads:
 *   paragraphs: [text, …]      maxSet: '1.8'            host: 'Word' | null
 *   search: 'fuzzy' | 'literal' | 'caret' | 'broken'    annotations: true | false
 *   tokensIncludeMark: false   popupReplaces: false
 */
(function () {
  'use strict';
  var cfg = window.__MOCK__ || {};
  var ZW = '‌';
  var uid = 0;

  function newPara(text) {
    return { id: 'p' + (++uid) + '-0000-0000', raw: text, cur: text, ranges: [], annotations: [], alive: true };
  }

  var doc = {
    paras: (cfg.paragraphs || []).map(newPara),
    tracking: 'Off',
    selection: null,
    selectedText: null,
    handlers: { changed: [], added: [], deleted: [], selection: [], annClicked: [], annPopup: [] },
    annSeq: 0,
    calls: { sync: 0, search: 0, insert: 0, tokens: 0, annotate: 0 }
  };

  function supported(v) {
    var max = String(cfg.maxSet || '1.8').split('.').map(Number), want = String(v).split('.').map(Number);
    return want[0] < max[0] || (want[0] === max[0] && want[1] <= max[1]);
  }

  function need(v, what) {
    if (!supported(v)) { throw new Error('ApiNotFound: ' + what + ' needs WordApi ' + v); }
  }

  // ── context ──
  function Context() {
    this.queue = [];
    this.document = new DocumentProxy(this);
  }
  Context.prototype.sync = function () {
    var self = this;
    return new Promise(function (resolve, reject) {
      setTimeout(function () {
        doc.calls.sync++;
        var q = self.queue;
        self.queue = [];
        try {
          q.forEach(function (op) { op(); });
          resolve();
        } catch (e) {
          reject(e);
        }
      }, 0);
    });
  };

  function loadable(obj, ctx, read) {
    obj._v = {};
    obj.load = function (props) {
      String(props || '').split(',').forEach(function (name) {
        name = name.trim().replace(/^items\//, '');
        if (name) { ctx.queue.push(function () { read(name); }); }
      });
      return obj;
    };
  }

  function prop(obj, name) {
    Object.defineProperty(obj, name, {
      get: function () {
        if (!(name in obj._v)) { throw new Error('PropertyNotLoaded: ' + name); }
        return obj._v[name];
      },
      configurable: true
    });
  }

  function clientResult(ctx, compute) {
    var r = { _has: false };
    Object.defineProperty(r, 'value', {
      get: function () {
        if (!r._has) { throw new Error('ValueNotLoaded'); }
        return r._value;
      }
    });
    ctx.queue.push(function () { r._value = compute(); r._has = true; });
    return r;
  }

  // ── document ──
  function DocumentProxy(ctx) {
    var self = this;
    this.ctx = ctx;
    this.body = new BodyProxy(ctx);
    loadable(this, ctx, function (name) {
      if (name === 'changeTrackingMode') { need('1.4', name); self._v[name] = doc.tracking; }
    });
    Object.defineProperty(this, 'changeTrackingMode', {
      get: function () {
        if (!('changeTrackingMode' in self._v)) { throw new Error('PropertyNotLoaded: changeTrackingMode'); }
        return self._v.changeTrackingMode;
      },
      set: function (v) { ctx.queue.push(function () { need('1.4', 'changeTrackingMode'); doc.tracking = v; }); }
    });
    var ev = function (list, set) {
      return { add: function (h) { if (set) { need(set, 'event'); } list.push(h); return { remove: function () {} }; } };
    };
    this.onParagraphChanged = ev(doc.handlers.changed, '1.6');
    this.onParagraphAdded = ev(doc.handlers.added, '1.6');
    this.onParagraphDeleted = ev(doc.handlers.deleted, '1.6');
    this.onAnnotationClicked = ev(doc.handlers.annClicked, '1.7');
    this.onAnnotationPopupAction = ev(doc.handlers.annPopup, '1.8');
  }
  DocumentProxy.prototype.getParagraphByUniqueLocalId = function (id) {
    need('1.6', 'getParagraphByUniqueLocalId');
    var p = new ParagraphProxy(this.ctx, null);
    this.ctx.queue.push(function () {
      var m = doc.paras.filter(function (x) { return x.id === id; })[0];
      if (!m) { var e = new Error('ItemNotFound'); e.code = 'ItemNotFound'; throw e; }
      p.m = m;
    });
    return p;
  };
  DocumentProxy.prototype.getSelection = function () {
    var ctx = this.ctx;
    return {
      paragraphs: {
        getFirst: function () {
          var p = new ParagraphProxy(ctx, null);
          ctx.queue.push(function () { p.m = (doc.selection && doc.selection.m) || doc.paras[0]; });
          return p;
        }
      }
    };
  };

  function BodyProxy(ctx) {
    var self = this;
    this.ctx = ctx;
    this.paragraphs = new ParagraphCollection(ctx);
    loadable(this, ctx, function (name) {
      if (name === 'text') { self._v.text = doc.paras.map(function (p) { return p.raw; }).join('\r'); }
    });
    prop(this, 'text');
  }
  BodyProxy.prototype.getReviewedText = function () {
    need('1.4', 'getReviewedText');
    return clientResult(this.ctx, function () { return doc.paras.map(function (p) { return p.cur; }).join('\r'); });
  };

  function ParagraphCollection(ctx) {
    var self = this;
    this.ctx = ctx;
    this._items = null;
    this._props = [];
    this.load = function (props) {
      var names = String(props || '').split(',').map(function (s) { return s.trim().replace(/^items\//, ''); }).filter(Boolean);
      ctx.queue.push(function () {
        self._items = doc.paras.map(function (m) { return new ParagraphProxy(ctx, m); });
        self._items.forEach(function (p) { names.forEach(function (n) { p._read(n); }); });
      });
      return self;
    };
    Object.defineProperty(this, 'items', {
      get: function () {
        if (!self._items) { throw new Error('PropertyNotLoaded: items'); }
        return self._items;
      }
    });
  }

  // ── paragraph ──
  function ParagraphProxy(ctx, m) {
    var self = this;
    this.ctx = ctx;
    this.m = m;
    this._read = function (name) {
      if (!self.m || !self.m.alive) { throw new Error('ItemNotFound'); }
      if (name === 'text') { self._v.text = self.m.raw; }
      if (name === 'uniqueLocalId') { need('1.6', name); self._v.uniqueLocalId = self.m.id; }
    };
    loadable(this, ctx, this._read);
    prop(this, 'text');
    prop(this, 'uniqueLocalId');
  }
  ParagraphProxy.prototype.getReviewedText = function () {
    need('1.4', 'getReviewedText');
    var self = this;
    return clientResult(this.ctx, function () { return self.m.cur; });
  };
  ParagraphProxy.prototype.select = function () {
    var self = this;
    this.ctx.queue.push(function () { setSelection(self.m, 0, self.m.raw.length); });
  };
  ParagraphProxy.prototype.search = function (q, opts) {
    var self = this;
    return new RangeCollection(this.ctx, function () { return find(self.m, 0, self.m.raw.length, q, opts); });
  };
  ParagraphProxy.prototype.getTextRanges = function (marks) {
    need('1.3', 'getTextRanges');
    var self = this;
    return new RangeCollection(this.ctx, function () {
      doc.calls.tokens++;
      var out = [], t = self.m.raw, start = 0;
      for (var i = 0; i < t.length; i++) {
        if (marks.indexOf(t.charAt(i)) !== -1) { out.push([start, i + 1]); start = i + 1; }
      }
      if (start < t.length) { out.push([start, t.length]); }
      return out.map(function (se, k) {
        return { m: self.m, s: se[0], e: se[1], mark: !!cfg.tokensIncludeMark && k === out.length - 1 };
      });
    });
  };
  ParagraphProxy.prototype.getAnnotations = function () {
    need('1.7', 'getAnnotations');
    var self = this, ctx = this.ctx;
    var col = { _items: null };
    col.load = function () {
      ctx.queue.push(function () {
        if (cfg.annotations === false) { throw new Error('GeneralException: annotations need a subscription'); }
        col._items = self.m.annotations.map(function (a) {
          return { id: a.id, delete: function () { ctx.queue.push(function () { self.m.annotations = self.m.annotations.filter(function (x) { return x.id !== a.id; }); }); } };
        });
      });
      return col;
    };
    Object.defineProperty(col, 'items', { get: function () { if (!col._items) { throw new Error('PropertyNotLoaded: items'); } return col._items; } });
    return col;
  };
  ParagraphProxy.prototype.insertAnnotations = function (set) {
    need('1.7', 'insertAnnotations');
    var self = this;
    return clientResult(this.ctx, function () {
      if (cfg.annotations === false) { throw new Error('GeneralException: annotations need a subscription'); }
      doc.calls.annotate++;
      return set.critiques.map(function (c) {
        var a = { id: 'a' + (++doc.annSeq), critique: c, text: self.m.raw.substr(c.start, c.length) };
        self.m.annotations.push(a);
        return a.id;
      });
    });
  };

  // ── ranges ──
  function RangeCollection(ctx, compute) {
    var self = this;
    this.ctx = ctx;
    this._items = null;
    ctx.queue.push(function () {
      self._items = compute().map(function (x) { return new RangeProxy(ctx, x.m, x.s, x.e, x.mark); });
    });
    this.load = function (props) {
      var names = String(props || '').split(',').map(function (s) { return s.trim().replace(/^items\//, ''); }).filter(Boolean);
      ctx.queue.push(function () { self._items.forEach(function (r) { names.forEach(function (n) { r._read(n); }); }); });
      return self;
    };
    Object.defineProperty(this, 'items', {
      get: function () {
        if (!self._items) { throw new Error('PropertyNotLoaded: items'); }
        return self._items;
      }
    });
  }

  function RangeProxy(ctx, m, s, e, mark) {
    var self = this;
    this.ctx = ctx;
    this.m = m; this.s = s; this.e = e; this.mark = !!mark; this.valid = true;
    if (m) { m.ranges.push(this); }
    this._read = function (name) {
      if (name === 'text') {
        if (!self.valid) { throw new Error('InvalidObjectPath'); }
        self._v.text = self.m.raw.slice(self.s, self.e) + (self.mark ? '\r' : '');
      }
    };
    loadable(this, ctx, this._read);
    prop(this, 'text');
  }
  RangeProxy.prototype.select = function () {
    var self = this;
    this.ctx.queue.push(function () { setSelection(self.m, self.s, self.e); });
  };
  RangeProxy.prototype.search = function (q, opts) {
    var self = this;
    return new RangeCollection(this.ctx, function () { return find(self.m, self.s, self.e, q, opts); });
  };
  RangeProxy.prototype.expandTo = function (other) {
    need('1.3', 'expandTo');
    var self = this;
    var r = new RangeProxy(this.ctx, null, 0, 0, false);
    this.ctx.queue.push(function () {
      r.m = self.m; r.s = Math.min(self.s, other.s); r.e = Math.max(self.e, other.e); r.mark = self.mark || other.mark;
      r.m.ranges.push(r);
    });
    return r;
  };
  RangeProxy.prototype.insertText = function (text, where) {
    var self = this;
    if (where !== 'Replace') { throw new Error('mock: only Replace is implemented'); }
    this.ctx.queue.push(function () {
      if (!self.valid) { throw new Error('InvalidObjectPath'); }
      doc.calls.insert++;
      replaceInPara(self.m, self.s, self.e, text, self);
      if (self.mark) { mergeWithNext(self.m); }
    });
    return self;
  };

  function setSelection(m, s, e) {
    doc.selection = { m: m, s: s, e: e };
    doc.selectedText = m.raw.slice(s, e);
    setTimeout(function () { doc.handlers.selection.forEach(function (h) { h({}); }); }, 0);
  }

  function replaceInPara(m, s, e, text, self) {
    var old = m.raw.slice(s, e);
    var tracked = doc.tracking !== 'Off';
    // The reviewed ("current") text loses the old words either way.
    var cs = m.cur.indexOf(old, Math.max(0, s - (m.raw.length - m.cur.length)));
    if (cs === -1) { cs = m.cur.indexOf(old); }
    if (cs !== -1) { m.cur = m.cur.slice(0, cs) + text + m.cur.slice(cs + old.length); }
    var inserted = tracked ? old + text : text;      // a tracked deletion stays in the raw text
    m.raw = m.raw.slice(0, s) + inserted + m.raw.slice(e);
    var delta = inserted.length - (e - s);
    m.ranges.forEach(function (r) {
      if (r === self) { r.e = s + inserted.length; return; }
      if (r.s >= e) { r.s += delta; r.e += delta; }
      else if (r.e > s) { r.valid = false; }
    });
    m.annotations = [];
    fire('changed', m);
  }

  function mergeWithNext(m) {
    var i = doc.paras.indexOf(m), next = doc.paras[i + 1];
    if (!next) { return; }
    m.raw += next.raw; m.cur += next.cur;
    next.alive = false;
    doc.paras.splice(i + 1, 1);
  }

  function fire(kind, m) {
    setTimeout(function () {
      doc.handlers[kind].forEach(function (h) { h({ type: kind, uniqueLocalIds: m ? [m.id] : [] }); });
    }, 0);
  }

  // ── Find, in the flavours the task pane must survive ──
  function decodeCarets(q) {
    var out = '', i = 0;
    while (i < q.length) {
      var c = q.charAt(i);
      if (c === '^' && i + 1 < q.length) {
        var n = q.charAt(i + 1);
        out += n === '^' ? '^' : n === 'x' ? ZW : n === 'z' ? '‍' : n === 's' ? ' ' : n === 't' ? '\t' : '�';
        i += 2;
      } else {
        out += c;
        i++;
      }
    }
    return out;
  }

  function find(m, from, to, q) {
    doc.calls.search++;
    if (q.length > 255 || q.length === 0) { var e = new Error('SearchStringInvalidOrTooLong'); e.code = 'SearchStringInvalidOrTooLong'; throw e; }
    var mode = cfg.search || 'fuzzy';
    var text = m.raw.slice(from, to), hits = [];
    if (mode === 'broken') { return hits; }
    if (mode === 'caret' && q.indexOf(ZW) !== -1) { return hits; }   // a literal half-space never matches
    var needle = decodeCarets(q);
    if (mode === 'fuzzy') {
      // Control characters are ignored on both sides, as with "Match control" switched off.
      var map = [], plain = '';
      for (var i = 0; i < text.length; i++) {
        if (text.charAt(i) !== ZW) { map.push(i); plain += text.charAt(i); }
      }
      var n2 = needle.split(ZW).join('');
      if (!n2) { return hits; }
      var at = plain.indexOf(n2);
      while (at !== -1) {
        hits.push({ m: m, s: from + map[at], e: from + map[at + n2.length - 1] + 1 });
        at = plain.indexOf(n2, at + n2.length);
      }
      return hits;
    }
    var j = text.indexOf(needle);
    while (j !== -1) {
      hits.push({ m: m, s: from + j, e: from + j + needle.length });
      j = text.indexOf(needle, j + needle.length);
    }
    return hits;
  }

  // ── globals ──
  window.Word = {
    run: function (fn) {
      var ctx = new Context();
      return Promise.resolve().then(function () { return fn(ctx); }).then(function (r) {
        return ctx.sync().then(function () { return r; });
      });
    }
  };

  window.Office = {
    HostType: { Word: 'Word' },
    EventType: { DocumentSelectionChanged: 'documentSelectionChanged' },
    onReady: function (cb) {
      var info = { host: cfg.host === undefined ? 'Word' : cfg.host, platform: 'PC' };
      return new Promise(function (resolve) {
        setTimeout(function () { if (cb) { cb(info); } resolve(info); }, 0);
      });
    },
    context: {
      requirements: { isSetSupported: function (name, v) { return name === 'WordApi' && supported(v); } },
      document: { addHandlerAsync: function (type, h) { doc.handlers.selection.push(h); } }
    }
  };

  // ── what the tests poke ──
  window.__mock = {
    doc: doc,
    texts: function () { return doc.paras.map(function (p) { return p.raw; }); },
    current: function () { return doc.paras.map(function (p) { return p.cur; }); },
    type: function (i, text) { doc.paras[i].raw = text; doc.paras[i].cur = text; doc.paras[i].ranges = []; fire('changed', doc.paras[i]); },
    addParagraph: function (text) { var m = newPara(text); doc.paras.push(m); fire('added', m); },
    moveCursor: function (i) { setSelection(doc.paras[i], 0, 0); },
    annotations: function () {
      var out = [];
      doc.paras.forEach(function (p, i) { p.annotations.forEach(function (a) { out.push({ para: i, id: a.id, text: a.text, critique: a.critique }); }); });
      return out;
    },
    popup: function (id, action) {
      var found = null;
      doc.paras.forEach(function (p) { p.annotations.forEach(function (a) { if (a.id === id) { found = { p: p, a: a }; } }); });
      if (found && action === 'Accept' && cfg.popupReplaces) {
        var c = found.a.critique;
        replaceInPara(found.p, c.start, c.start + c.length, c.popupOptions.suggestions[0], null);
      }
      return Promise.all(doc.handlers.annPopup.map(function (h) {
        return h({ id: id, action: action, critiqueSuggestion: found && found.a.critique.popupOptions ? found.a.critique.popupOptions.suggestions[0] : '' });
      }));
    },
    clickAnnotation: function (id) { doc.handlers.annClicked.forEach(function (h) { h({ id: id }); }); }
  };
}());
