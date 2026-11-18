/*
 * Virastar — Word task pane.
 *
 * Reads the paragraphs of the open document, checks them with fa_lint.py (running in a
 * worker, inside this browser), lists what it finds, and applies a fix when asked.
 * Nothing is sent to a server.
 *
 * Offsets cannot address text in Word, so a fix is located by searching the paragraph for
 * the whole words around it and is applied only when the text found is exactly the text
 * expected. When it cannot be located safely it is left alone and reported.
 */
/* global Office, Word, VirastarCore */
(function () {
  'use strict';

  var Core = VirastarCore;
  var faNum = Core.faNum;
  var query = new URLSearchParams(location.search);
  var $ = function (id) { return document.getElementById(id); };

  var LEVEL_NAME = { E: 'خطا', W: 'هشدار', C: 'وارسی' };
  var LEVEL_COLOR = { E: 'Red', W: 'Blue', C: 'Lavender' };
  var STORE_KEY = 'virastar.opts.v1';
  var PAGE = 200;

  var S = {
    ready: false,
    has: {},
    opts: { live: true, track: false, bulkW: false, follow: true, strict: false, verbless: false, underline: false, house: '' },
    house: null,
    filter: { E: true, W: true, C: true },
    paras: [],            // [{i, id, text, raw, pending}]
    cache: new Map(),     // paragraph text -> findings
    cons: [],             // [{pi, f}] document-wide consistency findings
    items: [],            // [{pi, f, key}]
    ignored: new Set(),
    limit: PAGE,
    scanning: false, applying: false, dirty: false, timer: 0, selTimer: 0,
    annSig: new Map(), annMap: new Map(), annEvents: false,
    engine: null
  };

  // ───────────────────────── small helpers ─────────────────────────

  function el(tag, cls, text) {
    var n = document.createElement(tag);
    if (cls) { n.className = cls; }
    if (text != null) { n.textContent = text; }
    return n;
  }

  function setStatus(text, kind) {
    var s = $('status');
    s.textContent = text;
    s.className = 'status' + (kind ? ' ' + kind : '');
  }

  function note(text) {
    var n = $('note');
    n.textContent = text || '';
    n.hidden = !text;
  }

  function reportError(err, prefix) {
    if (window.console) { console.error(err); }
    var msg = (err && (err.message || err.code)) || String(err);
    setStatus((prefix || 'خطا') + ': ' + msg, 'err');
  }

  function StaleError() { this.stale = true; this.message = 'stale'; }

  function loadOpts() {
    try {
      var saved = JSON.parse(localStorage.getItem(STORE_KEY) || '{}');
      Object.keys(S.opts).forEach(function (k) {
        if (typeof saved[k] === typeof S.opts[k]) { S.opts[k] = saved[k]; }
      });
    } catch (e) { /* storage may be unavailable; defaults stay */ }
    try { S.house = S.opts.house ? JSON.parse(S.opts.house) : null; } catch (e2) { S.house = null; }
  }

  function saveOpts() {
    try { localStorage.setItem(STORE_KEY, JSON.stringify(S.opts)); } catch (e) { /* not fatal */ }
  }

  function lintOpts() {
    return { house: S.house, strict: S.opts.strict, verbless: S.opts.verbless };
  }

  // ───────────────────────── the lint engine ─────────────────────────

  var worker = null, seq = 0, waiting = new Map();

  function rpc(cmd, args) {
    return new Promise(function (resolve, reject) {
      var id = ++seq;
      waiting.set(id, { resolve: resolve, reject: reject });
      worker.postMessage({ id: id, cmd: cmd, args: args });
    });
  }

  function startEngine() {
    worker = new Worker('lint-worker.js');
    worker.onmessage = function (ev) {
      var m = ev.data || {}, w = waiting.get(m.id);
      if (!w) { return; }
      waiting.delete(m.id);
      if (m.ok) { w.resolve(m.result); } else { w.reject(new Error(m.error)); }
    };
    worker.onerror = function (ev) {
      waiting.forEach(function (w) { w.reject(new Error(ev.message || 'worker')); });
      waiting.clear();
    };
    return rpc('init', {
      pyodideBase: query.get('pyodide') || '',
      lintUrl: new URL(query.get('lint') || '../fa_lint.py', location.href).href
    });
  }

  function toFinding(a) {
    return { code: a[0], level: a[1], s: a[2], e: a[3], msg: a[4], rep: a[5], auto: !!a[6] };
  }

  // ───────────────────────── reading the document ─────────────────────────

  /**
   * One entry per paragraph of the body. `text` is what gets checked; `raw` is what Word
   * reports as the paragraph text. They differ only where tracked changes are waiting to
   * be accepted, and such paragraphs are shown but not edited.
   */
  function readParagraphs() {
    return Word.run(function (ctx) {
      var body = ctx.document.body, ps = body.paragraphs, current = null;
      ps.load(S.has.ids ? 'text,uniqueLocalId' : 'text');
      if (S.has.reviewed) {
        body.load('text');
        current = body.getReviewedText('Current');
      }
      return ctx.sync().then(function () {
        if (!current || current.value === body.text) { return null; }
        var revs = ps.items.map(function (p) { return p.getReviewedText('Current'); });
        return ctx.sync().then(function () { return revs; });
      }).then(function (revs) {
        return ps.items.map(function (p, i) {
          var raw = p.text || '';
          var text = revs ? (revs[i].value || '') : raw;
          return { i: i, id: S.has.ids ? p.uniqueLocalId : null, text: text, raw: raw, pending: text !== raw };
        });
      });
    }).catch(function (err) {
      // Fall back one capability at a time; older or partial hosts still get a plain read.
      if (S.has.reviewed) { S.has.reviewed = false; return readParagraphs(); }
      if (S.has.ids) { S.has.ids = false; return readParagraphs(); }
      throw err;
    });
  }

  function schedule(ms) {
    clearTimeout(S.timer);
    S.timer = setTimeout(scan, ms == null ? 700 : ms);
  }

  function scan() {
    if (!S.ready) { return Promise.resolve(); }
    if (S.scanning || S.applying) { S.dirty = true; return Promise.resolve(); }
    S.scanning = true;
    var paras;
    return readParagraphs().then(function (ps) {
      paras = ps;
      var need = [], seen = new Set();
      paras.forEach(function (p) {
        if (p.text.trim() && !S.cache.has(p.text) && !seen.has(p.text)) {
          seen.add(p.text);
          need.push(p.text);
        }
      });
      var chain = Promise.resolve();
      var step = 120;
      for (var i = 0; i < need.length; i += step) {
        (function (from) {
          chain = chain.then(function () {
            var chunk = need.slice(from, from + step);
            if (need.length > step) {
              setStatus('بازبینی ' + faNum(Math.min(from + step, need.length)) + ' از ' + faNum(need.length) + ' بند…');
            }
            return rpc('lint', { texts: chunk, opts: lintOpts() }).then(function (res) {
              chunk.forEach(function (t, k) { S.cache.set(t, res[k].map(toFinding)); });
            });
          });
        }(i));
      }
      return chain;
    }).then(function () {
      return rpc('cons', { texts: paras.map(function (p) { return p.text; }) }).then(function (res) {
        return res.map(function (x) { return { pi: x[0], f: toFinding(x[1]) }; });
      }, function () { return []; });
    }).then(function (cons) {
      S.paras = paras;
      S.cons = cons;
      if (S.cache.size > 5000) {
        var keep = new Map();
        paras.forEach(function (p) { if (S.cache.has(p.text)) { keep.set(p.text, S.cache.get(p.text)); } });
        S.cache = keep;
      }
      render();
      if (S.opts.underline) { return syncAnnotations(false); }
    }).catch(function (err) {
      reportError(err, 'بازبینی انجام نشد');
    }).then(function () {
      S.scanning = false;
      if (S.dirty) { S.dirty = false; schedule(250); }
    });
  }

  // ───────────────────────── list of findings ─────────────────────────

  function findingsOf(pi) {
    var p = S.paras[pi];
    var fs = (S.cache.get(p.text) || []).slice();
    S.cons.forEach(function (c) { if (c.pi === pi) { fs.push(c.f); } });
    fs.sort(function (x, y) { return x.s - y.s || 'EWC'.indexOf(x.level) - 'EWC'.indexOf(y.level); });
    return fs;
  }

  function collect() {
    var items = [];
    S.paras.forEach(function (p, pi) {
      findingsOf(pi).forEach(function (f) {
        var key = Core.keyOf(f.code, p.text, f.s, f.e);
        if (!S.ignored.has(key)) { items.push({ pi: pi, f: f, key: key }); }
      });
    });
    return items;
  }

  function isBulk(it) {
    if (!it.f.auto || S.paras[it.pi].pending) { return false; }
    return S.opts.bulkW ? true : it.f.level !== 'W';
  }

  function suggestionOf(it) {
    var f = it.f, p = S.paras[it.pi];
    if (f.rep == null) { return null; }
    var eds = Core.planEdits(p.text, [{ s: f.s, e: f.e, rep: f.rep }]);
    return eds.length ? eds[0] : null;
  }

  function card(it) {
    var f = it.f, p = S.paras[it.pi];
    var c = el('article', 'card lv-' + f.level);
    c.tabIndex = 0;
    c.setAttribute('role', 'button');
    c.setAttribute('aria-label', LEVEL_NAME[f.level] + ': ' + f.msg);
    c.dataset.key = it.key;

    var head = el('div', 'head');
    head.appendChild(el('span', 'badge', LEVEL_NAME[f.level]));
    head.appendChild(el('span', 'code', f.code));
    c.appendChild(head);
    c.appendChild(el('p', 'msg', f.msg));

    var ed = suggestionOf(it);
    var span = { s: f.s, e: f.e };
    if (ed && p.text.slice(f.s, f.e).trim().length < 2) { span = { s: ed.a, e: ed.b }; }
    var sn = Core.snippet(p.text, span.s, span.e, 26);
    var ctx = el('p', 'ctx');
    ctx.appendChild(document.createTextNode(sn.before));
    ctx.appendChild(el('mark', null, sn.mid));
    ctx.appendChild(document.createTextNode(sn.after));
    c.appendChild(ctx);

    var acts = el('div', 'acts');
    if (ed && !p.pending) {
      var sug = el('p', 'sug', 'پیشنهاد: ');
      sug.appendChild(el('b', null, ed.neu.trim() ? '«' + ed.neu + '»' : 'برداشته شود'));
      c.appendChild(sug);
      var apply = el('button', 'apply', 'اعمال');
      apply.type = 'button';
      apply.addEventListener('click', function (ev) { ev.stopPropagation(); applyItems([it]); });
      acts.appendChild(apply);
    } else if (f.rep != null && p.pending) {
      c.appendChild(el('p', 'sug', 'این بند تغییرِ ردگیری‌شده دارد؛ نخست آن را بپذیرید یا رد کنید.'));
    }
    var skip = el('button', 'ghost', 'نادیده');
    skip.type = 'button';
    skip.addEventListener('click', function (ev) { ev.stopPropagation(); S.ignored.add(it.key); render(); });
    acts.appendChild(skip);
    c.appendChild(acts);

    c.addEventListener('click', function () { showInDocument(it); });
    c.addEventListener('keydown', function (ev) {
      if (ev.key === 'Enter' && ev.target === c) { showInDocument(it); }
    });
    return c;
  }

  function render() {
    var items = collect();
    S.items = items;
    var counts = { E: 0, W: 0, C: 0 };
    items.forEach(function (it) { counts[it.f.level]++; });
    $('nE').textContent = faNum(counts.E);
    $('nW').textContent = faNum(counts.W);
    $('nC').textContent = faNum(counts.C);

    var bulk = items.filter(isBulk).length;
    var btn = $('btnFixAll');
    btn.disabled = bulk === 0;
    btn.textContent = bulk
      ? 'درست‌کردن ' + faNum(bulk) + (S.opts.bulkW ? ' موردِ دارای جایگزین' : ' خطای مکانیکی')
      : 'خطای مکانیکی‌ای نیست';

    var list = $('list');
    list.textContent = '';
    var vis = items.filter(function (it) { return S.filter[it.f.level]; });
    if (!vis.length) {
      list.appendChild(el('p', 'empty', items.length ? 'همهٔ یافته‌ها با صافیِ بالا پنهان شده است.' : 'چیزی برای اصلاح پیدا نشد.'));
    }
    var group = null, lastPi = -1, parasWith = 0;
    vis.slice(0, S.limit).forEach(function (it) {
      if (it.pi !== lastPi) {
        group = el('section', 'group');
        group.dataset.pi = String(it.pi);
        group.appendChild(el('h2', null, 'بند ' + faNum(it.pi + 1) + ' · ' + S.paras[it.pi].text.slice(0, 70)));
        list.appendChild(group);
        lastPi = it.pi;
      }
      group.appendChild(card(it));
    });
    var seenPara = new Set();
    items.forEach(function (it) { seenPara.add(it.pi); });
    parasWith = seenPara.size;
    if (vis.length > S.limit) {
      var more = el('button', 'more', 'نمایش ' + faNum(Math.min(PAGE, vis.length - S.limit)) + ' مورد دیگر');
      more.type = 'button';
      more.addEventListener('click', function () { S.limit += PAGE; render(); });
      list.appendChild(more);
    }
    if (items.length) {
      setStatus(faNum(items.length) + ' مورد در ' + faNum(parasWith) + ' بند');
    } else {
      setStatus(S.paras.length ? 'متن بازبینی شد.' : 'سند خالی است.', 'ok');
    }
    if (S.paras.some(function (p) { return p.pending; })) {
      note('در این سند تغییرِ ردگیری‌شدهٔ پذیرفته‌نشده هست. ویراستار متنِ پس از تغییرها را می‌خواند و به بندهایی که چنین تغییری دارند دست نمی‌زند.');
    } else if ($('note').dataset.sticky !== '1') {
      note('');
    }
  }

  function markCurrent(pi) {
    var groups = document.querySelectorAll('.group');
    var hit = null;
    Array.prototype.forEach.call(groups, function (g) {
      var on = g.dataset.pi === String(pi);
      g.classList.toggle('current', on);
      if (on) { hit = g; }
    });
    if (hit && typeof hit.scrollIntoView === 'function') {
      var r = hit.getBoundingClientRect();
      if (r.top < 0 || r.top > window.innerHeight - 60) { hit.scrollIntoView({ block: 'start' }); }
    }
  }

  // ───────────────────────── finding text in Word ─────────────────────────

  function getParagraph(ctx, pi) {
    var p = S.paras[pi];
    if (!p) { return Promise.reject(new StaleError()); }
    var para, ps;
    if (p.id) {
      para = ctx.document.getParagraphByUniqueLocalId(p.id);
      para.load('text');
    } else {
      ps = ctx.document.body.paragraphs;
      ps.load('text');
    }
    return ctx.sync().then(function () {
      if (!p.id) { para = ps.items[p.i]; }
      if (!para || para.text !== p.raw) { throw new StaleError(); }
      return para;
    }, function () { throw new StaleError(); });
  }

  /** Search the paragraph for each edit's old text; keep a range only on an exact, countable match. */
  function searchPass(ctx, para, list, variant) {
    var cols = list.map(function (ed) {
      var c = para.search(Core.searchVariants(ed.old)[variant], { matchCase: true });
      c.load('text');
      return c;
    });
    return ctx.sync().then(function () {
      list.forEach(function (ed, i) {
        var exact = cols[i].items.filter(function (r) { return r.text === ed.old; });
        if (exact.length === ed.n && exact[ed.k]) { ed.range = exact[ed.k]; }
      });
    });
  }

  function locate(ctx, para, P, edits) {
    var todo = edits.filter(function (ed) { return !ed.tooLong; });
    var chain = Promise.resolve();
    [0, 1].forEach(function (variant) {
      chain = chain.then(function () {
        var pend = todo.filter(function (ed) { return !ed.range && Core.searchVariants(ed.old).length > variant; });
        if (!pend.length) { return; }
        return searchPass(ctx, para, pend, variant).catch(function () {
          // One unsearchable string fails the whole batch; retry one by one.
          var one = Promise.resolve();
          pend.forEach(function (ed) {
            one = one.then(function () { return searchPass(ctx, para, [ed], variant).catch(function () {}); });
          });
          return one;
        });
      });
    });
    // Last resort: split the paragraph at spaces and address whole tokens.
    return chain.then(function () {
      var rest = todo.filter(function (ed) { return !ed.range; });
      if (!rest.length || typeof para.getTextRanges !== 'function') { return; }
      var tokens = para.getTextRanges([' '], false);
      tokens.load('text');
      return ctx.sync().then(function () {
        var offs = Core.mapTokenRanges(tokens.items.map(function (r) { return r.text; }), P);
        if (!offs) { return; }
        var picked = [];
        rest.forEach(function (ed) {
          var ij = Core.tokensCovering(offs, ed.a, ed.b);
          if (!ij) { return; }
          var A = offs[ij[0]][0], B = offs[ij[1]][1];
          var r = ij[0] === ij[1] ? tokens.items[ij[0]] : tokens.items[ij[0]].expandTo(tokens.items[ij[1]]);
          r.load('text');
          picked.push({ ed: ed, r: r, old: P.slice(A, B), neu: P.slice(A, ed.a) + ed.neu + P.slice(ed.b, B), a: A, b: B });
        });
        if (!picked.length) { return; }
        return ctx.sync().then(function () {
          picked.forEach(function (x) {
            // Strict on purpose: a token that carries the paragraph mark must never be replaced.
            if (x.r.text !== x.old) { return; }
            x.ed.range = x.r; x.ed.old = x.old; x.ed.neu = x.neu; x.ed.a = x.a; x.ed.b = x.b;
          });
        });
      }).catch(function () {});
    });
  }

  function guard(promise) {
    return promise.catch(function (err) {
      if (err && err.stale) {
        setStatus('متن پس از آخرین بازبینی عوض شده بود؛ دوباره بازبینی می‌شود.');
        schedule(50);
      } else {
        reportError(err);
      }
    });
  }

  function showInDocument(it) {
    if (S.applying) { return; }
    var f = it.f, p = S.paras[it.pi];
    Array.prototype.forEach.call(document.querySelectorAll('.card.focus'), function (c) { c.classList.remove('focus'); });
    var mine = document.querySelector('.card[data-key="' + cssEscape(it.key) + '"]');
    if (mine) { mine.classList.add('focus'); }
    return guard(Word.run(function (ctx) {
      return getParagraph(ctx, it.pi).then(function (para) {
        if (p.pending) { para.select(); return ctx.sync(); }
        var mid = p.text.slice(f.s, f.e);
        var edits = Core.planEdits(p.text, [{ s: f.s, e: f.e, rep: mid }], { keepNoop: true });
        return locate(ctx, para, p.text, edits).then(function () {
          var ed = edits[0];
          if (!ed || !ed.range) { para.select(); return ctx.sync(); }
          if (!mid.trim() || mid === ed.old || mid.length > Core.MAX_SEARCH || Core.occurrences(ed.old, mid).length !== 1) {
            ed.range.select();
            return ctx.sync();
          }
          // Narrow the selection from the whole words to the finding itself when that is unambiguous.
          var inner = ed.range.search(Core.searchVariants(mid)[0], { matchCase: true });
          inner.load('text');
          return ctx.sync().then(function () {
            var exact = inner.items.filter(function (r) { return r.text === mid; });
            (exact.length === 1 ? exact[0] : ed.range).select();
            return ctx.sync();
          }, function () {
            ed.range.select();
            return ctx.sync();
          });
        });
      });
    }));
  }

  function cssEscape(s) {
    return (window.CSS && CSS.escape) ? CSS.escape(s) : s.replace(/["\\]/g, '\\$&');
  }

  // ───────────────────────── applying fixes ─────────────────────────

  function beginTracking(ctx) {
    var noop = function () { return Promise.resolve(); };
    if (!S.has.tracking) { return Promise.resolve({ end: noop, tracked: false }); }
    var doc = ctx.document;
    doc.load('changeTrackingMode');
    return ctx.sync().then(function () {
      var prev = doc.changeTrackingMode;
      if (prev !== 'Off') { return { end: noop, tracked: true }; }
      if (!S.opts.track) { return { end: noop, tracked: false }; }
      doc.changeTrackingMode = 'TrackAll';
      return ctx.sync().then(function () {
        return {
          tracked: true,
          end: function () { doc.changeTrackingMode = prev; return ctx.sync(); }
        };
      });
    }, function () { return { end: noop, tracked: false }; });
  }

  function applyInParagraph(ctx, pi, its, ps, stats) {
    var p = S.paras[pi], para, getP;
    if (p.id) {
      para = ctx.document.getParagraphByUniqueLocalId(p.id);
      para.load('text');
      getP = ctx.sync();
    } else {
      para = ps.items[p.i];
      getP = Promise.resolve();
    }
    return getP.then(function () {
      if (!para || para.text !== p.raw || p.pending) { stats.stale += its.length; return; }
      var edits = Core.planEdits(p.text, its.map(function (it) { return { s: it.f.s, e: it.f.e, rep: it.f.rep }; }));
      return locate(ctx, para, p.text, edits).then(function () {
        var good = edits.filter(function (ed) { return ed.range; });
        edits.forEach(function (ed) { if (!ed.range) { stats.failed += ed.fix.length; } });
        if (!good.length) { return; }
        good.forEach(function (ed) { ed.range.insertText(ed.neu, 'Replace'); });
        var after = S.has.reviewed ? para.getReviewedText('Current') : null;
        para.load('text');
        return ctx.sync().then(function () {
          good.forEach(function (ed) { stats.done += ed.fix.length; });
          var actual = after ? after.value : para.text;
          if (actual !== Core.applyEdits(p.text, good)) { stats.odd++; }
        });
      });
    }).catch(function (err) {
      if (window.console) { console.error(err); }
      stats.failed += its.length;
    });
  }

  function applyItems(items) {
    if (S.applying || !items.length) { return Promise.resolve(); }
    S.applying = true;
    setStatus('در حال اصلاح…');
    var stats = { done: 0, failed: 0, stale: 0, odd: 0, tracked: false };
    var byPara = new Map();
    items.forEach(function (it) {
      if (it.f.rep == null) { return; }
      if (!byPara.has(it.pi)) { byPara.set(it.pi, []); }
      byPara.get(it.pi).push(it);
    });
    return Word.run(function (ctx) {
      return beginTracking(ctx).then(function (tr) {
        stats.tracked = tr.tracked;
        var ps = null, first = Promise.resolve();
        if (!S.has.ids) {
          ps = ctx.document.body.paragraphs;
          ps.load('text');
          first = ctx.sync();
        }
        var chain = first;
        byPara.forEach(function (its, pi) {
          chain = chain.then(function () { return applyInParagraph(ctx, pi, its, ps, stats); });
        });
        return chain.then(tr.end, function (err) { return tr.end().then(function () { throw err; }); });
      });
    }).catch(function (err) {
      reportError(err, 'اصلاح انجام نشد');
    }).then(function () {
      S.applying = false;
      var parts = [];
      if (stats.done) { parts.push(faNum(stats.done) + ' مورد درست شد' + (stats.tracked ? ' و در «ردگیری تغییرها» ثبت شد' : '')); }
      if (stats.failed) { parts.push(faNum(stats.failed) + ' مورد را نتوانستم با اطمینان در متن پیدا کنم و دست نزدم'); }
      if (stats.stale) { parts.push(faNum(stats.stale) + ' مورد در بندی بود که پس از بازبینی عوض شده است'); }
      if (stats.odd) { parts.push('در ' + faNum(stats.odd) + ' بند نتیجه با انتظار یکی نبود؛ همان بند را نگاه کنید و اگر لازم شد Ctrl+Z بزنید'); }
      if (parts.length) { setStatus(parts.join('؛ ') + '.', stats.failed || stats.odd ? 'err' : 'ok'); }
      S.dirty = false;
      return scanKeepingStatus(parts.length > 0);
    });
  }

  /** Rescan after an edit without wiping the message that says what the edit did. */
  function scanKeepingStatus(keep) {
    var s = $('status'), text = s.textContent, cls = s.className;
    return scan().then(function () {
      if (keep) { s.textContent = text; s.className = cls; }
    });
  }

  // ───────────────────────── underlines in the document (experimental) ─────────────────────────

  function critiqueOf(f) {
    var c = { colorScheme: LEVEL_COLOR[f.level] || 'Blue', start: f.s, length: f.e - f.s };
    if (S.has.popup && f.rep) {
      c.popupOptions = {
        brandingTextResourceId: 'Vir.Brand',
        titleResourceId: 'Vir.Title.' + f.level,
        subtitleResourceId: 'Vir.Sub',
        suggestions: [f.rep]
      };
    }
    return c;
  }

  function annotateParagraph(ctx, para, pi, clearAll) {
    var p = S.paras[pi];
    var fs = [];
    if (!clearAll && !p.pending) {
      fs = Core.nonOverlapping(findingsOf(pi).filter(function (f) {
        return S.filter[f.level] && !S.ignored.has(Core.keyOf(f.code, p.text, f.s, f.e));
      }));
    }
    var key = p.id || ('#' + p.i);
    var sig = fs.length ? p.text + '\u0001' + fs.map(function (f) { return f.code + ':' + f.s + ':' + f.e; }).join(',') : '';
    if ((S.annSig.get(key) || '') === sig) { return Promise.resolve(); }
    var old = para.getAnnotations();
    old.load('id');
    return ctx.sync().then(function () {
      old.items.forEach(function (a) { S.annMap.delete(a.id); a.delete(); });
      var res = fs.length ? para.insertAnnotations({ critiques: fs.map(critiqueOf) }) : null;
      return ctx.sync().then(function () {
        if (res) {
          res.value.forEach(function (id, k) {
            S.annMap.set(id, { pi: pi, key: Core.keyOf(fs[k].code, p.text, fs[k].s, fs[k].e) });
          });
        }
        S.annSig.set(key, sig);
      });
    });
  }

  function syncAnnotations(clearAll) {
    if (!S.has.annot) { return Promise.resolve(); }
    return Word.run(function (ctx) {
      var ps = ctx.document.body.paragraphs;
      ps.load('text');
      return ctx.sync().then(function () {
        var chain = Promise.resolve();
        S.paras.forEach(function (p, pi) {
          var para = ps.items[p.i];
          if (!para || para.text !== p.raw) { return; }
          chain = chain.then(function () { return annotateParagraph(ctx, para, pi, clearAll); });
        });
        return chain;
      });
    }).catch(function (err) {
      if (window.console) { console.error(err); }
      S.opts.underline = false;
      $('optUnderline').checked = false;
      saveOpts();
      $('note').dataset.sticky = '1';
      note('خط‌کشی در متن در این ورد کار نکرد و خاموش شد. این امکان به اشتراک مایکروسافت ۳۶۵ نیاز دارد؛ فهرستِ همین پنجره بی‌آن هم کامل است.');
    });
  }

  function itemOfAnnotation(id) {
    var m = S.annMap.get(id);
    if (!m) { return null; }
    for (var i = 0; i < S.items.length; i++) {
      if (S.items[i].pi === m.pi && S.items[i].key === m.key) { return S.items[i]; }
    }
    return null;
  }

  function registerAnnotationEvents() {
    if (S.annEvents || !S.has.annot) { return Promise.resolve(); }
    return Word.run(function (ctx) {
      ctx.document.onAnnotationClicked.add(function (args) {
        var it = itemOfAnnotation(args.id);
        var c = it && document.querySelector('.card[data-key="' + cssEscape(it.key) + '"]');
        if (c) {
          Array.prototype.forEach.call(document.querySelectorAll('.card.focus'), function (x) { x.classList.remove('focus'); });
          c.classList.add('focus');
          c.scrollIntoView({ block: 'center' });
        }
        return Promise.resolve();
      });
      if (S.has.popup) {
        ctx.document.onAnnotationPopupAction.add(function (args) {
          var it = itemOfAnnotation(args.id);
          if (it && args.action !== 'Accept') { S.ignored.add(it.key); }
          // If Word has already put the suggestion in, the paragraph no longer matches and
          // applyItems leaves it alone; otherwise the fix is applied here.
          var next = (it && args.action === 'Accept') ? applyItems([it]) : Promise.resolve();
          return next.then(function () { schedule(100); });
        });
      }
      return ctx.sync();
    }).then(function () { S.annEvents = true; }, function () {});
  }

  // ───────────────────────── live events ─────────────────────────

  function onDocChanged() {
    if (S.opts.live) { schedule(); }
    return Promise.resolve();
  }

  function followCursor() {
    if (S.scanning || S.applying || !S.paras.length) { return; }
    Word.run(function (ctx) {
      var para = ctx.document.getSelection().paragraphs.getFirst();
      para.load(S.has.ids ? 'uniqueLocalId' : 'text');
      return ctx.sync().then(function () {
        var pi = -1;
        for (var i = 0; i < S.paras.length; i++) {
          if (S.has.ids ? S.paras[i].id === para.uniqueLocalId : S.paras[i].raw === para.text) { pi = i; break; }
        }
        markCurrent(pi);
      });
    }).catch(function () {});
  }

  function onSelectionChanged() {
    clearTimeout(S.selTimer);
    S.selTimer = setTimeout(function () {
      if (!S.has.paraEvents && S.opts.live) { schedule(600); }
      if (S.opts.follow) { followCursor(); }
    }, 350);
  }

  function registerEvents() {
    var p = Promise.resolve();
    if (S.has.ids) {
      p = Word.run(function (ctx) {
        ctx.document.onParagraphChanged.add(onDocChanged);
        ctx.document.onParagraphAdded.add(onDocChanged);
        ctx.document.onParagraphDeleted.add(onDocChanged);
        return ctx.sync();
      }).then(function () { S.has.paraEvents = true; }, function () { S.has.paraEvents = false; });
    }
    return p.then(function () {
      try {
        Office.context.document.addHandlerAsync(Office.EventType.DocumentSelectionChanged, onSelectionChanged);
      } catch (e) { /* the list still refreshes with the button */ }
    });
  }

  // ───────────────────────── controls ─────────────────────────

  function bind() {
    $('btnScan').addEventListener('click', function () { S.limit = PAGE; setStatus('بازبینی…'); scan(); });
    $('btnSettings').addEventListener('click', function () {
      var box = $('settings');
      box.hidden = !box.hidden;
      this.setAttribute('aria-expanded', String(!box.hidden));
    });
    $('btnFixAll').addEventListener('click', function () { applyItems(S.items.filter(isBulk)); });

    Array.prototype.forEach.call(document.querySelectorAll('.chip'), function (chip) {
      chip.addEventListener('click', function () {
        var lv = chip.dataset.level;
        S.filter[lv] = !S.filter[lv];
        chip.setAttribute('aria-pressed', String(S.filter[lv]));
        S.limit = PAGE;
        render();
        if (S.opts.underline) { syncAnnotations(false); }
      });
    });

    var toggles = { optLive: 'live', optTrack: 'track', optBulkW: 'bulkW', optFollow: 'follow', optStrict: 'strict', optVerbless: 'verbless', optUnderline: 'underline' };
    Object.keys(toggles).forEach(function (id) {
      var box = $(id), key = toggles[id];
      box.checked = !!S.opts[key];
      box.addEventListener('change', function () {
        S.opts[key] = box.checked;
        saveOpts();
        if (key === 'strict' || key === 'verbless') { S.cache.clear(); scan(); }
        if (key === 'live' && box.checked) { scan(); }
        if (key === 'bulkW') { render(); }
        if (key === 'underline') {
          $('note').dataset.sticky = '';
          if (box.checked) {
            registerAnnotationEvents().then(function () { return syncAnnotations(false); });
          } else {
            syncAnnotations(true);
          }
        }
      });
    });
    if (!S.has.tracking) { $('optTrack').disabled = true; }
    if (!S.has.annot) { $('rowUnderline').style.display = 'none'; S.opts.underline = false; }

    $('houseJson').value = S.opts.house || '';
    $('btnHouseSave').addEventListener('click', function () {
      var raw = $('houseJson').value.trim(), msg = $('houseMsg');
      if (!raw) { return clearHouse(); }
      rpc('house', { json: raw }).then(function (problem) {
        if (problem) {
          msg.textContent = 'پذیرفته نشد: ' + problem;
          msg.className = 'hint err';
          return;
        }
        S.opts.house = raw;
        S.house = JSON.parse(raw);
        saveOpts();
        msg.textContent = 'ذخیره شد.';
        msg.className = 'hint';
        S.cache.clear();
        scan();
      });
    });
    $('btnHouseClear').addEventListener('click', clearHouse);
  }

  function clearHouse() {
    $('houseJson').value = '';
    $('houseMsg').textContent = 'شیوهٔ خانگی برداشته شد.';
    $('houseMsg').className = 'hint';
    S.opts.house = '';
    S.house = null;
    saveOpts();
    S.cache.clear();
    scan();
  }

  // ───────────────────────── start ─────────────────────────

  function start(info) {
    if (!info || info.host !== 'Word' || typeof Word === 'undefined') {
      $('outside').style.display = 'block';
      return;
    }
    $('app').style.display = '';
    var sup = function (v) {
      try { return Office.context.requirements.isSetSupported('WordApi', v); } catch (e) { return false; }
    };
    S.has = { tracking: sup('1.4'), reviewed: sup('1.4'), ids: sup('1.6'), annot: sup('1.7'), popup: sup('1.8'), paraEvents: false };
    loadOpts();
    bind();
    startEngine().then(function (engine) {
      S.engine = engine;
      S.ready = true;
      $('about').textContent = 'fa_lint ' + engine.version + ' · ' + faNum(engine.rules) + ' قاعده';
      return registerEvents();
    }).then(function () {
      if (S.opts.underline) { registerAnnotationEvents(); }
      return scan();
    }).catch(function (err) {
      reportError(err, 'موتور بازبینی بارگیری نشد. اینترنت را بررسی کنید و پنجره را دوباره باز کنید');
    });
  }

  window.VirastarPane = { state: S, scan: scan, applyItems: applyItems, showInDocument: showInDocument };

  if (typeof Office === 'undefined') {
    start(null);
  } else {
    Office.onReady(start);
  }
}());
