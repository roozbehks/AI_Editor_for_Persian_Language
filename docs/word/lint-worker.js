/*
 * Virastar lint worker.
 * Runs fa_lint.py, unchanged, inside Pyodide so the task pane stays responsive.
 * The document text never leaves this browser: there is no server behind the add-in.
 *
 * Messages in:  {id, cmd: 'init', args: {pyodideBase, lintUrl}}
 *               {id, cmd: 'lint', args: {texts: [..], opts: {house, strict, verbless}}}
 *               {id, cmd: 'cons', args: {texts: [..]}}
 *               {id, cmd: 'house', args: {json: '...'}}
 * Messages out: {id, ok: true, result} or {id, ok: false, error}
 */
/* global importScripts, loadPyodide */
'use strict';

var DEFAULT_PYODIDE = 'https://cdn.jsdelivr.net/pyodide/v0.26.4/full/';
var py = null;

var GLUE = [
  'import json, re, sys',
  'sys.path.insert(0, "/home/pyodide")',
  'import fa_lint',
  '',
  'def _u16(text):',
  '    """Python counts code points, JavaScript counts UTF-16 units. Map only when they differ."""',
  '    if all(ord(c) < 0x10000 for c in text):',
  '        return None',
  '    m, n = [0], 0',
  '    for c in text:',
  '        n += 2 if ord(c) >= 0x10000 else 1',
  '        m.append(n)',
  '    return m',
  '',
  'def _pack(f, m):',
  '    code, level, s, e, msg, rep = f',
  '    if m:',
  '        s, e = m[s], m[e]',
  '    auto = rep is not None and (level in "EW" or code in fa_lint.AUTOFIX_C)',
  '    return [code, level, s, e, msg, rep, auto]',
  '',
  'def run_lint(texts_json, opts_json):',
  '    texts = json.loads(texts_json)',
  '    o = json.loads(opts_json)',
  '    cfg = o.get("house") or {}',
  '    out = []',
  '    for t in texts:',
  '        m = _u16(t)',
  '        fs = fa_lint.lint(t, cfg, bool(o.get("strict")), bool(o.get("verbless")), False)',
  '        out.append([_pack(f, m) for f in fs])',
  '    return json.dumps(out, ensure_ascii=False)',
  '',
  'def run_cons(texts_json):',
  '    """Spelling consistency across the whole document; one finding per pair of forms."""',
  '    texts = json.loads(texts_json)',
  '    starts, pos = [], 0',
  '    for t in texts:',
  '        starts.append(pos)',
  '        pos += len(t) + 1',
  '    joined = "\\n".join(texts)',
  '    out = []',
  '    for f in fa_lint.consistency_findings(fa_lint.mask(joined)):',
  '        i = max(k for k, st in enumerate(starts) if st <= f[2])',
  '        local = (f[0], f[1], f[2] - starts[i], f[3] - starts[i], f[4], None)',
  '        out.append([i, _pack(local, _u16(texts[i]))])',
  '    return json.dumps(out, ensure_ascii=False)',
  '',
  'def check_house(js):',
  '    """Empty string when the house style compiles, otherwise the reason it does not."""',
  '    try:',
  '        cfg = json.loads(js)',
  '        if not isinstance(cfg, dict):',
  '            return "باید یک شیء JSON باشد"',
  '        fa_lint.house_rules(cfg)',
  '        fa_lint.lint("آزمون", cfg)',
  '        return ""',
  '    except Exception as ex:',
  '        return str(ex)',
  ''
].join('\n');

function init(args) {
  var base = (args && args.pyodideBase) || DEFAULT_PYODIDE;
  if (base.charAt(base.length - 1) !== '/') { base += '/'; }
  importScripts(base + 'pyodide.js');
  return loadPyodide({ indexURL: base }).then(function (p) {
    py = p;
    return fetch(args.lintUrl, { cache: 'no-cache' });
  }).then(function (r) {
    if (!r.ok) { throw new Error('fa_lint.py: HTTP ' + r.status); }
    return r.text();
  }).then(function (src) {
    py.FS.mkdirTree('/home/pyodide');
    py.FS.writeFile('/home/pyodide/fa_lint.py', src);
    py.runPython(GLUE);
    return { version: py.runPython('fa_lint.__version__'), rules: py.runPython('len(fa_lint.RULES)') };
  });
}

function call(name, a, b) {
  var fn = py.globals.get(name);
  try {
    return b === undefined ? fn(a) : fn(a, b);
  } finally {
    fn.destroy();
  }
}

self.onmessage = function (ev) {
  var m = ev.data || {};
  var done = function (result) { self.postMessage({ id: m.id, ok: true, result: result }); };
  var fail = function (err) { self.postMessage({ id: m.id, ok: false, error: String((err && err.message) || err) }); };
  try {
    if (m.cmd === 'init') {
      init(m.args).then(done, fail);
    } else if (!py) {
      fail('not-ready');
    } else if (m.cmd === 'lint') {
      done(JSON.parse(call('run_lint', JSON.stringify(m.args.texts), JSON.stringify(m.args.opts || {}))));
    } else if (m.cmd === 'cons') {
      done(JSON.parse(call('run_cons', JSON.stringify(m.args.texts))));
    } else if (m.cmd === 'house') {
      done(call('check_house', m.args.json));
    } else {
      fail('unknown command: ' + m.cmd);
    }
  } catch (err) {
    fail(err);
  }
};
