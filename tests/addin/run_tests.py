#!/usr/bin/env python3
"""Integration tests for the Word task pane, run against a stand-in for Office.js.

They load the real task pane in headless Chrome, with the real fa_lint.py running in
Pyodide, and replace only Office.js by tests/addin/mock-office.js. They prove the logic of
the task pane; they do not prove anything about the real Word.

Setup (once):
    pip install playwright && playwright install chromium
    cd tests/addin && npm install
Run from the repository root:
    python3 tests/addin/run_tests.py
Environment: PYODIDE_DIR (default tests/addin/node_modules/pyodide), CHROME (a Chrome or
Chromium binary; default: Playwright's own).
"""
import asyncio
import functools
import http.server
import json
import os
import socketserver
import sys
import tempfile
import threading

from playwright.async_api import async_playwright

HERE = os.path.dirname(os.path.abspath(__file__))
REPO = os.path.dirname(os.path.dirname(HERE))
PYODIDE = os.environ.get('PYODIDE_DIR', os.path.join(HERE, 'node_modules', 'pyodide'))
CHROME = os.environ.get('CHROME')
OFFICE_JS = 'https://appsforoffice.microsoft.com/lib/1/hosted/office.js'

sys.path.insert(0, os.path.join(REPO, 'persian-writing', 'scripts'))
import fa_lint as F  # noqa: E402

ZW = '‌'
DOC = [
    'اين كتاب ها را می خوانم و میروم .',
    'لازم به ذکر است که این فرایند زمان بر می باشد.',
    'می رود و باز می رود و می' + ZW + 'ماند.',
    '',
    'متن درست و بی' + ZW + 'خطا.',
    'سلام 😀 دوست من ، حتما بیا.',
]


def fix_level(text, levels, cfg=None):
    """What repeated presses of the bulk button should leave behind."""
    for _ in range(6):
        fs = [f for f in F.lint(text, cfg or {}, False, False, False)
              if f[5] is not None and (f[1] in levels or f[0] in F.AUTOFIX_C)]
        new, n = F.apply_fixes(text, fs)
        if not n or new == text:
            break
        text = new
    return text


def count(text, level, cfg=None):
    return sum(1 for f in F.lint(text, cfg or {}, False, False, False) if f[1] == level)


class Quiet(http.server.SimpleHTTPRequestHandler):
    def log_message(self, *a):
        pass


def serve():
    root = tempfile.mkdtemp(prefix='virastar-test-')
    os.symlink(os.path.join(REPO, 'docs'), os.path.join(root, 'docs'))
    os.symlink(PYODIDE, os.path.join(root, '__pyodide'))
    handler = functools.partial(Quiet, directory=root)
    socketserver.TCPServer.allow_reuse_address = True
    httpd = socketserver.ThreadingTCPServer(('127.0.0.1', 0), handler)
    threading.Thread(target=httpd.serve_forever, daemon=True).start()
    return httpd, 'http://127.0.0.1:%d' % httpd.server_address[1]


class Run:
    def __init__(self):
        self.passed = 0
        self.failed = []

    def check(self, name, cond, detail=''):
        if cond:
            self.passed += 1
            print('  ok   ', name)
        else:
            self.failed.append(name)
            print('  FAIL ', name, '|', detail)


async def open_pane(browser, base, mock, errors):
    ctx = await browser.new_context(viewport={'width': 360, 'height': 900})
    page = await ctx.new_page()
    page.on('pageerror', lambda e: errors.append('PAGEERROR ' + str(e)))
    page.on('console', lambda m: errors.append(m.text) if m.type == 'error' and 'favicon' not in m.text else None)
    mock_js = open(os.path.join(HERE, 'mock-office.js'), encoding='utf-8').read()

    async def route(r):
        url = r.request.url
        if url == OFFICE_JS:
            await r.fulfill(status=200, body=mock_js, content_type='text/javascript')
        elif url.startswith('http://127.0.0.1'):
            await r.continue_()
        elif 'fonts.googleapis.com' in url:
            await r.fulfill(status=200, body='', content_type='text/css')   # the tests run offline
        else:
            await r.abort()
    await page.route('**/*', route)
    await page.add_init_script('window.__MOCK__ = ' + json.dumps(mock, ensure_ascii=False) + ';')
    await page.goto(base + '/docs/word/taskpane.html?pyodide=/__pyodide/')
    return ctx, page


async def settle(page, ms=1400):
    """Wait until no scan or edit is running and the debounce timers have fired."""
    await page.wait_for_timeout(ms)
    await page.wait_for_function('() => window.VirastarPane && !VirastarPane.state.scanning && !VirastarPane.state.applying', timeout=60000)
    await page.wait_for_timeout(250)


async def ready(page):
    await page.wait_for_function('() => window.VirastarPane && VirastarPane.state.ready', timeout=120000)
    await settle(page)


async def press_bulk(page, times=6):
    for _ in range(times):
        if await page.locator('#btnFixAll').is_disabled():
            break
        await page.click('#btnFixAll')
        await settle(page)


async def scenario_default(t, browser, base):
    print('default: Word with every capability, Find that ignores the half-space')
    errors = []
    ctx, page = await open_pane(browser, base, {'paragraphs': DOC}, errors)
    await ready(page)
    if os.environ.get('SHOT'):
        await page.screenshot(path=os.environ['SHOT'], full_page=True)
    cards = await page.locator('.card').count()
    want = sum(len(F.lint(p, {}, False, False, False)) for p in DOC if p.strip()) + len(F.consistency_findings(F.mask('\n'.join(DOC))))
    t.check('lists every finding of fa_lint', cards == want, '%d cards, fa_lint %d' % (cards, want))
    nE = await page.inner_text('#nE')
    t.check('error count is shown in Persian digits', nE == str(sum(count(p, 'E') for p in DOC)).translate(F.FA_DIGITS), nE)

    # Offsets after an emoji: the finding must sit on the comma, not two characters before it.
    key = await page.evaluate('''() => { const it = VirastarPane.state.items.find(i => i.pi === 5 && i.f.code === 'SP-PUNC');
        return it ? VirastarPane.state.paras[5].text.slice(it.f.s, it.f.e) : null; }''')
    t.check('offsets survive a character outside the BMP', key == ' ،', repr(key))

    # Click a card: the words are selected in the document.
    await page.locator('.card', has_text='CH-AR').first.click()
    await settle(page, 500)
    sel = await page.evaluate('() => __mock.doc.selectedText')
    t.check('clicking a card selects the text in the document', sel in ('ي', 'اين', 'ك', 'كتاب'), repr(sel))

    # Apply one fix.
    before = await page.evaluate('() => __mock.texts()[0]')
    await page.locator('.card', has_text='CH-AR').first.locator('button.apply').click()
    await settle(page)
    after = await page.evaluate('() => __mock.texts()[0]')
    t.check('applying one fix changes only that character', after != before and after.replace('ی', 'ي', 1) == before or after.replace('ک', 'ك', 1) == before, after)

    # A fix inside a repeated phrase lands on the right occurrence.
    it = await page.evaluate('''() => { const s = VirastarPane.state; const p = s.paras[2].text; const second = p.indexOf('می رود', 5);
        const item = s.items.find(i => i.pi === 2 && i.f.code === 'ZW-MI' && i.f.s >= second); if (!item) return null;
        return VirastarPane.applyItems([item]).then(() => __mock.texts()[2]); }''')
    t.check('a repeated phrase is fixed at the occurrence that was chosen', it == 'می رود و باز می' + ZW + 'رود و می' + ZW + 'ماند.', repr(it))

    # The bulk button, pressed until nothing mechanical is left.
    await settle(page)
    await press_bulk(page)
    texts = await page.evaluate('() => __mock.texts()')
    expected = [fix_level(p, 'E') for p in DOC]
    t.check('bulk fix leaves exactly what fa_lint --fix leaves for level E', texts == expected, json.dumps([a for a, b in zip(texts, expected) if a != b], ensure_ascii=False))
    t.check('no mechanical error remains', all(count(x, 'E') == 0 for x in texts))
    t.check('paragraphs were not merged or lost', len(texts) == len(DOC))

    # Live: typing into the document refreshes the list.
    await page.evaluate("() => __mock.type(4, 'جملهٔ تازه با كاف عربی.')")
    await settle(page, 1800)
    live = await page.locator('.group[data-pi="4"] .card').count()
    t.check('typing in the document refreshes the list by itself', live >= 1, str(live))
    await page.evaluate("() => __mock.moveCursor(4)")
    await settle(page, 900)
    t.check('the list follows the cursor', await page.locator('.group.current[data-pi="4"]').count() == 1)

    # Level filter and ignore.
    n_before = await page.locator('.card').count()
    await page.click('.chip.lv-W')
    n_filtered = await page.locator('.card').count()
    await page.click('.chip.lv-W')
    t.check('a level chip hides and shows its cards', n_filtered < n_before and await page.locator('.card').count() == n_before)
    await page.locator('.card button.ghost').first.click()
    t.check('«نادیده» removes one card', await page.locator('.card').count() == n_before - 1)

    # House style from the settings panel.
    await page.click('#btnSettings')
    await page.fill('#houseJson', json.dumps({'words': {'جمله': 'گزاره'}}, ensure_ascii=False))
    await page.click('#btnHouseSave')
    await settle(page)
    t.check('a pasted house style adds its own findings', await page.locator('.card', has_text='HOUSE').count() >= 1)
    await page.fill('#houseJson', '{"prefer": [{"pattern": "("}]}')
    await page.click('#btnHouseSave')
    await page.wait_for_timeout(600)
    t.check('a broken house style is refused with a reason', 'پذیرفته نشد' in await page.inner_text('#houseMsg'))
    t.check('no script errors', not errors, '; '.join(errors)[:300])
    await ctx.close()


async def scenario_search(t, browser, base, mode, extra=None):
    print('Find flavour: %s%s' % (mode, ' + paragraph mark inside the last token' if extra else ''))
    errors = []
    mock = {'paragraphs': DOC, 'search': mode}
    mock.update(extra or {})
    ctx, page = await open_pane(browser, base, mock, errors)
    await ready(page)
    await press_bulk(page)
    texts = await page.evaluate('() => __mock.texts()')
    expected = [fix_level(p, 'E') for p in DOC]
    calls = await page.evaluate('() => __mock.doc.calls')
    if extra:
        t.check('nothing is merged when a token carries the paragraph mark', len(texts) == len(DOC), str(len(texts)))
        t.check('what could not be located safely is reported, not forced', 'نتوانستم' in await page.inner_text('#status') or texts == expected)
    else:
        t.check('bulk fix reaches the same text', texts == expected, json.dumps([a for a, b in zip(texts, expected) if a != b], ensure_ascii=False))
    if mode == 'broken':
        t.check('falls back to token ranges when Find returns nothing', calls['tokens'] > 0, str(calls))
    t.check('no script errors', not errors, '; '.join(errors)[:300])
    await ctx.close()


async def scenario_old(t, browser, base):
    print('older Word: WordApi 1.3 only (no ids, no paragraph events, no change tracking API)')
    errors = []
    ctx, page = await open_pane(browser, base, {'paragraphs': DOC, 'maxSet': '1.3'}, errors)
    await ready(page)
    t.check('finds things without the newer APIs', await page.locator('.card').count() > 5)
    t.check('track-changes switch is disabled', await page.locator('#optTrack').is_disabled())
    t.check('underline switch is hidden', not await page.locator('#rowUnderline').is_visible())
    await press_bulk(page)
    texts = await page.evaluate('() => __mock.texts()')
    t.check('bulk fix works by paragraph position', texts == [fix_level(p, 'E') for p in DOC])
    await page.evaluate("() => { __mock.doc.paras[4].raw = 'متن تازه با كاف.'; __mock.doc.paras[4].cur = __mock.doc.paras[4].raw; __mock.moveCursor(4); }")
    await settle(page, 2200)
    t.check('a moved cursor triggers the rescan when paragraph events are missing', await page.locator('.group[data-pi="4"] .card').count() >= 1)
    t.check('no script errors', not errors, '; '.join(errors)[:300])
    await ctx.close()


async def scenario_tracking(t, browser, base):
    print('track changes')
    errors = []
    ctx, page = await open_pane(browser, base, {'paragraphs': DOC}, errors)
    await ready(page)
    await page.click('#btnSettings')
    await page.check('#optTrack')
    await press_bulk(page, times=1)
    raw = await page.evaluate('() => __mock.texts()')
    cur = await page.evaluate('() => __mock.current()')
    one_pass = []
    for p in DOC:
        fs = [f for f in F.lint(p, {}, False, False, False) if f[5] is not None and (f[1] == 'E' or f[0] in F.AUTOFIX_C)]
        one_pass.append(F.apply_fixes(p, fs)[0])
    t.check('tracked edits give the same reviewed text', cur == one_pass, json.dumps([a for a, b in zip(cur, one_pass) if a != b], ensure_ascii=False))
    t.check('the old words stay in the document as tracked deletions', raw != cur)
    t.check('change tracking is switched back off afterwards', await page.evaluate('() => __mock.doc.tracking') == 'Off')
    t.check('paragraphs with pending changes are flagged', await page.evaluate('() => VirastarPane.state.paras.some(p => p.pending)'))
    t.check('the note about pending changes is shown', 'ردگیری' in await page.inner_text('#note'))
    await press_bulk(page, times=1)
    t.check('paragraphs with pending changes are not edited again', await page.evaluate('() => __mock.texts()') == raw)
    t.check('no script errors', not errors, '; '.join(errors)[:300])
    await ctx.close()


async def scenario_underline(t, browser, base):
    print('underlines in the document')
    for replaces in (False, True):
        errors = []
        ctx, page = await open_pane(browser, base, {'paragraphs': DOC, 'popupReplaces': replaces}, errors)
        await ready(page)
        await page.click('#btnSettings')
        await page.check('#optUnderline')
        await settle(page, 2500)
        anns = await page.evaluate('() => __mock.annotations()')
        ok = all(a['text'] == DOC[a['para']].encode('utf-16-le')[a['critique']['start'] * 2:(a['critique']['start'] + a['critique']['length']) * 2].decode('utf-16-le') for a in anns)
        t.check('critiques are inserted at the right offsets', len(anns) > 5 and ok, str(len(anns)))
        target = next(a for a in anns if a['para'] == 0 and a['critique'].get('popupOptions') and a['text'] == 'ي')
        await page.evaluate('(id) => __mock.popup(id, "Accept")', target['id'])
        await settle(page, 1500)
        first = await page.evaluate('() => __mock.texts()[0]')
        t.check('accepting a popup suggestion fixes it exactly once (Word %s the text itself)' % ('replaces' if replaces else 'does not replace'),
                first == DOC[0].replace('ي', 'ی', 1), first)
        t.check('no script errors', not errors, '; '.join(errors)[:300])
        await ctx.close()
    errors = []
    ctx, page = await open_pane(browser, base, {'paragraphs': DOC, 'annotations': False}, errors)
    await ready(page)
    await page.click('#btnSettings')
    await page.check('#optUnderline')
    await settle(page, 2000)
    t.check('when Word refuses annotations the switch turns itself off', not await page.locator('#optUnderline').is_checked())
    t.check('and says why', 'اشتراک' in await page.inner_text('#note'))
    t.check('the list keeps working', await page.locator('.card').count() > 5)
    await ctx.close()


async def scenario_outside(t, browser, base):
    print('opened outside Word')
    errors = []
    ctx, page = await open_pane(browser, base, {'host': None}, errors)
    await page.wait_for_timeout(800)
    t.check('says it must be opened inside Word', await page.locator('#outside').is_visible() and not await page.locator('#app').is_visible())
    await ctx.close()


async def main():
    if not os.path.exists(os.path.join(PYODIDE, 'pyodide.js')):
        sys.exit('Pyodide not found in %s (see the top of this file).' % PYODIDE)
    httpd, base = serve()
    t = Run()
    async with async_playwright() as p:
        kw = {'args': ['--no-sandbox']}
        if CHROME:
            kw['executable_path'] = CHROME
        browser = await p.chromium.launch(**kw)
        await scenario_default(t, browser, base)
        await scenario_search(t, browser, base, 'literal')
        await scenario_search(t, browser, base, 'caret')
        await scenario_search(t, browser, base, 'broken')
        await scenario_search(t, browser, base, 'broken', {'tokensIncludeMark': True})
        await scenario_old(t, browser, base)
        await scenario_tracking(t, browser, base)
        await scenario_underline(t, browser, base)
        await scenario_outside(t, browser, base)
        await browser.close()
    httpd.shutdown()
    print('\n%d passed, %d failed' % (t.passed, len(t.failed)))
    for name in t.failed:
        print('  failed:', name)
    sys.exit(1 if t.failed else 0)


if __name__ == '__main__':
    asyncio.run(main())
