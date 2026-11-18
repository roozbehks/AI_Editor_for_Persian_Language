# ویراستار هوشواره

**AI Editor for Persian Language** — rules, a linter and a Word add-in that make AI-written (and human-written) Persian read like edited Persian: Academy orthography, half-spaces, punctuation, and the sentence and wording faults Persian editors correct. Everything runs locally; the docs are in Persian. Try it in the browser at <https://roozbehks.github.io/AI_Editor_for_Persian_Language/>.

---

هوش مصنوعی فارسی را روان می‌نویسد، ولی ویراسته نمی‌نویسد: نیم‌فاصله جا می‌افتد، جمله بوی ترجمه می‌دهد و «می‌باشد» و «توسط» همه‌جا هست. این مخزن قاعده‌های ویراستاران فارسی را به شکلی درآورده است که هوشواره و آدم هر دو بتوانند به کارش ببندند، بی‌آنکه هر بار شیوه‌نامه‌ها از نو خوانده شود.

## چهار راه استفاده

| راه | برای چه کسی | از کجا |
|---|---|---|
| **دستور کوتاه** | هر هوشواره‌ای: یک صفحه قاعده که در دستورهای سفارشی می‌گذارید | [`persian-writing/PROMPT.md`](persian-writing/PROMPT.md) |
| **مهارت Claude** | کسی که با Claude می‌نویسد: قاعده‌های هسته، شش راهنما و ابزار وارسی در یک بسته | پوشهٔ [`persian-writing/`](persian-writing) |
| **ابزار وارسی** | هر کس پایتون ۳ دارد، یا فقط یک مرورگر | [`fa_lint.py`](persian-writing/scripts/fa_lint.py) یا [صفحهٔ پروژه](https://roozbehks.github.io/AI_Editor_for_Persian_Language/#try) |
| **افزونهٔ ورد** | کسی که در ورد مایکروسافت می‌نویسد و پیشنهاد همزمان می‌خواهد | [`word-addin/`](word-addin) |

### دستور کوتاه

متن درون `PROMPT.md` را در دستورهای سفارشی هوشواره‌تان یا در آغاز گفت‌وگو بگذارید.

### مهارت Claude

پوشهٔ `persian-writing` را فشرده کنید، یا بستهٔ آماده را از [صفحهٔ پروژه](https://roozbehks.github.io/AI_Editor_for_Persian_Language/#download) بردارید، و در Claude از بخش Customize و سپس Skills بارگذاری کنید. در Claude Code پوشه را در `.claude/skills/` بگذارید. شرح کامل در [`persian-writing/README.md`](persian-writing/README.md) است.

### ابزار وارسی

```bash
python3 persian-writing/scripts/fa_lint.py متن.txt                  # گزارش
python3 persian-writing/scripts/fa_lint.py متن.txt --fix --inplace   # اصلاح خودکار خطاهای مکانیکی
python3 persian-writing/scripts/fa_lint.py متن.txt --house my.json   # با شیوهٔ خانگی
python3 persian-writing/scripts/fa_lint.py --list-rules             # فهرست قاعده‌ها
```

فقط کتابخانهٔ استاندارد پایتون را می‌خواهد. پرونده‌های `txt` و `md` و `html` را کامل می‌خواند، از `py` و `json` فقط رشته‌های فارسی را، و `docx` را اگر `python-docx` نصب باشد. سه سطح دارد: **E** خطای مکانیکی که خودکار درست می‌شود، **W** آنچه به‌احتمال زیاد باید عوض شود، و **C** نامزدی که داوری می‌خواهد.

### افزونهٔ ورد

پنجره‌ای کناری در ورد که هنگام نوشتن فهرست یافته‌ها را تازه می‌کند، جای هر مورد را در متن نشان می‌دهد و اصلاحش را اعمال می‌کند؛ اگر بخواهید با «ردگیری تغییرها». راهنمای نصب و محدودیت‌ها در [`word-addin/README.md`](word-addin/README.md) است. این افزونه هنوز در وردِ واقعی آزموده نشده است.

## اولویت با فرهنگستان

مرجع نخست این مجموعه فرهنگستان زبان و ادب فارسی است: «دستور خط فارسی»، «فرهنگ املایی خط فارسی» و واژه‌های مصوب.

- جایی که فرهنگستان یک صورت دارد همان می‌آید: «فناوری»، «ایدئال»، «ویدئو»، «حتی»، «به‌عنوان»، «اینکه».
- جایی که فرهنگستان دو صورت را برابر می‌داند، ابزار فقط ناهمسانی درون یک متن را نشان می‌دهد: «فرایند» و «فرآیند»، «پایین» و «پائین»، «گفت‌وگو» و «گفتگو».
- شیوهٔ خانگیِ هر نویسنده و ناشر جای خالیِ فرهنگستان را پر می‌کند. اگر کسی آگاهانه از فرهنگستان جدا می‌شود، باید همان را صریح در شیوهٔ خانگی‌اش بنویسد.
- کتاب‌های درست‌نویسی و شیوه‌نامهٔ ناشران پس از این‌ها می‌آیند.

فهرست منابع و جاهایی که با هم نمی‌خوانند در [`references/manabe.md`](persian-writing/references/manabe.md) است. قاعده‌ها از منابع گرفته و از نو نوشته شده است؛ متن خودِ منابع در این مخزن نیست.

## در مخزن چه هست

| پوشه | کار |
|---|---|
| [`persian-writing/`](persian-writing) | مهارت: `SKILL.md`، شش راهنما در `references/`، ابزارها در `scripts/`، و `PROMPT.md` |
| [`word-addin/`](word-addin) | مانیفست افزونهٔ ورد و راهنمای نصب |
| [`docs/`](docs) | صفحهٔ پروژه برای GitHub Pages، و پرونده‌های افزونهٔ ورد در `docs/word/` |
| [`tools/`](tools) | ساخت صفحه و مانیفست |
| [`tests/`](tests) | آزمون منطق افزونه |

`docs/index.html` و `docs/kit.json` و `docs/fa_lint.py` و `docs/word/manifest.xml` ساخته می‌شوند. پس از هر تغییر در `persian-writing/` یا مانیفست این را اجرا کنید:

```bash
pip install markdown
python3 tools/build_site.py
```

## راه‌اندازی صفحهٔ پروژه

صفحهٔ پروژه و افزونهٔ ورد هر دو از GitHub Pages داده می‌شوند. یک بار باید روشنش کرد: در مخزن به **Settings** و سپس **Pages** بروید، در **Build and deployment** گزینهٔ **Deploy from a branch** را برگزینید، شاخهٔ `main` و پوشهٔ `/docs` را انتخاب کنید و **Save** را بزنید. چند دقیقه بعد صفحه در نشانی بالا باز می‌شود.

## آزمون

```bash
python3 persian-writing/scripts/fa_lint.py --selftest    # قاعده‌های ابزار
node --test tests/core.test.js                           # منطق متنیِ افزونه
python3 tests/addin/run_tests.py                         # پنجرهٔ افزونه؛ پیش‌نیازها در سرِ همان پرونده
```

ابزار روی سه مجموعه سنجیده شده است. در ۴۴ متن منتشرشده، حدود ۱۲۰ هزار نویسه، خطای مکانیکیِ نادرست گزارش نکرد. در ۸۲۹ جملهٔ نادرست مطبوعاتی که صورت درستشان هم در دست بود، در حدود ۱۵ درصد نشانهٔ خطای نگارشی را گرفت؛ بقیه خطای معنا و منطق بود. و در ۳۳۱ جفت «پیش و پس» از ویرایش‌های یک نویسنده نشان داد که ویرایش سبکی را قاعدهٔ نوشتاری می‌گیرد، نه ابزار. پس ابزار تور ایمنی است و جای ویراستار را نمی‌گیرد.

## پروانه

پروانهٔ انتشار هنوز برگزیده نشده است. تا آن زمان همهٔ حقوق برای صاحب مخزن محفوظ است.

## سپاس

گردآوری و تدوین: RKSVisions، با همکاری Claude. قاعده‌ها وام‌دار فرهنگستان زبان و ادب فارسی و ویراستارانی است که نامشان در فهرست منابع آمده است.
