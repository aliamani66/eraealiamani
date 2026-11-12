<div dir="rtl" style="direction: rtl; text-align: right; font-family: IRANYekanX, Vazirmatn, 'Segoe UI', Tahoma, sans-serif; line-height: 1.8;">
<style>pre { direction: ltr; text-align: left; white-space: pre-wrap; } code { direction: ltr; unicode-bidi: isolate; }</style>

# پرامپت‌های تصویر اسلایدها — ارائه‌ی «BREAKING CHANGES»

- این فایل فهرست پرامپت‌هایی است که با آن‌ها تصویر پس‌زمینه‌ی اسلایدها را می‌سازید: ۲۵ تصویر، برای هر کدام یک پرامپت. ۲۲ تای اول ساخته شده‌اند و در `new/slides/` هستند؛ سه تای آخر (`23_colonnade`، `24_two_chairs` و `25_student_desk`) برای اسلایدهای تازه‌اند و هنوز باید ساخته شوند.
- توضیح‌ها فارسی‌اند، اما خود پرامپت‌ها انگلیسی‌اند، چون مدل‌های تصویرساز با انگلیسی بهتر کار می‌کنند.
- هر پرامپت خودکفاست (سبک و قاعده‌ی «بدون متن» داخل خودش تکرار شده)؛ هر بار فقط همان یکی را کپی کنید و در مدل تصویرساز بچسبانید.
- تصویر را با نسبت `16:9` و دست‌کم `1920×1080` پیکسل بسازید و دقیقاً با نام همان بخش در پوشه‌ی `new/slides/` ذخیره کنید (`.jpg` یا `.png`، هر دو درست است).
- بعد از ذخیره‌ی هر چند تصویر، از ریشه‌ی ریپو دستور `.venv\Scripts\python.exe new\build_deck.py` را اجرا کنید تا دک دوباره ساخته شود (قبلش فایل را در پاورپوینت ببندید).
- اسلایدی که تصویرش هنوز نیست هم ساخته می‌شود و پس‌زمینه‌ی تیره‌ی طراحی‌شده می‌گیرد؛ پس می‌توانید تصویرها را کم‌کم و به هر ترتیبی بسازید.
- پیشنهاد: اول `01_cover` را بسازید و بقیه را با آن هم‌سبک کنید؛ دلیلش در بخش بعد آمده است.

## قاعده‌های مشترک

- **یک دنیای بصری:** همه‌ی تصویرها یک عکس ثابت، تاریک، آرام و سینمایی از یک شیء یا یک صحنه‌ی ساده‌اند؛ فوتورئال و مینیمال، با فضای خالی فراوان. نه علمی‌تخیلی، نه نئون و سایبرپانک، نه ایلاستریشن و صحنه‌ی شلوغ؛ هولوگرام، ربات و برد مدار درخشان هم نداریم.
- **کینتسوگی:** کینتسوگی یعنی ترمیم چیز شکسته با درز طلای واقعی. حرف ارائه هم همین است: خرابی‌ها (اینسیدنت‌ها و تغییرهای هویتی) ما را بهتر کردند. پس هر چیزی که در تصویر ترک خورده یا ترمیم شده، با طلای صیقلی پر شده است؛ در پرامپت‌ها این را توصیف کرده‌ایم و فقط به اسم کینتسوگی بسنده نکرده‌ایم.
- **مواد، نور و رنگ:** سطح‌های مشکی مات، ابسیدین و سنگ اسلیت تیره؛ یک نور گرم اصلی و سایه‌های عمیق؛ طلای واقعی (نه رنگ زرد) تنها رنگ پرجان است. نور سرد (سیان مایل به سفید) فقط جایی است که صفحه‌نمایش یا چراغ دیتاسنتر خود موضوع باشد؛ رنگ‌های دیگر فقط همان استثناهایی‌اند که در بخش خودشان آمده (آبی غروب، سپیده‌دم، چراغ‌های کهربایی).
- **دوربین:** عکس واقعی با لنز ۵۰ تا ۸۵ میلی‌متر، عمق میدان کم و دانه‌ی ظریف فیلم.
- **جای متن:** اسلایدها راست‌به‌چپ‌اند و متن سمت راست می‌نشیند؛ پس مگر خلافش نوشته شده باشد، سوژه در یک‌سوم چپ است و دوسوم راست تاریک، خالی و کم‌کنتراست می‌ماند. فقط `12_handover` (دست‌ها پایین و وسط) و `22_gold_network` (الگو روی کل کادر) استثنا هستند.
- **ممنوع‌ها:** هیچ نوشته‌ای در تصویر نباشد (حرف، عدد، لوگو، رابط کاربری، برچسب، صفحه‌ی خوانا، عطف کتاب)؛ چهره‌ی قابل‌تشخیص نباشد (آدم فقط به شکل دست یا سایه‌ای دور و از پشت)؛ و هیچ صحنه‌ی خشن یا هیجانی نباشد. همه‌ی متن‌ها را بعداً در پاورپوینت، به‌صورت متن واقعی، اضافه می‌کنید.

بلوک زیر همان STYLE مشترک است: خط اول سبک و خط دوم قاعده‌ی «بدون متن». این دو خط به انتهای هر پرامپت چسبانده شده‌اند، پس لازم نیست جدا بچسبانید؛ اینجا برای مرجع است و برای وقتی که بخواهید سبک همه‌ی تصویرها را یک‌جا عوض کنید. فقط در هشت صحنه‌ی باز (راهرو، دفتر، تالار، پل، اتاق عملیات، پنجره) به‌جای `still-life` و لنز ۸۵ میلی‌متری، لنز ۳۵ میلی‌متری آمده تا کادر باز بماند.

```text
Photorealistic cinematic still-life photograph, 85mm lens, shallow depth of field, fine film grain, matte black and dark slate materials, one warm key light, deep shadows, real polished gold as the only saturated colour, minimalist, 16:9 landscape.
Absolutely no text: no letters, numbers, logos, labels or watermarks anywhere, and no faces.
```

اگر مدل فیلدی برای negative prompt دارد، بلوک NEGATIVE را در آن بگذارید؛ اگر ندارد (مثل بیشتر چت‌بات‌های تصویرساز) لازم نیست، چون قاعده‌ی «بدون متن» داخل خود پرامپت‌هاست. واژه‌ی people را عمداً در آن نیاورده‌ایم، چون دو تصویر دست یا سایه‌ی آدم دارند (`12_handover` و `14_dusk_office`).

```text
text, letters, numbers, words, captions, logos, brand names, watermark, signature, user interface, readable screens, labels, signs, faces, portraits, crowds, neon, cyberpunk, sci-fi, hologram, robot, glowing circuit board, illustration, cartoon, 3D render, CGI, oversaturated colours, yellow paint, cluttered busy composition, fire, explosion, smoke, rubble, blood, distorted hands, extra fingers
```

**ترتیب کار:** اول `01_cover` را بسازید و تا وقتی از جنس سنگ، طلا و نورش راضی نیستید جلو نروید. اگر مدل «تصویر مرجع» (style / image reference) می‌پذیرد، همین تصویر را برای ساخت همه‌ی تصویرهای بعدی به‌عنوان مرجع سبک بدهید تا مواد و نور در کل ارائه یک‌دست بماند. اگر مدل ویرایش تصویر (edit / variation) دارد، `22_gold_network` را هم از روی `01_cover` بسازید تا دقیقاً همان سنگ باشد.

**اندازه:** اگر مدل مستقیماً `1920×1080` نمی‌دهد، همان `16:9` را در بزرگ‌ترین اندازه‌ی ممکن بسازید و با ابزار upscale به دست‌کم `1920×1080` برسانید.

**ایراد رایج در همه‌ی تصویرها:** اگر مدل سوژه را وسط گذاشت، آخر پرامپت اضافه کنید `subject small and placed in the left third, the right two-thirds empty dark negative space`؛ اگر آدم یا چهره اضافه کرد، `no people, no faces, nobody present`.

## جدول خلاصه

| # | فایل | اسلاید | توضیح یک‌خطی |
|---|---|---|---|
| ۱ | `01_cover.jpg` | ۲ و ۴۷ | سنگ مشکی با یک ترک طلایی؛ عنوان و پایان |
| ۲ | `02_java.jpg` | ۵ | فنجان مشکی سالم کنار کیبورد؛ پنج سال دولوپر جاوا |
| ۳ | `03_door.jpg` | ۶ | در سنگین دیتاسنتر که نیمه‌باز مانده؛ رفتن به عملیات |
| ۴ | `04_queue.jpg` | ۷ | ردیف صندلی‌های خالی تا یک در بسته؛ صف پشت در عملیات |
| ۵ | `05_key.jpg` | ۸ و ۹ | کلید برنجی روی سنگی با ترک طلایی؛ یک خط اشتباه |
| ۶ | `06_scale.jpg` | ۱۰ | ترازوی برنجی کج با تیر ترمیم‌شده با طلا؛ لودبالانسر |
| ۷ | `07_manuals.jpg` | ۱۲ | پشته‌ی کتاب‌های فنی کهنه و دیکشنری زیر چراغ مطالعه؛ مستندات وب‌لاجیک |
| ۸ | `08_aisle.jpg` | ۱۳ و ۱۵ | راهروی دیتاسنتر که تاریکی به چراغ‌های سبز دور می‌رسد؛ خاموشی ۷۰ سرور |
| ۹ | `09_pipes.jpg` | ۱۶ | دیوار لوله‌های مشکی موازی با درزهای طلایی؛ Git و پایپ‌لاین |
| ۱۰ | `10_lens.jpg` | ۱۷ | پرتوهای پراکنده که از یک عدسی می‌گذرند و یک پرتوی طلایی می‌شوند؛ لاگ یک‌جا |
| ۱۱ | `11_workbench.jpg` | ۱۹ | میز کار با پنج جعبه‌ی یکسان زیر نور گرم؛ مسیر پنج‌مرحله‌ای R&D |
| ۱۲ | `12_handover.jpg` | ۲۰ | دو دست و یک باتوم مشکی با درز طلایی؛ بساز، تحویل بده، رها کن |
| ۱۳ | `13_pillar.jpg` | ۲۱ | تخته‌سنگی عظیم روی یک ستون باریک؛ SPOF |
| ۱۴ | `14_dusk_office.jpg` | ۲۲ | دفتر خالی در غروب با یک میز روشن و سایه‌ی یک نفر؛ همکار جدید |
| ۱۵ | `15_bridge.jpg` | ۲۹ | پلی که تقریباً ساخته شده و فقط دهانه‌ی آخرش مانده؛ سه سال آماده‌سازی |
| ۱۶ | `16_blackout.jpg` | ۳۰ | تالار تاریک دیتاسنتر با چند چراغ اضطراری کهربایی؛ از دست رفتن زیرساخت |
| ۱۷ | `17_night_ops.jpg` | ۳۱ | میز بلند اتاق عملیات در شب با مانیتورهای روشن؛ سه شب مهاجرت |
| ۱۸ | `18_padlock.jpg` | ۳۲ | قفل برنجی کوچک روی در عظیم گاوصندوق؛ یک دسترسی |
| ۱۹ | `19_dawn.jpg` | ۳۳ و ۳۴ | پنجره‌ی اتاق عملیات پیش از طلوع با اولین خط نارنجی؛ انتظار بار صبح |
| ۲۰ | `20_graph.jpg` | ۳۵ و ۳۷ | شبکه‌ای از سیم و مهره‌ی طلا روی مشکی؛ ISR و ایجنت‌ها |
| ۲۱ | `21_box.jpg` | ۴۲ | جعبه‌ی مشکی بسته با نور طلایی از زیر درش؛ نسخه‌ی بعدی |
| ۲۲ | `22_gold_network.jpg` | ۴۶ | شبکه‌ی ترک‌های طلایی روی همان سنگ؛ هر شکستی چیزی اضافه کرد |
| ۲۳ | `23_colonnade.jpg` | ۲۵ | همان تخته‌سنگِ اسلاید SPOF، این بار روی هفت ستون؛ مدیر شدن (تازه) |
| ۲۴ | `24_two_chairs.jpg` | ۲۷ | دو صندلی خالی روبه‌روی هم زیر یک چراغ؛ واحد کار عوض شد (تازه) |
| ۲۵ | `25_student_desk.jpg` | ۴۱ | یک صندلی دسته‌دار کلاس با دفتری سفید؛ دوباره مبتدی شدن (تازه) |

## پرامپت‌ها

### `01_cover.jpg`

**اسلاید:** ۲ (عنوان) و ۴۷ (پرسش و پاسخ، کم‌نور)

**کاربرد:** مخاطب باید از همان ثانیه‌ی اول حس کند چیزی شکسته و با طلا ترمیم شده است؛ این تصویر لحن بصری کل ارائه را تعیین می‌کند.

**پرامپت:** (اول این را بسازید و برای بقیه مرجع قرار دهید.)

```text
A wide slab of matte black obsidian stone seen from slightly above, split by exactly one jagged crack filled with polished gold, like kintsugi repair. The crack runs through the left third of the frame, from the top-left corner down toward the bottom-centre, glinting under a warm raking light. The right two-thirds is smooth, empty, dark stone. Photorealistic cinematic still-life photograph, 85mm lens, shallow depth of field, fine film grain, matte black and dark slate materials, one warm key light, deep shadows, real polished gold as the only saturated colour, minimalist, 16:9 landscape. Absolutely no text: no letters, numbers, logos, labels or watermarks anywhere, and no faces.
```

**ترکیب‌بندی:** ترک در یک‌سوم چپ از گوشه‌ی بالا به سمت پایین-وسط می‌رود؛ دوسوم راست باید سنگ تیره و خالی بماند، چون عنوان آنجا می‌نشیند.

**اگر خوب درنیامد:** معمولاً مدل چند ترک می‌سازد یا ترک را وسط کادر می‌گذارد؛ به پرامپت اضافه کنید `only one single crack, no other cracks, the right two-thirds completely empty`.

### `02_java.jpg`

**اسلاید:** ۵ (نسخه‌ی ۱.۰.۰؛ پنج سال دولوپر جاوا)

**کاربرد:** حس پنج سال آرام و پایدار: همه‌چیز سالم است و هیچ ترکی نیست، و فنجان قهوه همان شوخی جاواست.

**پرامپت:**

```text
An intact matte black ceramic coffee cup, perfectly smooth with no cracks, a gentle curl of steam rising, on a dark desk in the left third of the frame. Beside it, softly out of focus, the edge of a keyboard with blank unmarked keys. A warm golden desk-lamp glow lights the cup. The right two-thirds is dark, empty desk. Photorealistic cinematic still-life photograph, 85mm lens, shallow depth of field, fine film grain, matte black and dark slate materials, one warm key light, deep shadows, real polished gold as the only saturated colour, minimalist, 16:9 landscape. Absolutely no text: no letters, numbers, logos, labels or watermarks anywhere, and no faces.
```

**ترکیب‌بندی:** فنجان در یک‌سوم چپ و لبه‌ی محو کیبورد کنارش است؛ دوسوم راست میز تاریک و خالی می‌ماند.

**اگر خوب درنیامد:** مدل معمولاً روی کلیدهای کیبورد حرف و عدد می‌نویسد یا به فنجان ترک و خط طلایی می‌دهد؛ اضافه کنید `keyboard keys completely blank and unmarked, the cup perfectly smooth and whole with no gold lines`.

### `03_door.jpg`

**اسلاید:** ۶ (رفتن از توسعه به عملیات؛ خیلی کم‌نور می‌شود)

**کاربرد:** آستانه‌ی یک تصمیم بزرگ را نشان می‌دهد: دری سنگین که نیمه‌باز است و پشتش ناشناخته.

**پرامپت:**

```text
A heavy brushed-steel datacenter door stands slightly ajar at the far end of a long dark corridor, in the left third. A thin blade of cold cyan-white light, the only cold light, spills through the gap across the floor; a brass handle catches a faint gold highlight. Corridor walls fall into deep shadow; the right two-thirds is dark and empty. Photorealistic cinematic photograph, 35mm lens, fine film grain, matte black and dark slate materials, one warm key light, deep shadows, real polished gold as the only saturated colour, minimalist, 16:9 landscape. Absolutely no text: no letters, numbers, logos, labels or watermarks anywhere, and no faces.
```

**ترکیب‌بندی:** در در یک‌سوم چپ و انتهای راهرو است؛ دیوار و کف سمت راست تاریک و خالی می‌ماند.

**اگر خوب درنیامد:** مدل راهرو را آبی نئونی و علمی‌تخیلی می‌سازد یا در را وسط می‌گذارد؛ اضافه کنید `realistic plain corridor, no neon, no sci-fi glow, the door small and far away in the left third`.

### `04_queue.jpg`

**اسلاید:** ۷ (صف پشت در عملیات برای کوئری و لاگ)

**کاربرد:** مخاطب باید حس صف و معطلی را ببیند: صندلی‌های خالی جای آدم‌هایی‌اند که پشت یک در بسته منتظرند.

**پرامپت:**

```text
A long row of identical empty matte black waiting chairs lined along a dim corridor wall, receding in strong one-point perspective toward the left, where a single closed door waits at the end with a thin line of warm gold light glowing under it. The nearest chairs and the whole right side of the frame fall into deep, empty darkness. Photorealistic cinematic photograph, 35mm lens, fine film grain, matte black and dark slate materials, one warm key light, deep shadows, real polished gold as the only saturated colour, minimalist, 16:9 landscape. Absolutely no text: no letters, numbers, logos, labels or watermarks anywhere, and no faces.
```

**ترکیب‌بندی:** چشم‌انداز به سمت چپ فرو می‌رود و در بسته در انتهای چپ است؛ نزدیک‌ترین صندلی‌ها و سمت راست در تاریکی می‌مانند.

**اگر خوب درنیامد:** ممکن است آدم‌ها روی صندلی‌ها بنشینند یا راهرو قرینه و وسط‌چین شود؛ اضافه کنید `all chairs completely empty, no people, vanishing point in the left third, right half in darkness`.

### `05_key.jpg`

**اسلاید:** ۸ و ۹ (یک خط پسورد غلط و یک ساعت قطعی بانک؛ کم‌نور)

**کاربرد:** یک شیء کوچک روی سطحی ترک‌خورده، یعنی کوچک‌ترین تغییر دستی هم می‌تواند چیز بزرگی را بشکند.

**پرامپت:**

```text
Macro close-up of a single old brass key with a smooth, plain bow, lying on a matte black stone surface in the lower-left of the frame, catching a warm highlight. A thin crack filled with polished gold, like kintsugi repair, runs across the stone beneath it. The upper and right parts of the frame stay dark and empty. Photorealistic cinematic still-life photograph, 85mm lens, shallow depth of field, fine film grain, matte black and dark slate materials, one warm key light, deep shadows, real polished gold as the only saturated colour, minimalist, 16:9 landscape. Absolutely no text: no letters, numbers, logos, labels or watermarks anywhere, and no faces.
```

**ترکیب‌بندی:** کلید پایین-چپ و ترک طلایی زیر آن است؛ بالا و راست تاریک و خالی می‌ماند.

**اگر خوب درنیامد:** کلید بزرگ و وسط‌چین می‌شود یا روی آن حکاکی و نوشته می‌افتد؛ اضافه کنید `small key in the lower-left corner, smooth bow with no engraving, the rest of the frame empty`.

### `06_scale.jpg`

**اسلاید:** ۱۰ (تغییر الگوریتم لودبالانسر روی سیستم زنده)

**کاربرد:** تعادلی که با یک تغییر به هم خورد؛ درز طلایی یعنی همین خطا بخشی از قصه‌ی بهتر شدن است.

**پرامپت:**

```text
A small antique brass balance scale tipped hard to one side: the pan holding a black pebble rests low on a matte black stone surface, the other pan hangs high and empty. One arm of the beam is repaired with a seam of polished gold, like kintsugi repair. The scale stands in the left third; the right two-thirds is dark and empty. Photorealistic cinematic still-life photograph, 85mm lens, shallow depth of field, fine film grain, matte black and dark slate materials, one warm key light, deep shadows, real polished gold as the only saturated colour, minimalist, 16:9 landscape. Absolutely no text: no letters, numbers, logos, labels or watermarks anywhere, and no faces.
```

**ترکیب‌بندی:** ترازوی کوچک در یک‌سوم چپ است؛ بقیه‌ی کادر تاریک و خالی می‌ماند.

**اگر خوب درنیامد:** ترازو متعادل و قرینه کشیده می‌شود یا خیلی بزرگ است؛ اضافه کنید `the beam tilted at a steep angle, one pan touching the stone, the other pan high in the air, small scale in the left third`.

### `07_manuals.jpg`

**اسلاید:** ۱۲ (خواندن دو سه هزار صفحه مستند وب‌لاجیک با دیکشنری)

**کاربرد:** شب‌های طولانی یادگیری با سماجت، نه با استعداد؛ تلاشی آرام و بی‌ادعا.

**پرامپت:**

```text
A tall, slightly leaning stack of thick, worn technical manuals with plain blank spines and unmarked covers, in the left third, on a dark desk at night. Beside it lies a small open pocket dictionary with blank, blurred pages and gilded edges catching the light. One brass desk lamp casts a pool of warm gold light; the right two-thirds is deep shadow, empty. Photorealistic cinematic still-life photograph, 85mm lens, shallow depth of field, fine film grain, matte black and dark slate materials, one warm key light, deep shadows, real polished gold as the only saturated colour, minimalist, 16:9 landscape. Absolutely no text: no letters, numbers, logos, labels or watermarks anywhere, and no faces.
```

**ترکیب‌بندی:** پشته‌ی کتاب‌ها و دیکشنری در یک‌سوم چپ‌اند؛ دوسوم راست در سایه و خالی می‌ماند.

**اگر خوب درنیامد:** روی عطف کتاب‌ها و صفحه‌ی دیکشنری متن بی‌معنی می‌آید؛ اضافه کنید `book spines completely plain and blank, no lettering, no titles, dictionary pages blank and out of focus`.

### `08_aisle.jpg`

**اسلاید:** ۱۳ و ۱۵ (حلقه‌ی for که ۷۰ سرور را یکی‌یکی خاموش کرد؛ کم‌نور)

**کاربرد:** تاریکی ردیف‌به‌ردیف جلو می‌آید و چراغ‌های سبز دور آخرین نشانه‌های زنده‌ی سیستم‌اند.

**پرامپت:**

```text
A long datacenter aisle between tall server racks in strict one-point perspective, vanishing point just left of centre. The foreground racks are completely dark and unlit; far down the aisle the last racks still show tiny green status lights, as if the darkness is advancing toward them. Polished floor, realistic, no neon. The right side sinks into deep shadow. Photorealistic cinematic photograph, 35mm lens, fine film grain, matte black and dark slate materials, one warm key light, deep shadows, real polished gold as the only saturated colour, minimalist, 16:9 landscape. Absolutely no text: no letters, numbers, logos, labels or watermarks anywhere, and no faces.
```

**ترکیب‌بندی:** نقطه‌ی گریز کمی چپ‌تر از وسط است؛ راک‌های جلو کاملاً تاریک و سمت راست در سایه‌ی عمیق می‌ماند.

**اگر خوب درنیامد:** مدل تونلی آبی نئونی و علمی‌تخیلی می‌سازد یا همه‌ی چراغ‌ها را روشن می‌کند؛ اضافه کنید `realistic datacenter, foreground racks pitch black with no lights at all, only tiny green LEDs far away, no blue neon`.

### `09_pipes.jpg`

**اسلاید:** ۱۶ (اسکریپت‌ها به Git و پایپ‌لاین جنکینز آمدند؛ زیر متن کم‌نور)

**کاربرد:** نظم و تکرارپذیری: لوله‌ها موازی و تمیزند و درزهایشان را طلا ترمیم کرده، یعنی نظمی که از دل خرابی بیرون آمد.

**پرامپت:**

```text
A wall of orderly, perfectly parallel matte black industrial pipes running horizontally across the whole frame, evenly spaced, each with a small brass valve, every joint traced with a thin seam of polished gold, like kintsugi repair. Clean, symmetrical, unmarked, soft warm light grazing the surface, an even low-contrast texture from edge to edge, no focal point. Photorealistic cinematic still-life photograph, 85mm lens, shallow depth of field, fine film grain, matte black and dark slate materials, one warm key light, deep shadows, real polished gold as the only saturated colour, minimalist, 16:9 landscape. Absolutely no text: no letters, numbers, logos, labels or watermarks anywhere, and no faces.
```

**ترکیب‌بندی:** بافتی یکنواخت روی کل کادر و بدون کانون؛ چون زیر متن می‌رود، کنتراست باید ملایم بماند.

**اگر خوب درنیامد:** روی لوله‌ها برچسب و عدد می‌آید یا لوله‌ها درهم و شلوغ می‌شوند؛ اضافه کنید `plain unmarked pipes, no stencils, no labels, no gauges, strictly parallel and evenly spaced`.

### `10_lens.jpg`

**اسلاید:** ۱۷ (لاگ‌های همه‌ی سرورها در یک‌جا)

**کاربرد:** هرج‌ومرج لاگ‌های پراکنده از یک عدسی می‌گذرد و تبدیل به یک پرتوی روشن و خوانا می‌شود.

**پرامپت:**

```text
Dozens of thin scattered rays of light enter from the far left, converge through one glass lens on a black stand in the left third, and leave it as a single clean golden beam travelling right across a black background, fading into darkness. Faint haze makes the rays visible. Real glass, not a digital graphic. The right half stays black and empty. Photorealistic cinematic still-life photograph, 85mm lens, shallow depth of field, fine film grain, matte black and dark slate materials, one warm key light, deep shadows, real polished gold as the only saturated colour, minimalist, 16:9 landscape. Absolutely no text: no letters, numbers, logos, labels or watermarks anywhere, and no faces.
```

**ترکیب‌بندی:** عدسی در یک‌سوم چپ است و پرتوی طلایی به سمت راست می‌رود و محو می‌شود.

**اگر خوب درنیامد:** تصویر شبیه گرافیک دیجیتال یا لیزر نئونی می‌شود، یا پرتو تا لبه‌ی راست می‌رسد؛ اضافه کنید `real glass lens on a stand, soft natural haze, the beam fades out before the right edge, no lens flare, no neon`.

### `11_workbench.jpg`

**اسلاید:** ۱۹ (مسیر پنج‌مرحله‌ای R&D؛ زیر نمودار کم‌نور)

**کاربرد:** حس نظم و تکرارپذیری: هر فناوری از همین میز و از همین پنج مرحله‌ی یکسان می‌گذرد.

**پرامپت:**

```text
A dark workshop bench of old black-stained wood with exactly five identical small matte black boxes in an evenly spaced straight row across the middle of the frame, lit by one narrow strip of warm gold light above them. Each lid is edged with a hairline of polished gold. Calm, orderly, plain dark background above and below. Photorealistic cinematic still-life photograph, 85mm lens, shallow depth of field, fine film grain, matte black and dark slate materials, one warm key light, deep shadows, real polished gold as the only saturated colour, minimalist, 16:9 landscape. Absolutely no text: no letters, numbers, logos, labels or watermarks anywhere, and no faces.
```

**ترکیب‌بندی:** پنج جعبه در یک ردیف افقی در میانه‌ی کادرند و تصویر یکنواخت و آرام است، چون زیر نمودار می‌رود.

**اگر خوب درنیامد:** تعداد جعبه‌ها چهار یا شش می‌شود یا روی جعبه‌ها نوشته می‌آید؛ اضافه کنید `exactly five identical boxes in one straight row, count them, blank lids, nothing else on the bench`.

### `12_handover.jpg`

**اسلاید:** ۲۰ (بساز، تحویل بده، رها کن)

**کاربرد:** لحظه‌ی تحویل و رها کردن، آرام و با اعتماد: چیزی که ساخته‌ایم از دستی به دستی دیگر می‌رود.

**پرامپت:**

```text
Two bare hands handing over a matte black relay baton, the baton marked only by a single thin seam of polished gold, like kintsugi repair. The hands and baton fill only the bottom third of the frame, centred, lit by a warm rim light from behind. No watches, no rings, no sleeves. The upper two-thirds is pure dark, empty space. Photorealistic cinematic still-life photograph, 85mm lens, shallow depth of field, fine film grain, matte black and dark slate materials, one warm key light, deep shadows, real polished gold as the only saturated colour, minimalist, 16:9 landscape. Absolutely no text: no letters, numbers, logos, labels or watermarks anywhere, and no faces.
```

**ترکیب‌بندی:** استثنا: دست‌ها در یک‌سوم پایین و وسط کادرند و دوسوم بالا باید کاملاً تاریک و خالی بماند، چون متن بزرگ و وسط‌چین آنجا می‌نشیند.

**اگر خوب درنیامد:** دست‌ها بد درمی‌آیند (انگشت اضافه)، خیلی بزرگ می‌شوند یا ساعت مچی با عدد می‌گیرند؛ اضافه کنید `exactly two hands with five fingers each, bare wrists, no watch, hands small in the lower third of the frame`.

### `13_pillar.jpg`

**اسلاید:** ۲۱ (یک آدم به‌عنوان تنها نقطه‌ی شکست، SPOF)

**کاربرد:** سنگینی همه‌چیز روی یک ستون باریک، یعنی همان «قلقش دست منه» که به SPOF تبدیل می‌شود.

**پرامپت:**

```text
A massive slab of dark stone rests on one single thin stone column in a vast dark hall, the column in the left third of the frame, the far end of the slab vanishing into darkness. A beam of warm gold light falls from above onto the column, a hairline crack filled with polished gold running down it. The right two-thirds is dark and empty. Photorealistic cinematic still-life photograph, 85mm lens, shallow depth of field, fine film grain, matte black and dark slate materials, one warm key light, deep shadows, real polished gold as the only saturated colour, minimalist, 16:9 landscape. Absolutely no text: no letters, numbers, logos, labels or watermarks anywhere, and no faces.
```

**ترکیب‌بندی:** ستون در یک‌سوم چپ است و نور از بالا می‌تابد؛ دوسوم راست تاریک و خالی می‌ماند.

**اگر خوب درنیامد:** مدل چند ستون می‌سازد یا تالاری شبیه معبد؛ اضافه کنید `exactly one single thin column supports the slab, no other columns, the hall otherwise empty`.

### `14_dusk_office.jpg`

**اسلاید:** ۲۲ (نسخه‌ی ۳.۱.۰؛ عصری که همکار جدید پایپ‌لاین‌ها را سریع‌تر و تمیزتر بازنویسی می‌کرد)

**کاربرد:** حس دیدن کار خودت در دست دیگری، بدون دشمنی: یک غروب آرام که در آن چیزی عوض می‌شود.

**پرامپت:**

```text
An empty open-plan office at dusk, rows of dark desks, tall windows glowing deep blue. In the mid-distance, left of centre, one desk with a blurred, unreadable glowing monitor and the silhouette of a person working, seen from behind. An empty chair stands out of focus in the foreground. Cool blue dusk light, one warm amber desk lamp. The lower-right is dark. Photorealistic cinematic photograph, 35mm lens, fine film grain, matte black and dark slate materials, one warm key light, deep shadows, real polished gold as the only saturated colour, minimalist, 16:9 landscape. Absolutely no text: no letters, numbers, logos, labels or watermarks anywhere, and no faces.
```

**ترکیب‌بندی:** میز روشن کمی چپ‌تر از وسط و دورتر است و صندلی خالی در پیش‌زمینه؛ گوشه‌ی پایین-راست تاریک و خالی می‌ماند.

**اگر خوب درنیامد:** چهره، آدم رو به دوربین یا نوشته روی مانیتور می‌آید؛ اضافه کنید `the person only a dark silhouette seen strictly from behind, monitor a soft blank glow with nothing readable`.

### `15_bridge.jpg`

**اسلاید:** ۲۹ (نسخه‌ی ۵.۰.۰؛ سه سال آماده‌سازی، قبل از مهاجرت)

**کاربرد:** سه سال کارِ واقعی که بیشترِ پل را ساخته است؛ فقط دهانه‌ی آخر، یعنی خودِ مهاجرت، مانده.

**پرامپت:**

```text
A nearly finished steel bridge enters from the left edge of the frame and reaches across a deep foggy gap at night, stopping a few metres short of the far side, the final span missing. Warm amber construction lights glow along the deck and dissolve into thick fog. The gap, the mist and the far cliff on the right are dark, quiet and empty. Photorealistic cinematic photograph, 35mm lens, fine film grain, matte black and dark slate materials, one warm key light, deep shadows, real polished gold as the only saturated colour, minimalist, 16:9 landscape. Absolutely no text: no letters, numbers, logos, labels or watermarks anywhere, and no faces.
```

**ترکیب‌بندی:** پل از سمت چپ وارد کادر می‌شود؛ شکاف و مه در سمت راست تاریک و خالی می‌مانند.

**اگر خوب درنیامد:** مدل پل را کامل می‌کند و دو طرف شکاف را به هم می‌رساند؛ اضافه کنید `the bridge clearly ends in mid-air with a visible empty gap of several metres before the far side`.

### `16_blackout.jpg`

**اسلاید:** ۳۰ (خرداد ۱۴۰۴؛ کل زیرساخت از بین رفت)

**کاربرد:** سکوت و تاریکی جدی و باوقار: نه دراما و نه ترس، فقط حس از دست رفتن.

**پرامپت:**

```text
A large datacenter hall in near-total darkness, long rows of silent server racks fading into black, lit only by a few small amber emergency lights high on the walls and a faint warm reflection on the polished floor. Perfectly still air, intact and clean: no fire, no smoke, no rubble, no people. Sober and silent. The right side of the frame is the darkest. Photorealistic cinematic photograph, 35mm lens, fine film grain, matte black and dark slate materials, one warm key light, deep shadows, real polished gold as the only saturated colour, minimalist, 16:9 landscape. Absolutely no text: no letters, numbers, logos, labels or watermarks anywhere, and no faces.
```

**ترکیب‌بندی:** چراغ‌های کهربایی کم و کوچک‌اند؛ سمت راست تاریک‌ترین بخش کادر می‌ماند.

**اگر خوب درنیامد:** مدل دود، آتش، آوار یا آدم اضافه می‌کند؛ اضافه کنید `intact undamaged hall, clean polished floor, no smoke, no fire, no debris, nobody present` و در بخش negative هم `people` را بگذارید.

### `17_night_ops.jpg`

**اسلاید:** ۳۱ (سه شب مهاجرت)

**کاربرد:** خستگی و تعهد شب‌های طولانی کار، بدون قهرمان‌سازی؛ جای خالی آدم‌ها خودش حرف می‌زند.

**پرامپت:**

```text
A long desk in an operations room at night, a row of monitors glowing cold cyan-white with only soft, blurred, unreadable light, several used coffee cups, and a dark jacket hanging over an empty chair. A single brass desk lamp adds a pool of warm gold light. Nobody is present. Wide, evenly composed, low-contrast, tired late-night mood, the far monitors dissolving into soft focus. Photorealistic cinematic photograph, 35mm lens, fine film grain, matte black and dark slate materials, one warm key light, deep shadows, real polished gold as the only saturated colour, minimalist, 16:9 landscape. Absolutely no text: no letters, numbers, logos, labels or watermarks anywhere, and no faces.
```

**ترکیب‌بندی:** کادر یکنواخت و بدون کانون قوی است، چون زیر متن کم‌نور می‌شود.

**اگر خوب درنیامد:** روی مانیتورها نمودار و متن بی‌معنی می‌آید یا آدمی با چهره پشت میز می‌نشیند؛ اضافه کنید `monitors show only a soft blurred glow with no readable content, empty room, nobody present`.

### `18_padlock.jpg`

**اسلاید:** ۳۲ (بدترین لحظه: یک دسترسی که نزدیک بود همه‌چیز را متوقف کند)

**کاربرد:** تناقض تلخ و بامزه‌ی ماجرا: سه سال آماده‌سازی، و مانع آخر یک قفل کوچک.

**پرامپت:**

```text
A tiny brass padlock hanging from a small steel hasp on an enormous, smooth matte black vault door that fills the whole frame, a narrow spotlight from above striking only the padlock so the brass glows warm gold. The padlock sits in the left third. The vast door has no dial, no wheel, no markings and fades into darkness; the right two-thirds is dark and empty. Photorealistic cinematic still-life photograph, 85mm lens, shallow depth of field, fine film grain, matte black and dark slate materials, one warm key light, deep shadows, real polished gold as the only saturated colour, minimalist, 16:9 landscape. Absolutely no text: no letters, numbers, logos, labels or watermarks anywhere, and no faces.
```

**ترکیب‌بندی:** قفل در یک‌سوم چپ و زیر نور نقطه‌ای است؛ بقیه‌ی در عظیم در تاریکی فرو می‌رود و سمت راست خالی می‌ماند.

**اگر خوب درنیامد:** روی در گاوصندوق صفحه‌ی عدد و چرخ ظاهر می‌شود؛ اضافه کنید `smooth plain door surface, no dial, no wheel, no numbers, no engraving, the padlock small in the left third`.

### `19_dawn.jpg`

**اسلاید:** ۳۳ و ۳۴ (۰۶:۰۰ تا ۰۶:۳۰؛ اول کم‌نور، بعد روشن‌تر)

**کاربرد:** انتظار پیش از طلوع و بار صبحی که هنوز نیامده: امیدی کم‌رنگ، نه جشن پیروزی.

**پرامپت:**

```text
View from inside a dark operations room toward a wide window at pre-dawn: deep blue sky with the first thin orange line of light on the horizon above a distant city of tiny dim lights. Dark silhouettes of monitors and chair backs line the bottom edge of the frame. The horizon sits in the lower half; the upper half is calm, smooth, empty sky. Photorealistic cinematic photograph, 35mm lens, fine film grain, matte black and dark slate materials, one warm key light, deep shadows, real polished gold as the only saturated colour, minimalist, 16:9 landscape. Absolutely no text: no letters, numbers, logos, labels or watermarks anywhere, and no faces.
```

**ترکیب‌بندی:** افق در نیمه‌ی پایین است و سیلوئت مانیتورها لبه‌ی پایین را می‌گیرد؛ نیمه‌ی بالا آسمان آرام و خالی می‌ماند.

**اگر خوب درنیامد:** مدل خورشید، ماه یا ستاره اضافه می‌کند یا آسمان را زیادی نارنجی می‌کند؛ اضافه کنید `only one thin orange line on the horizon, the sky otherwise deep blue and empty, no sun disc, no moon, no stars`.

### `20_graph.jpg`

**اسلاید:** ۳۵ (ISR؛ پلتفرم ارکستریشن با ایجنت‌های AI) و ۳۷ (دو عدد قبل و بعد؛ خیلی کم‌نور)

**کاربرد:** هماهنگی چیزهای زیاد مثل نخ‌هایی که به هم بسته شده‌اند: آینده‌نگر، اما دست‌ساز و گرم، نه علمی‌تخیلی.

**پرامپت:**

```text
Fine gold wire threads strung between tiny polished gold beads that glow softly, forming a branching network on a black velvet surface. The network is dense on the left and thins out to a few lines and beads toward the right, where the black stays empty. Macro photograph, warm rim light on the wire, real metal objects, not a digital graphic, no screen, no interface. Photorealistic cinematic still-life photograph, 85mm lens, shallow depth of field, fine film grain, matte black and dark slate materials, one warm key light, deep shadows, real polished gold as the only saturated colour, minimalist, 16:9 landscape. Absolutely no text: no letters, numbers, logos, labels or watermarks anywhere, and no faces.
```

**ترکیب‌بندی:** شبکه در سمت چپ متراکم است و به سمت راست رقیق می‌شود؛ راست‌ترین بخش سیاه و خالی می‌ماند.

**اگر خوب درنیامد:** مدل رابط کاربری یا گرافی دیجیتال با گره‌های نئونی می‌سازد؛ اضافه کنید `physical gold wire and tiny gold beads photographed on black fabric, macro, no screen, no UI, no glow effects, no lens flare`.

### `21_box.jpg`

**اسلاید:** ۴۲ (نسخه‌ی ۶.۰.۰؛ هنوز منتشر نشده)

**کاربرد:** کنجکاوی و انتظار: چیزی بسته و ناشناخته که هنوز نوشته نشده، ولی از زیر درش روشنایی بیرون می‌زند.

**پرامپت:**

```text
A sealed matte black box with a flat closed lid rests on a low stone pedestal in a dark room, in the left third. A single thin line of warm golden light leaks out from under the lid and spills softly onto the pedestal. The box has no handle, label or markings. The right two-thirds is dark, empty room with faint haze. Photorealistic cinematic still-life photograph, 85mm lens, shallow depth of field, fine film grain, matte black and dark slate materials, one warm key light, deep shadows, real polished gold as the only saturated colour, minimalist, 16:9 landscape. Absolutely no text: no letters, numbers, logos, labels or watermarks anywhere, and no faces.
```

**ترکیب‌بندی:** جعبه در یک‌سوم چپ و روی پایه است؛ دوسوم راست اتاق تاریک و خالی می‌ماند.

**اگر خوب درنیامد:** در جعبه باز می‌شود یا نور از بالای آن می‌تابد؛ اضافه کنید `lid fully closed, the light only a hairline leaking out from under the lid, nothing else glowing`.

### `22_gold_network.jpg`

**اسلاید:** ۴۶ (جمله‌ی آخر: هر شکستی چیزی اضافه کرد)

**کاربرد:** در پایان، همان سنگ شکسته‌ی روز اول پر از ترمیم‌های طلایی است: شکست‌ها آخر کار نبودند و چیزی به ما اضافه کردند.

**پرامپت:** (اگر مدل ویرایش تصویر دارد، `01_cover` را به‌عنوان مرجع بدهید و `Keep the same stone surface, lighting and camera as the reference image.` را به ابتدای پرامپت اضافه کنید.)

```text
A wide slab of matte black obsidian stone seen from slightly above, crossed by many fine cracks filled with polished gold, like kintsugi repair, branching like river deltas and tree roots across the entire frame and softly glowing with warm light. A full-frame pattern, but the horizontal centre band is slightly darker and calmer, with fewer, thinner gold lines. Photorealistic cinematic still-life photograph, 85mm lens, shallow depth of field, fine film grain, matte black and dark slate materials, one warm key light, deep shadows, real polished gold as the only saturated colour, minimalist, 16:9 landscape. Absolutely no text: no letters, numbers, logos, labels or watermarks anywhere, and no faces.
```

**ترکیب‌بندی:** استثنا: الگو روی کل کادر پخش است، ولی نوار افقی وسط (جای دو خط متن) باید کمی تاریک‌تر و آرام‌تر بماند.

**اگر خوب درنیامد:** ترک‌ها در مرکز هم پررنگ و شلوغ می‌شوند و زیر متن می‌روند یا طلا مثل نئون می‌درخشد؛ اضافه کنید `fewer, thinner and dimmer gold lines in the horizontal centre band, gold as real metal, not neon`.

### `23_colonnade.jpg`

**اسلاید:** ۲۵ (نسخه‌ی ۴.۰.۰؛ مدیر شدن)

**کاربرد:** قافیه‌ی بصری با `13_pillar`: همان تالار و همان تخته‌سنگِ اسلاید SPOF، ولی این بار بار روی یک ستون نیست و روی هفت ستون نشسته است. هفت، تعداد اعضای تیم است.

**پرامپت:** (اگر مدل تصویر مرجع می‌پذیرد، `13_pillar` را به‌عنوان مرجع بدهید و `Keep the same hall, stone, lighting and camera as the reference image.` را به ابتدای پرامپت اضافه کنید.)

```text
A massive slab of dark stone like a low roof in a vast dark hall, resting on exactly seven slender plain stone columns standing in one single row. The row of columns sits in the left third of the frame and recedes slightly into the darkness, the far end of the slab vanishing. A beam of warm gold light falls from above onto the nearest columns; one column has a hairline crack filled with polished gold, like kintsugi repair. The right two-thirds is dark and empty. Photorealistic cinematic still-life photograph, 85mm lens, shallow depth of field, fine film grain, matte black and dark slate materials, one warm key light, deep shadows, real polished gold as the only saturated colour, minimalist, 16:9 landscape. Absolutely no text: no letters, numbers, logos, labels or watermarks anywhere, and no faces.
```

**ترکیب‌بندی:** ردیف ستون‌ها در یک‌سوم چپ است و نور از بالا می‌تابد؛ دوسوم راست تاریک و خالی می‌ماند.

**اگر خوب درنیامد:** مدل تعداد ستون‌ها را کم‌وزیاد می‌کند یا تالاری شبیه معبد یونانی می‌سازد؛ اضافه کنید `exactly seven plain slender columns in a single row, no capitals, no ornaments, not a temple, the hall otherwise empty`. تعداد ستون‌ها مهم است؛ اگر شش یا هشت تا شد دوباره بسازید.

### `24_two_chairs.jpg`

**اسلاید:** ۲۷ («واحد کار عوض شد»؛ مصاحبه‌ها و آموزش)

**کاربرد:** میز کارِ این نسخه: دو صندلی روبه‌روی هم. کار دیگر پشت مانیتور نیست؛ در گفت‌وگو با یک آدم دیگر است. قافیه‌ای با `11_workbench`.

**پرامپت:**

```text
Two simple empty matte black chairs facing each other across a small bare dark table in a dark quiet room, placed in the left third of the frame. A single warm pendant lamp hangs above the table and casts a soft pool of golden light on the tabletop; on the table lies one closed plain black notebook with a thin polished gold pen beside it. Nobody is present. The right two-thirds is dark, empty room. Photorealistic cinematic still-life photograph, 85mm lens, shallow depth of field, fine film grain, matte black and dark slate materials, one warm key light, deep shadows, real polished gold as the only saturated colour, minimalist, 16:9 landscape. Absolutely no text: no letters, numbers, logos, labels or watermarks anywhere, and no faces.
```

**ترکیب‌بندی:** دو صندلی و میز در یک‌سوم چپ‌اند؛ دوسوم راست اتاق تاریک و خالی می‌ماند.

**اگر خوب درنیامد:** مدل آدم روی صندلی می‌نشاند، لپ‌تاپ و مانیتور روی میز می‌گذارد یا صحنه را شبیه اتاق بازجویی می‌کند؛ اضافه کنید `both chairs completely empty, no people, no laptop, no screen, warm and calm like a quiet meeting corner, not an interrogation room`.

### `25_student_desk.jpg`

**اسلاید:** ۴۱ («صندلی شاگردی»؛ دوباره مبتدی شدن)

**کاربرد:** هزینه‌ی هر تغییر بزرگ: دوباره سر کلاس نشستن. یک صندلی دسته‌دار کلاس، خالی و منتظر؛ آرام و محترمانه، نه نوستالژیک و بچگانه.

**پرامپت:**

```text
A single empty classroom chair with a small attached writing tablet arm, matte black metal and dark wood, standing alone in a dark empty classroom, in the left third of the frame, seen from a three-quarter angle. On the writing tablet lies an open notebook with completely blank pages and a thin polished gold pencil. A warm beam of light from a high window falls on the chair. The right two-thirds is dark and empty, with only a faint hint of other chairs lost in shadow. Photorealistic cinematic still-life photograph, 85mm lens, shallow depth of field, fine film grain, matte black and dark slate materials, one warm key light, deep shadows, real polished gold as the only saturated colour, minimalist, 16:9 landscape. Absolutely no text: no letters, numbers, logos, labels or watermarks anywhere, and no faces.
```

**ترکیب‌بندی:** صندلی در یک‌سوم چپ است و نور از بالا روی آن می‌افتد؛ دوسوم راست تاریک و خالی می‌ماند.

**اگر خوب درنیامد:** مدل تخته‌سیاهِ پر از نوشته، کلاسی پر از صندلی یا دفتری با خط و نوشته می‌سازد؛ اضافه کنید `no blackboard, no writing anywhere, notebook pages completely blank, only one chair clearly visible, the rest of the room in darkness`.

## اسلایدهایی که عمداً عکس ندارند

این اسلایدها عمداً تایپوگرافیک یا سیاه‌اند (کد، نمودار یا یک جمله روی زمینه‌ی سیاه) و هیچ تصویری برایشان ساخته نمی‌شود:

- اسلاید ۱: کد Pull Request و دو دکمه‌ی Approve و Request changes
- اسلاید ۳: فایل CHANGELOG
- اسلاید ۴: کارت‌های MAJOR، MINOR و PATCH
- اسلاید ۱۱: کامنت‌های بسته‌شده
- اسلاید ۱۴: هفتاد چراغ سرور
- اسلاید ۱۸: Trigger ≠ Root cause
- اسلاید ۲۳: یک جمله روی زمینه‌ی سیاه
- اسلاید ۲۴: فایل CODEOWNERS و «من نه.»
- اسلاید ۲۶: نمودار scale-up و scale-out
- اسلاید ۲۸: کامنت «مدیر شد؛ دیگه فنی نیست» و یک تست قرمز
- اسلاید ۳۸: «۳ ثانیه» و کامنت
- اسلاید ۳۹: دو جمله با علامت مساوی
- اسلاید ۴۰: «۱۶ سال، یک کار»
- اسلاید ۴۳: جدول Added و Removed
- اسلاید ۴۴: جمله‌هایی که شکست
- اسلاید ۴۵: CHANGELOG شما

اسلاید ۳۶ (ویدئوی دمو) هم در این فایل تصویری ندارد.

</div>
