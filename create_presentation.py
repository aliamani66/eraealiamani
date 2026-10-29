import os
import shutil
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN
from pptx.enum.shapes import MSO_SHAPE
from pptx.oxml import parse_xml

# Initialize presentation
prs = Presentation()
# Set 16:9 widescreen dimensions (13.333 x 7.5 inches)
prs.slide_width = Inches(13.333)
prs.slide_height = Inches(7.5)

blank_slide_layout = prs.slide_layouts[6]

slides_data = [
    {
        "file": "slide1.jpg",
        "title": "اسلاید ۱ — پرده اول: افتتاحیه، استعاره دایناسورها و دوقلوهای ترس",
        "badge": None,
        "notes": (
            "📌 راهنمای استیج: ورود آرام به وسط سن، لبخند، مکث ۲ ثانیه‌ای، نگاه متمرکز به جمعیت و شروع گرم و بدون تکلف.\n\n"
            "🎙️ نکات کلیدی گفتار:\n"
            "- ۶۶ میلیون سال پیش و بزرگ‌ترین درس بقا روی زمین.\n"
            "- تعریف دایناسور بودن در دنیای مهندسی: ۵ تا ۱۰ سال تسلط بر یک ابزار، کشیدن دیوار بتنی دور کدها و ساختن حاشیه امن کاذب («قلقش فقط دست خودمه»).\n"
            "- قانون بقا: انقراض غول‌های صلب و مغرور، و بقای پستانداران کوچک، چابک و تطبیق‌پذیر.\n"
            "- معرفی «دوقلوهای ترس» در طول ۱۶ سال مسیر حرفه‌ای:\n"
            "  ۱. قل اول: ترس از اشتباه کردن و ضایع شدن (Fear of Incompetence)\n"
            "  ۲. قل دوم: ترس از حذف شدن و بی‌مصرف شدن (Fear of Irrelevance)\n"
            "- دعوت مخاطب به دیدن نقشه راه ۱۶ ساله..."
        )
    },
    {
        "file": "career_roadmap_highway.jpg",
        "title": "اسلاید ۲ — نقشه راه ۱۶ ساله: از دایناسور جاوا تا زیست‌بوم هوش مصنوعی (۱۳۸۸ تا ۱۴۰۳)",
        "badge": None,
        "notes": (
            "📌 راهنمای استیج: ایستادن در سمت چپ سن، اشاره دست به امتداد جاده نوری، ایجاد حس یک سفر پرماجرا.\n\n"
            "🎙️ نکات کلیدی گفتار:\n"
            "- این بزرگراه ۱۶ ساله، نقشه راه عبور من از این دو ترس در دنیای واقعی مهندسی است.\n"
            "- ۷ ایستگاه کلیدی و ۷ پوست‌اندازی بزرگ:\n"
            "  • ۱۳۸۸: دژ جاوا و برنامه‌نویسی ناب\n"
            "  • ۱۳۹۴: تابوشکنی و جهش جسورانه به دنیای عملیات\n"
            "  • ۱۳۹۶: میدان جنگ ابزارها، حلقه For و ضرورت Jenkins\n"
            "  • ۱۳۹۷: تجمیع داده‌ها و لاگ‌ها در پلتفرم ELK\n"
            "  • ۱۳۹۸: تشکیل تیم R&D و قانون طلایی «بساز و رها کن»\n"
            "  • ۱۴۰۰: نقطه اوج و بحران ساعت ۶ صبح بانک در کوبرنتیز\n"
            "  • ۱۴۰۳+: شهاب‌سنگ AI و پلتفرم خودکارسازی هوشمند ISR\n"
            "- تعامل با ارائه: کلیک روی هر ایستگاه برای ورود به سال مربوطه."
        )
    },
    {
        "file": "slide2.jpg",
        "title": "اسلاید ۳ — پرده دوم (۱۳۸۸ تا ۱۳۹۴): امپراتوری جاوا و ورود به دنیای عملیات",
        "badge": "📍 ایستگاه ۱ (۱۳۸۸): امپراتوری جاوا و ورود به دنیای عملیات",
        "notes": (
            "📌 راهنمای استیج: حرکت به سمت راست سن، لحن صمیمی و نوستالژیک با طنز خودزنی مهندسی.\n\n"
            "🎙️ نکات کلیدی گفتار:\n"
            "- ۶ سال برنامه‌نویسی ناب و خالص جاوا از سال ۸۸ تا ۹۴ در اوج حاشیه امن.\n"
            "- سال ۹۴ و تصمیم به رفتن به عملیات (برچسب افت کلاس و خودکشی شغلی در نگاه برنامه‌نویس‌های آن زمان).\n"
            "- بیدار شدن دوقلوهای ترس: ترس از دستور اشتباه در پروداکشن زنده + ترس از نابودی هویت دولوپری.\n"
            "- ورود به عملیات با ذهن و سلاح یک برنامه‌نویس.\n"
            "- ایده آقای شریفی (Query Executor) و توسعه قابلیت SSH و لایو تیل (Live Tail) لاگ‌ها با جاوا توسط علی؛ جمع شدن صف پشت در عملیات."
        )
    },
    {
        "file": "slide3.jpg",
        "title": "اسلاید ۴ — پرده سوم (۱۳۹۵): وب‌لاجیک (WebLogic) در اسکیل Enterprise و ماژول Authenticator",
        "badge": "📍 ایستگاه ۲ (۱۳۹۵): وب‌لاجیک انترپرایز و ماژول اختصاصی Authenticator",
        "notes": (
            "📌 راهنمای استیج: قدم زدن روی سن، روایت پرانرژی میدان جنگ ابزارها.\n\n"
            "🎙️ نکات کلیدی گفتار:\n"
            "- سال ۱۳۹۵: نیاز به مهاجرت از تام‌کت به وب‌لاجیک برای پایداری و High Availability واقعی در اسکیل بانکی.\n"
            "- مقاومت تیم تولید: «وب‌لاجیک سخته و ناشناخته است، ما دست نمی‌زنیم!»\n"
            "- خواندن ۲ الی ۳ هزار صفحه داکیومنت سنگین انگلیسی وب‌لاجیک با زبان ضعیف و دیکشنری.\n"
            "- نوشتن ماژول احراز هویت اختصاصی وب‌لاجیک و بالا آمدن سیستم.\n"
            "- درس بزرگ: بعدها با آمدن Cloud-Native فهمیدیم حتی وب‌لاجیک با آن عظمت هم دائمی نیست و باید رها شود."
        )
    },
    {
        "file": "slide4.jpg",
        "title": "اسلاید ۵ — پرده سوم (۱۳۹۶): حلقه For مرگبار و ضرورت اتوماسیون با Jenkins",
        "badge": "📍 ایستگاه ۳ (۱۳۹۶): فاجعه حلقه For و ضرورت اتوماسیون با Jenkins",
        "notes": (
            "📌 راهنمای استیج: لبخند و لحن طنزآمیز، همراه کردن تماشاگران با استرس‌های واقعی پروداکشن.\n\n"
            "🎙️ نکات کلیدی گفتار:\n"
            "- سال ۱۳۹۶: مهاجرت از NFS به کپی مستقیم با اسکریپت و حلقه for روی لیست IP سرورها.\n"
            "- فاجعه ارسال اشتباهی دستور shutdown داخل حلقه for و خاموش شدن یکی‌یکی سرورهای دیتاسنتر!\n"
            "- ضرورت اتوماسیون با Git و Jenkins و مقاومت شدید همکاران از ترس بیکار شدن و حذف شدن.\n"
            "- نوشتن فریم‌ورک اتوماسیون با زبان Groovy در جنکینز برای راحت کردن کار بچه‌ها، نه گرفتن جای آن‌ها."
        )
    },
    {
        "file": "slide5.jpg",
        "title": "اسلاید ۶ — پرده سوم (۱۳۹۷): نجات از جزیره‌های لاگ و مهاجرت به ELK",
        "badge": "📍 ایستگاه ۴ (۱۳۹۷): نجات از جزیره‌های لاگ و مهاجرت به پلتفرم متمرکز ELK",
        "notes": (
            "📌 راهنمای استیج: لحن رو به جلو، نشان دادن غلبه بر هرج‌ومرج داده‌ها.\n\n"
            "🎙️ نکات کلیدی گفتار:\n"
            "- سال ۱۳۹۷: پایان دوران SSH زدن دستی به ۲۰ سرور مختلف برای پیدا کردن یک خطا.\n"
            "- نوشتن کانفیگ پیچیده ۵۰۰-۶۰۰ خطی Filebeat برای هدایت متمرکز تمام لاگ‌ها به ELK.\n"
            "- نفس کشیدن سرورها و شکل‌گیری نیاز به یک تحول عمیق‌تر در ساختار سازمان."
        )
    },
    {
        "file": "slide6.jpg",
        "title": "اسلاید ۷ — پرده چهارم (۱۳۹۸ تا ۱۴۰۰): تیم تحقیق و توسعه (R&D Ops) و قانون طلایی «رها کردن»",
        "badge": "📍 ایستگاه ۵ (۱۳۹۸): تیم R&D Ops و فرهنگ طلایی «بساز، تحویل بده و رها کن»",
        "notes": (
            "📌 راهنمای استیج: ایستادن در مرکز سن، لحن مقتدر در بیان اصول سازمانی، سپس لحن صمیمی و صادقانه در روایت ترس درونی.\n\n"
            "🎙️ نکات کلیدی گفتار:\n"
            "- سال ۱۳۹۸ و ابتکار آقای شریفی: تشکیل «تیم تولیدِ عملیات» (۴ برنامه‌نویس باسابقه + ۱ مهندس دوآپس) و تکامل به R&D Ops.\n"
            "- چرخه ۵ مرحله‌ای: ۱. تحلیل نیاز و فیدبک ۲. تحقیق و POC ۳. مستندسازی ۴. آماده‌سازی پروداکشن ۵. آموزش و تحویل قطعی به عملیات.\n"
            "- قانون طلایی سازمان: «حق ندارید هیچ سرویسی که بالا می‌آورید رو خودتون نگه دارید! بسازید، تحویل بدید و رها کنید!» (تفکیک کیفیت از مالکیت).\n"
            "- سال ۱۳۹۹ و مواجهه با قل دوم ترس: دیدن همکار جوانی که پایپ‌لاین‌های من رو سریع‌تر بازنویسی می‌کرد؛ غلبه بر وسوسه انحصارطلبی و رفتن سراغ ساخت ابزار توپولوژی خودکار میکروسرویس‌ها."
        )
    },
    {
        "file": "slide7.jpg",
        "title": "اسلاید ۸ — پرده پنجم (سال ۱۴۰۰ - نقطه اوج): ساعت ۰۶:۰۰ صبح بانک و امتحان نهایی",
        "badge": "📍 ایستگاه ۶ (۱۴۰۰): نقطه اوج — ساعت ۰۶:۰۰ صبح بانک و امتحان نهایی در کوبرنتیز",
        "notes": (
            "📌 راهنمای استیج: ضربان صدا بالا، ایجاد حس تعلیق و هیجان در سالن.\n\n"
            "🎙️ نکات کلیدی گفتار:\n"
            "- سال ۱۴۰۰: ۳ سال برنامه‌ریزی محتاطانه برای مهاجرت به کوبرنتیز، اما وقوع اینسیدنت فاجعه‌بار و خوابیدن زیرساخت سنتی.\n"
            "- فشرده شدن پروژه ۳ ساله در ۳ شبانه‌روز بی‌خوابی مطلق و مهاجرت بیگ‌بنگ Core بانک به کوبرنتیز.\n"
            "- ساعت ۶:۱۵ صبح، لرزش دست روی کیبورد قبل از باز شدن شعب بانک و هجوم دوباره دوقلوهای ترس.\n"
            "- غلبه بر فلج ذهنی با تمرکز روی لاگ‌ها و قدم‌های کوچک؛ ساعت ۶:۳۰ صبح، سبز ماندن لاگ‌ها و نجات سیستم!"
        )
    },
    {
        "file": "slide8.jpg",
        "title": "اسلاید ۹ — پرده ششم (۱۴۰۳ و آینده): پلتفرم خودکارسازی هوشمند عملیات (Autonomous ISR Platform)",
        "badge": "📍 ایستگاه ۷ (۱۴۰۳+): پلتفرم خودکارسازی هوشمند ISR و دستیاران هوش مصنوعی",
        "notes": (
            "📌 راهنمای استیج: نشان دادن جهش تکنولوژی امروز، انتقال حس سرعت و هوشمندی.\n\n"
            "🎙️ نکات کلیدی گفتار:\n"
            "- رسیدن به زمان حال (۱۴۰۳): توسعه پلتفرم هوشمند ISR در تیم تحقیق و توسعه.\n"
            "- ارکستراسیون فرآیندها با n8n، تریگر خودکار جاب‌های Jenkins، تگ زدن اتوماتیک ورژن‌های Git و تحلیل لاگ‌ها و Auto-Remediation با ایجنت‌های هوش مصنوعی.\n"
            "- انجام کارهایی که روزی ساعت‌ها جلسه و شیفت شبانه می‌طلبید، در چند ثانیه."
        )
    },
    {
        "file": "slide9.jpg",
        "title": "اسلاید ۱۰ — پرده ششم (ادامه): شهاب‌سنگ هوش مصنوعی (AI) و دوراهی سرنوشت‌ساز بقا",
        "badge": "☄️ پرده پایانی: شهاب‌سنگ هوش مصنوعی (AI) و دوراهی سرنوشت‌ساز بقا",
        "notes": (
            "📌 راهنمای استیج: اشاره مستقیم به تصویر شهاب‌سنگ و تقابل دو گروه، برانگیختن همذات‌پنداری عمیق در سالن.\n\n"
            "🎙️ نکات کلیدی گفتار:\n"
            "- بیدار شدن همگانی دوقلوهای ترس: امروز همه ما (از جونیور تا سینیور) پشت ChatGPT و Cursor دلمان لرزیده.\n"
            "- تقابل دو رویکرد در تصویر:\n"
            "  • سمت چپ (دایناسورها): سنگر گرفتن در قلعه قدیمی و بغل کردن سرورها و کابل‌های سنتی از ترس تغییر.\n"
            "  • سمت راست (پستانداران تطبیق‌پذیر): رفتن به استقبال تغییر با آغوش باز و یادگیری مداوم.\n"
            "- ماشین کد می‌زند و پایپ‌لاین ران می‌کند، اما هرگز نمی‌تواند «شجاعت» داشته باشد و پوست بیندازد."
        )
    },
    {
        "file": "slide10.jpg",
        "title": "اسلاید ۱۱ — پایان: پیام طلایی بقا — دایناسور نباشید؛ پوست بیندازید و زنده بمانید",
        "badge": "🌍 جمع‌بندی: پیام طلایی بقا — دایناسور نباشید؛ تغییر کنید و زنده بمانید",
        "notes": (
            "📌 راهنمای استیج: ایستادن در مرکز سن، لحن الهام‌بخش و مقتدر.\n\n"
            "🎙️ نکات کلیدی گفتار:\n"
            "- از دوقلوهای ترس فرار نکنید؛ آن‌ها نشانه این هستند که در لبه رشدتان ایستاده‌اید.\n"
            "- دست ترس‌هایتان را بگیرید و با هم بپرید وسط دنیای جدید.\n\n"
            "جمله طلایی پایانی:\n"
            "«دایناسور نباشید؛ تغییر کنید، پوست بیندازید و زنده بمانید.»\n\n"
            "*(سکوت ۳ ثانیه‌ای — نگاه مستقیم به سالن)*\n"
            "خیلی ممنونم از همگی."
        )
    }
]

slides_dir = r"d:\work\erae-ali\slides"
created_slides = []

# Phase 1: Create all slides with full-screen background pictures and speech notes
for idx, item in enumerate(slides_data):
    img_path = os.path.join(slides_dir, item["file"])
    if not os.path.exists(img_path):
        print(f"Warning: {img_path} not found!")
        continue
    
    slide = prs.slides.add_slide(blank_slide_layout)
    created_slides.append(slide)
    
    # Add full-screen 16:9 image
    slide.shapes.add_picture(img_path, Inches(0), Inches(0), Inches(13.333), Inches(7.5))
    
    # Add smooth cinematic fade transition to each slide
    trans_xml = '<p:transition xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" spd="med"><p:fade/></p:transition>'
    slide._element.append(parse_xml(trans_xml))
    
    # Add speaker notes
    notes_slide = slide.notes_slide
    text_frame = notes_slide.notes_text_frame
    text_frame.text = f"{item['title']}\n\n{item['notes']}"

roadmap_slide = created_slides[1]  # Slide 2

# Helper to format Persian RTL text inside text frame
def style_rtl_text(shape, text, font_size_pt=12, is_bold=True, text_color=RGBColor(255, 255, 255), align=PP_ALIGN.CENTER):
    tf = shape.text_frame
    tf.word_wrap = True
    p = tf.paragraphs[0]
    p.text = text
    p.font.name = 'Vazirmatn'
    p.font.size = Pt(font_size_pt)
    p.font.bold = is_bold
    p.font.color.rgb = text_color
    p.alignment = align
    pPr = p._p.get_or_add_pPr()
    pPr.set('rtl', '1')

# Helper to create styled clickable buttons
def add_nav_button(slide, left, top, width, height, text, target_slide, font_size_pt=12, bg_color=RGBColor(15, 23, 42), border_color=RGBColor(6, 182, 212)):
    btn = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, left, top, width, height)
    btn.fill.solid()
    btn.fill.fore_color.rgb = bg_color
    btn.line.color.rgb = border_color
    btn.line.width = Pt(1.5)
    style_rtl_text(btn, text, font_size_pt=font_size_pt, is_bold=True, text_color=RGBColor(255, 255, 255), align=PP_ALIGN.CENTER)
    btn.click_action.target_slide = target_slide
    return btn

# Helper to create transparent clickable hotspots for Slide 2
def add_transparent_hotspot(slide, x_px, y_px, w_px, h_px, target_slide):
    # Convert pixels to inches for 1376x768 -> 13.333x7.5 inches
    left = Inches((x_px / 1376.0) * 13.333)
    top = Inches((y_px / 768.0) * 7.5)
    width = Inches((w_px / 1376.0) * 13.333)
    height = Inches((h_px / 768.0) * 7.5)
    
    shape = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, left, top, width, height)
    spPr = shape._element.spPr
    for child in list(spPr):
        if child.tag.endswith('solidFill') or child.tag.endswith('ln'):
            spPr.remove(child)
            
    # Add near-zero alpha fill (catches mouse clicks without visual obstruction)
    fill_xml = parse_xml(
        '<a:solidFill xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">'
        '<a:srgbClr val="FFFFFF"><a:alpha val="100"/></a:srgbClr>'
        '</a:solidFill>'
    )
    spPr.append(fill_xml)
    ln_xml = parse_xml(
        '<a:ln xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">'
        '<a:noFill/>'
        '</a:ln>'
    )
    spPr.append(ln_xml)
    shape.click_action.target_slide = target_slide
    return shape

# Phase 2: Add Hotspots & interactive elements on Slide 2 (Roadmap)
station_hotspots_px = [
    # (x, y, w, h, target_slide_index)
    (40, 280, 175, 480, 2),    # Station 1: 1388 Java -> Slide 3
    (215, 220, 135, 440, 3),   # Station 2: 1394 Ops & WebLogic -> Slide 4
    (350, 200, 150, 400, 4),   # Station 3: 1396 Jenkins -> Slide 5
    (500, 190, 195, 390, 5),   # Station 4: 1397 ELK -> Slide 6
    (700, 180, 235, 400, 6),   # Station 5: 1398 R&D Ops -> Slide 7
    (940, 170, 185, 380, 7),   # Station 6: 1400 K8s Bank -> Slide 8
    (1130, 150, 180, 370, 8),  # Station 7: 1403+ AI ISR -> Slide 9
]

for x, y, w, h, target_idx in station_hotspots_px:
    add_transparent_hotspot(roadmap_slide, x, y, w, h, created_slides[target_idx])

# Add subtle navigation controls on Slide 2:
# 1. Direct jump to finale button at top-left
add_nav_button(
    roadmap_slide,
    left=Inches(0.5),
    top=Inches(0.35),
    width=Inches(3.4),
    height=Inches(0.52),
    text="☄️ ورود مستقیم به پرده پایانی (AI) ◀",
    target_slide=created_slides[9],  # Slide 10
    font_size_pt=11.5,
    bg_color=RGBColor(26, 16, 37),
    border_color=RGBColor(236, 72, 153)
)

# 2. Subtle interactive hint bar at bottom-center
hint_box = roadmap_slide.shapes.add_shape(
    MSO_SHAPE.ROUNDED_RECTANGLE,
    Inches(3.8),
    Inches(6.85),
    Inches(5.7),
    Inches(0.45)
)
hint_box.fill.solid()
hint_box.fill.fore_color.rgb = RGBColor(10, 15, 28)
hint_box.line.color.rgb = RGBColor(30, 58, 95)
hint_box.line.width = Pt(1)
style_rtl_text(
    hint_box,
    "🖱️ برای ورود به هر دوره شغلی، روی ایستگاه مورد نظر در بزرگراه کلیک کنید",
    font_size_pt=11,
    is_bold=False,
    text_color=RGBColor(148, 163, 184),
    align=PP_ALIGN.CENTER
)

# Phase 3: Add Badges, Back-to-Roadmap, and Next/Prev buttons on Deep-Dive Slides (Slides 3-9)
for idx in range(2, 9):  # indices 2 to 8 -> Slides 3 to 9
    slide = created_slides[idx]
    badge_text = slides_data[idx]["badge"]
    
    # 1. Back to Roadmap Button (Top-Left)
    add_nav_button(
        slide,
        left=Inches(0.5),
        top=Inches(0.35),
        width=Inches(3.1),
        height=Inches(0.52),
        text="🗺️ بازگشت به بزرگراه مسیر شغلی",
        target_slide=roadmap_slide,
        font_size_pt=12,
        bg_color=RGBColor(15, 23, 42),
        border_color=RGBColor(6, 182, 212)
    )
    
    # 2. Station Badge Indicator (Top-Right)
    badge = slide.shapes.add_shape(
        MSO_SHAPE.ROUNDED_RECTANGLE,
        Inches(4.5),
        Inches(0.35),
        Inches(8.3),
        Inches(0.52)
    )
    badge.fill.solid()
    badge.fill.fore_color.rgb = RGBColor(10, 15, 30)
    badge.line.color.rgb = RGBColor(56, 189, 248)
    badge.line.width = Pt(1.5)
    style_rtl_text(
        badge,
        badge_text,
        font_size_pt=12,
        is_bold=True,
        text_color=RGBColor(241, 245, 249),
        align=PP_ALIGN.CENTER
    )
    
    # 3. Bottom Previous Station Button (Bottom-Right)
    if idx > 2:
        prev_target = created_slides[idx - 1]
        add_nav_button(
            slide,
            left=Inches(10.6),
            top=Inches(6.8),
            width=Inches(2.2),
            height=Inches(0.45),
            text="▶ ایستگاه قبلی",
            target_slide=prev_target,
            font_size_pt=11,
            bg_color=RGBColor(15, 23, 42),
            border_color=RGBColor(100, 116, 139)
        )
        
    # 4. Bottom Next Station Button (Bottom-Left)
    next_target = created_slides[idx + 1]
    next_text = "ایستگاه بعدی ◀" if idx < 8 else "☄️ پرده پایانی: شهاب‌سنگ AI ◀"
    next_border = RGBColor(6, 182, 212) if idx < 8 else RGBColor(236, 72, 153)
    add_nav_button(
        slide,
        left=Inches(0.5),
        top=Inches(6.8),
        width=Inches(3.1 if idx == 8 else 2.2),
        height=Inches(0.45),
        text=next_text,
        target_slide=next_target,
        font_size_pt=11,
        bg_color=RGBColor(15, 23, 42),
        border_color=next_border
    )

# Phase 4: Navigation on Conclusion Slides (Slide 10 & 11)
# Slide 10: AI Meteor
slide10 = created_slides[9]
add_nav_button(
    slide10,
    left=Inches(0.5),
    top=Inches(0.35),
    width=Inches(2.5),
    height=Inches(0.52),
    text="🗺️ بزرگراه مسیر شغلی",
    target_slide=roadmap_slide,
    font_size_pt=12,
    bg_color=RGBColor(15, 23, 42),
    border_color=RGBColor(6, 182, 212)
)
badge10 = slide10.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(3.8), Inches(0.35), Inches(9.0), Inches(0.52))
badge10.fill.solid()
badge10.fill.fore_color.rgb = RGBColor(26, 16, 37)
badge10.line.color.rgb = RGBColor(236, 72, 153)
badge10.line.width = Pt(1.5)
style_rtl_text(badge10, slides_data[9]["badge"], font_size_pt=12, is_bold=True, text_color=RGBColor(253, 242, 248), align=PP_ALIGN.CENTER)

add_nav_button(
    slide10,
    left=Inches(0.5),
    top=Inches(6.8),
    width=Inches(2.5),
    height=Inches(0.45),
    text="🌍 پیام پایانی بقا ◀",
    target_slide=created_slides[10],
    font_size_pt=11,
    bg_color=RGBColor(15, 23, 42),
    border_color=RGBColor(34, 197, 94)
)
add_nav_button(
    slide10,
    left=Inches(10.6),
    top=Inches(6.8),
    width=Inches(2.2),
    height=Inches(0.45),
    text="▶ ایستگاه ۷ (۱۴۰۳+)",
    target_slide=created_slides[8],
    font_size_pt=11,
    bg_color=RGBColor(15, 23, 42),
    border_color=RGBColor(100, 116, 139)
)

# Slide 11: Final Message
slide11 = created_slides[10]
add_nav_button(
    slide11,
    left=Inches(0.5),
    top=Inches(0.35),
    width=Inches(2.5),
    height=Inches(0.52),
    text="🗺️ بزرگراه مسیر شغلی",
    target_slide=roadmap_slide,
    font_size_pt=12,
    bg_color=RGBColor(15, 23, 42),
    border_color=RGBColor(6, 182, 212)
)
badge11 = slide11.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(3.8), Inches(0.35), Inches(9.0), Inches(0.52))
badge11.fill.solid()
badge11.fill.fore_color.rgb = RGBColor(10, 25, 20)
badge11.line.color.rgb = RGBColor(34, 197, 94)
badge11.line.width = Pt(1.5)
style_rtl_text(badge11, slides_data[10]["badge"], font_size_pt=12, is_bold=True, text_color=RGBColor(240, 253, 244), align=PP_ALIGN.CENTER)

add_nav_button(
    slide11,
    left=Inches(10.6),
    top=Inches(6.8),
    width=Inches(2.2),
    height=Inches(0.45),
    text="▶ شهاب‌سنگ AI",
    target_slide=created_slides[9],
    font_size_pt=11,
    bg_color=RGBColor(15, 23, 42),
    border_color=RGBColor(100, 116, 139)
)

# Save presentations
output_interactive = r"d:\work\erae-ali\TechTalk_Dinosaurs_Survival_Interactive.pptx"
output_pptx_v1 = r"d:\work\erae-ali\TechTalk_Dinosaurs_Survival.pptx"
output_pptx_v2 = r"d:\work\erae-ali\TechTalk_Dinosaurs_Survival_v2.pptx"

prs.save(output_interactive)
print(f"Saved: {output_interactive}")

try:
    shutil.copy2(output_interactive, output_pptx_v1)
    print(f"Updated: {output_pptx_v1}")
except Exception as e:
    print(f"Could not update {output_pptx_v1}: {e}")

try:
    shutil.copy2(output_interactive, output_pptx_v2)
    print(f"Updated: {output_pptx_v2}")
except Exception as e:
    print(f"Note: {output_pptx_v2} is currently open in PowerPoint. Close it if you wish to overwrite it.")

