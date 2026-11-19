import os
import sys
import shutil
import pptx
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN
from pptx.enum.shapes import MSO_SHAPE

def build_lecture_m7_6_1():
    template_path = r'd:\ხელოვნური ინტელექტის ინსტრუმენტები\Skillwill Presentation Example .pptx'
    asset_dir = r'd:\ხელოვნური ინტელექტის ინსტრუმენტები\assets'
    img_icon_title = os.path.join(asset_dir, 'icon_title.png')

    prs = Presentation(template_path)
    prs.slide_width = Inches(19.0)
    prs.slide_height = Inches(10.6875)

    COLOR_RED = RGBColor(0xDF, 0x38, 0x26)
    COLOR_CARD_RED = RGBColor(0xFD, 0xED, 0xEB)
    COLOR_BORDER_RED = RGBColor(0xF5, 0xC2, 0xBE)
    COLOR_DARK = RGBColor(0x20, 0x21, 0x24)
    COLOR_MUTED = RGBColor(0x5F, 0x63, 0x68)
    COLOR_TEXT_MUTED = RGBColor(0x70, 0x75, 0x7A)
    COLOR_BG_LIGHT = RGBColor(0xF8, 0xF9, 0xFA)
    COLOR_BORDER = RGBColor(0xEA, 0xEE, 0xF2)
    COLOR_WHITE = RGBColor(0xFF, 0xFF, 0xFF)
    COLOR_LINE = RGBColor(0xDA, 0xDC, 0xE0)
    COLOR_GREEN_BG = RGBColor(0xEA, 0xF8, 0xEA)
    COLOR_GREEN_BOR = RGBColor(0xC2, 0xE8, 0xC2)
    COLOR_GREEN_TEXT = RGBColor(0x13, 0x73, 0x33)

    FONT_ARIAL = 'Arial'

    # -------------------------------------------------------------------------
    # SLIDE 1: Cover Slide
    # -------------------------------------------------------------------------
    s1 = prs.slides[0]
    for shape in s1.shapes:
        if shape.has_text_frame and shape.text_frame.text.strip():
            tf = shape.text_frame
            tf.clear()
            p = tf.paragraphs[0]
            p.text = "6.1 - ამოცანის 4 ძირითადი ტიპი და შაბლონების ბიბლიოთეკა"
            p.alignment = PP_ALIGN.CENTER
            p.font.name = FONT_ARIAL
            p.font.size = Pt(32)
            p.font.bold = True
            p.font.color.rgb = COLOR_RED

            p2 = tf.add_paragraph()
            p2.text = "Summarize, Infer, Transform და Expand არქიტექტურა, სპეციალიზებული შაბლონები და ტიპური შეცდომები"
            p2.alignment = PP_ALIGN.CENTER
            p2.font.name = FONT_ARIAL
            p2.font.size = Pt(18)
            p2.font.color.rgb = COLOR_MUTED
            p2.space_before = Pt(14)
            break

    # Clear template slide 1
    s2 = prs.slides[1]
    for sp in list(s2.shapes):
        sp.element.getparent().remove(sp.element)

    blank_layout = prs.slide_layouts[6]
    created_slides = [s2]
    for _ in range(7):
        created_slides.append(prs.slides.add_slide(blank_layout))

    # Move closing slide to the end
    closing_elem = prs.slides._sldIdLst[2]
    prs.slides._sldIdLst.remove(closing_elem)
    prs.slides._sldIdLst.append(closing_elem)
    s10 = prs.slides[9]

    for shape in s10.shapes:
        if shape.has_text_frame and shape.text_frame.text.strip():
            for p in shape.text_frame.paragraphs:
                p.font.name = FONT_ARIAL

    # Header function with NO corner artifact and NO footer text, Arial font
    def add_header(slide, title, category="ᲛᲝᲓᲣᲚᲘ 7 • PROMPT ENGINEERING • 6.1"):
        if os.path.exists(img_icon_title):
            slide.shapes.add_picture(img_icon_title, Inches(1.0), Inches(0.85), width=Inches(0.42), height=Inches(0.45))

        tb = slide.shapes.add_textbox(Inches(1.55), Inches(0.6), Inches(17.0), Inches(1.2))
        tf = tb.text_frame
        tf.word_wrap = True
        tf.margin_left = tf.margin_top = tf.margin_right = tf.margin_bottom = 0

        p0 = tf.paragraphs[0]
        p0.text = category.upper()
        p0.font.name = FONT_ARIAL
        p0.font.size = Pt(12)
        p0.font.bold = True
        p0.font.color.rgb = COLOR_RED

        p1 = tf.add_paragraph()
        p1.text = title
        p1.font.name = FONT_ARIAL
        p1.font.size = Pt(28)
        p1.font.bold = True
        p1.font.color.rgb = COLOR_DARK
        p1.space_before = Pt(2)

        line = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(1.0), Inches(1.95), Inches(17.0), Inches(0.02))
        line.fill.solid()
        line.fill.fore_color.rgb = COLOR_LINE
        line.line.fill.background()

    card_w = Inches(5.4); card_h = Inches(5.4); card_y = Inches(2.35)
    xs_3 = [Inches(1.0), Inches(6.8), Inches(12.6)]
    step_w = Inches(3.95); step_h = Inches(5.4); step_y = Inches(2.35)
    xs_4 = [Inches(1.0), Inches(5.35), Inches(9.7), Inches(14.05)]
    bot_y = Inches(8.0)

    # -------------------------------------------------------------------------
    # SLIDE 2: LLM Task Taxonomy: 4 Core Operations
    # -------------------------------------------------------------------------
    s2 = created_slides[0]
    add_header(s2, "LLM ამოცანების ტაქსონომია: 4 ბირთვული საინჟინრო ოპერაცია", "ᲢᲐᲥᲡᲝᲜᲝᲛᲘᲐ • 4 CORE OPERATIONS")

    ops_s2 = [
        ("📑 SUMMARIZE", "შეკუმშვა & კონცენტრაცია", "ინფორმაციის სინთეზი", [
            ("დანიშნულება", "დიდი მოცულობის ტექსტის შემოკლება არსებითი შინაარსის, ფაქტებისა და კონტექსტის სრული შენარჩუნებით."),
            ("ტიპური ქეისები", "Executive Summary, შეხვედრების ოქმები, სამეცნიერო სტატიების TL;DR, დიალოგის ისტორიის შეკუმშვა."),
            ("ფოკუსი", "დაბალი ტემპერატურა (~0.2), მკაცრი ლიმიტი სიგრძეზე და ფაქტობრივი ერთგულება წყაროსთან.")
        ]),
        ("🔍 INFER", "ანალიზი & ექსტრაქცია", "იმპლიციტური ცოდნა", [
            ("დანიშნულება", "ტექსტში დამალული აზრის, სენტიმენტის, განზრახვის ან სტრუქტურირებული ობიექტების ამოღება."),
            ("ტიპური ქეისები", "სენტიმენტ-ანალიზი, თემატიკის დეტექცია, Entity Extraction, მომხმარებლის უკმაყოფილების გაზომვა."),
            ("ფოკუსი", "მკაცრი JSON Schema, ტიპიზებული ველები, Null Handling და 0.0 ტემპერატურა.")
        ]),
        ("🔄 TRANSFORM", "გარდაქმნა & თარგმანი", "შინაარსის შენარჩუნება", [
            ("დანიშნულება", "ტექსტის ენობრივი გარსის, სტილის ან მონაცემთა ფორმატის შეცვლა ინფორმაციის უცვლელობით."),
            ("ტიპური ქეისები", "ენობრივი თარგმანი, Markdown/HTML ➔ JSON კონვერტაცია, კოდის რეფაქტორინგი, ტონის ცვლილება."),
            ("ფოკუსი", "ინფორმაციის კონსერვაცია — არაფერი დაიკარგოს და არაფერი ზედმეტი არ დაემატოს.")
        ]),
        ("🌱 EXPAND", "გაშლა & გენერაცია", "კონტექსტის შექმნა", [
            ("დანიშნულება", "მოკლე მონახაზიდან ან პარამეტრებიდან სრულფასოვანი, დეტალური და მდიდარი ტექსტის შექმნა."),
            ("ტიპური ქეისები", "პერსონალიზებული ელფოსტები, Brainstorming, ანგარიშების დეტალიზაცია, კრეატიული მონახაზები."),
            ("ფოკუსი", "კონტროლირებადი ტემპერატურა (0.7) და მკაცრი უარყოფითი წესები Fluff-ისა და სიყალბის წინააღმდეგ.")
        ])
    ]

    for i, (badge, title, subtitle, items) in enumerate(ops_s2):
        bg_col = COLOR_CARD_RED if i == 0 else COLOR_BG_LIGHT
        bor_col = COLOR_BORDER_RED if i == 0 else COLOR_BORDER
        c = s2.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, xs_4[i], step_y, step_w, step_h)
        c.fill.solid(); c.fill.fore_color.rgb = bg_col
        c.line.color.rgb = bor_col; c.line.width = Pt(1.5)

        tb = s2.shapes.add_textbox(xs_4[i] + Inches(0.25), step_y + Inches(0.25), step_w - Inches(0.5), step_h - Inches(0.5))
        tf = tb.text_frame; tf.word_wrap = True; tf.margin_left = tf.margin_top = 0

        p = tf.paragraphs[0]; p.text = badge.upper()
        p.font.name = FONT_ARIAL; p.font.size = Pt(11); p.font.bold = True
        p.font.color.rgb = COLOR_RED if i == 0 else COLOR_MUTED

        p = tf.add_paragraph(); p.text = title
        p.font.name = FONT_ARIAL; p.font.size = Pt(15); p.font.bold = True; p.font.color.rgb = COLOR_DARK; p.space_before = Pt(4)

        p = tf.add_paragraph(); p.text = subtitle
        p.font.name = FONT_ARIAL; p.font.size = Pt(11); p.font.color.rgb = COLOR_TEXT_MUTED; p.space_before = Pt(2)

        for heading, body in items:
            p = tf.add_paragraph(); p.text = f"• {heading}: {body}"
            p.font.name = FONT_ARIAL; p.font.size = Pt(11); p.font.color.rgb = COLOR_DARK; p.space_before = Pt(6)

    # Bottom takeaway
    bb = s2.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(1.0), bot_y, Inches(17.0), Inches(2.0))
    bb.fill.solid(); bb.fill.fore_color.rgb = COLOR_WHITE
    bb.line.color.rgb = COLOR_BORDER; bb.line.width = Pt(1.2)

    tbb = s2.shapes.add_textbox(Inches(1.4), bot_y + Inches(0.3), Inches(16.2), Inches(1.4))
    tfb = tbb.text_frame; tfb.word_wrap = True; tfb.margin_left = tfb.margin_top = 0
    p = tfb.paragraphs[0]; p.text = "💡 მთავარი არსი (ტაქსონომიური მოდელი):"
    p.font.name = FONT_ARIAL; p.font.size = Pt(14); p.font.bold = True; p.font.color.rgb = COLOR_RED
    p = tfb.add_paragraph()
    p.text = "ნებისმიერი რთული საწარმოო AI სისტემა (Support ასისტენტიდან Enterprise RAG-მდე) აგებულია ამ 4 ოპერაციის კომბინაციით; თითოეულ ოპერაციას სჭირდება სპეციფიკური პრომპტის არქიტექტურა, პარამეტრები და დამცავი შეზღუდვები."
    p.font.name = FONT_ARIAL; p.font.size = Pt(13); p.font.color.rgb = COLOR_MUTED; p.space_before = Pt(4)

    # -------------------------------------------------------------------------
    # SLIDE 3: Operation 1: Summarize (Compression & Targeted Summaries)
    # -------------------------------------------------------------------------
    s3 = created_slides[1]
    add_header(s3, "Summarize: ინფორმაციის შეკუმშვა და მიზნობრივი რეზიუმირება", "ᲝᲞᲔᲠᲐᲪᲘᲐ 1 • SUMMARIZATION PATTERNS")

    sum_s3 = [
        ("🎯 ARCHITECTURAL PATTERNS", "შეკუმშვის პატერნები", "სტრუქტურული ტიპები", [
            ("Executive Summary", "დიდი ბიზნეს-დოკუმენტის შეჯამება მკაცრად 3 საკვანძო პუნქტად მენეჯმენტისთვის."),
            ("Action-Item Extraction", "შეხვედრის ტრანსკრიპტიდან გადაწყვეტილებების, შემსრულებლებისა და ვადების იზოლაცია."),
            ("მკაცრი სიგრძის ლიმიტები", "სიგრძის დაფიქსირება სიტყვების ან წინადადებების რაოდენობით (მაგ. 'მაქსიმუმ 50 სიტყვა').")
        ]),
        ("⚙️ PARAMETERS & CONTROLS", "პარამეტრების მართვა", "ჰალუცინაციების მინიმიზაცია", [
            ("დაბალი ტემპერატურა (0.1 – 0.3)", "მინიმალური ალბათური რყევა ფაქტობრივი მონაცემების ზუსტად გადმოსაცემად."),
            ("Context Grounding", "მკაცრი უარყოფითი წესი: 'დაეყრდენი მხოლოდ მოწოდებულ ტექსტს; არ დაამატო გარე ფაქტები'."),
            ("სტრუქტურული გამოსავალი", "წინასწარ განსაზღვრული Markdown სათაურები ან Bullet-point ფორმატი.")
        ]),
        ("👥 TARGETED SUMMARIZATION", "მიზნობრივი შეჯამება", "აუდიტორიაზე მორგება", [
            ("პერსპექტივის ფოკუსირება", "ერთი და იგივე ხელშეკრულების შეჯამება ფინანსური დირექტორისთვის, იურისტისთვის ან დეველოპერისთვის."),
            ("ფილტრაცია ინტერესის მიხედვით", "მოდელი უგულებელყოფს მეორეხარისხოვან თემებს და ყურადღებას ამახვილებს კონკრეტულ როლზე."),
            ("დროის ეკონომია", "მომხმარებელი იღებს მხოლოდ მისთვის კრიტიკულ გადაწყვეტილებებს და არა ზოგად შინაარსს.")
        ])
    ]

    for i, (badge, title, subtitle, items) in enumerate(sum_s3):
        bg_col = COLOR_CARD_RED if i == 2 else COLOR_BG_LIGHT
        bor_col = COLOR_BORDER_RED if i == 2 else COLOR_BORDER
        c = s3.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, xs_3[i], card_y, card_w, card_h)
        c.fill.solid(); c.fill.fore_color.rgb = bg_col
        c.line.color.rgb = bor_col; c.line.width = Pt(1.5)

        tb = s3.shapes.add_textbox(xs_3[i] + Inches(0.35), card_y + Inches(0.35), card_w - Inches(0.7), card_h - Inches(0.7))
        tf = tb.text_frame; tf.word_wrap = True; tf.margin_left = tf.margin_top = 0

        p = tf.paragraphs[0]; p.text = badge.upper()
        p.font.name = FONT_ARIAL; p.font.size = Pt(11); p.font.bold = True
        p.font.color.rgb = COLOR_RED if i == 2 else COLOR_MUTED

        p = tf.add_paragraph(); p.text = title
        p.font.name = FONT_ARIAL; p.font.size = Pt(17); p.font.bold = True; p.font.color.rgb = COLOR_DARK; p.space_before = Pt(4)

        p = tf.add_paragraph(); p.text = subtitle
        p.font.name = FONT_ARIAL; p.font.size = Pt(12); p.font.color.rgb = COLOR_TEXT_MUTED; p.space_before = Pt(2)

        for heading, body in items:
            p = tf.add_paragraph(); p.text = f"• {heading}: {body}"
            p.font.name = FONT_ARIAL; p.font.size = Pt(12); p.font.color.rgb = COLOR_DARK; p.space_before = Pt(8)

    # Bottom takeaway
    bb = s3.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(1.0), bot_y, Inches(17.0), Inches(2.0))
    bb.fill.solid(); bb.fill.fore_color.rgb = COLOR_WHITE
    bb.line.color.rgb = COLOR_BORDER; bb.line.width = Pt(1.2)

    tbb = s3.shapes.add_textbox(Inches(1.4), bot_y + Inches(0.3), Inches(16.2), Inches(1.4))
    tfb = tbb.text_frame; tfb.word_wrap = True; tfb.margin_left = tfb.margin_top = 0
    p = tfb.paragraphs[0]; p.text = "🎯 საინჟინრო პრინციპი:"
    p.font.name = FONT_ARIAL; p.font.size = Pt(14); p.font.bold = True; p.font.color.rgb = COLOR_RED
    p = tfb.add_paragraph()
    p.text = "Summarize არ არის უბრალოდ ტექსტის შემოკლება; საუკეთესო საწარმოო რეზიუმე ყოველთვის მიზნობრივია (Targeted) და ეფუძნება მკაცრ შეზღუდვას: 'არ ახსენო არცერთი ფაქტი, რომელიც წყაროში პირდაპირ არ წერია'."
    p.font.name = FONT_ARIAL; p.font.size = Pt(13); p.font.color.rgb = COLOR_MUTED; p.space_before = Pt(4)

    # -------------------------------------------------------------------------
    # SLIDE 4: Operation 2: Infer (Analysis, Sentiment & Extraction)
    # -------------------------------------------------------------------------
    s4 = created_slides[2]
    add_header(s4, "Infer: ტექსტის ანალიზი, სენტიმენტის დეტექცია და ექსტრაქცია", "ᲝᲞᲔᲠᲐᲪᲘᲐ 2 • INFERENCE & EXTRACTION")

    infer_s4 = [
        ("❤️ SENTIMENT & EMOTION", "სენტიმენტ-ანალიზი", "ტონისა და ემოციის გაზომვა", [
            ("3-დონიანი სენტიმენტი", "Positive, Negative, Neutral — მკაფიოდ გაწერილი საზღვრებითა და კრიტერიუმებით."),
            ("მრავალგანზომილებიანი ემოციები", "ბრაზი, იმედგაცრუება, მადლიერება, გადაუდებლობა (Urgency ქულა 1-დან 5-მდე)."),
            ("SLA მარშრუტიზაცია", "გაბრაზებული მომხმარებლის მოთხოვნა ავტომატურად იღებს მაღალ პრიორიტეტს სისტემაში.")
        ]),
        ("🧩 MULTI-FIELD EXTRACTION", "ერთდროული ექსტრაქცია", "ინფორმაციის სტრუქტურირება", [
            ("ერთიანი გამოძახება", "ერთ პრომპტში ერთდროულად ამოიღება: თემატიკა, პროდუქტის სახელი, შეცდომის ტიპი და განზრახვა."),
            ("მკაცრი JSON Schema", "ველები სრულად ტიპიზებულია Pydantic/Zod სქემით, რაც გამორიცხავს პარსინგის ერორებს."),
            ("Null Value Handling", "მკაფიო წესი იმ შემთხვევისთვის, როცა რომელიმე ველი შემავალ ტექსტში საერთოდ არ მოიპოვება.")
        ]),
        ("⚖️ STANCE & FACT CHECK", "პოზიციის დეტექცია", "მტკიცებულებათა ანალიზი", [
            ("Stance Detection", "ავტორის პოზიცია კონკრეტულ საკითხზე — ეთანხმება, ეწინააღმდეგება თუ ნეიტრალურია."),
            ("Intent Classification", "მომხმარებელს სურს ყიდვა, გამოწერის გაუქმება, ტექნიკური დახმარება თუ პრეტენზია."),
            ("Explicit Grounding", "მოდელს ევალება მიუთითოს ტექსტის ზუსტი ციტატა, რომელმაც განაპირობა დასკვნა.")
        ])
    ]

    for i, (badge, title, subtitle, items) in enumerate(infer_s4):
        bg_col = COLOR_CARD_RED if i == 1 else COLOR_BG_LIGHT
        bor_col = COLOR_BORDER_RED if i == 1 else COLOR_BORDER
        c = s4.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, xs_3[i], card_y, card_w, card_h)
        c.fill.solid(); c.fill.fore_color.rgb = bg_col
        c.line.color.rgb = bor_col; c.line.width = Pt(1.5)

        tb = s4.shapes.add_textbox(xs_3[i] + Inches(0.35), card_y + Inches(0.35), card_w - Inches(0.7), card_h - Inches(0.7))
        tf = tb.text_frame; tf.word_wrap = True; tf.margin_left = tf.margin_top = 0

        p = tf.paragraphs[0]; p.text = badge.upper()
        p.font.name = FONT_ARIAL; p.font.size = Pt(11); p.font.bold = True
        p.font.color.rgb = COLOR_RED if i == 1 else COLOR_MUTED

        p = tf.add_paragraph(); p.text = title
        p.font.name = FONT_ARIAL; p.font.size = Pt(17); p.font.bold = True; p.font.color.rgb = COLOR_DARK; p.space_before = Pt(4)

        p = tf.add_paragraph(); p.text = subtitle
        p.font.name = FONT_ARIAL; p.font.size = Pt(12); p.font.color.rgb = COLOR_TEXT_MUTED; p.space_before = Pt(2)

        for heading, body in items:
            p = tf.add_paragraph(); p.text = f"• {heading}: {body}"
            p.font.name = FONT_ARIAL; p.font.size = Pt(12); p.font.color.rgb = COLOR_DARK; p.space_before = Pt(8)

    # Bottom takeaway
    bb = s4.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(1.0), bot_y, Inches(17.0), Inches(2.0))
    bb.fill.solid(); bb.fill.fore_color.rgb = COLOR_WHITE
    bb.line.color.rgb = COLOR_BORDER; bb.line.width = Pt(1.2)

    tbb = s4.shapes.add_textbox(Inches(1.4), bot_y + Inches(0.3), Inches(16.2), Inches(1.4))
    tfb = tbb.text_frame; tfb.word_wrap = True; tfb.margin_left = tfb.margin_top = 0
    p = tfb.paragraphs[0]; p.text = "💡 მთავარი არსი:"
    p.font.name = FONT_ARIAL; p.font.size = Pt(14); p.font.bold = True; p.font.color.rgb = COLOR_RED
    p = tfb.add_paragraph()
    p.text = "Infer ამოცანებში მოდელი მოქმედებს როგორც ანალიტიკოსი; ყველაზე ეფექტური საინჟინრო მიდგომაა მრავალველიანი ექსტრაქცია ერთიან JSON ობიექტად, სადაც თითოეულ ველს თან ახლავს ციტატა და თავდაჯერებულობის ქულა."
    p.font.name = FONT_ARIAL; p.font.size = Pt(13); p.font.color.rgb = COLOR_MUTED; p.space_before = Pt(4)

    # -------------------------------------------------------------------------
    # SLIDE 5: Operation 3: Transform (Format Conversion, Tone Shift & Refactoring)
    # -------------------------------------------------------------------------
    s5 = created_slides[3]
    add_header(s5, "Transform: ფორმატის კონვერტაცია, ტონის შეცვლა და ლოკალიზაცია", "ᲝᲞᲔᲠᲐᲪᲘᲐ 3 • TRANSFORMATION & REFACTORING")

    trans_s5 = [
        ("🔄 FORMAT CONVERSION", "ფორმატის გარდაქმნა", "სტრუქტურული ტრანსფორმაცია", [
            ("ტექსტიდან სტრუქტურაში", "თავისუფალი ტექსტის ან ცხრილის გადაყვანა ვალიდურ JSON, XML, CSV ან SQL ჩანაწერებად."),
            ("მარკირების ცვლილება", "HTML დოკუმენტის სუფთა Markdown-ად ქცევა ან ტექნიკური ლოგების JSON მასივად გარდაქმნა."),
            ("სინტაქსური ერთგვაროვნება", "მონაცემთა ტიპების (String, Integer, Boolean) ზუსტი შენარჩუნება გარდაქმნის პროცესში.")
        ]),
        ("🎭 TONE & REGISTER SHIFT", "ტონის ტრანსფორმაცია", "სტილისტური მოდიფიკაცია", [
            ("დეველოპერიდან კლიენტამდე", "ტექნიკური Bug Report-ის გარდაქმნა მომხმარებლისთვის გასაგებ, ემპათიურ შეტყობინებად."),
            ("ფორმალურობის რეგულირება", "შიდა სასაუბრო ჩანაწერიდან ოფიციალური ბიზნეს-კორესპონდენციის ფორმულირება."),
            ("ბრენდის ხმის დაცვა", "კომპანიის სტილისტური გზამკვლევის (Corporate Tone of Voice) ავტომატური დაცვა.")
        ]),
        ("🌐 LOCALIZATION & CODE", "ლოკალიზაცია და კოდი", "ენობრივი და ლოგიკური ადაპტაცია", [
            ("კულტურული თარგმანი", "იდიომებისა და ადგილობრივი ნორმების ადაპტაცია პირდაპირი მექანიკური კალკირების ნაცვლად."),
            ("კოდის ენობრივი მიგრაცია", "Python ფუნქციების გადაყვანა TypeScript-ში ან SQL მოთხოვნების ოპტიმიზაცია ახალ დიალექტზე."),
            ("დოკუმენტაციის რეფაქტორინგი", "მოძველებული API დოკუმენტაციის მოდერნიზაცია OpenAPI 3.1 სტანდარტების მიხედვით.")
        ])
    ]

    for i, (badge, title, subtitle, items) in enumerate(trans_s5):
        bg_col = COLOR_CARD_RED if i == 0 else COLOR_BG_LIGHT
        bor_col = COLOR_BORDER_RED if i == 0 else COLOR_BORDER
        c = s5.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, xs_3[i], card_y, card_w, card_h)
        c.fill.solid(); c.fill.fore_color.rgb = bg_col
        c.line.color.rgb = bor_col; c.line.width = Pt(1.5)

        tb = s5.shapes.add_textbox(xs_3[i] + Inches(0.35), card_y + Inches(0.35), card_w - Inches(0.7), card_h - Inches(0.7))
        tf = tb.text_frame; tf.word_wrap = True; tf.margin_left = tf.margin_top = 0

        p = tf.paragraphs[0]; p.text = badge.upper()
        p.font.name = FONT_ARIAL; p.font.size = Pt(11); p.font.bold = True
        p.font.color.rgb = COLOR_RED if i == 0 else COLOR_MUTED

        p = tf.add_paragraph(); p.text = title
        p.font.name = FONT_ARIAL; p.font.size = Pt(17); p.font.bold = True; p.font.color.rgb = COLOR_DARK; p.space_before = Pt(4)

        p = tf.add_paragraph(); p.text = subtitle
        p.font.name = FONT_ARIAL; p.font.size = Pt(12); p.font.color.rgb = COLOR_TEXT_MUTED; p.space_before = Pt(2)

        for heading, body in items:
            p = tf.add_paragraph(); p.text = f"• {heading}: {body}"
            p.font.name = FONT_ARIAL; p.font.size = Pt(12); p.font.color.rgb = COLOR_DARK; p.space_before = Pt(8)

    # Bottom takeaway
    bb = s5.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(1.0), bot_y, Inches(17.0), Inches(2.0))
    bb.fill.solid(); bb.fill.fore_color.rgb = COLOR_WHITE
    bb.line.color.rgb = COLOR_BORDER; bb.line.width = Pt(1.2)

    tbb = s5.shapes.add_textbox(Inches(1.4), bot_y + Inches(0.3), Inches(16.2), Inches(1.4))
    tfb = tbb.text_frame; tfb.word_wrap = True; tfb.margin_left = tfb.margin_top = 0
    p = tfb.paragraphs[0]; p.text = "🎯 საინჟინრო პრინციპი:"
    p.font.name = FONT_ARIAL; p.font.size = Pt(14); p.font.bold = True; p.font.color.rgb = COLOR_RED
    p = tfb.add_paragraph()
    p.text = "Transform ოპერაციის მთავარი საინჟინრო კანონია 'ინფორმაციის კონსერვაცია' — არაფერი უნდა დაიკარგოს და არაფერი ზედმეტი არ უნდა დაემატოს; იცვლება მხოლოდ ფორმა, სინტაქსი და ენობრივი გარსი."
    p.font.name = FONT_ARIAL; p.font.size = Pt(13); p.font.color.rgb = COLOR_MUTED; p.space_before = Pt(4)

    # -------------------------------------------------------------------------
    # SLIDE 6: Operation 4: Expand (Controlled Expansion & Mitigating Fluff)
    # -------------------------------------------------------------------------
    s6 = created_slides[4]
    add_header(s6, "Expand: კონტროლირებადი გაშლა, პერსონალიზაცია და იდეაცია", "ᲝᲞᲔᲠᲐᲪᲘᲐ 4 • CONTROLLED EXPANSION")

    exp_s6 = [
        ("✉️ PERSONALIZED OUTREACH", "პერსონალიზებული პასუხები", "კონტექსტური გენერაცია", [
            ("მოკლე ბრიფიდან წერილამდე", "მომხმარებლის პროფილის, აქტივობისა და ისტორიის საფუძველზე სრულფასოვანი პასუხის შექმნა."),
            ("დინამიკური არგუმენტაცია", "მომხმარებლის კონკრეტულ საჭიროებებზე მორგებული შეთავაზებები და გადაწყვეტის გზები."),
            ("ავთენტური კომუნიკაცია", "ცოცხალი, პროფესიული ტონი რობოტული და შაბლონური ფრაზების გარეშე.")
        ]),
        ("💡 BRAINSTORMING & IDEATION", "იდეების გენერაცია", "შემოქმედებითი გაფართოება", [
            ("მრავალმხრივი პერსპექტივები", "პროდუქტის ფუნქციონალის ან მარკეტინგული კამპანიის 10 ალტერნატიული იდეის გენერირება."),
            ("მართული კრეატივი", "იდეების გენერაცია მკაცრად განსაზღვრული ბიუჯეტის, ტექნოლოგიური სტეკისა და აუდიტორიის ფარგლებში."),
            ("ტემპერატურის კონტროლი", "საშუალო/მაღალი ტემპერატურა (0.7 – 0.8) ვარიაციულობისა და მრავალფეროვნების გასაზრდელად.")
        ]),
        ("🛡️ MITIGATING FLUFF", "გაშლის შეზღუდვები", "წყლის დასხმის პრევენცია", [
            ("Over-Expansion რისკი", "მოდელის მიდრეკილება, დაწეროს უსარგებლო შესავალი და ბანალური ფრაზები (Fluff)."),
            ("Sycophancy (მლიქვნელობა)", "მოდელის ზედმეტი თავაზიანობა და პირფერობა მომხმარებლის მიმართ არსზე გადასვლის ნაცვლად."),
            ("მკაცრი უარყოფითი წესი", "სისტემური აკრძალვა: 'არ გამოიყენო შესავალი და თავაზიანი ფრაზები; მაქსიმუმ 3 აბზაცი; პირდაპირ გადადი არსზე'.")
        ])
    ]

    for i, (badge, title, subtitle, items) in enumerate(exp_s6):
        bg_col = COLOR_CARD_RED if i == 2 else COLOR_BG_LIGHT
        bor_col = COLOR_BORDER_RED if i == 2 else COLOR_BORDER
        c = s6.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, xs_3[i], card_y, card_w, card_h)
        c.fill.solid(); c.fill.fore_color.rgb = bg_col
        c.line.color.rgb = bor_col; c.line.width = Pt(1.5)

        tb = s6.shapes.add_textbox(xs_3[i] + Inches(0.35), card_y + Inches(0.35), card_w - Inches(0.7), card_h - Inches(0.7))
        tf = tb.text_frame; tf.word_wrap = True; tf.margin_left = tf.margin_top = 0

        p = tf.paragraphs[0]; p.text = badge.upper()
        p.font.name = FONT_ARIAL; p.font.size = Pt(11); p.font.bold = True
        p.font.color.rgb = COLOR_RED if i == 2 else COLOR_MUTED

        p = tf.add_paragraph(); p.text = title
        p.font.name = FONT_ARIAL; p.font.size = Pt(17); p.font.bold = True; p.font.color.rgb = COLOR_DARK; p.space_before = Pt(4)

        p = tf.add_paragraph(); p.text = subtitle
        p.font.name = FONT_ARIAL; p.font.size = Pt(12); p.font.color.rgb = COLOR_TEXT_MUTED; p.space_before = Pt(2)

        for heading, body in items:
            p = tf.add_paragraph(); p.text = f"• {heading}: {body}"
            p.font.name = FONT_ARIAL; p.font.size = Pt(12); p.font.color.rgb = COLOR_DARK; p.space_before = Pt(8)

    # Bottom takeaway
    bb = s6.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(1.0), bot_y, Inches(17.0), Inches(2.0))
    bb.fill.solid(); bb.fill.fore_color.rgb = COLOR_WHITE
    bb.line.color.rgb = COLOR_BORDER; bb.line.width = Pt(1.2)

    tbb = s6.shapes.add_textbox(Inches(1.4), bot_y + Inches(0.3), Inches(16.2), Inches(1.4))
    tfb = tbb.text_frame; tfb.word_wrap = True; tfb.margin_left = tfb.margin_top = 0
    p = tfb.paragraphs[0]; p.text = "💡 მთავარი არსი:"
    p.font.name = FONT_ARIAL; p.font.size = Pt(14); p.font.bold = True; p.font.color.rgb = COLOR_RED
    p = tfb.add_paragraph()
    p.text = "Expand ყველაზე სარისკო ოპერაციაა ჰალუცინაციებისა და ზედმეტი ტექსტის (Fluff) თვალსაზრისით; წარმატებული Expand ყოველთვის მოითხოვს მკაცრ უარყოფით შეზღუდვებსა და სტრუქტურულ ჩარჩოს."
    p.font.name = FONT_ARIAL; p.font.size = Pt(13); p.font.color.rgb = COLOR_MUTED; p.space_before = Pt(4)

    # -------------------------------------------------------------------------
    # SLIDE 7: Configuration Matrix of the 4 Tasks (Table)
    # -------------------------------------------------------------------------
    s7 = created_slides[5]
    add_header(s7, "4 ამოცანის შედარებითი კონფიგურაციის მატრიცა (System Blueprint)", "ᲨᲔᲓᲐᲠᲔᲑᲘᲗᲘ ᲐᲜᲐᲚᲘᲖᲘ • CONFIGURATION MATRIX")

    # Table: 5 rows (Header + 4 operations), 5 columns
    tbl_left = Inches(1.0); tbl_top = Inches(2.35); tbl_w = Inches(17.0); tbl_h = Inches(5.2)
    tbl_shape = s7.shapes.add_table(5, 5, tbl_left, tbl_top, tbl_w, tbl_h)
    table = tbl_shape.table

    col_widths = [Inches(2.6), Inches(4.0), Inches(2.8), Inches(3.2), Inches(4.4)]
    for ci, w in enumerate(col_widths):
        table.columns[ci].width = w

    headers = ["ამოცანის ტიპი", "ოპერაციის მიზანი", "ოპტიმალური Temp", "ჰალუცინაციის რისკი", "საკვანძო საინჟინრო შეზღუდვა"]
    for ci, h in enumerate(headers):
        cell = table.cell(0, ci)
        cell.fill.solid(); cell.fill.fore_color.rgb = COLOR_DARK
        tf = cell.text_frame; tf.word_wrap = True; tf.margin_left = tf.margin_right = Inches(0.12)
        p = tf.paragraphs[0]; p.text = h; p.font.name = FONT_ARIAL; p.font.size = Pt(12); p.font.bold = True; p.font.color.rgb = COLOR_WHITE

    matrix_rows = [
        ("Summarize (შეკუმშვა)", "ტექსტის შემოკლება და მთავარი აზრის ექსტრაქცია", "0.1 – 0.3 (დაბალი)", "საშუალო (ფაქტების გაზვიადება)", "მკაცრი სიგრძის ლიმიტი და წყაროზე მიჯაჭვულობა"),
        ("Infer (ანალიზი)", "იმპლიციტური ცოდნის, სენტიმენტისა და ველების ამოღება", "0.0 – 0.2 (მინიმალური)", "დაბალი (თუ სქემა მკაცრია)", "Pydantic/JSON Schema, Null Handling და ვალიდაცია"),
        ("Transform (გარდაქმნა)", "ფორმატის, სტილის ან ენის შეცვლა შინაარსის უცვლელობით", "0.1 – 0.4 (დაბალი)", "მინიმალური", "ინფორმაციის კონსერვაცია (აკრძალულია რამის დამატება)"),
        ("Expand (გაშლა)", "მოკლე ბრიფიდან დეტალური პასუხის ან იდეების გენერაცია", "0.7 – 0.9 (მაღალი)", "მაღალი (წყლის დასხმა, ფანტაზია)", "უარყოფითი წესები Fluff-ზე და აბზაცების მაქსიმალური რაოდენობა")
    ]

    for ri, row_data in enumerate(matrix_rows):
        row_bg = COLOR_WHITE if ri % 2 == 0 else COLOR_BG_LIGHT
        for ci, val in enumerate(row_data):
            cell = table.cell(ri + 1, ci)
            cell.fill.solid(); cell.fill.fore_color.rgb = row_bg
            tf = cell.text_frame; tf.word_wrap = True; tf.margin_left = tf.margin_right = Inches(0.12)
            p = tf.paragraphs[0]; p.text = val; p.font.name = FONT_ARIAL; p.font.size = Pt(11)
            if ci == 0:
                p.font.bold = True; p.font.color.rgb = COLOR_DARK
            elif ci == 2:
                p.font.color.rgb = COLOR_RED; p.font.bold = True
            elif ci == 4:
                p.font.bold = True; p.font.color.rgb = COLOR_DARK
            else:
                p.font.color.rgb = COLOR_MUTED

    # Bottom takeaway
    bb = s7.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(1.0), bot_y, Inches(17.0), Inches(2.0))
    bb.fill.solid(); bb.fill.fore_color.rgb = COLOR_WHITE
    bb.line.color.rgb = COLOR_BORDER; bb.line.width = Pt(1.2)

    tbb = s7.shapes.add_textbox(Inches(1.4), bot_y + Inches(0.3), Inches(16.2), Inches(1.4))
    tfb = tbb.text_frame; tfb.word_wrap = True; tfb.margin_left = tfb.margin_top = 0
    p = tfb.paragraphs[0]; p.text = "💡 შედარების მთავარი დასკვნა:"
    p.font.name = FONT_ARIAL; p.font.size = Pt(14); p.font.bold = True; p.font.color.rgb = COLOR_RED
    p = tfb.add_paragraph()
    p.text = "თითოეულ ამოცანას აქვს თავისი ოპტიმალური კონფიგურაცია; Infer და Summarize მოითხოვს მინიმალურ ტემპერატურასა და ფაქტობრივ სიზუსტეს, ხოლო Expand მოითხოვს მკაცრ სტრუქტურულ შეზღუდვებსა და Fluff-ის ბლოკირებას."
    p.font.name = FONT_ARIAL; p.font.size = Pt(13); p.font.color.rgb = COLOR_MUTED; p.space_before = Pt(4)

    # -------------------------------------------------------------------------
    # SLIDE 8: Common Pitfalls by Task Type
    # -------------------------------------------------------------------------
    s8 = created_slides[6]
    add_header(s8, "ტიპური შეცდომები 4 ბირთვულ ოპერაციაში და მათი პრევენცია", "ᲐᲜᲢᲘᲞᲐᲢᲔᲠᲜᲔᲑᲘ • COMMON PITFALLS BY TYPE")

    pitfalls_s8 = [
        ("❌ SUMMARIZE PITFALL", "Hallucinated Generalization", "ფაქტების გამოგონება", [
            ("შეცდომა", "მოდელი აკეთებს დაუსაბუთებელ განზოგადებებს და ამატებს ფაქტებს, რომლებიც ორიგინალში საერთოდ არ ეწერა."),
            ("პრევენცია", "ინსტრუქცია: 'შეაჯამე მხოლოდ ტექსტში პირდაპირ ნახსენები ფაქტები; ნუ გამოიტან ნაადრევ დასკვნებს'.")
        ]),
        ("❌ INFER PITFALL", "Over-Interpretation", "განზრახვის გაზვიადება", [
            ("შეცდომა", "მომხმარებლის მარტივ შეკითხვაში მოდელი ხედავს აგრესიას ან პირიქით, აიგნორებს სერიოზულ რისკს."),
            ("პრევენცია", "მკაფიო რუბრიკა, კლასიფიკაციის ზღვრები და 3-Shot მაგალითები თითოეული კატეგორიისთვის.")
        ]),
        ("❌ TRANSFORM PITFALL", "Loss of Content / Filler", "ინფორმაციის დანაკარგი", [
            ("შეცდომა", "გარდაქმნისას იკარგება კრიტიკული დეტალები ან მოდელი ამატებს თავაზიან ფრაზებს JSON-ის წინ."),
            ("პრევენცია", "Prefill ტექნიკა ({) და შეზღუდვა: 'დააბრუნე მხოლოდ გარდაქმნილი მონაცემი შესავლის გარეშე'.")
        ]),
        ("❌ EXPAND PITFALL", "Generic Fluff & Sycophancy", "წყლის დასხმა და პირფერობა", [
            ("შეცდომა", "გრძელი, უშინაარსო ტექსტი, ბანალური კლიშეები და მომხმარებლის ზედმეტი მლიქვნელობა."),
            ("პრევენცია", "აკრძალვა: 'არ გამოიყენო სიტყვები: რა თქმა უნდა, მოხარული ვარ; მაქსიმუმ 150 სიტყვა; პირდაპირ გადადი არსზე'.")
        ])
    ]

    for i, (badge, title, subtitle, items) in enumerate(pitfalls_s8):
        bg_col = COLOR_CARD_RED
        bor_col = COLOR_BORDER_RED
        c = s8.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, xs_4[i], step_y, step_w, step_h)
        c.fill.solid(); c.fill.fore_color.rgb = bg_col
        c.line.color.rgb = bor_col; c.line.width = Pt(1.5)

        tb = s8.shapes.add_textbox(xs_4[i] + Inches(0.25), step_y + Inches(0.25), step_w - Inches(0.5), step_h - Inches(0.5))
        tf = tb.text_frame; tf.word_wrap = True; tf.margin_left = tf.margin_top = 0

        p = tf.paragraphs[0]; p.text = badge.upper()
        p.font.name = FONT_ARIAL; p.font.size = Pt(11); p.font.bold = True
        p.font.color.rgb = COLOR_RED

        p = tf.add_paragraph(); p.text = title
        p.font.name = FONT_ARIAL; p.font.size = Pt(15); p.font.bold = True; p.font.color.rgb = COLOR_DARK; p.space_before = Pt(4)

        p = tf.add_paragraph(); p.text = subtitle
        p.font.name = FONT_ARIAL; p.font.size = Pt(11); p.font.color.rgb = COLOR_TEXT_MUTED; p.space_before = Pt(2)

        for heading, body in items:
            p = tf.add_paragraph(); p.text = f"• {heading}: {body}"
            p.font.name = FONT_ARIAL; p.font.size = Pt(11); p.font.color.rgb = COLOR_DARK; p.space_before = Pt(8)

    # Bottom takeaway
    bb = s8.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(1.0), bot_y, Inches(17.0), Inches(2.0))
    bb.fill.solid(); bb.fill.fore_color.rgb = COLOR_WHITE
    bb.line.color.rgb = COLOR_BORDER; bb.line.width = Pt(1.2)

    tbb = s8.shapes.add_textbox(Inches(1.4), bot_y + Inches(0.3), Inches(16.2), Inches(1.4))
    tfb = tbb.text_frame; tfb.word_wrap = True; tfb.margin_left = tfb.margin_top = 0
    p = tfb.paragraphs[0]; p.text = "💡 მთავარი არსი:"
    p.font.name = FONT_ARIAL; p.font.size = Pt(14); p.font.bold = True; p.font.color.rgb = COLOR_RED
    p = tfb.add_paragraph()
    p.text = "ანტიპატერნების ცოდნა საშუალებას გაძლევთ წინასწარ ჩააშენოთ თავდაცვითი შეზღუდვები შაბლონში, რაც 90%-ით ამცირებს მოდელის არასასურველ ქცევასა და ჰალუცინაციებს."
    p.font.name = FONT_ARIAL; p.font.size = Pt(13); p.font.color.rgb = COLOR_MUTED; p.space_before = Pt(4)

    # -------------------------------------------------------------------------
    # SLIDE 9: Enterprise Template Library & Transition to Lecture 6.2
    # -------------------------------------------------------------------------
    s9 = created_slides[7]
    add_header(s9, "შაბლონების საწარმოო ბიბლიოთეკა და გადასვლა Reasoning რეჟიმებზე", "ᲡᲐᲬᲐᲠᲛᲝᲝ ᲑᲘᲑᲚᲘᲝᲗᲔᲙᲐ • ENTERPRISE TEMPLATE BLUEPRINT")

    box_w = Inches(8.35); box_h = Inches(5.4)

    # Best Practices Box
    c_rec = s9.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(1.0), card_y, box_w, box_h)
    c_rec.fill.solid(); c_rec.fill.fore_color.rgb = COLOR_GREEN_BG
    c_rec.line.color.rgb = COLOR_GREEN_BOR; c_rec.line.width = Pt(1.5)

    tb_rec = s9.shapes.add_textbox(Inches(1.35), card_y + Inches(0.35), box_w - Inches(0.7), box_h - Inches(0.7))
    tf_rec = tb_rec.text_frame; tf_rec.word_wrap = True; tf_rec.margin_left = tf_rec.margin_top = 0
    p = tf_rec.paragraphs[0]; p.text = "✅ ENTERPRISE BLUEPRINT (საუკეთესო პატერნები)"
    p.font.name = FONT_ARIAL; p.font.size = Pt(13); p.font.bold = True; p.font.color.rgb = COLOR_GREEN_TEXT

    p = tf_rec.add_paragraph(); p.text = "მოდულური შაბლონების არქიტექტურა"
    p.font.name = FONT_ARIAL; p.font.size = Pt(16); p.font.bold = True; p.font.color.rgb = COLOR_DARK; p.space_before = Pt(4)

    rec_items = [
        ("Modular Task Templates", "ცალკეული, სპეციალიზებული შაბლონები თითოეული ოპერაციისთვის ერთიან Prompt Registry-ში."),
        ("Strict XML Delimiters", "შემავალი მონაცემების მკაცრი შეფუთვა ტეგებში (<source_text>, <target_format>)."),
        ("Programmatic Parameter Tuning", "ტემპერატურისა და Top-P-ის ავტომატური პროგრამული კონფიგურაცია ამოცანის ტიპის მიხედვით."),
        ("Regression Benchmarking", "თითოეული ტიპის შაბლონის სისტემატური ტესტირება შესაბამის Golden Evaluation Set-ზე.")
    ]
    for head, body in rec_items:
        p = tf_rec.add_paragraph(); p.text = f"• {head}: {body}"
        p.font.name = FONT_ARIAL; p.font.size = Pt(11); p.font.color.rgb = COLOR_DARK; p.space_before = Pt(6)

    # Next Steps Box
    c_next = s9.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(9.65), card_y, box_w, box_h)
    c_next.fill.solid(); c_next.fill.fore_color.rgb = COLOR_BG_LIGHT
    c_next.line.color.rgb = COLOR_BORDER; c_next.line.width = Pt(1.5)

    tb_next = s9.shapes.add_textbox(Inches(10.0), card_y + Inches(0.35), box_w - Inches(0.7), box_h - Inches(0.7))
    tf_next = tb_next.text_frame; tf_next.word_wrap = True; tf_next.margin_left = tf_next.margin_top = 0
    p = tf_next.paragraphs[0]; p.text = "🧠 ADVANCED REASONING (რთული ლოგიკა)"
    p.font.name = FONT_ARIAL; p.font.size = Pt(13); p.font.bold = True; p.font.color.rgb = COLOR_RED

    p = tf_next.add_paragraph(); p.text = "როდესაც 4 ოპერაცია საკმარისი აღარ არის"
    p.font.name = FONT_ARIAL; p.font.size = Pt(16); p.font.bold = True; p.font.color.rgb = COLOR_DARK; p.space_before = Pt(4)

    next_items = [
        ("CoT Paradigm Shift", "როდესაც ამოცანა მოითხოვს მრავალსაფეხურიან ლოგიკურ მსჯელობას, მარტივი ტრანსფორმაცია მარცხდება."),
        ("Reasoning Models 2026", "Frontier მოდელების (GPT-6 Sol, Claude 5.x Extended Thinking) შიდა მსჯელობის ტოკენების ძალა."),
        ("Cost & Latency Trade-offs", "როდის ღირს Reasoning მოდელების ჩართვა და როდის არის ის რესურსების გაუმართლებელი ფლანგვა."),
        ("10 ლოგიკური ამოცანის ანალიზი", "პრაქტიკული ექსპერიმენტი: როგორ ცვლის Chain-of-Thought შედეგის სიზუსტეს.")
    ]
    for head, body in next_items:
        p = tf_next.add_paragraph(); p.text = f"• {head}: {body}"
        p.font.name = FONT_ARIAL; p.font.size = Pt(11); p.font.color.rgb = COLOR_DARK; p.space_before = Pt(6)

    # Bottom takeaway & transition to Lecture 6.2
    bb = s9.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(1.0), bot_y, Inches(17.0), Inches(2.0))
    bb.fill.solid(); bb.fill.fore_color.rgb = COLOR_WHITE
    bb.line.color.rgb = COLOR_BORDER; bb.line.width = Pt(1.2)

    tbb = s9.shapes.add_textbox(Inches(1.4), bot_y + Inches(0.25), Inches(16.2), Inches(1.5))
    tfb = tbb.text_frame; tfb.word_wrap = True; tfb.margin_left = tfb.margin_top = 0
    p = tfb.paragraphs[0]; p.text = "💡 საინჟინრო შეჯამება და გადასვლა ლექციაზე 6.2:"
    p.font.name = FONT_ARIAL; p.font.size = Pt(14); p.font.bold = True; p.font.color.rgb = COLOR_RED
    p = tfb.add_paragraph()
    p.text = "4 საბაზისო ოპერაცია ქმნის AI აპლიკაციების საინჟინრო ბირთვს. შემდეგ ლექციაში (6.2: Chain-of-Thought, Reasoning მოდელები და Extended Thinking) შევისწავლით: როგორ ვაიძულოთ მოდელი იაზროვნოს ნაბიჯ-ნაბიჯ, როგორ მუშაობს Extended Thinking 2026 წელს და 10 ლოგიკური ამოცანის ექსპერიმენტულ ანალიზს CoT-ით და უიმისოდ."
    p.font.name = FONT_ARIAL; p.font.size = Pt(12.5); p.font.color.rgb = COLOR_MUTED; p.space_before = Pt(4)

    # -------------------------------------------------------------------------
    # FINAL POST-PROCESSING: Strict 100% Arial & Clean Notes
    # -------------------------------------------------------------------------
    for slide in prs.slides:
        if slide.has_notes_slide:
            notes_tf = slide.notes_slide.notes_text_frame
            notes_tf.text = ""

        for shape in slide.shapes:
            if shape.has_text_frame:
                for paragraph in shape.text_frame.paragraphs:
                    paragraph.font.name = FONT_ARIAL
                    for run in paragraph.runs:
                        run.font.name = FONT_ARIAL
            elif shape.has_table:
                for row in shape.table.rows:
                    for cell in row.cells:
                        for paragraph in cell.text_frame.paragraphs:
                            paragraph.font.name = FONT_ARIAL
                            for run in paragraph.runs:
                                run.font.name = FONT_ARIAL

    # -------------------------------------------------------------------------
    # SAVE PPTX AND DISTRIBUTE TO TARGET DIRECTORIES
    # -------------------------------------------------------------------------
    filename = "6.1 - ამოცანის 4 ძირითადი ტიპი და შაბლონების ბიბლიოთეკა.pptx"
    out_master = os.path.join(r'd:\ხელოვნური ინტელექტის ინსტრუმენტები', filename)
    prs.save(out_master)
    print(f"Master PPTX generated successfully: {out_master}")

    target_dirs = [
        r'd:\ხელოვნური ინტელექტის ინსტრუმენტები\Prompt ინჟინერია და უსაფრთხოება',
        r'D:\Prompt ინჟინერია და უსაფრთხოება',
        r'G:\My Drive\Skillwill\Prompt ინჟინერია და უსაფრთხოება',
        r'G:\My Drive\Prompt ინჟინერია და უსაფრთხოება'
    ]

    for td in target_dirs:
        try:
            if not os.path.exists(td):
                os.makedirs(td, exist_ok=True)
            dst_file = os.path.join(td, filename)
            shutil.copy2(out_master, dst_file)
            print(f"Copied to target: {dst_file}")
        except Exception as e:
            print(f"Error copying to {td}: {e}")

if __name__ == '__main__':
    build_lecture_m7_6_1()
