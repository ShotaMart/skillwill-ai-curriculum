import os
import sys
import shutil
import pptx
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN
from pptx.enum.shapes import MSO_SHAPE

def build_lecture_4():
    template_path = r'd:\ხელოვნური ინტელექტის ინსტრუმენტები\Skillwill Presentation Example .pptx'
    asset_dir = r'd:\ხელოვნური ინტელექტის ინსტრუმენტები\assets'
    img_icon_title = os.path.join(asset_dir, 'icon_title.png')
    logo_dir = os.path.join(asset_dir, 'logos')

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
    COLOR_GREEN = RGBColor(0x13, 0x73, 0x33)
    COLOR_GREEN_BG = RGBColor(0xE6, 0xF4, 0xEA)
    COLOR_BLUE = RGBColor(0x1A, 0x73, 0xE8)
    COLOR_BLUE_BG = RGBColor(0xE8, 0xF0, 0xFE)

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
            p.text = "4 - Claude სრულად I: ინტერფეისი, ფაილები და Artifacts"
            p.alignment = PP_ALIGN.CENTER
            p.font.name = FONT_ARIAL
            p.font.size = Pt(38)
            p.font.bold = True
            p.font.color.rgb = COLOR_RED

            p2 = tf.add_paragraph()
            p2.text = "Anthropic-ის უახლესი AI ასისტენტის შესაძლებლობები და სამუშაო ფუნქციონალი"
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
    def add_header(slide, title, category="ᲛᲝᲓᲣᲚᲘ 5 • ᲨᲔᲡᲐᲕᲐᲚᲘ AI • 4"):
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

    # -------------------------------------------------------------------------
    # SLIDE 2: Anthropic & Claude Philosophy
    # -------------------------------------------------------------------------
    s2 = created_slides[0]
    add_header(s2, "Anthropic და Claude (უსაფრთხოება და ანალიტიკური სიღრმე)", "ფილოსოფია • რით განსხვავდება Claude სხვებისგან?")

    box_w = Inches(8.35); box_h = Inches(5.4); box_y = Inches(2.35)

    # Left: Company Profile
    b1 = s2.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(1.0), box_y, box_w, box_h)
    b1.fill.solid(); b1.fill.fore_color.rgb = COLOR_BG_LIGHT
    b1.line.color.rgb = COLOR_BORDER; b1.line.width = Pt(1.5)

    tb1 = s2.shapes.add_textbox(Inches(1.4), box_y + Inches(0.4), box_w - Inches(0.8), box_h - Inches(0.8))
    tf1 = tb1.text_frame; tf1.word_wrap = True; tf1.margin_left = tf1.margin_top = 0
    p = tf1.paragraphs[0]; p.text = "ᲙᲝᲛᲞᲐᲜᲘᲘᲡ ᲞᲠᲝᲤᲘᲚᲘ"
    p.font.name = FONT_ARIAL; p.font.size = Pt(12); p.font.bold = True; p.font.color.rgb = COLOR_MUTED

    p = tf1.add_paragraph(); p.text = "Anthropic: AI Safety Lab"
    p.font.name = FONT_ARIAL; p.font.size = Pt(22); p.font.bold = True; p.font.color.rgb = COLOR_DARK; p.space_before = Pt(4)

    ant_points = [
        ("დაფუძნება და მისია", "OpenAI-ის ყოფილი წამყვანი მკვლევრების მიერ დაფუძნებული ლაბორატორია უსაფრთხო და სანდო AI-ს შესაქმნელად."),
        ("Constitutional AI", "მოდელი გაწვრთნილია მკაფიო ეთიკურ პრინციპებზე, რაც გამორიცხავს მავნე, აგრესიულ ან არაპროგნოზირებად ქცევას."),
        ("გლობალური ინფრასტრუქტურა", "Amazon (AWS) და Google Cloud-ის სტრატეგიული პარტნიორობა და მძლავრი გამოთვლითი რესურსები."),
        ("სამიზნე აუდიტორია", "პროფესიონალები, მკვლევრები და კომპანიები, რომლებსაც მაღალი სიზუსტე და მონაცემთა დაცვა სჭირდებათ.")
    ]
    for t, d in ant_points:
        p = tf1.add_paragraph(); p.text = f"• {t}: {d}"
        p.font.name = FONT_ARIAL; p.font.size = Pt(13); p.font.color.rgb = COLOR_TEXT_MUTED; p.space_before = Pt(8)

    # Right: Unique Strengths
    b2 = s2.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(9.65), box_y, box_w, box_h)
    b2.fill.solid(); b2.fill.fore_color.rgb = COLOR_CARD_RED
    b2.line.color.rgb = COLOR_BORDER_RED; b2.line.width = Pt(1.5)

    tb2 = s2.shapes.add_textbox(Inches(10.05), box_y + Inches(0.4), box_w - Inches(0.8), box_h - Inches(0.8))
    tf2 = tb2.text_frame; tf2.word_wrap = True; tf2.margin_left = tf2.margin_top = 0
    p = tf2.paragraphs[0]; p.text = "CLAUDE-ᲘᲡ 3 ᲣᲛᲗᲐᲕᲠᲔᲡᲘ ᲫᲐᲚᲐ"
    p.font.name = FONT_ARIAL; p.font.size = Pt(12); p.font.bold = True; p.font.color.rgb = COLOR_RED

    p = tf2.add_paragraph(); p.text = "რატომ ირჩევენ პროფესიონალები?"
    p.font.name = FONT_ARIAL; p.font.size = Pt(22); p.font.bold = True; p.font.color.rgb = COLOR_DARK; p.space_before = Pt(4)

    str_points = [
        ("Natural & Nuanced Language (ბუნებრივი ენა)", "აკადემიური, გამართული და ზრდილობიანი ტონი. ქართულ ენაზე ავლენს უმაღლეს სტილისტურ და გრამატიკულ სისწორეს."),
        ("Massive Context Window (200K+ ტოკენი)", "ერთ საუბარში თავისუფლად იტევს 500+ გვერდიან წიგნს, კონტრაქტს ან კოდის მთლიან რეპოზიტორიას."),
        ("Logic & Coding Leader (ლოგიკისა და კოდის ლიდერი)", "გლობალურ ინდუსტრიულ ბენჩმარკებში რთული მსჯელობისა და პროგრამირების აღიარებული ჩემპიონი.")
    ]
    for t, d in str_points:
        p = tf2.add_paragraph(); p.text = f"⭐ {t}"
        p.font.name = FONT_ARIAL; p.font.size = Pt(13); p.font.bold = True; p.font.color.rgb = COLOR_DARK; p.space_before = Pt(8)
        p = tf2.add_paragraph(); p.text = d
        p.font.name = FONT_ARIAL; p.font.size = Pt(12); p.font.color.rgb = COLOR_TEXT_MUTED; p.space_before = Pt(2)

    # Bottom takeaway
    bot_y = Inches(8.0)
    bb = s2.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(1.0), bot_y, Inches(17.0), Inches(2.0))
    bb.fill.solid(); bb.fill.fore_color.rgb = COLOR_WHITE
    bb.line.color.rgb = COLOR_BORDER; bb.line.width = Pt(1.2)

    tbb = s2.shapes.add_textbox(Inches(1.4), bot_y + Inches(0.3), Inches(16.2), Inches(1.4))
    tfb = tbb.text_frame; tfb.word_wrap = True; tfb.margin_left = tfb.margin_top = 0
    p = tfb.paragraphs[0]; p.text = "💡 მთავარი არსი (მარტივი ანალოგია):"
    p.font.name = FONT_ARIAL; p.font.size = Pt(14); p.font.bold = True; p.font.color.rgb = COLOR_RED
    p = tfb.add_paragraph()
    p.text = "თუ ტრადიციული ჩატბოტები მოკლე კითხვა-პასუხზე არიან ორიენტირებულნი, Claude მუშაობს როგორც ანალიტიკოსი პარტნიორი: კითხულობს დიდ დოკუმენტებს მთლიანობაში, იაზრებს კონტექსტს და არ ავიწყდება დეტალები."
    p.font.name = FONT_ARIAL; p.font.size = Pt(13); p.font.color.rgb = COLOR_DARK; p.space_before = Pt(4)

    # -------------------------------------------------------------------------
    # SLIDE 3: Claude 5.x Model Family (2026 Baseline: Haiku, Sonnet, Opus)
    # -------------------------------------------------------------------------
    s3 = created_slides[1]
    add_header(s3, "Claude 5.x მოდელების ოჯახი (Haiku, Sonnet, Opus)", "მოდელები • 2026 წლის აქტუალური თაობა და დანიშნულება")

    models_claude = [
        {
            "tag": "01 • ᲡᲘᲡᲬᲠᲐᲤᲔ & ᲔᲙᲝᲜᲝᲛᲘᲐ",
            "name": "Claude 5.2 Haiku",
            "role": "ულტრა-სწრაფი და მსუბუქი მოდელი",
            "desc": "შექმნილია წამიერი რეაგირებისთვის. იდეალურია ყოველდღიური რუტინული ამოცანებისთვის, მომხმარებელთა მხარდაჭერის ჩატბოტებისთვის და მოკლე ტექსტური შეჯამებისთვის.",
            "pros": "მინიმალური ლატენტობა (სისწრაფე), ეკონომიური რესურსი",
            "color": COLOR_BG_LIGHT, "bcolor": COLOR_BORDER, "tcolor": COLOR_MUTED
        },
        {
            "tag": "02 • ᲝᲥᲠᲝᲡ ᲡᲢᲐᲜᲓᲐᲠᲢᲘ (ᲤᲚᲐᲒᲛᲐᲜᲘ)",
            "name": "Claude 5.5 Sonnet",
            "role": "ინდუსტრიის მთავარი სამუშაო ფლაგმანი",
            "desc": "მთავარი სამუშაო მოდელი 95% პროფესიული ამოცანისთვის. აერთიანებს მაქსიმალურ ანალიტიკურ სიღრმეს, უმაღლეს კოდინგს, Artifacts-ის გენერირებასა და სწრაფ პასუხს.",
            "pros": "ოპტიმალური ინტელექტი, სისწრაფე და სიზუსტე",
            "color": COLOR_CARD_RED, "bcolor": COLOR_BORDER_RED, "tcolor": COLOR_RED
        },
        {
            "tag": "03 • ᲛᲐᲥᲡᲘᲛᲐᲚᲣᲠᲘ ᲘᲜᲢᲔᲚᲔᲥᲢᲘ",
            "name": "Claude 5.5 Opus",
            "role": "ურთულესი კვლევებისა და მეცნიერების მოდელი",
            "desc": "ყველაზე მძლავრი მოდელი უკიდურესად რთული მათემატიკური, ლოგიკური და სამეცნიერო პრობლემებისთვის, სადაც სიჩქარეზე მნიშვნელოვანი უზუსტესი მსჯელობაა.",
            "pros": "მაქსიმალური მსჯელობის სიღრმე და აბსტრაქცია",
            "color": COLOR_BG_LIGHT, "bcolor": COLOR_BORDER, "tcolor": COLOR_MUTED
        }
    ]

    card_w = Inches(5.4); card_gap = Inches(0.4); start_x = Inches(1.0); card_y = Inches(2.35)
    for i, mod in enumerate(models_claude):
        cx = start_x + i * (card_w + card_gap)

        c = s3.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, cx, card_y, card_w, Inches(7.7))
        c.fill.solid(); c.fill.fore_color.rgb = mod["color"]
        c.line.color.rgb = mod["bcolor"]; c.line.width = Pt(1.5)

        # Centered square logo 1.0" x 1.0"
        lp = os.path.join(logo_dir, 'claude.png')
        if os.path.exists(lp):
            s3.shapes.add_picture(lp, cx + Inches(0.4), card_y + Inches(0.4), width=Inches(1.0), height=Inches(1.0))

        tb = s3.shapes.add_textbox(cx + Inches(1.6), card_y + Inches(0.4), card_w - Inches(2.0), Inches(1.2))
        tf = tb.text_frame; tf.word_wrap = True; tf.margin_left = tf.margin_top = 0
        p = tf.paragraphs[0]; p.text = mod["tag"]
        p.font.name = FONT_ARIAL; p.font.size = Pt(11); p.font.bold = True; p.font.color.rgb = mod["tcolor"]

        p = tf.add_paragraph(); p.text = mod["name"]
        p.font.name = FONT_ARIAL; p.font.size = Pt(20); p.font.bold = True; p.font.color.rgb = COLOR_DARK; p.space_before = Pt(2)

        # Body text
        tb_body = s3.shapes.add_textbox(cx + Inches(0.4), card_y + Inches(1.8), card_w - Inches(0.8), Inches(5.3))
        tfb = tb_body.text_frame; tfb.word_wrap = True; tfb.margin_left = tfb.margin_top = 0

        p = tfb.paragraphs[0]; p.text = f"🎯 დანიშნულება:\n{mod['role']}"
        p.font.name = FONT_ARIAL; p.font.size = Pt(14); p.font.bold = True; p.font.color.rgb = COLOR_DARK

        p = tfb.add_paragraph(); p.text = f"\n{mod['desc']}"
        p.font.name = FONT_ARIAL; p.font.size = Pt(13); p.font.color.rgb = COLOR_TEXT_MUTED; p.space_before = Pt(4)

        p = tfb.add_paragraph(); p.text = f"\n⭐ მთავარი უპირატესობა:\n{mod['pros']}"
        p.font.name = FONT_ARIAL; p.font.size = Pt(13); p.font.bold = True; p.font.color.rgb = COLOR_RED if i == 1 else COLOR_MUTED; p.space_before = Pt(8)

    # -------------------------------------------------------------------------
    # SLIDE 4: Workspace Functional Architecture (NO Screenshots, Pure Structure!)
    # -------------------------------------------------------------------------
    s4 = created_slides[2]
    add_header(s4, "Claude.ai სამუშაო გარემოს ძირითადი ფუნქციები", "ფუნქციონალი • სამუშაო შესაძლებლობების რუკა")

    ui_pillars = [
        {
            "tag": "01 • ᲛᲝᲓᲔᲚᲘᲡ ᲛᲐᲠᲗᲕᲐ",
            "title": "Model Selection & Controls",
            "sub": "მოდელის შერჩევა ამოცანის მიხედვით",
            "points": [
                "სასურველი მოდელის (Sonnet, Haiku, Opus) მარტივი გადართვა ჩატის დაწყებამდე.",
                "Sonnet — ნაგულისხმევი არჩევანი ანალიტიკისა და კოდინგისთვის.",
                "Haiku — გადართვა წამიერი და მარტივი რუტინული ამოცანებისას.",
                "კონტექსტის ზომისა და რესურსების ოპტიმიზაცია."
            ]
        },
        {
            "tag": "02 • ᲤᲐᲘᲚᲔᲑᲘᲡ ᲨᲔᲧᲕᲐᲜᲐ",
            "title": "Multimodal Input (Paperclip)",
            "sub": "ფაილების ატვირთვა პირდაპირ ჩატში",
            "points": [
                "დოკუმენტების მიმაგრება: PDF, Word, Excel, CSV, TXT, კოდის ფაილები.",
                "სურათებისა და დიაგრამების ატვირთვა (Vision ანალიზი).",
                "რამდენიმე ფაილის ერთდროული ატვირთვა ჯვარედინი შედარებისთვის.",
                "ფაილის წამიერი ინდექსაცია კონტექსტის მეხსიერებაში."
            ]
        },
        {
            "tag": "03 • ᲩᲐᲢᲘᲡ ᲛᲐᲠᲗᲕᲐ",
            "title": "Chat Management & History",
            "sub": "დიალოგების ორგანიზება და ძიება",
            "points": [
                "საუბრების ავტომატური დახარისხება და დასათაურება თემატიკის მიხედვით.",
                "სწრაფი ძიება (Ctrl+K) წინა ჩატების ისტორიაში საკვანძო სიტყვებით.",
                "ჩატის გადარქმევა, დაარქივება და წაშლა.",
                "ახალი სუფთა სესიის დაწყება კონტექსტის გასასუფთავებლად."
            ]
        },
        {
            "tag": "04 • ᲞᲐᲡᲣᲮᲘᲡ ᲤᲝᲠᲛᲐᲢᲘᲠᲔᲑᲐ",
            "title": "Rich Canvas & Output",
            "sub": "სტრუქტურირებული პასუხები",
            "points": [
                "პასუხების ფორმატირება Markdown სტილით (სათაურები, ბულეტები, ხაზგასმა).",
                "კოდის ბლოკები ენის ავტომატური ამოცნობითა და 1-კლიკიანი კოპირებით.",
                "დიდი ანალიტიკური ცხრილები და მონაცემთა შედარება.",
                "Artifacts რეჟიმის ავტომატური გააქტიურება ცალკე ფანჯარაში."
            ]
        }
    ]

    card_w = Inches(8.35); card_h = Inches(3.6); card_gap_x = Inches(0.4); card_gap_y = Inches(0.4)
    for i, pil in enumerate(ui_pillars):
        row = i // 2
        col = i % 2
        cx = Inches(1.0) + col * (card_w + card_gap_x)
        cy = Inches(2.35) + row * (card_h + card_gap_y)

        c = s4.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, cx, cy, card_w, card_h)
        c.fill.solid()
        c.fill.fore_color.rgb = COLOR_CARD_RED if i == 0 else COLOR_BG_LIGHT
        c.line.color.rgb = COLOR_BORDER_RED if i == 0 else COLOR_BORDER
        c.line.width = Pt(1.5)

        tb = s4.shapes.add_textbox(cx + Inches(0.4), cy + Inches(0.3), card_w - Inches(0.8), card_h - Inches(0.6))
        tf = tb.text_frame; tf.word_wrap = True; tf.margin_left = tf.margin_top = 0

        p = tf.paragraphs[0]; p.text = pil["tag"]
        p.font.name = FONT_ARIAL; p.font.size = Pt(11); p.font.bold = True
        p.font.color.rgb = COLOR_RED if i == 0 else COLOR_MUTED

        p = tf.add_paragraph(); p.text = pil["title"]
        p.font.name = FONT_ARIAL; p.font.size = Pt(17); p.font.bold = True; p.font.color.rgb = COLOR_DARK; p.space_before = Pt(2)

        for pt in pil["points"]:
            p = tf.add_paragraph(); p.text = f"• {pt}"
            p.font.name = FONT_ARIAL; p.font.size = Pt(12); p.font.color.rgb = COLOR_TEXT_MUTED; p.space_before = Pt(4)

    # -------------------------------------------------------------------------
    # SLIDE 5: Multimodal File Handling (PDF, Excel, Images)
    # -------------------------------------------------------------------------
    s5 = created_slides[3]
    add_header(s5, "ფაილების დამუშავება: PDF, მონაცემთა ცხრილები და სურათები", "მონაცემები • რა დოკუმენტების „წაკითხვა“ შეუძლია Claude-ს?")

    file_types = [
        {
            "tag": "01 • ᲓᲝᲙᲣᲛᲔᲜᲢᲔᲑᲘ",
            "title": "PDF, Word & Text",
            "sub": "დიდი ტექსტების ანალიტიკა და შეჯამება",
            "points": [
                "100+ გვერდიანი ხელშეკრულებების, კვლევებისა და წიგნების წამიერი დამუშავება.",
                "ძირითადი თეზისების, იურიდიული რისკებისა და საკვანძო პირობების ამოღება.",
                "კითხვა-პასუხი დოკუმენტის შინაარსზე (ზუსტი მითითებებითა და ციტატებით).",
                "დოკუმენტის სტილისტური გადამუშავება და ენობრივი თარგმანი."
            ]
        },
        {
            "tag": "02 • ᲛᲝᲜᲐᲪᲔᲛᲔᲑᲘ",
            "title": "Excel & CSV ცხრილები",
            "sub": "სტრუქტურირებული მონაცემები",
            "points": [
                "ფინანსური რეპორტების, გაყიდვებისა და მომხმარებელთა სიების წაკითხვა.",
                "მონაცემთა გაფილტვრა, დაჯგუფება და ტენდენციების აღმოჩენა.",
                "მზა ფორმულების (Excel / SQL) და ანალიტიკური დასკვნების მომზადება.",
                "ახალი, გასუფთავებული და ფორმატირებული ცხრილების გენერირება."
            ]
        },
        {
            "tag": "03 • ᲕᲘᲖᲣᲐᲚᲘ",
            "title": "Vision (სურათები & ფოტო)",
            "sub": "გრაფიკებისა და ჩანაწერების ამოცნობა",
            "points": [
                "რთული დიაგრამების, ინფოგრაფიკისა და გრაფიკების ახსნა ციფრებში.",
                "ხელნაწერი ტექსტების, ჩეკებისა და დოკუმენტების ფოტოების გაციფრულება (OCR).",
                "UI დიზაინისა და ვებსაიტის მონახაზის (Wireframe) გაანალიზება.",
                "ტექნიკური სქემებისა და ნახაზების შინაარსობრივი წაკითხვა."
            ]
        }
    ]

    card_w = Inches(5.4); card_gap = Inches(0.4); start_x = Inches(1.0); card_y = Inches(2.35)
    for i, ft in enumerate(file_types):
        cx = start_x + i * (card_w + card_gap)

        c = s5.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, cx, card_y, card_w, Inches(7.7))
        c.fill.solid()
        c.fill.fore_color.rgb = COLOR_CARD_RED if i == 0 else COLOR_BG_LIGHT
        c.line.color.rgb = COLOR_BORDER_RED if i == 0 else COLOR_BORDER
        c.line.width = Pt(1.5)

        tb = s5.shapes.add_textbox(cx + Inches(0.4), card_y + Inches(0.4), card_w - Inches(0.8), Inches(6.9))
        tf = tb.text_frame; tf.word_wrap = True; tf.margin_left = tf.margin_top = 0

        p = tf.paragraphs[0]; p.text = ft["tag"]
        p.font.name = FONT_ARIAL; p.font.size = Pt(11); p.font.bold = True
        p.font.color.rgb = COLOR_RED if i == 0 else COLOR_MUTED

        p = tf.add_paragraph(); p.text = ft["title"]
        p.font.name = FONT_ARIAL; p.font.size = Pt(22); p.font.bold = True; p.font.color.rgb = COLOR_DARK; p.space_before = Pt(4)

        p = tf.add_paragraph(); p.text = ft["sub"]
        p.font.name = FONT_ARIAL; p.font.size = Pt(13); p.font.bold = True; p.font.color.rgb = COLOR_MUTED; p.space_before = Pt(2)

        for pt in ft["points"]:
            p = tf.add_paragraph(); p.text = f"• {pt}"
            p.font.name = FONT_ARIAL; p.font.size = Pt(13); p.font.color.rgb = COLOR_TEXT_MUTED; p.space_before = Pt(8)

    # -------------------------------------------------------------------------
    # SLIDE 6: What are Artifacts? (Conceptual Architecture, NO Screenshots)
    # -------------------------------------------------------------------------
    s6 = created_slides[4]
    add_header(s6, "რა არის Artifacts? — ინტერაქტიული სამუშაო სივრცე", "რევოლუციური ფუნქცია • გაყოფილი ეკრანის პარადიგმა")

    top_y = Inches(2.35)
    tb_intro6 = s6.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(1.0), top_y, Inches(17.0), Inches(1.5))
    tb_intro6.fill.solid(); tb_intro6.fill.fore_color.rgb = COLOR_CARD_RED
    tb_intro6.line.color.rgb = COLOR_BORDER_RED; tb_intro6.line.width = Pt(1.5)

    tb6_t = s6.shapes.add_textbox(Inches(1.4), top_y + Inches(0.25), Inches(16.2), Inches(1.0))
    tf6 = tb6_t.text_frame; tf6.word_wrap = True; tf6.margin_left = tf6.margin_top = 0
    p = tf6.paragraphs[0]; p.text = "🚀 პარადიგმის ცვლილება: ჩატბოტიდან სრულფასოვან სამუშაო სივრცემდე"
    p.font.name = FONT_ARIAL; p.font.size = Pt(14); p.font.bold = True; p.font.color.rgb = COLOR_RED
    p = tf6.add_paragraph()
    p.text = "ტრადიციულ ჩატში გრძელი კოდი ან დოკუმენტი დიალოგის ნაკადში იკარგებოდა. Artifacts ეკრანს ორად ყოფს: მარცხნივ მიმდინარეობს ჩატი, ხოლო მარჯვნივ იხსნება დამოუკიდებელი, ცოცხალი ინტერაქტიული პანელი!"
    p.font.name = FONT_ARIAL; p.font.size = Pt(13); p.font.color.rgb = COLOR_DARK; p.space_before = Pt(4)

    # 3 core pillars of Artifacts
    art_pillars = [
        {
            "tag": "01 • ᲠᲝᲓᲘᲡ ᲘᲥᲛᲜᲔᲑᲐ?",
            "title": "Self-Contained Content",
            "sub": "დამოუკიდებელი ობიექტი",
            "desc": "არტეფაქტი იქმნება მაშინ, როდესაც კონტენტი არის სრულყოფილი და გამოსაყენებელი: კოდი, დიაგრამა, ვებგვერდი ან დიდი დოკუმენტი (15+ ხაზი), რომელსაც მომხმარებელი შემდგომ გამოიყენებს."
        },
        {
            "tag": "02 • ᲠᲝᲒᲝᲠ ᲛᲣᲨᲐᲝᲑᲡ?",
            "title": "Live Interactive Preview",
            "sub": "ცოცხალი გაშვება ბრაუზერში",
            "desc": "HTML, React, CSS, SVG და Mermaid დიაგრამები პირდაპირ ეკრანზე მუშაობს. მომხმარებელს შეუძლია ღილაკებზე დაჭერა, მონაცემების შეყვანა და შედეგის მომენტალური ნახვა."
        },
        {
            "tag": "03 • ᲕᲔᲠᲡᲘᲔᲑᲘᲡ ᲙᲝᲜᲢᲠᲝᲚᲘ",
            "title": "Iterative Versioning",
            "sub": "მუდმივი გაუმჯობესება",
            "desc": "თქვენ შეგიძლიათ ჩატში მისცეთ დამატებითი შენიშვნა („შეცვალე ფერი“, „დაამატე ღილაკი“). Claude განაახლებს მარჯვენა არტეფაქტს, ძველი ვერსიები კი ისტორიაში ინახება."
        }
    ]

    card_w = Inches(5.4); card_gap = Inches(0.4); start_x = Inches(1.0); card_y = Inches(4.2)
    for i, col in enumerate(art_pillars):
        cx = start_x + i * (card_w + card_gap)

        c = s6.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, cx, card_y, card_w, Inches(5.85))
        c.fill.solid(); c.fill.fore_color.rgb = COLOR_BG_LIGHT
        c.line.color.rgb = COLOR_BORDER; c.line.width = Pt(1.5)

        tb = s6.shapes.add_textbox(cx + Inches(0.4), card_y + Inches(0.4), card_w - Inches(0.8), Inches(5.0))
        tf = tb.text_frame; tf.word_wrap = True; tf.margin_left = tf.margin_top = 0

        p = tf.paragraphs[0]; p.text = col["tag"]
        p.font.name = FONT_ARIAL; p.font.size = Pt(11); p.font.bold = True; p.font.color.rgb = COLOR_MUTED

        p = tf.add_paragraph(); p.text = col["title"]
        p.font.name = FONT_ARIAL; p.font.size = Pt(19); p.font.bold = True; p.font.color.rgb = COLOR_DARK; p.space_before = Pt(4)

        p = tf.add_paragraph(); p.text = col["sub"]
        p.font.name = FONT_ARIAL; p.font.size = Pt(13); p.font.bold = True; p.font.color.rgb = COLOR_RED; p.space_before = Pt(2)

        p = tf.add_paragraph(); p.text = f"\n{col['desc']}"
        p.font.name = FONT_ARIAL; p.font.size = Pt(13); p.font.color.rgb = COLOR_TEXT_MUTED; p.space_before = Pt(6)

    # -------------------------------------------------------------------------
    # SLIDE 7: Artifacts 3 Main Categories
    # -------------------------------------------------------------------------
    s7 = created_slides[5]
    add_header(s7, "Artifacts-ის 3 ძირითადი კატეგორია და შესაძლებლობები", "პრაქტიკა • რა იქმნება არტეფაქტის პანელში?")

    use_cases = [
        {
            "tag": "01 • ᲕᲘᲖᲣᲐᲚᲘᲖᲐᲪᲘᲐ",
            "title": "Interactive Diagrams",
            "geo": "(ინტერაქტიული დიაგრამები)",
            "desc": "ბიზნეს-პროცესების ნაკადები, ორგანიზაციული სტრუქტურები და ლოგიკური ბლოკ-სქემები (Mermaid, SVG).",
            "meta": "მაგალითი: გაყიდვების პროცესის სქემა, მომხმარებლის გზა (User Journey), ალგორითმის რუკა."
        },
        {
            "tag": "02 • ᲓᲝᲙᲣᲛᲔᲜᲢᲐᲪᲘᲐ",
            "title": "Standalone Documents",
            "geo": "(სრულყოფილი დოკუმენტები)",
            "desc": "სრული წესდებები, ინსტრუქციები, სტატიები და ხელშეკრულებები, რომლებიც მზადაა ჩამოსატვირთად და გასაზიარებლად.",
            "meta": "მაგალითი: კომპანიის შიდა რეგლამენტი, პროექტის ტექნიკური დავალება (PRD), ანალიტიკური რეპორტი."
        },
        {
            "tag": "03 • ᲞᲠᲝᲢᲝᲢᲘᲞᲔᲑᲘ",
            "title": "Live Code & Apps",
            "geo": "(ცოცხალი აპლიკაციები & კოდი)",
            "desc": "ინტერაქტიული კალკულატორები, ვიზუალური დაშბორდები, მინი-თამაშები და ვებსაიტის ინტერფეისები React/HTML/JS-ში.",
            "meta": "მაგალითი: იპოთეკური სესხის ცოცხალი კალკულატორი, ფასების ცხრილი, ინტერაქტიული ქვიზი."
        }
    ]

    card_w = Inches(5.4); card_gap = Inches(0.4); start_x = Inches(1.0); card_y = Inches(2.35)
    for i, uc in enumerate(use_cases):
        cx = start_x + i * (card_w + card_gap)

        c = s7.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, cx, card_y, card_w, Inches(7.7))
        c.fill.solid()
        c.fill.fore_color.rgb = COLOR_CARD_RED if i == 0 else COLOR_BG_LIGHT
        c.line.color.rgb = COLOR_BORDER_RED if i == 0 else COLOR_BORDER
        c.line.width = Pt(1.5)

        tb = s7.shapes.add_textbox(cx + Inches(0.4), card_y + Inches(0.4), card_w - Inches(0.8), Inches(6.9))
        tf = tb.text_frame; tf.word_wrap = True; tf.margin_left = tf.margin_top = 0

        p = tf.paragraphs[0]; p.text = uc["tag"]
        p.font.name = FONT_ARIAL; p.font.size = Pt(11); p.font.bold = True
        p.font.color.rgb = COLOR_RED if i == 0 else COLOR_MUTED

        p = tf.add_paragraph(); p.text = uc["title"]
        p.font.name = FONT_ARIAL; p.font.size = Pt(22); p.font.bold = True; p.font.color.rgb = COLOR_DARK; p.space_before = Pt(4)

        p = tf.add_paragraph(); p.text = uc["geo"]
        p.font.name = FONT_ARIAL; p.font.size = Pt(13); p.font.bold = True; p.font.color.rgb = COLOR_MUTED; p.space_before = Pt(2)

        p = tf.add_paragraph(); p.text = f"\n{uc['desc']}"
        p.font.name = FONT_ARIAL; p.font.size = Pt(13); p.font.color.rgb = COLOR_TEXT_MUTED; p.space_before = Pt(6)

        p = tf.add_paragraph(); p.text = f"\n🎯 {uc['meta']}"
        p.font.name = FONT_ARIAL; p.font.size = Pt(12); p.font.bold = True; p.font.color.rgb = COLOR_DARK; p.space_before = Pt(8)

    # -------------------------------------------------------------------------
    # SLIDE 8: Web Search vs Internal Knowledge (Bilingual Terminology)
    # -------------------------------------------------------------------------
    s8 = created_slides[6]
    add_header(s8, "Web Search (ცოცხალი ვებ-ძიება) vs შიდა ცოდნა", "სიზუსტე • როგორ იგებს Claude უახლეს ფაქტებს?")

    col_w = Inches(8.35); col_h = Inches(5.4); col_y = Inches(2.35)

    # Left: Internal Knowledge
    b1 = s8.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(1.0), col_y, col_w, col_h)
    b1.fill.solid(); b1.fill.fore_color.rgb = COLOR_BG_LIGHT
    b1.line.color.rgb = COLOR_BORDER; b1.line.width = Pt(1.5)

    tb1 = s8.shapes.add_textbox(Inches(1.4), col_y + Inches(0.4), col_w - Inches(0.8), col_h - Inches(0.8))
    tf1 = tb1.text_frame; tf1.word_wrap = True; tf1.margin_left = tf1.margin_top = 0

    p = tf1.paragraphs[0]; p.text = "01 • ᲨᲘᲓᲐ ᲒᲐᲬᲕᲠᲗᲜᲘᲚᲘ ᲪᲝᲓᲜᲐ"
    p.font.name = FONT_ARIAL; p.font.size = Pt(12); p.font.bold = True; p.font.color.rgb = COLOR_MUTED

    p = tf1.add_paragraph(); p.text = "Internal Knowledge (მოდელის მეხსიერება)"
    p.font.name = FONT_ARIAL; p.font.size = Pt(20); p.font.bold = True; p.font.color.rgb = COLOR_DARK; p.space_before = Pt(4)

    int_points = [
        ("Knowledge Cutoff (ცოდნის ზღვარი)", "მოდელმა იცის მხოლოდ ის ფაქტები, რომლებზეც გაიწვრთნა. უახლესი მოვლენები მისთვის უცნობია."),
        ("როდის არის იდეალური?", "ლოგიკა, კოდირება, წერილების შედგენა, აბსტრაქტული აზროვნება და უნივერსალური პრინციპები."),
        ("რისკი", "უახლეს ფაქტებზე დასმულ კითხვაზე შესაძლოა დაუშვას ჰალუცინაცია.")
    ]
    for h, d in int_points:
        p = tf1.add_paragraph(); p.text = f"• {h}: {d}"
        p.font.name = FONT_ARIAL; p.font.size = Pt(13); p.font.color.rgb = COLOR_TEXT_MUTED; p.space_before = Pt(8)

    # Right: Real-Time Web Search
    b2 = s8.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(9.65), col_y, col_w, col_h)
    b2.fill.solid(); b2.fill.fore_color.rgb = COLOR_CARD_RED
    b2.line.color.rgb = COLOR_BORDER_RED; b2.line.width = Pt(1.5)

    tb2 = s8.shapes.add_textbox(Inches(10.05), col_y + Inches(0.4), col_w - Inches(0.8), col_h - Inches(0.8))
    tf2 = tb2.text_frame; tf2.word_wrap = True; tf2.margin_left = tf2.margin_top = 0

    p = tf2.paragraphs[0]; p.text = "02 • ᲠᲔᲐᲚᲣᲠ ᲓᲠᲝᲨᲘ ᲫᲘᲔᲑᲐ"
    p.font.name = FONT_ARIAL; p.font.size = Pt(12); p.font.bold = True; p.font.color.rgb = COLOR_RED

    p = tf2.add_paragraph(); p.text = "Web Search (ცოცხალი ვებ-ძიება)"
    p.font.name = FONT_ARIAL; p.font.size = Pt(20); p.font.bold = True; p.font.color.rgb = COLOR_DARK; p.space_before = Pt(4)

    web_points = [
        ("როდის ირთვება?", "როდესაც მომხმარებლის კითხვა ეხება მიმდინარე მოვლენებს, აქციების ფასებს, ახალ კანონებს ან უახლეს ამბებს."),
        ("Citations (წყაროების ციტირება)", "Claude პასუხში თითოეულ ფაქტს ურთავს ოფიციალურ, დაწკაპუნებად ვებ-ბმულს."),
        ("სიზუსტის გარანტია", "მკვეთრად ამცირებს ჰალუცინაციის რისკს და საშუალებას გაძლევთ, ინფორმაცია 1 წამში თავად გადაამოწმოთ.")
    ]
    for h, d in web_points:
        p = tf2.add_paragraph(); p.text = f"⭐ {h}: {d}"
        p.font.name = FONT_ARIAL; p.font.size = Pt(13); p.font.color.rgb = COLOR_DARK; p.space_before = Pt(8)

    # Bottom comparison summary
    bot_y = Inches(8.0)
    bb8 = s8.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(1.0), bot_y, Inches(17.0), Inches(2.0))
    bb8.fill.solid(); bb8.fill.fore_color.rgb = COLOR_WHITE
    bb8.line.color.rgb = COLOR_BORDER; bb8.line.width = Pt(1.2)

    tbb8 = s8.shapes.add_textbox(Inches(1.4), bot_y + Inches(0.3), Inches(16.2), Inches(1.4))
    tfb8 = tbb8.text_frame; tfb8.word_wrap = True; tfb8.margin_left = tfb8.margin_top = 0
    p = tfb8.paragraphs[0]; p.text = "💡 მთავარი არსი (მარტივი ანალოგია):"
    p.font.name = FONT_ARIAL; p.font.size = Pt(14); p.font.bold = True; p.font.color.rgb = COLOR_RED
    p = tfb8.add_paragraph()
    p.text = "შიდა ცოდნა არის ადამიანის საკუთარი მეხსიერება, ხოლო Web Search — შესაძლებლობა, რომ უცნობი ან ახალი საკითხის გასარკვევად წამიერად ჩაიხედოს უახლეს ენციკლოპედიაში. ორივეს კომბინაცია იძლევა უმაღლეს სიზუსტეს."
    p.font.name = FONT_ARIAL; p.font.size = Pt(13); p.font.color.rgb = COLOR_DARK; p.space_before = Pt(4)

    # -------------------------------------------------------------------------
    # SLIDE 9: 5 Core Working Scenarios (Syllabus Roadmap, NO Leaking to Lecture 5!)
    # -------------------------------------------------------------------------
    s9 = created_slides[7]
    add_header(s9, "5 ტიპური სამუშაო ამოცანა Claude-ში", "სამუშაო რუკა • Claude I-ის პრაქტიკული შესაძლებლობები")

    tasks_5 = [
        ("01", "PDF დოკუმენტის შეჯამება და ანალიტიკა", "დიდი მოცულობის ხელშეკრულებების, კვლევებისა და წიგნების წამიერი დამუშავება, ძირითადი თეზისების ამოღება და იურიდიული რისკების გამოვლენა."),
        ("02", "მონაცემთა ცხრილის წაკითხვა და ვიზუალიზაცია", "Excel/CSV ფაილების იმპორტი, ფინანსური მონაცემების გაფილტვრა, ძირითადი ტენდენციების აღმოჩენა და სტრუქტურული ცხრილების გენერირება."),
        ("03", "სურათისა და ნახაზის აღწერა (Vision & OCR)", "ინფოგრაფიკის, ტექნიკური ნახაზების, ჩეკებისა და ხელნაწერი დოკუმენტების ფოტოების გაციფრულება და შინაარსობრივი ანალიზი."),
        ("04", "ვებ-ძიება დამოწმებული წყაროებით (Citations)", "ინტერნეტში ცოცხალი ძიების გააქტიურება მიმდინარე მოვლენებზე, ოფიციალური წყაროების ციტირებითა და ჰალუცინაციების მინიმიზაციით."),
        ("05", "ინტერაქტიული Artifact-ის შექმნა", "ცოცხალი ინტერაქტიული აპლიკაციის (კალკულატორი), ბიზნეს-პროცესის დიაგრამის (Mermaid) ან სრულყოფილი წესდების შექმნა ცალკე სამუშაო პანელში.")
    ]

    card_h = Inches(1.35); card_gap = Inches(0.2); start_y = Inches(2.35)
    for i, (num, tname, tdesc) in enumerate(tasks_5):
        cy = start_y + i * (card_h + card_gap)

        c = s9.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(1.0), cy, Inches(17.0), card_h)
        c.fill.solid()
        c.fill.fore_color.rgb = COLOR_CARD_RED if i == 0 else COLOR_BG_LIGHT
        c.line.color.rgb = COLOR_BORDER_RED if i == 0 else COLOR_BORDER
        c.line.width = Pt(1.5)

        # Number pill on left
        pill = s9.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(1.3), cy + Inches(0.3), Inches(1.1), Inches(0.75))
        pill.fill.solid()
        pill.fill.fore_color.rgb = COLOR_RED if i == 0 else COLOR_WHITE
        pill.line.color.rgb = COLOR_BORDER_RED if i == 0 else COLOR_BORDER
        pill.line.width = Pt(1.2)

        tb_p = s9.shapes.add_textbox(Inches(1.3), cy + Inches(0.3), Inches(1.1), Inches(0.75))
        tf_p = tb_p.text_frame; tf_p.word_wrap = False; tf_p.margin_left = tf_p.margin_top = 0
        p = tf_p.paragraphs[0]; p.text = num; p.alignment = PP_ALIGN.CENTER
        p.font.name = FONT_ARIAL; p.font.size = Pt(18); p.font.bold = True
        p.font.color.rgb = COLOR_WHITE if i == 0 else COLOR_DARK

        # Text on right
        tb_txt = s9.shapes.add_textbox(Inches(2.7), cy + Inches(0.2), Inches(15.0), Inches(0.95))
        tf_t = tb_txt.text_frame; tf_t.word_wrap = True; tf_t.margin_left = tf_t.margin_top = 0

        p = tf_t.paragraphs[0]; p.text = tname
        p.font.name = FONT_ARIAL; p.font.size = Pt(16); p.font.bold = True; p.font.color.rgb = COLOR_DARK

        p = tf_t.add_paragraph(); p.text = tdesc
        p.font.name = FONT_ARIAL; p.font.size = Pt(12); p.font.color.rgb = COLOR_TEXT_MUTED; p.space_before = Pt(3)

    # -------------------------------------------------------------------------
    # SAVE PPTX AND DISTRIBUTE TO TARGET DIRECTORIES
    # -------------------------------------------------------------------------
    filename = "4 - Claude სრულად I - ინტერფეისი, ფაილები და Artifacts.pptx"
    out_master = os.path.join(r'd:\ხელოვნური ინტელექტის ინსტრუმენტები', filename)
    prs.save(out_master)
    print("Master PPTX for Lecture 4 generated successfully!")

    target_dirs = [
        r'G:\My Drive\შესავალი ხელოვნურ ინტელექტში',
        r'G:\My Drive\Skillwill\შესავალი ხელოვნურ ინტელექტში',
        r'D:\შესავალი ხელოვნურ ინტელექტში',
        r'd:\ხელოვნური ინტელექტის ინსტრუმენტები\შესავალი ხელოვნურ ინტელექტში'
    ]

    for td in target_dirs:
        try:
            if not os.path.exists(td):
                os.makedirs(td, exist_ok=True)
            dst_file = os.path.join(td, filename)
            shutil.copy2(out_master, dst_file)
            print(f"Copied to Drive target: {dst_file}")
        except Exception as e:
            print(f"Error copying to {td}: {e}")

if __name__ == '__main__':
    build_lecture_4()
