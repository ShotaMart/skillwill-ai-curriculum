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
    screenshot_dir = os.path.join(asset_dir, 'screenshots')

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
    COLOR_TERRACOTTA = RGBColor(0xD9, 0x77, 0x06)

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
            p2.text = "Anthropic-ის წამყვანი AI ასისტენტის შესაძლებლობები და სამუშაო გარემო"
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
    # SLIDE 2: Who is Anthropic and what makes Claude unique?
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
        ("დაფუძნება", "2021 წელს OpenAI-ის ყოფილი წამყვანი მკვლევრების (Dario & Daniela Amodei) მიერ."),
        ("მისია & პრინციპი", "Constitutional AI — ეთიკური, პროგნოზირებადი და ადამიანზე ორიენტირებული უსაფრთხოების სტანდარტები."),
        ("გლობალური პარტნიორები", "Amazon (AWS) და Google — მილიარდობით დოლარის ინვესტიცია და ინფრასტრუქტურული მხარდაჭერა."),
        ("ფოკუსი", "არა უბრალოდ ჩატბოტი, არამედ სიღრმისეული ანალიტიკის, კვლევებისა და კოდინგის ინსტრუმენტი.")
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
        ("Natural & Nuanced Language (ბუნებრივი ენა)", "ყველაზე ნაკლებად რობოტული, აკადემიური და ზრდილობიანი ტონი. ქართულ ენაზე ავლენს უმაღლეს სტილისტურ გამართულობას."),
        ("Massive Context Window (200K ტოკენი)", "ერთ საუბარში თავისუფლად იტევს 500+ გვერდიან წიგნს ან მთლიან კორპორაციულ კვლევას."),
        ("Logic & Coding Leader (ლოგიკისა და კოდის ლიდერი)", "მსოფლიო ბენჩმარკებში (SWE-bench, HumanEval) Claude 3.5 Sonnet უპირობო პირველ ადგილს იკავებს.")
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
    p.text = "თუ ChatGPT არის „სწრაფი ენციკლოპედისტი“, Claude არის „ანალიტიკოსი მკვლევარი“. ის არ ჩქარობს ზედაპირული პასუხის გაცემას, ღრმად ჩადის კონტექსტში და კითხულობს უზარმაზარ დოკუმენტებს ისე, რომ დეტალები არ ავიწყდება."
    p.font.name = FONT_ARIAL; p.font.size = Pt(13); p.font.color.rgb = COLOR_DARK; p.space_before = Pt(4)

    # -------------------------------------------------------------------------
    # SLIDE 3: Claude 3.5 Family (Haiku, Sonnet, Opus)
    # -------------------------------------------------------------------------
    s3 = created_slides[1]
    add_header(s3, "Claude 3.5 მოდელების ოჯახი (ვის რა როლი აქვს?)", "მოდელები • სისწრაფის, ფასისა და ინტელექტის ბალანსი")

    models_claude = [
        {
            "tag": "01 • ᲡᲘᲡᲬᲠᲐᲤᲔ & ᲔᲙᲝᲜᲝᲛᲘᲐ",
            "name": "Claude 3.5 Haiku",
            "role": "მსუბუქი და ულტრა-სწრაფი ასისტენტი",
            "desc": "შექმნილია წამიერი რეაგირებისთვის. იდეალურია ყოველდღიური რუტინული ამოცანებისთვის, მომხმარებელთა ჩატბოტებისთვის და მოკლე ტექსტების დასამუშავებლად.",
            "pros": "უმაღლესი სისწრაფე, მინიმალური ფასი",
            "color": COLOR_BG_LIGHT, "bcolor": COLOR_BORDER, "tcolor": COLOR_MUTED
        },
        {
            "tag": "02 • ᲝᲥᲠᲝᲡ ᲡᲢᲐᲜᲓᲐᲠᲢᲘ (ᲤᲚᲐᲒᲛᲐᲜᲘ)",
            "name": "Claude 3.5 Sonnet",
            "role": "ინდუსტრიის წამყვანი ინტელექტი",
            "desc": "მთავარი სამუშაო მოდელი. აერთიანებს მაქსიმალურ ანალიტიკურ სიღრმეს, შეუდარებელ კოდინგს, Artifacts-ის გენერირებასა და სწრაფ პასუხს.",
            "pros": "ოპტიმალური არჩევანი 95% პროფესიული ამოცანისთვის",
            "color": COLOR_CARD_RED, "bcolor": COLOR_BORDER_RED, "tcolor": COLOR_RED
        },
        {
            "tag": "03 • ᲛᲐᲥᲡᲘᲛᲐᲚᲣᲠᲘ ᲡᲘᲫᲚᲘᲔᲠᲔ",
            "name": "Claude 3 Opus",
            "role": "ურთულესი კვლევების მოდელი",
            "desc": "ყველაზე მძიმე მათემატიკური და სამეცნიერო ამოცანებისთვის. ამუშავებს უზარმაზარ აბსტრაქტულ მონაცემებს, სადაც სიჩქარეზე მნიშვნელოვანი უზუსტესი ლოგიკაა.",
            "pros": "მაქსიმალური მსჯელობის სიღრმე",
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
    # SLIDE 4: Workspace Interface Anatomy (REAL UI SCREENSHOT!)
    # -------------------------------------------------------------------------
    s4 = created_slides[2]
    add_header(s4, "Claude.ai სამუშაო ინტერფეისის ანატომია", "ინტერფეისი • რეალური სამუშაო გარემოს გაცნობა")

    # Left: Real screenshot card (Inches 10.0 x Inches 7.7)
    scr_card_w = Inches(10.2); scr_card_h = Inches(7.7); scr_x = Inches(1.0); scr_y = Inches(2.35)
    sc_box = s4.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, scr_x, scr_y, scr_card_w, scr_card_h)
    sc_box.fill.solid(); sc_box.fill.fore_color.rgb = COLOR_WHITE
    sc_box.line.color.rgb = COLOR_BORDER; sc_box.line.width = Pt(1.5)

    real_ui_img = os.path.join(screenshot_dir, 'claude_ui.png')
    if os.path.exists(real_ui_img):
        # Insert real UI screenshot inside box with padding
        s4.shapes.add_picture(real_ui_img, scr_x + Inches(0.2), scr_y + Inches(0.8), width=scr_card_w - Inches(0.4), height=Inches(6.5))

    tb_ui_label = s4.shapes.add_textbox(scr_x + Inches(0.4), scr_y + Inches(0.2), scr_card_w - Inches(0.8), Inches(0.5))
    tf_ui_l = tb_ui_label.text_frame; tf_ui_l.word_wrap = True; tf_ui_l.margin_left = tf_ui_l.margin_top = 0
    p = tf_ui_l.paragraphs[0]; p.text = "🖥️ CLAUDE.AI — ოფიციალური სამუშაო გარემო"
    p.font.name = FONT_ARIAL; p.font.size = Pt(13); p.font.bold = True; p.font.color.rgb = COLOR_DARK

    # Right: UI Breakdown Callouts (Inches 6.5 x 7.7)
    call_x = Inches(11.5); call_w = Inches(6.5); call_y = Inches(2.35)
    call_box = s4.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, call_x, call_y, call_w, scr_card_h)
    call_box.fill.solid(); call_box.fill.fore_color.rgb = COLOR_BG_LIGHT
    call_box.line.color.rgb = COLOR_BORDER; call_box.line.width = Pt(1.5)

    tb_call = s4.shapes.add_textbox(call_x + Inches(0.4), call_y + Inches(0.4), call_w - Inches(0.8), scr_card_h - Inches(0.8))
    tfc = tb_call.text_frame; tfc.word_wrap = True; tfc.margin_left = tfc.margin_top = 0

    p = tfc.paragraphs[0]; p.text = "ᲘᲜᲢᲔᲠᲤᲔᲘᲡᲘᲡ 4 ᲫᲘᲠᲘᲗᲐᲓᲘ ᲖᲝᲜᲐ"
    p.font.name = FONT_ARIAL; p.font.size = Pt(13); p.font.bold = True; p.font.color.rgb = COLOR_RED

    ui_zones = [
        ("01 • Prompt Bar (შესაყვანი ველი)", "მთავარი არეალი შეკითხვის დასაწერად, მოდელის ასარჩევად (Sonnet/Haiku) და კონტექსტის სამართავად."),
        ("02 • Paperclip (ფაილების მიმაგრება)", "დოკუმენტების (PDF, Word, Excel) და სურათების მარტივი ატვირთვა პირდაპირ ჩატში."),
        ("03 • Sidebar (ნავიგაცია და არქივი)", "მარცხენა პანელი: წინა დიალოგების ისტორია, სწრაფი ძიება (Ctrl+K) და ახალი ჩატის დაწყება."),
        ("04 • Chat Canvas (დიალოგის ველი)", "სუფთა ინტერფეისი, სადაც პასუხები ფორმატირდება სათაურებით, ცხრილებითა და კოდის ბლოკებით.")
    ]
    for h, d in ui_zones:
        p = tfc.add_paragraph(); p.text = f"\n📌 {h}"
        p.font.name = FONT_ARIAL; p.font.size = Pt(13); p.font.bold = True; p.font.color.rgb = COLOR_DARK; p.space_before = Pt(8)
        p = tfc.add_paragraph(); p.text = d
        p.font.name = FONT_ARIAL; p.font.size = Pt(12); p.font.color.rgb = COLOR_TEXT_MUTED; p.space_before = Pt(2)

    # -------------------------------------------------------------------------
    # SLIDE 5: Multimodal File Handling (PDF, Excel, Images)
    # -------------------------------------------------------------------------
    s5 = created_slides[3]
    add_header(s5, "ფაილების დამუშავება (Multimodal: PDF, Excel და სურათები)", "მონაცემები • რა დოკუმენტების „წაკითხვა“ შეუძლია Claude-ს?")

    file_types = [
        {
            "tag": "01 • ᲓᲝᲙᲣᲛᲔᲜᲢᲔᲑᲘ",
            "title": "PDF, Word & Text",
            "sub": "ტექსტური ანალიტიკა და შეჯამება",
            "points": [
                "100+ გვერდიანი ხელშეკრულებების, კვლევებისა და წიგნების წამიერი დამუშავება.",
                "ძირითადი თეზისების, იურიდიული რისკებისა და საკვანძო პირობების ამოღება.",
                "კითხვა-პასუხი დოკუმენტის შინაარსზე (ზუსტი მითითებებით).",
                "დოკუმენტის სტილისტური გადამუშავება და ენობრივი თარგმანი."
            ]
        },
        {
            "tag": "02 • ᲛᲝᲜᲐᲪᲔᲛᲔᲑᲘ",
            "title": "Excel & CSV ცხრილები",
            "sub": "სტრუქტურირებული ანალიზი",
            "points": [
                "ფინანსური რეპორტების, გაყიდვებისა და მომხმარებელთა სიების წაკითხვა.",
                "მონაცემთა გაფილტვრა, დაჯგუფება და ტენდენციების აღმოჩენა.",
                "მზა ფორმულების (Excel / SQL) და ანალიტიკური დასკვნების მომზადება.",
                "ახალი, გასუფთავებული ცხრილების შექმნა."
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
                "ტექნიკური სქემების წაკითხვა."
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
    # SLIDE 6: What are Artifacts? (REAL SPLIT-SCREEN SCREENSHOT!)
    # -------------------------------------------------------------------------
    s6 = created_slides[4]
    add_header(s6, "რა არის Artifacts? — ინტერაქტიული სამუშაო სივრცე", "რევოლუციური ფუნქცია • გაყოფილი ეკრანის პარადიგმა")

    top_y = Inches(2.35)
    tb_intro6 = s6.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(1.0), top_y, Inches(17.0), Inches(1.2))
    tb_intro6.fill.solid(); tb_intro6.fill.fore_color.rgb = COLOR_CARD_RED
    tb_intro6.line.color.rgb = COLOR_BORDER_RED; tb_intro6.line.width = Pt(1.5)

    tb6_t = s6.shapes.add_textbox(Inches(1.4), top_y + Inches(0.18), Inches(16.2), Inches(0.85))
    tf6 = tb6_t.text_frame; tf6.word_wrap = True; tf6.margin_left = tf6.margin_top = 0
    p = tf6.paragraphs[0]; p.text = "🚀 პარადიგმის ცვლილება: ჩატბოტიდან სრულფასოვან სამუშაო სივრცემდე"
    p.font.name = FONT_ARIAL; p.font.size = Pt(14); p.font.bold = True; p.font.color.rgb = COLOR_RED
    p = tf6.add_paragraph(); p.text = "ტრადიციულ ჩატში გრძელი კოდი ან დოკუმენტი დიალოგის ნაკადში იკარგებოდა. Artifacts ეკრანს ორად ყოფს: მარცხნივ ჩატია, მარჯვნივ კი ცალკე ინტერაქტიული პანელი!"
    p.font.name = FONT_ARIAL; p.font.size = Pt(13); p.font.color.rgb = COLOR_DARK; p.space_before = Pt(2)

    # Real Screenshot on Left (Inches 10.2 x 6.3)
    art_box_x = Inches(1.0); art_box_y = Inches(3.75); art_box_w = Inches(10.2); art_box_h = Inches(6.3)
    art_box = s6.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, art_box_x, art_box_y, art_box_w, art_box_h)
    art_box.fill.solid(); art_box.fill.fore_color.rgb = COLOR_WHITE
    art_box.line.color.rgb = COLOR_BORDER; art_box.line.width = Pt(1.5)

    art_real_img = os.path.join(screenshot_dir, 'artifacts_demo.png')
    if os.path.exists(art_real_img):
        s6.shapes.add_picture(art_real_img, art_box_x + Inches(0.2), art_box_y + Inches(0.2), width=art_box_w - Inches(0.4), height=art_box_h - Inches(0.4))

    # Right: 3 Key Artifact Features (Inches 6.5 x 6.3)
    feat_x = Inches(11.5); feat_w = Inches(6.5); feat_y = Inches(3.75)
    feat_box = s6.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, feat_x, feat_y, feat_w, art_box_h)
    feat_box.fill.solid(); feat_box.fill.fore_color.rgb = COLOR_BG_LIGHT
    feat_box.line.color.rgb = COLOR_BORDER; feat_box.line.width = Pt(1.5)

    tb_feat = s6.shapes.add_textbox(feat_x + Inches(0.4), feat_y + Inches(0.4), feat_w - Inches(0.8), art_box_h - Inches(0.8))
    tff = tb_feat.text_frame; tff.word_wrap = True; tff.margin_left = tff.margin_top = 0

    p = tff.paragraphs[0]; p.text = "ARTIFACTS-ᲘᲡ 3 ᲣᲞᲘᲠᲐᲢᲔᲡᲝᲑᲐ"
    p.font.name = FONT_ARIAL; p.font.size = Pt(12); p.font.bold = True; p.font.color.rgb = COLOR_RED

    art_advantages = [
        ("Live Preview (ცოცხალი შედეგი)", "HTML/JS კოდი, დიაგრამები და ვებ-კომპონენტები პირდაპირ ეკრანზე მუშაობს და იტესტება რეალურ დროში."),
        ("Independent Object (დამოუკიდებელი ობიექტი)", "არტეფაქტი არსებობს როგორც დამოუკიდებელი დოკუმენტი — მისი რედაქტირება, კოპირება და ექსპორტი ხდება 1 კლიკით."),
        ("Version History (ვერსიების ისტორია)", "Claude-ს ყოველი დამატებითი თხოვნა არტეფაქტს აახლებს ახალ ვერსიად, რაც ძველ ვერსიაზე დაბრუნების საშუალებას გაძლევთ.")
    ]
    for h, d in art_advantages:
        p = tff.add_paragraph(); p.text = f"\n✅ {h}"
        p.font.name = FONT_ARIAL; p.font.size = Pt(13); p.font.bold = True; p.font.color.rgb = COLOR_DARK; p.space_before = Pt(8)
        p = tff.add_paragraph(); p.text = d
        p.font.name = FONT_ARIAL; p.font.size = Pt(12); p.font.color.rgb = COLOR_TEXT_MUTED; p.space_before = Pt(2)

    # -------------------------------------------------------------------------
    # SLIDE 7: Artifacts Use Cases (Diagrams, Docs, Prototypes)
    # -------------------------------------------------------------------------
    s7 = created_slides[5]
    add_header(s7, "Artifacts-ის რეალური შესაძლებლობები (Use Cases)", "პრაქტიკა • რა იქმნება არტეფაქტის ფანჯარაში?")

    use_cases = [
        {
            "tag": "01 • ᲕᲘᲖᲣᲐᲚᲘᲖᲐᲪᲘᲐ",
            "title": "Interactive Diagrams",
            "geo": "(ინტერაქტიული დიაგრამები)",
            "desc": "ბიზნეს-პროცესების ნაკადები, ორგანიზაციული სტრუქტურები და არქიტექტურული სქემები (Mermaid, SVG).",
            "meta": "მაგალითი: გაყიდვების პროცესის სქემა, მომხმარებლის გზა (User Journey)."
        },
        {
            "tag": "02 • ᲓᲝᲙᲣᲛᲔᲜᲢᲐᲪᲘᲐ",
            "title": "Ready Documents",
            "geo": "(სრულყოფილი დოკუმენტები)",
            "desc": "სრული წესდებები, ინსტრუქციები, სტატიები და ხელშეკრულებები, რომლებიც მზადაა ჩამოსატვირთად და გასაზიარებლად.",
            "meta": "მაგალითი: კომპანიის შიდა რეგლამენტი, პროექტის ტექნიკური დავალება (PRD)."
        },
        {
            "tag": "03 • ᲞᲠᲝᲢᲝᲢᲘᲞᲔᲑᲘ",
            "title": "Live Code & Apps",
            "geo": "(ცოცხალი აპლიკაციები & კოდი)",
            "desc": "ინტერაქტიული კალკულატორები, ვიზუალური დაშბორდები, მინი-თამაშები და ვებსაიტის ინტერფეისები React/HTML/JS-ში.",
            "meta": "მაგალითი: იპოთეკური სესხის ცოცხალი კალკულატორი, ფასების ცხრილი."
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
        ("როდის არის იდეალური?", "ლოგიკა, კოდირება, წერილების შედგენა, აბსტრაქტული აზროვნება და უნივერსალური ფაქტები."),
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
    # SLIDE 9: Ecosystem & Access (Claude.ai, Workspace, Enterprise)
    # -------------------------------------------------------------------------
    s9 = created_slides[7]
    add_header(s9, "Claude-ის ეკოსისტემა: პლატფორმები და ხელმისაწვდომობა", "პროდუქტები • სად და როგორ ვიყენებთ Claude-ს")

    eco_columns = [
        {
            "cat": "01 • CLAUDE.AI (ᲡᲐᲛᲝᲛᲮᲛᲐᲠᲔᲑᲚᲝ)",
            "title": "Desktop & Mobile",
            "sub": "ყოველდღიური სამუშაო გარემო",
            "tools": [
                ("claude.png", "Claude.ai (Web App)", "ბრაუზერში მუშაობა (Free და Pro პაკეტები)"),
                ("claude.png", "Claude Desktop (Mac/Win)", "დამოუკიდებელი აპლიკაცია თქვენს კომპიუტერზე")
            ]
        },
        {
            "cat": "02 • WORKSPACE & STORAGE",
            "title": "Cloud Workspaces",
            "sub": "მონაცემების ინტეგრაცია",
            "tools": [
                ("claude.png", "Artifacts Environment", "ინტერაქტიული დოკუმენტებისა და კოდის მართვა"),
                ("google_cloud.png", "Google Drive Integration", "ფაილების პირდაპირი იმპორტი Google Drive-იდან")
            ]
        },
        {
            "cat": "03 • ENTERPRISE & SECURITY",
            "title": "Developer & Team",
            "sub": "კორპორაციული უსაფრთხოება",
            "tools": [
                ("claude.png", "Anthropic API", "მოდელების ინტეგრაცია საკუთარ პროგრამებში"),
                ("aws.png", "AWS Bedrock / Data Boundary", "კორპორაციული მონაცემების სრული დაცვა და კონფიდენციალურობა")
            ]
        }
    ]

    col_w = Inches(5.4); col_gap = Inches(0.4); start_x = Inches(1.0)
    for i, col in enumerate(eco_columns):
        cx = start_x + i * (col_w + col_gap)

        cb = s9.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, cx, Inches(2.35), col_w, Inches(7.7))
        cb.fill.solid()
        cb.fill.fore_color.rgb = COLOR_CARD_RED if i == 0 else COLOR_BG_LIGHT
        cb.line.color.rgb = COLOR_BORDER_RED if i == 0 else COLOR_BORDER
        cb.line.width = Pt(1.5)

        tb_head = s9.shapes.add_textbox(cx + Inches(0.35), Inches(2.55), col_w - Inches(0.7), Inches(1.6))
        tf = tb_head.text_frame; tf.word_wrap = True; tf.margin_left = tf.margin_top = 0

        p = tf.paragraphs[0]; p.text = col["cat"]
        p.font.name = FONT_ARIAL; p.font.size = Pt(11); p.font.bold = True
        p.font.color.rgb = COLOR_RED if i == 0 else COLOR_MUTED

        p = tf.add_paragraph(); p.text = col["title"]
        p.font.name = FONT_ARIAL; p.font.size = Pt(20); p.font.bold = True; p.font.color.rgb = COLOR_DARK; p.space_before = Pt(2)

        p = tf.add_paragraph(); p.text = col["sub"]
        p.font.name = FONT_ARIAL; p.font.size = Pt(12); p.font.color.rgb = COLOR_TEXT_MUTED; p.space_before = Pt(3)

        card_h = Inches(1.9)
        spacing = Inches(2.2)
        base_y = Inches(4.5)

        for j, (l_file, t_name, t_desc) in enumerate(col["tools"]):
            item_y = base_y + j * spacing

            t_card = s9.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, cx + Inches(0.3), item_y, col_w - Inches(0.6), card_h)
            t_card.fill.solid(); t_card.fill.fore_color.rgb = COLOR_WHITE
            t_card.line.color.rgb = COLOR_BORDER_RED if i == 0 else COLOR_BORDER
            t_card.line.width = Pt(1.2)

            # Square Logo 1.1" x 1.1"
            lp = os.path.join(logo_dir, l_file)
            logo_sz = Inches(1.1)
            logo_y = item_y + (card_h - logo_sz) / 2
            if os.path.exists(lp):
                s9.shapes.add_picture(lp, cx + Inches(0.5), logo_y, width=logo_sz, height=logo_sz)

            tb_item = s9.shapes.add_textbox(cx + Inches(1.8), item_y + Inches(0.3), col_w - Inches(2.2), card_h - Inches(0.5))
            tf = tb_item.text_frame; tf.word_wrap = True; tf.margin_left = tf.margin_top = 0

            p = tf.paragraphs[0]; p.text = t_name
            p.font.name = FONT_ARIAL; p.font.size = Pt(14); p.font.bold = True; p.font.color.rgb = COLOR_DARK

            p = tf.add_paragraph(); p.text = t_desc
            p.font.name = FONT_ARIAL; p.font.size = Pt(11); p.font.color.rgb = COLOR_TEXT_MUTED; p.space_before = Pt(3)

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
