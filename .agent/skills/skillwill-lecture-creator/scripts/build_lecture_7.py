import os
import sys
import shutil
import pptx
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN
from pptx.enum.shapes import MSO_SHAPE

def build_lecture_7():
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
            p.text = "7 - Gemini სრულად I: Gemini, Workspace, AI Studio და პარამეტრები"
            p.alignment = PP_ALIGN.CENTER
            p.font.name = FONT_ARIAL
            p.font.size = Pt(34)
            p.font.bold = True
            p.font.color.rgb = COLOR_RED

            p2 = tf.add_paragraph()
            p2.text = "Google-ის AI ეკოსისტემა, მულტიმოდალური მოდელები და LLM-ის მართვის პარამეტრები"
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
    def add_header(slide, title, category="ᲛᲝᲓᲣᲚᲘ 5 • GOOGLE GEMINI • 7"):
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

    box_w = Inches(8.35); box_h = Inches(5.4); box_y = Inches(2.35)
    card_w = Inches(5.4); card_h = Inches(5.4); card_y = Inches(2.35)
    xs_3 = [Inches(1.0), Inches(6.8), Inches(12.6)]
    step_w = Inches(3.95); step_h = Inches(5.4); step_y = Inches(2.35)
    xs_4 = [Inches(1.0), Inches(5.35), Inches(9.7), Inches(14.05)]
    bot_y = Inches(8.0)

    # -------------------------------------------------------------------------
    # SLIDE 2: Google's AI Vision & Gemini Ecosystem Map (2026)
    # -------------------------------------------------------------------------
    s2 = created_slides[0]
    add_header(s2, "Google-ის AI ხედვა და Gemini ეკოსისტემის რუკა (2026)", "ეკოსისტემა • სამომხმარებლო, საოფისე და დეველოპერული გარემო")

    cards_s2 = [
        ("📱 CONSUMER", "Gemini App (სამომხმარებლო)", "gemini.google.com", [
            ("ყოველდღიური ასისტენტი", "ვებ და მობილური აპლიკაცია კითხვა-პასუხისთვის, იდეების გენერაციისა და ძიებისთვის."),
            ("Gemini Live", "რეალურ დროში ხმოვანი კომუნიკაცია და ვიზუალური აღქმა კამერით."),
            ("Google Extensions", "პირდაპირი კავშირი YouTube, Google Maps, Flights და Hotels სერვისებთან.")
        ]),
        ("💼 WORKSPACE", "Google Workspace ინტეგრაცია", "საოფისე პროდუქტიულობა", [
            ("ჩაშენებული ასისტენტი", "Gemini ინტეგრირებულია Google Docs, Sheets, Slides, Drive და Gmail აპლიკაციებში."),
            ("დოკუმენტების სინთეზი", "Drive-ის ფაილების პირდაპირი ანალიზი და საქმიანი მიმოწერის ავტომატური დრაფტები."),
            ("კორპორაციული უსაფრთხოება", "მონაცემები დაცულია Google Workspace-ის მკაცრი უსაფრთხოების სტანდარტებით.")
        ]),
        ("💻 DEVELOPER", "Google AI Studio", "დეველოპერული პლატფორმა", [
            ("პროტოტიპირების IDE", "სწრაფი ვებ-ინტერფეისი მოდელების ტესტირებისთვის, პრომპტინგისა და შედარებისთვის."),
            ("პარამეტრების სრული კონტროლი", "Temperature, Top-P, Top-K, Max Tokens და Thinking Budget-ის მართვა."),
            ("API და ექსპორტი", "ერთი დაწკაპუნებით კოდის გენერაცია (Python, JavaScript, cURL) და Vertex AI-ზე მიგრაცია.")
        ])
    ]

    for i, (badge, title, subtitle, items) in enumerate(cards_s2):
        bg_col = COLOR_CARD_RED if i == 1 else COLOR_BG_LIGHT
        bor_col = COLOR_BORDER_RED if i == 1 else COLOR_BORDER
        c = s2.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, xs_3[i], card_y, card_w, card_h)
        c.fill.solid(); c.fill.fore_color.rgb = bg_col
        c.line.color.rgb = bor_col; c.line.width = Pt(1.5)

        tb = s2.shapes.add_textbox(xs_3[i] + Inches(0.35), card_y + Inches(0.35), card_w - Inches(0.7), card_h - Inches(0.7))
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
    bb = s2.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(1.0), bot_y, Inches(17.0), Inches(2.0))
    bb.fill.solid(); bb.fill.fore_color.rgb = COLOR_WHITE
    bb.line.color.rgb = COLOR_BORDER; bb.line.width = Pt(1.2)

    tbb = s2.shapes.add_textbox(Inches(1.4), bot_y + Inches(0.3), Inches(16.2), Inches(1.4))
    tfb = tbb.text_frame; tfb.word_wrap = True; tfb.margin_left = tfb.margin_top = 0
    p = tfb.paragraphs[0]; p.text = "💡 მთავარი არსი (მარტივი ანალოგია):"
    p.font.name = FONT_ARIAL; p.font.size = Pt(14); p.font.bold = True; p.font.color.rgb = COLOR_RED
    p = tfb.add_paragraph()
    p.text = "Gemini არის Google-ის ერთიანი ეკოსისტემის ინტელექტუალური შრე: როგორც Android აერთიანებს სმარტფონის აპლიკაციებს, ისე Gemini აკავშირებს Google-ის ყველა საოფისე დოკუმენტს, მეილსა და დეველოპერულ ხელსაწყოს ერთიან გონიერ სივრცეში."
    p.font.name = FONT_ARIAL; p.font.size = Pt(13); p.font.color.rgb = COLOR_MUTED; p.space_before = Pt(4)

    # -------------------------------------------------------------------------
    # SLIDE 3: Gemini Models Architecture & Generations (2026)
    # -------------------------------------------------------------------------
    s3 = created_slides[1]
    add_header(s3, "Gemini მოდელების არქიტექტურა და თაობები (2026)", "არქიტექტურა • Gemini 4 ოჯახი, მულტიმოდალობა და 2M+ კონტექსტი")

    models_s3 = [
        ("Gemini 4 Carbon", "Frontier Reasoning", [
            ("დანიშნულება", "Google-ის ყველაზე მძლავრი მოდელი რთული სამეცნიერო კვლევებისთვის, კიბერუსაფრთხოებისა და პროგრამირებისთვის."),
            ("Deep Investigation", "გამოიყენებს გაფართოებულ Thinking Budget-ს ურთულესი პრობლემების გადასაჭრელად.")
        ]),
        ("Gemini 4 Argon & Hydrogen", "სამუშაო ფლაგმანები", [
            ("Argon", "ოპტიმიზებული ფლაგმანი მაღალი სისწრაფითა და ანალიტიკური სიღრმით."),
            ("Hydrogen", "ავტონომიური აგენტური კოდირება, ინსტრუმენტების ორკესტრირება და სანდო დეველოპმენტი.")
        ]),
        ("Gemini 4 Neon", "მაღალი სიზუსტე", [
            ("მინიმალური ჰალუცინაციები", "ოპტიმიზებულია ფაქტობრივი სიზუსტისთვის (Low Fabrication Rate) და ამოცანების სწრაფი დასრულებისთვის."),
            ("საწარმოო გარემო", "იდეალურია იურიდიული და სამედიცინო დოკუმენტების ანალიზისთვის.")
        ]),
        ("Gemini 3.8 Flash", "სისწრაფე და მასშტაბი", [
            ("Workhorse მოდელი", "ულტრა-დაბალი დაყოვნება, ეკონომიური ტარიფი და მასიური გამტარუნარიანობა."),
            ("Massive Context (1M-2M+)", "ერთდროულად იტევს 1-საათიან ვიდეოს, 10-საათიან აუდიოს ან 700,000+ სიტყვას.")
        ])
    ]

    for i, (title, subtitle, items) in enumerate(models_s3):
        bg_col = COLOR_CARD_RED if i == 0 else COLOR_BG_LIGHT
        bor_col = COLOR_BORDER_RED if i == 0 else COLOR_BORDER
        c = s3.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, xs_4[i], step_y, step_w, step_h)
        c.fill.solid(); c.fill.fore_color.rgb = bg_col
        c.line.color.rgb = bor_col; c.line.width = Pt(1.5)

        tb = s3.shapes.add_textbox(xs_4[i] + Inches(0.25), step_y + Inches(0.35), step_w - Inches(0.5), step_h - Inches(0.7))
        tf = tb.text_frame; tf.word_wrap = True; tf.margin_left = tf.margin_top = 0

        p = tf.paragraphs[0]; p.text = f"ᲛᲝᲓᲔᲚᲘ {i+1}".upper()
        p.font.name = FONT_ARIAL; p.font.size = Pt(11); p.font.bold = True
        p.font.color.rgb = COLOR_RED if i == 0 else COLOR_MUTED

        p = tf.add_paragraph(); p.text = title
        p.font.name = FONT_ARIAL; p.font.size = Pt(16); p.font.bold = True; p.font.color.rgb = COLOR_DARK; p.space_before = Pt(4)

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
    p = tfb.paragraphs[0]; p.text = "💡 Native Multimodal ბუნება:"
    p.font.name = FONT_ARIAL; p.font.size = Pt(14); p.font.bold = True; p.font.color.rgb = COLOR_RED
    p = tfb.add_paragraph()
    p.text = "Gemini თავიდანვე შეიქმნა როგორც მულტიმოდალური მოდელი: ის ტექსტს, ვიდეოკადრებს, აუდიო ბგერებსა და სურათებს აღიქვამს ერთიან ენაზე, რაც საშუალებას აძლევს პირდაპირ „უყუროს“ 1-საათიან ლექციას ან ვიდეოჩანაწერს და წამებში უპასუხოს დეტალურ კითხვებს."
    p.font.name = FONT_ARIAL; p.font.size = Pt(13); p.font.color.rgb = COLOR_MUTED; p.space_before = Pt(4)

    # -------------------------------------------------------------------------
    # SLIDE 4: Google Workspace Integration (Docs, Gmail, Drive, Sheets)
    # -------------------------------------------------------------------------
    s4 = created_slides[2]
    add_header(s4, "Google Workspace ინტეგრაცია (Docs, Gmail, Drive და Sheets)", "პროდუქტიულობა • AI როგორც ჩაშენებული საოფისე ასისტენტი")

    # Left: Core Capabilities
    b1 = s4.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(1.0), box_y, box_w, box_h)
    b1.fill.solid(); b1.fill.fore_color.rgb = COLOR_BG_LIGHT
    b1.line.color.rgb = COLOR_BORDER; b1.line.width = Pt(1.5)

    tb1 = s4.shapes.add_textbox(Inches(1.4), box_y + Inches(0.4), box_w - Inches(0.8), box_h - Inches(0.8))
    tf1 = tb1.text_frame; tf1.word_wrap = True; tf1.margin_left = tf1.margin_top = 0
    p = tf1.paragraphs[0]; p.text = "ᲡᲐᲝᲤᲘᲡᲔ ᲤᲣᲜᲥᲪᲘᲝᲜᲐᲚᲘ"
    p.font.name = FONT_ARIAL; p.font.size = Pt(12); p.font.bold = True; p.font.color.rgb = COLOR_MUTED

    p = tf1.add_paragraph(); p.text = "აპლიკაციებში ჩაშენებული ასისტენტი"
    p.font.name = FONT_ARIAL; p.font.size = Pt(21); p.font.bold = True; p.font.color.rgb = COLOR_DARK; p.space_before = Pt(4)

    ws_points = [
        ("Google Docs & Drive", "დოკუმენტების გენერაცია, სტილის რედაქტირება და Drive-ში შენახული ათობით ფაილის შინაარსობრივი სინთეზი."),
        ("Gmail & კომუნიკაცია", "გრძელი მიმოწერის თრედების შეჯამება, პრიორიტეტული საკითხების გამოყოფა და პასუხის დრაფტების შექმნა."),
        ("Google Sheets & ანალიტიკა", "რთული ფორმულების ავტომატური გენერაცია, მონაცემთა დახარისხება და ცხრილების ვიზუალიზაცია."),
        ("Google Slides & პრეზენტაციები", "სლაიდების სტრუქტურის, ტექსტებისა და შესაბამისი ილუსტრაციების ავტომატური დაგენერირება.")
    ]
    for t, d in ws_points:
        p = tf1.add_paragraph(); p.text = f"• {t}: {d}"
        p.font.name = FONT_ARIAL; p.font.size = Pt(13); p.font.color.rgb = COLOR_TEXT_MUTED; p.space_before = Pt(8)

    # Right: Enterprise Security
    b2 = s4.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(9.65), box_y, box_w, box_h)
    b2.fill.solid(); b2.fill.fore_color.rgb = COLOR_CARD_RED
    b2.line.color.rgb = COLOR_BORDER_RED; b2.line.width = Pt(1.5)

    tb2 = s4.shapes.add_textbox(Inches(10.05), box_y + Inches(0.4), box_w - Inches(0.8), box_h - Inches(0.8))
    tf2 = tb2.text_frame; tf2.word_wrap = True; tf2.margin_left = tf2.margin_top = 0
    p = tf2.paragraphs[0]; p.text = "ᲙᲝᲠᲞᲝᲠᲐᲪᲘᲣᲚᲘ ᲣᲡᲐᲤᲠᲗᲮᲝᲔᲑᲐ"
    p.font.name = FONT_ARIAL; p.font.size = Pt(12); p.font.bold = True; p.font.color.rgb = COLOR_RED

    p = tf2.add_paragraph(); p.text = "Enterprise Data Boundary"
    p.font.name = FONT_ARIAL; p.font.size = Pt(21); p.font.bold = True; p.font.color.rgb = COLOR_DARK; p.space_before = Pt(4)

    sec_points = [
        ("მოდელის წვრთნის აკრძალვა", "თქვენი კომპანიის დოკუმენტები, ელფოსტა და მონაცემები არ გამოიყენება Google-ის AI მოდელების საწვრთნელად."),
        ("მონაცემთა იზოლაცია", "თითოეული ორგანიზაციის მონაცემები მკაცრად იზოლირებულია საკუთარ Workspace დომენში."),
        ("წვდომის უფლებების დაცვა (ACL)", "Gemini კითხულობს მხოლოდ იმ ფაილებს, რომლებზეც კონკრეტულ თანამშრომელს უკვე აქვს მინიჭებული უფლება."),
        ("საერთაშორისო სერტიფიკატები", "SOC 1/2/3, ISO/IEC 27001, HIPAA და GDPR შესაბამისობა.")
    ]
    for t, d in sec_points:
        p = tf2.add_paragraph(); p.text = f"⭐ {t}"
        p.font.name = FONT_ARIAL; p.font.size = Pt(13); p.font.bold = True; p.font.color.rgb = COLOR_DARK; p.space_before = Pt(8)
        p = tf2.add_paragraph(); p.text = d
        p.font.name = FONT_ARIAL; p.font.size = Pt(12); p.font.color.rgb = COLOR_TEXT_MUTED; p.space_before = Pt(2)

    # Bottom takeaway
    bb = s4.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(1.0), bot_y, Inches(17.0), Inches(2.0))
    bb.fill.solid(); bb.fill.fore_color.rgb = COLOR_WHITE
    bb.line.color.rgb = COLOR_BORDER; bb.line.width = Pt(1.2)

    tbb = s4.shapes.add_textbox(Inches(1.4), bot_y + Inches(0.3), Inches(16.2), Inches(1.4))
    tfb = tbb.text_frame; tfb.word_wrap = True; tfb.margin_left = tfb.margin_top = 0
    p = tfb.paragraphs[0]; p.text = "💡 ბიზნეს უპირატესობა:"
    p.font.name = FONT_ARIAL; p.font.size = Pt(14); p.font.bold = True; p.font.color.rgb = COLOR_RED
    p = tfb.add_paragraph()
    p.text = "Workspace-ის ინტეგრაცია გამორიცხავს ფაილების ხელით კოპირებას — Gemini პირდაპირ ხედავს თქვენს სამუშაო კონტექსტს და პასუხობს კომპანიის შიდა დოკუმენტებზე დაყრდნობით, უსაფრთხოების წესების სრული დაცვით."
    p.font.name = FONT_ARIAL; p.font.size = Pt(13); p.font.color.rgb = COLOR_MUTED; p.space_before = Pt(4)

    # -------------------------------------------------------------------------
    # SLIDE 5: Gemini Canvas and Deep Research
    # -------------------------------------------------------------------------
    s5 = created_slides[3]
    add_header(s5, "Gemini Canvas და Deep Research", "ინსტრუმენტები • ინტერაქტიული სამუშაო მაგიდა და ავტონომიური კვლევა")

    # Left: Gemini Canvas
    b1 = s5.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(1.0), box_y, box_w, box_h)
    b1.fill.solid(); b1.fill.fore_color.rgb = COLOR_BG_LIGHT
    b1.line.color.rgb = COLOR_BORDER; b1.line.width = Pt(1.5)

    tb1 = s5.shapes.add_textbox(Inches(1.4), box_y + Inches(0.4), box_w - Inches(0.8), box_h - Inches(0.8))
    tf1 = tb1.text_frame; tf1.word_wrap = True; tf1.margin_left = tf1.margin_top = 0
    p = tf1.paragraphs[0]; p.text = "ᲘᲜᲢᲔᲠᲐᲥᲢᲘᲣᲚᲘ ᲡᲐᲛᲣᲨᲐᲝ ᲓᲐᲤᲐ"
    p.font.name = FONT_ARIAL; p.font.size = Pt(12); p.font.bold = True; p.font.color.rgb = COLOR_MUTED

    p = tf1.add_paragraph(); p.text = "Gemini Canvas (ტექსტი და კოდი)"
    p.font.name = FONT_ARIAL; p.font.size = Pt(21); p.font.bold = True; p.font.color.rgb = COLOR_DARK; p.space_before = Pt(4)

    canvas_points = [
        ("Side-by-Side გარემო", "მარცხნივ საუბრის ჩატი, მარჯვნივ ინტერაქტიული დოკუმენტი ან პროგრამული კოდი."),
        ("Inline რედაქტირება", "კონკრეტული აბზაცის ან ფუნქციის მონიშვნა და მყისიერი შესწორება მთელი ტექსტის თავიდან გენერირების გარეშე."),
        ("სტილისა და ტონის სლაიდერები", "ტექსტის სირთულის, სიგრძისა და ფორმატის ერთი დაწკაპუნებით რეგულირება."),
        ("კოდის შესრულება და დიაგნოსტიკა", "კოდის სინტაქსის შემოწმება, ერორების პოვნა და ოპტიმიზაცია.")
    ]
    for t, d in canvas_points:
        p = tf1.add_paragraph(); p.text = f"• {t}: {d}"
        p.font.name = FONT_ARIAL; p.font.size = Pt(13); p.font.color.rgb = COLOR_TEXT_MUTED; p.space_before = Pt(8)

    # Right: Gemini Deep Research
    b2 = s5.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(9.65), box_y, box_w, box_h)
    b2.fill.solid(); b2.fill.fore_color.rgb = COLOR_CARD_RED
    b2.line.color.rgb = COLOR_BORDER_RED; b2.line.width = Pt(1.5)

    tb2 = s5.shapes.add_textbox(Inches(10.05), box_y + Inches(0.4), box_w - Inches(0.8), box_h - Inches(0.8))
    tf2 = tb2.text_frame; tf2.word_wrap = True; tf2.margin_left = tf2.margin_top = 0
    p = tf2.paragraphs[0]; p.text = "ᲐᲕᲢᲝᲜᲝᲛᲘᲣᲠᲘ ᲐᲒᲔᲜᲢᲣᲠᲘ ᲙᲕᲚᲔᲕᲐ (2026)"
    p.font.name = FONT_ARIAL; p.font.size = Pt(12); p.font.bold = True; p.font.color.rgb = COLOR_RED

    p = tf2.add_paragraph(); p.text = "Gemini Deep Research"
    p.font.name = FONT_ARIAL; p.font.size = Pt(21); p.font.bold = True; p.font.color.rgb = COLOR_DARK; p.space_before = Pt(4)

    deep_points = [
        ("მრავალეტაპიანი ავტონომიური ძიება", "მოდელი თავად ადგენს კვლევის გეგმას, აგენერირებს ათობით საძიებო მოთხოვნას და კითხულობს ასობით ვებ-გვერდსა და PDF-ს."),
        ("Google Search & Workspace სინთეზი", "აერთიანებს საჯარო ვებ-მონაცემებსა და თქვენს პირად Drive/Gmail დოკუმენტებს."),
        ("ფაქტების კროს-ვალიდაცია", "ამოწმებს ურთიერთსაწინააღმდეგო ინფორმაციას და ახდენს მონაცემების შეჯერებას."),
        ("სრულფასოვანი ანალიტიკური რეპორტი", "აგენერირებს აკადემიური დონის ვრცელ ანგარიშს ზუსტი ციტატებითა და ბმულებით.")
    ]
    for t, d in deep_points:
        p = tf2.add_paragraph(); p.text = f"⭐ {t}"
        p.font.name = FONT_ARIAL; p.font.size = Pt(13); p.font.bold = True; p.font.color.rgb = COLOR_DARK; p.space_before = Pt(8)
        p = tf2.add_paragraph(); p.text = d
        p.font.name = FONT_ARIAL; p.font.size = Pt(12); p.font.color.rgb = COLOR_TEXT_MUTED; p.space_before = Pt(2)

    # Bottom takeaway
    bb = s5.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(1.0), bot_y, Inches(17.0), Inches(2.0))
    bb.fill.solid(); bb.fill.fore_color.rgb = COLOR_WHITE
    bb.line.color.rgb = COLOR_BORDER; bb.line.width = Pt(1.2)

    tbb = s5.shapes.add_textbox(Inches(1.4), bot_y + Inches(0.3), Inches(16.2), Inches(1.4))
    tfb = tbb.text_frame; tfb.word_wrap = True; tfb.margin_left = tfb.margin_top = 0
    p = tfb.paragraphs[0]; p.text = "💡 კვლევისა და შექმნის სინერგია:"
    p.font.name = FONT_ARIAL; p.font.size = Pt(14); p.font.bold = True; p.font.color.rgb = COLOR_RED
    p = tfb.add_paragraph()
    p.text = "Deep Research ატარებს სიღრმისეულ კვლევას და აგროვებს ფაქტებს, ხოლო Canvas-ში მომხმარებელი და AI ერთობლივად ამუშავებენ საბოლოო დოკუმენტს ან სტრატეგიულ გეგმას."
    p.font.name = FONT_ARIAL; p.font.size = Pt(13); p.font.color.rgb = COLOR_MUTED; p.space_before = Pt(4)

    # -------------------------------------------------------------------------
    # SLIDE 6: Google AI Studio - Developer Prototyping Platform
    # -------------------------------------------------------------------------
    s6 = created_slides[4]
    add_header(s6, "Google AI Studio - დეველოპერული პლატფორმა", "დეველოპმენტი • მოდელების ტესტირება, პრომპტინგი და API ინტეგრაცია")

    cards_s6 = [
        ("🛠️ ENVIRONMENT", "რა არის AI Studio?", "პროტოტიპირების IDE", [
            ("სწრაფი ტესტირება", "ვებ-ინტერფეისი უახლესი Gemini მოდელების (Gemini 4, 3.8 Flash) მყისიერი შემოწმებისთვის."),
            ("Zero-Setup", "არ საჭიროებს რთულ ინფრასტრუქტურას — მუშაობა იწყება ბრაუზერიდან Google ანგარიშით."),
            ("API Keys Management", "უფასო და ფასიანი API გასაღებების მარტივი გენერაცია და ლიმიტების მონიტორინგი.")
        ]),
        ("📑 PROMPT MODES", "Prompt-ის რეჟიმები", "3 სამუშაო ფორმატი", [
            ("Chat Prompt", "მრავალრაუნდიანი დიალოგი კონტექსტის შენარჩუნებით (ჩატბოტების პროტოტიპი)."),
            ("Freeform Prompt", "ტექსტისა და მედიის თავისუფალი კომბინაცია Few-Shot მაგალითებით."),
            ("Structured Prompt", "მკაცრად განსაზღვრული ცხრილური ფორმატი (Input / Output წყვილები) მონაცემთა ექსტრაქციისთვის.")
        ]),
        ("🧠 SYSTEM INSTRUCTIONS", "System Instructions", "მოდელის ქცევის წესები", [
            ("გლობალური როლი", "განსაზღვრავს მოდელის პერსონას, ექსპერტიზასა და პროფესიულ ჩარჩოს."),
            ("მკაცრი შეზღუდვები", "რა არ უნდა გააკეთოს მოდელმა (მაგ. „არასოდეს გამოიყენო სავარაუდო ფაქტები“)."),
            ("Output Constraints", "პასუხების ფორმატირება მხოლოდ JSON, Markdown ან ცხრილის სახით.")
        ])
    ]

    for i, (badge, title, subtitle, items) in enumerate(cards_s6):
        bg_col = COLOR_CARD_RED if i == 1 else COLOR_BG_LIGHT
        bor_col = COLOR_BORDER_RED if i == 1 else COLOR_BORDER
        c = s6.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, xs_3[i], card_y, card_w, card_h)
        c.fill.solid(); c.fill.fore_color.rgb = bg_col
        c.line.color.rgb = bor_col; c.line.width = Pt(1.5)

        tb = s6.shapes.add_textbox(xs_3[i] + Inches(0.35), card_y + Inches(0.35), card_w - Inches(0.7), card_h - Inches(0.7))
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
    bb = s6.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(1.0), bot_y, Inches(17.0), Inches(2.0))
    bb.fill.solid(); bb.fill.fore_color.rgb = COLOR_WHITE
    bb.line.color.rgb = COLOR_BORDER; bb.line.width = Pt(1.2)

    tbb = s6.shapes.add_textbox(Inches(1.4), bot_y + Inches(0.3), Inches(16.2), Inches(1.4))
    tfb = tbb.text_frame; tfb.word_wrap = True; tfb.margin_left = tfb.margin_top = 0
    p = tfb.paragraphs[0]; p.text = "💡 დეველოპერული ძალა (Get Code):"
    p.font.name = FONT_ARIAL; p.font.size = Pt(14); p.font.bold = True; p.font.color.rgb = COLOR_RED
    p = tfb.add_paragraph()
    p.text = "AI Studio-ში აწყობილი ნებისმიერი პრომპტი ერთი დაწკაპუნებით გადადის მზა პროგრამულ კოდში (Python, JavaScript, cURL, Swift), რაც იდეალურ ხიდს ქმნის იდეიდან რეალურ აპლიკაციამდე."
    p.font.name = FONT_ARIAL; p.font.size = Pt(13); p.font.color.rgb = COLOR_MUTED; p.space_before = Pt(4)

    # -------------------------------------------------------------------------
    # SLIDE 7: LLM Parameters Control: Temperature & Creativity
    # -------------------------------------------------------------------------
    s7 = created_slides[5]
    add_header(s7, "LLM პარამეტრების მართვა: Temperature და კრეატიულობა", "პარამეტრები • ალბათობების განაწილების მათემატიკური რეგულირება")

    temp_s7 = [
        ("Temperature = 0.0", "დეტერმინისტული / მკაცრი", [
            ("როგორ მუშაობს", "მოდელი ყოველთვის ირჩევს უმაღლესი ალბათობის ტოკენს (Argmax)."),
            ("შედეგი", "პასუხი არის 100%-ით განმეორებადი, ზუსტი და ლოგიკური."),
            ("იდეალურია", "კოდირებისთვის, მონაცემთა ექსტრაქციისთვის, მათემატიკისთვის, იურიდიული ანალიზისთვის.")
        ]),
        ("Temperature = 0.7 - 1.0", "დაბალანსებული / ბუნებრივი", [
            ("როგორ მუშაობს", "ალბათობების განაწილება ბუნებრივია — მოდელი ზოგჯერ ირჩევს მეორე-მესამე ტოკენსაც."),
            ("შედეგი", "ტექსტი იკითხება ცოცხლად, მრავალფეროვნად და ადამიანურად."),
            ("იდეალურია", "საქმიანი მიმოწერისთვის, დოკუმენტების რეზიუმირებისთვის, ყოველდღიური ჩატისთვის.")
        ]),
        ("Temperature = 1.5 - 2.0", "მაქსიმალური კრეატიულობა", [
            ("როგორ მუშაობს", "ალბათობები გათანაბრებულია — დაბალი ალბათობის სიტყვებსაც ეძლევა შანსი (Softmax Flattening)."),
            ("შედეგი", "მოულოდნელი, ორიგინალური მეტაფორები და იდეები."),
            ("რისკები", "იზრდება ჰალუცინაციების, ლოგიკური შეცდომებისა და აბსურდული ფრაზების ალბათობა.")
        ])
    ]

    for i, (title, subtitle, items) in enumerate(temp_s7):
        bg_col = COLOR_CARD_RED if i == 0 else COLOR_BG_LIGHT
        bor_col = COLOR_BORDER_RED if i == 0 else COLOR_BORDER
        c = s7.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, xs_3[i], card_y, card_w, card_h)
        c.fill.solid(); c.fill.fore_color.rgb = bg_col
        c.line.color.rgb = bor_col; c.line.width = Pt(1.5)

        tb = s7.shapes.add_textbox(xs_3[i] + Inches(0.35), card_y + Inches(0.35), card_w - Inches(0.7), card_h - Inches(0.7))
        tf = tb.text_frame; tf.word_wrap = True; tf.margin_left = tf.margin_top = 0

        p = tf.paragraphs[0]; p.text = f"ᲠᲔᲟᲘᲛᲘ {i+1}".upper()
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
    bb = s7.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(1.0), bot_y, Inches(17.0), Inches(2.0))
    bb.fill.solid(); bb.fill.fore_color.rgb = COLOR_WHITE
    bb.line.color.rgb = COLOR_BORDER; bb.line.width = Pt(1.2)

    tbb = s7.shapes.add_textbox(Inches(1.4), bot_y + Inches(0.3), Inches(16.2), Inches(1.4))
    tfb = tbb.text_frame; tfb.word_wrap = True; tfb.margin_left = tfb.margin_top = 0
    p = tfb.paragraphs[0]; p.text = "💡 მთავარი არსი (მარტივი ანალოგია):"
    p.font.name = FONT_ARIAL; p.font.size = Pt(14); p.font.bold = True; p.font.color.rgb = COLOR_RED
    p = tfb.add_paragraph()
    p.text = "Temperature არის მოდელის „ფანტაზიის რეგულატორი“: 0-ზე მოდელი მკაცრი ბუღალტერია, რომელიც მხოლოდ ფაქტებს აღიარებს; 2-ზე კი თავისუფალი მწერალი, რომელიც ყველაზე უცნაურ და მოულოდნელ სიტყვებს ირჩევს."
    p.font.name = FONT_ARIAL; p.font.size = Pt(13); p.font.color.rgb = COLOR_MUTED; p.space_before = Pt(4)

    # -------------------------------------------------------------------------
    # SLIDE 8: Probability Filters: Top-P, Top-K, Max Tokens & Thinking Budget
    # -------------------------------------------------------------------------
    s8 = created_slides[6]
    add_header(s8, "ალბათური ფილტრები: Top-P, Top-K, Max Tokens და Thinking Budget", "პარამეტრები • გენერაციის შეზღუდვისა და ფილტრაციის მექანიზმები")

    filters_s8 = [
        ("Top-P (Nucleus Sampling)", "კუმულაციური ზღვარი", [
            ("როგორ მუშაობს", "ტოკენების ალბათობები ლაგდება კლებადობით და მოდელი ირჩევს მხოლოდ იმ ტოპ-ტოკენებს, რომელთა ჯამი აღწევს P მნიშვნელობას (მაგ. 0.95)."),
            ("დანიშნულება", "ჭრის ნაკლებად სავარაუდო სიტყვების „გრძელ კუდს“ და იცავს ტექსტს უაზრო ფრაზებისგან.")
        ]),
        ("Top-K Sampling", "ფიქსირებული არჩევანი", [
            ("როგორ მუშაობს", "მოდელი ზღუდავს კანდიდატთა სიას ზუსტად K ყველაზე ალბათური ტოკენით (მაგ. Top-K = 40)."),
            ("დანიშნულება", "გამორიცხავს აბსოლუტურად შეუსაბამო ტერმინებს და უზრუნველყოფს პასუხის კონცენტრირებულობას.")
        ]),
        ("Max Output Tokens", "სიგრძის ლიმიტი", [
            ("როგორ მუშაობს", "განსაზღვრავს ტოკენების მაქსიმალურ რაოდენობას, რაც მოდელმა შეიძლება გამოიტანოს ერთ პასუხში."),
            ("დანიშნულება", "იცავს სისტემას უსასრულო ლუპებისგან, აკონტროლებს API-ს ხარჯებს და პასუხის ლაკონიურობას.")
        ]),
        ("Thinking Budget", "მსჯელობის რესურსი (2026)", [
            ("როგორ მუშაობს", "არეგულირებს, რამდენ დროსა და გამოთვლით ტოკენებს ხარჯავს მოდელი ფიქრის პროცესში (None, Low, Medium, High)."),
            ("დანიშნულება", "რთულ ამოცანებში უზრუნველყოფს ღრმა მსჯელობას შეცდომების მინიმიზაციით.")
        ])
    ]

    for i, (title, subtitle, items) in enumerate(filters_s8):
        bg_col = COLOR_CARD_RED if i == 0 else COLOR_BG_LIGHT
        bor_col = COLOR_BORDER_RED if i == 0 else COLOR_BORDER
        c = s8.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, xs_4[i], step_y, step_w, step_h)
        c.fill.solid(); c.fill.fore_color.rgb = bg_col
        c.line.color.rgb = bor_col; c.line.width = Pt(1.5)

        tb = s8.shapes.add_textbox(xs_4[i] + Inches(0.25), step_y + Inches(0.35), step_w - Inches(0.5), step_h - Inches(0.7))
        tf = tb.text_frame; tf.word_wrap = True; tf.margin_left = tf.margin_top = 0

        p = tf.paragraphs[0]; p.text = f"ᲤᲘᲚᲢᲠᲘ {i+1}".upper()
        p.font.name = FONT_ARIAL; p.font.size = Pt(11); p.font.bold = True
        p.font.color.rgb = COLOR_RED if i == 0 else COLOR_MUTED

        p = tf.add_paragraph(); p.text = title
        p.font.name = FONT_ARIAL; p.font.size = Pt(16); p.font.bold = True; p.font.color.rgb = COLOR_DARK; p.space_before = Pt(4)

        p = tf.add_paragraph(); p.text = subtitle
        p.font.name = FONT_ARIAL; p.font.size = Pt(12); p.font.color.rgb = COLOR_TEXT_MUTED; p.space_before = Pt(2)

        for heading, body in items:
            p = tf.add_paragraph(); p.text = f"• {heading}: {body}"
            p.font.name = FONT_ARIAL; p.font.size = Pt(12); p.font.color.rgb = COLOR_DARK; p.space_before = Pt(8)

    # Bottom takeaway
    bb = s8.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(1.0), bot_y, Inches(17.0), Inches(2.0))
    bb.fill.solid(); bb.fill.fore_color.rgb = COLOR_WHITE
    bb.line.color.rgb = COLOR_BORDER; bb.line.width = Pt(1.2)

    tbb = s8.shapes.add_textbox(Inches(1.4), bot_y + Inches(0.3), Inches(16.2), Inches(1.4))
    tfb = tbb.text_frame; tfb.word_wrap = True; tfb.margin_left = tfb.margin_top = 0
    p = tfb.paragraphs[0]; p.text = "💡 პარამეტრების კომბინაციის წესი:"
    p.font.name = FONT_ARIAL; p.font.size = Pt(14); p.font.bold = True; p.font.color.rgb = COLOR_RED
    p = tfb.add_paragraph()
    p.text = "Temperature და Top-P ერთად მუშაობს: Temperature განსაზღვრავს ალბათობების განაწილებას, ხოლო Top-P და Top-K ადგენს დასაშვებ საზღვრებს. ზუსტი საინჟინრო ამოცანებისთვის: Temp=0.0, Top-P=0.95, High Thinking Budget."
    p.font.name = FONT_ARIAL; p.font.size = Pt(13); p.font.color.rgb = COLOR_MUTED; p.space_before = Pt(4)

    # -------------------------------------------------------------------------
    # SLIDE 9: Safety Settings & Responsible AI (Grounding & SynthID)
    # -------------------------------------------------------------------------
    s9 = created_slides[7]
    add_header(s9, "Safety Settings და პასუხისმგებლიანი AI", "უსაფრთხოება • ფილტრები, წყლის ნიშნები და ფაქტების შემოწმება")

    cards_s9 = [
        ("🛡️ SAFETY THRESHOLDS", "უსაფრთხოების ფილტრები", "AI Studio ზღვრები", [
            ("4 საკვანძო კატეგორია", "Harassment (შევიწროება), Hate Speech (სიძულვილის ენა), Sexually Explicit (ინტიმური), Dangerous Content (საფრთხე)."),
            ("ზღვრების რეგულირება", "Block None (დეველოპერული ტესტი), Block Few, Block Some, Block Most."),
            ("ბიზნეს ადაპტაცია", "სამედიცინო ან იურიდიული აპლიკაციებისთვის ფილტრების ზუსტი მორგება.")
        ]),
        ("🔍 GROUNDING", "Grounding with Google Search", "ფაქტების შემოწმება", [
            ("ცოცხალი Google Search", "მოდელი ამოწმებს საკუთარ პასუხებს Google-ის უახლეს საძიებო ინდექსში რეალურ დროში."),
            ("ციტირების ნიშნულები", "თითოეულ ფაქტს ეწერება პირდაპირი ბმული და წყაროს სანდოობის ქულა."),
            ("ჰალუცინაციების პრევენცია", "რადიკალურად ამცირებს მოგონილი მონაცემების, სახელებისა და თარიღების რისკს.")
        ]),
        ("💧 SYNTHID", "SynthID & Watermarking", "ციფრული იდენტიფიკაცია", [
            ("უხილავი წყლის ნიშანი", "Google DeepMind-ის ტექნოლოგია, რომელიც AI-გენერირებულ ტექსტში, სურათსა და აუდიოში ათავსებს უხილავ კოდს."),
            ("დეზინფორმაციის ბლოკირება", "მკაფიოდ ავლენს, ტექსტი დაწერილია ადამიანის თუ AI-ს მიერ."),
            ("მდგრადობა", "წყლის ნიშანი არ იშლება ტექსტის მცირე რედაქტირების ან სურათის კომპრესიის დროს.")
        ])
    ]

    for i, (badge, title, subtitle, items) in enumerate(cards_s9):
        bg_col = COLOR_CARD_RED if i == 1 else COLOR_BG_LIGHT
        bor_col = COLOR_BORDER_RED if i == 1 else COLOR_BORDER
        c = s9.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, xs_3[i], card_y, card_w, card_h)
        c.fill.solid(); c.fill.fore_color.rgb = bg_col
        c.line.color.rgb = bor_col; c.line.width = Pt(1.5)

        tb = s9.shapes.add_textbox(xs_3[i] + Inches(0.35), card_y + Inches(0.35), card_w - Inches(0.7), card_h - Inches(0.7))
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
    bb = s9.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(1.0), bot_y, Inches(17.0), Inches(2.0))
    bb.fill.solid(); bb.fill.fore_color.rgb = COLOR_WHITE
    bb.line.color.rgb = COLOR_BORDER; bb.line.width = Pt(1.2)

    tbb = s9.shapes.add_textbox(Inches(1.4), bot_y + Inches(0.3), Inches(16.2), Inches(1.4))
    tfb = tbb.text_frame; tfb.word_wrap = True; tfb.margin_left = tfb.margin_top = 0
    p = tfb.paragraphs[0]; p.text = "🎯 შემაჯამებელი დასკვნა:"
    p.font.name = FONT_ARIAL; p.font.size = Pt(14); p.font.bold = True; p.font.color.rgb = COLOR_RED
    p = tfb.add_paragraph()
    p.text = "Gemini და Google AI Studio დეველოპერს აძლევს სრულყოფილ ინსტრუმენტულ არსენალს: მულტიმოდალურ მოდელებს უზარმაზარი კონტექსტით, პარამეტრების მათემატიკურ მართვას და საწარმოო დონის უსაფრთხოებას ფაქტების ავტომატური შემოწმებით."
    p.font.name = FONT_ARIAL; p.font.size = Pt(13); p.font.color.rgb = COLOR_MUTED; p.space_before = Pt(4)

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
    filename = "7 - Gemini სრულად I - Gemini, Workspace, AI Studio და პარამეტრები.pptx"
    out_master = os.path.join(r'd:\ხელოვნური ინტელექტის ინსტრუმენტები', filename)
    prs.save(out_master)
    print(f"Master PPTX for Lecture 7 generated successfully: {out_master}")

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
            print(f"Copied to target: {dst_file}")
        except Exception as e:
            print(f"Error copying to {td}: {e}")

if __name__ == '__main__':
    build_lecture_7()
