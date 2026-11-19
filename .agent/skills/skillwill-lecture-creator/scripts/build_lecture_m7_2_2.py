import os
import sys
import shutil
import pptx
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN
from pptx.enum.shapes import MSO_SHAPE

def build_lecture_m7_2_2():
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
            p.text = "2.2 - კონტექსტის მართვა, PII შეზღუდვები და Prefill ტექნიკა"
            p.alignment = PP_ALIGN.CENTER
            p.font.name = FONT_ARIAL
            p.font.size = Pt(32)
            p.font.bold = True
            p.font.color.rgb = COLOR_RED

            p2 = tf.add_paragraph()
            p2.text = "როლის ეფექტიანობა, მონაცემთა უსაფრთხოება, ფორმატის კონტროლი და Output Steering"
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
    def add_header(slide, title, category="ᲛᲝᲓᲣᲚᲘ 7 • PROMPT ENGINEERING • 2.2"):
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
    # SLIDE 2: Paradigm Shift (Context Pollution vs Data Hygiene)
    # -------------------------------------------------------------------------
    s2 = created_slides[0]
    add_header(s2, "კონტექსტის დაბინძურება vs მიზანმიმართული მონაცემთა ჰიგიენა", "ᲞᲐᲠᲐᲓᲘᲒᲛᲐ • CONTEXT POLLUTION & PRIVACY RISKS")

    cards_s2 = [
        ("🗂️ CONTEXT POLLUTION", "კონტექსტის დაბინძურება", "ჭარბი და გაუფილტრავი მასალა", [
            ("Attention Dilution", "სრული ლოგებისა და პირადი მიმოწერის გაუფილტრავი ჩაყრა პრომპტში ფანტავს ყურადღების ვექტორებს."),
            ("ხარჯების ზრდა", "ათობით ათასი არარელევანტური ტოკენი ზრდის API გადასახადებს და იწვევს სისტემის დაყოვნებას (Latency)."),
            ("დაკარგული ფაქტები", "დიდ კონტექსტში კრიტიკული ფაქტები იკარგება (Needle-in-a-Haystack პრობლემა).")
        ]),
        ("🛑 PRIVACY & PII RISKS", "კონფიდენციალობის რისკი", "მონაცემთა გაჟონვის საფრთხე", [
            ("პირადი მონაცემები (PII)", "პასპორტის ნომრების, ტელეფონებისა და პირადი მისამართების გაგზავნა არღვევს GDPR-სა და EU AI Act-ს."),
            ("კორპორაციული საიდუმლოებები", "ფინანსური ანგარიშებისა და მომხმარებელთა სიების გაჟონვა პროვაიდერის ლოგებში."),
            ("Credentials Leakage", "API გასაღებებისა და პაროლების შემთხვევითი მოხვედრა LLM-ის შემავალ ნაკადში.")
        ]),
        ("🎯 TARGETED CONTEXT & STEERING", "საინჟინრო მართვა", "სუფთა მონაცემები & Prefill", [
            ("მონაცემთა ანონიმიზაცია", "სენსიტიური ველების ჩანაცვლება უსაფრთხო ტოკენებით ([USER_1], [ORG_A]) მოდელამდე მიღწევამდე."),
            ("მიზანმიმართული ექსტრაქცია", "მოდელს მიეწოდება მხოლოდ ამოცანისთვის კრიტიკულად აუცილებელი ფაქტები."),
            ("Output Steering (Prefill)", "პასუხის დასაწყისის კარნახი, რაც გარანტირებულად იცავს სასურველ ფორმატს.")
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
    p.text = "კონტექსტი პრომპტში არის საწვავი — ზედმეტი ან დაბინძურებული საწვავი ძრავს აფუჭებს; საინჟინრო ამოცანაა მოდელს მივაწოდოთ მხოლოდ გასუფთავებული, ანონიმიზებული და რელევანტური ფაქტები."
    p.font.name = FONT_ARIAL; p.font.size = Pt(13); p.font.color.rgb = COLOR_MUTED; p.space_before = Pt(4)

    # -------------------------------------------------------------------------
    # SLIDE 3: Anatomy: Role Efficacy (When It Works vs When It Fails)
    # -------------------------------------------------------------------------
    s3 = created_slides[1]
    add_header(s3, "როლის (Role) ანატომია: როდის მუშაობს და როდის არის ზედმეტი?", "ᲐᲜᲐᲢᲝᲛᲘᲐ • როლის რეალური წონა და მითები")

    role_s3 = [
        ("✅ როდის შველის როლი (High Value)", "რეალური ეფექტიანობა", [
            ("დარგობრივი ექსპერტიზა", "როდესაც ამოცანა მოითხოვს პროფესიულ ცოდნას (იურისტი, Cloud Security Architect, FinOps Lead)."),
            ("აუდიტორიის ენა", "მოდელი ირჩევს შესაბამის ტერმინოლოგიასა და აბსტრაქციის დონეს (C-Level vs Junior Dev)."),
            ("საერთაშორისო სტანდარტები", "როლი ააქტიურებს სპეციფიკურ ჩარჩოებს (ISO 27001, OWASP, GAAP, Clean Code).")
        ]),
        ("❌ როდის არის როლი ზედმეტი (No Value)", "ილუზიები და მითები", [
            ("მარტივი ოპერაციები", "ტექსტის გაწმენდა, პირდაპირი თარგმანი, JSON-ში გადაყვანა არ საჭიროებს როლს."),
            ("არითმეტიკული გამოთვლები", "„შენ ხარ აინშტაინი“ არ აუმჯობესებს მათემატიკურ გამოთვლებს და ლოგიკას."),
            ("შეუსაბამო მეტაფორები", "„შენ ხარ შექსპირი, დაწერე SQL ქვერი“ — იწვევს ლოგიკის დაბნევასა და შეცდომებს.")
        ]),
        ("⚖️ საინჟინრო ბალანსი (Role Precision)", "პროფესიული ფორმულირება", [
            ("ფუნქციური როლი", "„Senior React Architect“ ბევრად უკეთესია, ვიდრე „საუკეთესო პროგრამისტი სამყაროში“."),
            ("მოლოდინების გაწერა", "როლს თან უნდა ახლდეს მისი შეფასების კრიტერიუმები და პასუხისმგებლობის სფერო."),
            ("სისტემურ დონეზე ჩასმა", "როლი ყოველთვის უნდა გაიწეროს System Instructions-ში და არა User შეტყობინებაში.")
        ])
    ]

    for i, (title, subtitle, items) in enumerate(role_s3):
        bg_col = COLOR_CARD_RED if i == 0 else COLOR_BG_LIGHT
        bor_col = COLOR_BORDER_RED if i == 0 else COLOR_BORDER
        c = s3.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, xs_3[i], card_y, card_w, card_h)
        c.fill.solid(); c.fill.fore_color.rgb = bg_col
        c.line.color.rgb = bor_col; c.line.width = Pt(1.5)

        tb = s3.shapes.add_textbox(xs_3[i] + Inches(0.35), card_y + Inches(0.35), card_w - Inches(0.7), card_h - Inches(0.7))
        tf = tb.text_frame; tf.word_wrap = True; tf.margin_left = tf.margin_top = 0

        p = tf.paragraphs[0]; p.text = f"ᲐᲡᲞᲔᲥᲢᲘ {i+1}".upper()
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
    p = tfb.paragraphs[0]; p.text = "💡 როლის ოქროს წესი:"
    p.font.name = FONT_ARIAL; p.font.size = Pt(14); p.font.bold = True; p.font.color.rgb = COLOR_RED
    p = tfb.add_paragraph()
    p.text = "როლი არ არის მაგიური შელოცვა — ის მუშაობს მხოლოდ მაშინ, როდესაც განსაზღვრავს ექსპერტიზის დომენს, სამიზნე აუდიტორიას, შეფასების კრიტერიუმებსა და ტექნიკურ სტანდარტებს."
    p.font.name = FONT_ARIAL; p.font.size = Pt(13); p.font.color.rgb = COLOR_MUTED; p.space_before = Pt(4)

    # -------------------------------------------------------------------------
    # SLIDE 4: Mechanism: Data Protection & PII Sanitization
    # -------------------------------------------------------------------------
    s4 = created_slides[2]
    add_header(s4, "მონაცემთა ჰიგიენა: PII და კონფიდენციალობის ფილტრაცია", "ᲛᲔᲥᲐᲜᲘᲖᲛᲘ • მონაცემთა გაჟონვის პრევენცია და ანონიმიზაცია")

    # Left: Prohibited Data
    b1 = s4.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(1.0), box_y, box_w, box_h)
    b1.fill.solid(); b1.fill.fore_color.rgb = COLOR_BG_LIGHT
    b1.line.color.rgb = COLOR_BORDER; b1.line.width = Pt(1.5)

    tb1 = s4.shapes.add_textbox(Inches(1.4), box_y + Inches(0.4), box_w - Inches(0.8), box_h - Inches(0.8))
    tf1 = tb1.text_frame; tf1.word_wrap = True; tf1.margin_left = tf1.margin_top = 0
    p = tf1.paragraphs[0]; p.text = "WHAT NEVER ENTERS? (აკრძალული მონაცემები)"
    p.font.name = FONT_ARIAL; p.font.size = Pt(12); p.font.bold = True; p.font.color.rgb = COLOR_MUTED

    p = tf1.add_paragraph(); p.text = "რა არ უნდა მოხვდეს პრომპტში?"
    p.font.name = FONT_ARIAL; p.font.size = Pt(21); p.font.bold = True; p.font.color.rgb = COLOR_DARK; p.space_before = Pt(4)

    proh_points = [
        ("Personally Identifiable Information (PII)", "პირადი ნომრები, პასპორტის მონაცემები, ბარათის ნომრები, ტელეფონები, მისამართები."),
        ("კორპორაციული საიდუმლოებები", "ფინანსური ანგარიშები, მომხმარებელთა სიები, ჯერ გამოუქვეყნებელი სტრატეგიული გეგმები."),
        ("Secrets & Credentials", "API გასაღებები, პაროლები, SSH გასაღებები, Database Connection Strings."),
        ("სამედიცინო და იურიდიული მონაცემები", "პაციენტის ისტორია, ჯანმრთელობის ჩანაწერები და კონფიდენციალური სასამართლო საქმეები.")
    ]
    for t, d in proh_points:
        p = tf1.add_paragraph(); p.text = f"• {t}: {d}"
        p.font.name = FONT_ARIAL; p.font.size = Pt(13); p.font.color.rgb = COLOR_TEXT_MUTED; p.space_before = Pt(8)

    # Right: Sanitization Pipeline
    b2 = s4.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(9.65), box_y, box_w, box_h)
    b2.fill.solid(); b2.fill.fore_color.rgb = COLOR_CARD_RED
    b2.line.color.rgb = COLOR_BORDER_RED; b2.line.width = Pt(1.5)

    tb2 = s4.shapes.add_textbox(Inches(10.05), box_y + Inches(0.4), box_w - Inches(0.8), box_h - Inches(0.8))
    tf2 = tb2.text_frame; tf2.word_wrap = True; tf2.margin_left = tf2.margin_top = 0
    p = tf2.paragraphs[0]; p.text = "SANITIZATION PIPELINE (ანონიმიზაციის მექანიზმი)"
    p.font.name = FONT_ARIAL; p.font.size = Pt(12); p.font.bold = True; p.font.color.rgb = COLOR_RED

    p = tf2.add_paragraph(); p.text = "მონაცემთა გაწმენდის ეტაპები"
    p.font.name = FONT_ARIAL; p.font.size = Pt(21); p.font.bold = True; p.font.color.rgb = COLOR_DARK; p.space_before = Pt(4)

    sanit_points = [
        ("Token Replacement (ჩანაცვლება)", "john.doe@company.ge ➔ [USER_EMAIL_1], 010240... ➔ [USER_ID_1]."),
        ("Synthetic Data (სინთეზური მონაცემები)", "ტესტირებისთვის რეალური მონაცემების ნაცვლად ფიქტიური ჩანაწერების გამოყენება."),
        ("Regex & Automated Masking", "ავტომატური სკრიპტები (Presidio), რომლებიც პრომპტის გაგზავნამდე ჭრის სენსიტიურ პატერნებს."),
        ("De-anonymization (უკუჩანაცვლება)", "მოდელის პასუხში უსაფრთხო ტოკენების უკან ჩანაცვლება კომპანიის შიდა დაცულ სერვერზე.")
    ]
    for t, d in sanit_points:
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
    p = tfb.paragraphs[0]; p.text = "💡 უსაფრთხოების წესი:"
    p.font.name = FONT_ARIAL; p.font.size = Pt(14); p.font.bold = True; p.font.color.rgb = COLOR_RED
    p = tfb.add_paragraph()
    p.text = "ერთხელ გაგზავნილი მონაცემი უკან აღარ ბრუნდება — საწარმოო პრომპტინგში მოქმედებს Zero-Trust პრინციპი: არცერთი PII ტოკენი არ უნდა მოხვდეს მოდელის კონტექსტში გაფილტვრის გარეშე."
    p.font.name = FONT_ARIAL; p.font.size = Pt(13); p.font.color.rgb = COLOR_MUTED; p.space_before = Pt(4)

    # -------------------------------------------------------------------------
    # SLIDE 5: Ecosystem: 5 Dimensions of Format Control
    # -------------------------------------------------------------------------
    s5 = created_slides[3]
    add_header(s5, "ფორმატის კონტროლის 5 განზომილება: სტრუქტურა, სია, სიგრძე, ენა და ტონი", "ᲔᲙᲝᲡᲘᲡᲢᲔᲛᲐ • გამომავალი ნაკადის საინჟინრო პარამეტრები")

    format_s5 = [
        ("Structure (სტრუქტურა)", "მონაცემთა მოდელი", [
            ("Markdown ცხრილები", "სვეტების მკაცრი დასახელება და მონაცემთა შედარება."),
            ("JSON Schema", "სავალდებულო გასაღებები (Keys) და ტიპიზაცია (Types)."),
            ("Key-Value წყვილები", "მყისიერად წაკითხვადი ბლოკები სკრიპტებისთვის.")
        ]),
        ("Scope & Length (სიგრძე)", "მოცულობის საზღვრები", [
            ("ზუსტი მეტრიკა", "„ზუსტად 3 ბულეტი“, „მაქსიმუმ 100 სიტყვა“, „1 აბზაცი“."),
            ("Stop Sequences", "სიმბოლოების დაწესება გენერაციის მყისიერად შესაწყვეტად."),
            ("ტოკენების ეკონომია", "იცავს API ბიუჯეტს ზედმეტი ხარჯვისგან.")
        ]),
        ("Tone & Style (ტონი)", "კომუნიკაციის სტილი", [
            ("აკადემიური და ობიექტური", "ფაქტებზე ორიენტირებული, ემოციების გარეშე."),
            ("კორპორაციული / საქმიანი", "პროფესიული, თავაზიანი და დიპლომატიური."),
            ("ლაკონიური საინჟინრო", "მხოლოდ ფაქტები და გადაწყვეტილებები.")
        ]),
        ("Language & Terms (ენა)", "ტერმინოლოგიური სიზუსტე", [
            ("ორენოვანი სტანდარტი", "ქართული გრამატიკა ინგლისური IT ტერმინებით."),
            ("ტერმინების ლექსიკონი", "დარგობრივი სიტყვების მკაცრი დაცვა."),
            ("კოდირება და ფორმატი", "UTF-8 და სწორი სპეციალური სიმბოლოები.")
        ])
    ]

    for i, (title, subtitle, items) in enumerate(format_s5):
        bg_col = COLOR_CARD_RED if i == 0 else COLOR_BG_LIGHT
        bor_col = COLOR_BORDER_RED if i == 0 else COLOR_BORDER
        c = s5.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, xs_4[i], step_y, step_w, step_h)
        c.fill.solid(); c.fill.fore_color.rgb = bg_col
        c.line.color.rgb = bor_col; c.line.width = Pt(1.5)

        tb = s5.shapes.add_textbox(xs_4[i] + Inches(0.25), step_y + Inches(0.35), step_w - Inches(0.5), step_h - Inches(0.7))
        tf = tb.text_frame; tf.word_wrap = True; tf.margin_left = tf.margin_top = 0

        p = tf.paragraphs[0]; p.text = f"ᲒᲐᲜᲖᲝᲛᲘᲚᲔᲑᲐ {i+1}".upper()
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
    bb = s5.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(1.0), bot_y, Inches(17.0), Inches(2.0))
    bb.fill.solid(); bb.fill.fore_color.rgb = COLOR_WHITE
    bb.line.color.rgb = COLOR_BORDER; bb.line.width = Pt(1.2)

    tbb = s5.shapes.add_textbox(Inches(1.4), bot_y + Inches(0.3), Inches(16.2), Inches(1.4))
    tfb = tbb.text_frame; tfb.word_wrap = True; tfb.margin_left = tfb.margin_top = 0
    p = tfb.paragraphs[0]; p.text = "💡 ფორმატის მართვა:"
    p.font.name = FONT_ARIAL; p.font.size = Pt(14); p.font.bold = True; p.font.color.rgb = COLOR_RED
    p = tfb.add_paragraph()
    p.text = "ფორმატის კონტროლი არ არის მხოლოდ ესთეტიკა — ეს არის გარანტია, რომ LLM-ის გამოსავალი შეცდომის გარეშე წაიკითხება მომდევნო პროგრამული მოდულის ან მონაცემთა ბაზის მიერ."
    p.font.name = FONT_ARIAL; p.font.size = Pt(13); p.font.color.rgb = COLOR_MUTED; p.space_before = Pt(4)

    # -------------------------------------------------------------------------
    # SLIDE 6: Rules: Prefill Technique (Output Steering & Zero Preamble)
    # -------------------------------------------------------------------------
    s6 = created_slides[4]
    add_header(s6, "Prefill ტექნიკა: პასუხის დასაწყისის კარნახი და ფორმატის გარანტია", "ᲬᲔᲡᲔᲑᲘ ᲓᲐ ᲠᲘᲡᲙᲔᲑᲘ • Output Steering და ჰალუცინაციების ბლოკირება")

    prefill_s6 = [
        ("⚡ რა არის Prefill?", "The Concept", "პასუხის საწყისი სიმბოლოები", [
            ("Assistant Message შევსება", "ასისტენტის როლის წინასწარ შევსება პირველი სიმბოლოებით (მაგ. {, [{\\\"id\\\":, | No | Event |)."),
            ("Next-Token იძულება", "მოდელი იძულებულია გააგრძელოს დაწყებული სტრუქტურა ალბათური მექანიზმის გამო."),
            ("დეტერმინისტული დასაწყისი", "გამორიცხავს ალტერნატიულ, არასასურველ საწყის ფრაზებს.")
        ]),
        ("🎯 3 მთავარი უპირატესობა", "საინჟინრო მოგება", "სისუფთავე და სიზუსტე", [
            ("Zero Preamble (ნულოვანი წყალი)", "მოდელი ფიზიკურად ვეღარ წერს: „რა თქმა უნდა, სიამოვნებით დაგეხმარებით...“."),
            ("100% JSON გარანტია", "გამორიცხავს Markdown Backticks (```json) ბლოკებს, რაც ხშირად ტეხავს ავტომატურ პარსერებს."),
            ("Output Steering", "მოდელის მსჯელობის ნაკადის მიმართვა სასურველი ლოგიკური მიმართულებით.")
        ]),
        ("⚠️ პლატფორმული მხარდაჭერა", "Ecosystem Integration", "ინსტრუმენტული რეალიზაცია", [
            ("Anthropic Claude API", "პირდაპირ მხარდაჭერილია Assistant შეტყობინების ველში ტექსტის ჩაწერით."),
            ("OpenAI GPT-6", "ამ როლს ასრულებს Structured Outputs ფუნქცია (Strict JSON Schema დეკოდერის დონეზე)."),
            ("Open-Weights (vLLM / Ollama)", "მხარდაჭერილია გენერაციის პარამეტრებში საწყისი პრომპტის გაგრძელებით.")
        ])
    ]

    for i, (badge, title, subtitle, items) in enumerate(prefill_s6):
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
    p = tfb.paragraphs[0]; p.text = "💡 საინჟინრო ხრიკი:"
    p.font.name = FONT_ARIAL; p.font.size = Pt(14); p.font.bold = True; p.font.color.rgb = COLOR_RED
    p = tfb.add_paragraph()
    p.text = "თუ მოდელს პასუხს დააწყებინებთ სიმბოლოთი {, ის ფიზიკურად ვეღარ დაწერს შესავალ მისალმებას და მყისიერად დააბრუნებს სუფთა მონაცემთა ობიექტს. ეს არის პრომპტინგის ერთ-ერთი ყველაზე მძლავრი პრაქტიკული ტექნიკა."
    p.font.name = FONT_ARIAL; p.font.size = Pt(13); p.font.color.rgb = COLOR_MUTED; p.space_before = Pt(4)

    # -------------------------------------------------------------------------
    # SLIDE 7: Workspaces: Platform Implementation of Format Control
    # -------------------------------------------------------------------------
    s7 = created_slides[5]
    add_header(s7, "ფორმატის მართვა პლატფორმებში: Anthropic, OpenAI და Gemini", "ᲡᲐᲛᲣᲨᲐᲝ ᲒᲐᲠᲔᲛᲝ • როგორ ვაიძულებთ მოდელს ფორმატის დაცვას")

    plat_s7 = [
        ("Anthropic Console (Prefill Tool)", "console.anthropic.com", [
            ("Assistant Field", "ასისტენტის შეტყობინების ველში ჩაწერილი ტექსტი ხდება გენერაციის საწყისი წერტილი."),
            ("XML ტეგების დახურვა", "<analysis>\\n{ ➔ მოდელი ავტომატურად ხურავს }</analysis> სტრუქტურას."),
            ("Prompt Generator", "პრომპტის ავტომატური გენერაცია Prefill-ის ოპტიმალური პარამეტრებით.")
        ]),
        ("OpenAI Workbench (Structured Outputs)", "platform.openai.com", [
            ("Strict Mode JSON Schema", "JSON Schema-ს შეყვანა პარამეტრების პანელში მკაცრი ვალიდაციის რეჟიმით."),
            ("Grammar-Based Constrained Decoding", "მოდელის დეკოდერი მათემატიკურად ბლოკავს ნებისმიერ ტოკენს, რომელიც სქემას არღვევს."),
            ("Zero Hallucinated Keys", "გარანტირებულია, რომ გასაღებების დასახელება 100%-ით ემთხვევა სქემას.")
        ]),
        ("Google AI Studio (Response Schema)", "ai.google.dev", [
            ("Structured Response", "Response Schema ველში Pydantic ან JSON Schema-ს პირდაპირი მიბმა."),
            ("Structured Prompt Mode", "Input / Output ცხრილური ტესტირება Few-Shot მაგალითებით ფორმატის გასამყარებლად."),
            ("Regex Constraints", "მოდელის პასუხის შეზღუდვა რეგულარული გამოსახულებების მიხედვით.")
        ])
    ]

    for i, (title, subtitle, items) in enumerate(plat_s7):
        bg_col = COLOR_CARD_RED if i == 0 else COLOR_BG_LIGHT
        bor_col = COLOR_BORDER_RED if i == 0 else COLOR_BORDER
        c = s7.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, xs_3[i], card_y, card_w, card_h)
        c.fill.solid(); c.fill.fore_color.rgb = bg_col
        c.line.color.rgb = bor_col; c.line.width = Pt(1.5)

        tb = s7.shapes.add_textbox(xs_3[i] + Inches(0.35), card_y + Inches(0.35), card_w - Inches(0.7), card_h - Inches(0.7))
        tf = tb.text_frame; tf.word_wrap = True; tf.margin_left = tf.margin_top = 0

        p = tf.paragraphs[0]; p.text = f"ᲞᲚᲐᲢᲤᲝᲠᲛᲐ {i+1}".upper()
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
    p = tfb.paragraphs[0]; p.text = "💡 პლატფორმული სტანდარტი:"
    p.font.name = FONT_ARIAL; p.font.size = Pt(14); p.font.bold = True; p.font.color.rgb = COLOR_RED
    p = tfb.add_paragraph()
    p.text = "2026 წელს ფორმატის დაცვა აღარ არის პრომპტის „თხოვნა“ — ეს არის API პარამეტრი (Schema / Prefill), რომელიც მათემატიკურ გარანტიას იძლევა, რომ მოდელი არ გადაუხვევს სინტაქსურ წესებს."
    p.font.name = FONT_ARIAL; p.font.size = Pt(13); p.font.color.rgb = COLOR_MUTED; p.space_before = Pt(4)

    # -------------------------------------------------------------------------
    # SLIDE 8: Experiment: Traditional Request vs Prefill & Masking
    # -------------------------------------------------------------------------
    s8 = created_slides[6]
    add_header(s8, "პრაქტიკული შედარება: ჩვეულებრივი მოთხოვნა vs Prefill & Masking", "ᲔᲥᲡᲞᲔᲠᲘᲛᲔᲜᲢᲘ • INPUT / OUTPUT შედარება და ფორმატის სიზუსტე")

    # Top container: Side-by-side comparison box
    top_box_y = Inches(2.35)
    top_box_h = Inches(3.35)

    # Left: Unfiltered & Unsteered box
    b_unf = s8.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(1.0), top_box_y, Inches(8.35), top_box_h)
    b_unf.fill.solid(); b_unf.fill.fore_color.rgb = COLOR_BG_LIGHT
    b_unf.line.color.rgb = COLOR_BORDER; b_unf.line.width = Pt(1.5)

    tb_u = s8.shapes.add_textbox(Inches(1.35), top_box_y + Inches(0.25), Inches(7.65), top_box_h - Inches(0.5))
    tf_u = tb_u.text_frame; tf_u.word_wrap = True; tf_u.margin_left = tf_u.margin_top = 0
    p = tf_u.paragraphs[0]; p.text = "❌ ᲩᲕᲔᲣᲚᲔᲑᲠᲘᲕᲘ ᲛᲝᲗᲮᲝᲕᲜᲐ (PII ᲒᲐᲟᲝᲜᲕᲘᲗ)"
    p.font.name = FONT_ARIAL; p.font.size = Pt(11); p.font.bold = True; p.font.color.rgb = COLOR_MUTED

    p = tf_u.add_paragraph(); p.text = "„აი კლიენტის მეილი: shota@firm.ge, ID: 01024... შეამოწმე JSON-ში.“"
    p.font.name = FONT_ARIAL; p.font.size = Pt(14); p.font.bold = True; p.font.color.rgb = COLOR_DARK; p.space_before = Pt(4)

    unf_flaws = [
        ("კონფიდენციალობის დარღვევა", "PII მონაცემები გაიგზავნა მოდელთან და შეინახა პროვაიდერის სერვერებზე."),
        ("ფორმატის რღვევა (Preamble)", "პასუხი იწყება ფრაზით: „რა თქმა უნდა, აი თქვენი JSON:“ და Markdown ბლოკით ```json."),
        ("პროგრამული ჩავარდნა", "პარსერი ვარდება შეცდომით (SyntaxError), რადგან ტექსტი არ წარმოადგენს სუფთა JSON-ს.")
    ]
    for h, d in unf_flaws:
        p = tf_u.add_paragraph(); p.text = f"• {h}: {d}"
        p.font.name = FONT_ARIAL; p.font.size = Pt(12); p.font.color.rgb = COLOR_TEXT_MUTED; p.space_before = Pt(4)

    # Right: Masked & Prefilled box
    b_prf = s8.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(9.65), top_box_y, Inches(8.35), top_box_h)
    b_prf.fill.solid(); b_prf.fill.fore_color.rgb = COLOR_CARD_RED
    b_prf.line.color.rgb = COLOR_BORDER_RED; b_prf.line.width = Pt(1.5)

    tb_p = s8.shapes.add_textbox(Inches(10.0), top_box_y + Inches(0.25), Inches(7.65), top_box_h - Inches(0.5))
    tf_p = tb_p.text_frame; tf_p.word_wrap = True; tf_p.margin_left = tf_p.margin_top = 0
    p = tf_p.paragraphs[0]; p.text = "✅ ᲒᲐᲤᲘᲚᲢᲠᲣᲚᲘ ᲞᲠᲝᲛᲞᲢᲘ + PREFILL"
    p.font.name = FONT_ARIAL; p.font.size = Pt(11); p.font.bold = True; p.font.color.rgb = COLOR_RED

    p = tf_p.add_paragraph(); p.text = "„შეამოწმე [CLIENT_1] მონაცემები.“ + Assistant Prefill: {\"status\":"
    p.font.name = FONT_ARIAL; p.font.size = Pt(14); p.font.bold = True; p.font.color.rgb = COLOR_DARK; p.space_before = Pt(4)

    prf_strengths = [
        ("ნულოვანი PII რისკი", "მოდელი ხედავს მხოლოდ ანონიმიზებულ ტოკენებს ([CLIENT_1]), მონაცემები დაცულია."),
        ("მყისიერი სუფთა სინტაქსი", "პასუხი იწყება პირდაპირ: {\"status\": \"VALID\", \"risk\": \"LOW\"} ყოველგვარი შესავლის გარეშე."),
        ("100% ავტომატიზაცია", "გამოსავალი წამებში პარსირდება backend სერვისის მიერ დამატებითი წმენდის გარეშე.")
    ]
    for h, d in prf_strengths:
        p = tf_p.add_paragraph(); p.text = f"⭐ {h}: {d}"
        p.font.name = FONT_ARIAL; p.font.size = Pt(12); p.font.bold = (h == "100% ავტომატიზაცია"); p.font.color.rgb = COLOR_DARK; p.space_before = Pt(4)

    # Middle Model Cards (3 cards)
    mid_y = Inches(5.95)
    mid_h = Inches(1.8)
    models_exp = [
        ("OpenAI GPT-6", "openai.png", "Structured Outputs რეჟიმში 100%-ით იცავს სქემას, ნულოვანი Preamble და სუფთა JSON."),
        ("Claude 5.5 Sonnet", "claude.png", "Prefill-ით მყისიერად აგრძელებს JSON ობიექტს შესავლის გარეშე, იდეალური ანონიმიზაცია."),
        ("Gemini 2.5 Pro", "gemini.png", "Response Schema-თი აბრუნებს სუფთა JSON ობიექტს, უნაკლო ვალიდაცია Masked ტოკენებზე.")
    ]
    for i, (m_name, m_logo, m_desc) in enumerate(models_exp):
        c = s8.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, xs_3[i], mid_y, card_w, mid_h)
        c.fill.solid(); c.fill.fore_color.rgb = COLOR_WHITE
        c.line.color.rgb = COLOR_BORDER; c.line.width = Pt(1.2)

        logo_path = os.path.join(logo_dir, m_logo)
        if os.path.exists(logo_path):
            s8.shapes.add_picture(logo_path, xs_3[i] + Inches(0.3), mid_y + Inches(0.3), width=Inches(0.65), height=Inches(0.65))

        tb = s8.shapes.add_textbox(xs_3[i] + Inches(1.1), mid_y + Inches(0.2), card_w - Inches(1.3), mid_h - Inches(0.4))
        tf = tb.text_frame; tf.word_wrap = True; tf.margin_left = tf.margin_top = 0
        p = tf.paragraphs[0]; p.text = m_name
        p.font.name = FONT_ARIAL; p.font.size = Pt(14); p.font.bold = True; p.font.color.rgb = COLOR_DARK
        p = tf.add_paragraph(); p.text = m_desc
        p.font.name = FONT_ARIAL; p.font.size = Pt(11); p.font.color.rgb = COLOR_MUTED; p.space_before = Pt(3)

    # Bottom takeaway
    bb = s8.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(1.0), bot_y, Inches(17.0), Inches(2.0))
    bb.fill.solid(); bb.fill.fore_color.rgb = COLOR_WHITE
    bb.line.color.rgb = COLOR_BORDER; bb.line.width = Pt(1.2)

    tbb = s8.shapes.add_textbox(Inches(1.4), bot_y + Inches(0.3), Inches(16.2), Inches(1.4))
    tfb = tbb.text_frame; tfb.word_wrap = True; tfb.margin_left = tfb.margin_top = 0
    p = tfb.paragraphs[0]; p.text = "💡 ექსპერიმენტის დასკვნა:"
    p.font.name = FONT_ARIAL; p.font.size = Pt(14); p.font.bold = True; p.font.color.rgb = COLOR_RED
    p = tfb.add_paragraph()
    p.text = "PII-ის გაფილტვრა იცავს კომპანიას სამართლებრივი პასუხისმგებლობისგან, ხოლო Prefill / Schema ტექნიკა უზრუნველყოფს 100%-იან ტექნიკურ საიმედოობასა და ავტომატიზაციას."
    p.font.name = FONT_ARIAL; p.font.size = Pt(13); p.font.color.rgb = COLOR_MUTED; p.space_before = Pt(4)

    # -------------------------------------------------------------------------
    # SLIDE 9: Ecosystem: PII Masking & Data Sanitization Tools (2026)
    # -------------------------------------------------------------------------
    s9 = created_slides[7]
    add_header(s9, "მონაცემთა დაცვა და PII Masking ინსტრუმენტების რუკა (2026)", "ᲘᲜᲡᲢᲠᲣᲛᲔᲜᲢᲔᲑᲘ • ანონიმიზაცია, Gateway და უსაფრთხოების ფილტრები")

    mask_s9 = [
        ("🛡️ OPEN SOURCE PII MASKING", "ანონიმიზატორები", "მონაცემთა ფილტრაცია", [
            ("Microsoft Presidio", "ღია კოდის ინდუსტრიული სტანდარტი PII-ის ავტომატური დეტექციისა და ანონიმიზაციისთვის."),
            ("Spacy / HuggingFace NER", "დარგობრივი ენობრივი მოდელები სენსიტიური სახელებისა და ლოკაციების ამოსაცნობად."),
            ("Custom Regex Matchers", "სპეციალური რეგულარული გამოსახულებები ქართული პირადი ნომრებისა და IBAN-ისთვის.")
        ]),
        ("🌐 AI GATEWAYS & PROXIES", "უსაფრთხოების კარიბჭეები", "პრომპტის შემოწმება", [
            ("Portkey / Cloudflare Gateway", "პრომპტის რეალურ დროში შემოწმება PII-ზე, ტრაფიკის ლოგირება და ქეშირება."),
            ("Lakera Guard", "API ფილტრი, რომელიც ბლოკავს სენსიტიურ მონაცემებსა და Prompt Injection შეტევებს."),
            ("Parea AI / LiteLLM Proxy", "მრავალპროვაიდერიანი კარიბჭე მონაცემთა დაცვის ერთიანი პოლიტიკით.")
        ]),
        ("📋 ENTERPRISE COMPLIANCE", "კორპორაციული სტანდარტები", "რეგულაციები და სინთეზი", [
            ("Skyflow / Private AI", "საწარმოო დონის სერვისები GDPR, HIPAA და EU AI Act შესაბამისობის სრული გარანტიით."),
            ("Synthetic Data Vault (SDV)", "რეალური ბაზების გარდაქმნა უსაფრთხო სინთეზურ სატესტო მონაცემებად."),
            ("Audit Logging", "ყველა გაგზავნილი და მიღებული ტოკენის კრიპტოგრაფიული აუდიტის ლოგი.")
        ])
    ]

    for i, (badge, title, subtitle, items) in enumerate(mask_s9):
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
    p = tfb.paragraphs[0]; p.text = "💡 ინდუსტრიული ეკოსისტემა:"
    p.font.name = FONT_ARIAL; p.font.size = Pt(14); p.font.bold = True; p.font.color.rgb = COLOR_RED
    p = tfb.add_paragraph()
    p.text = "საწარმოო AI სისტემებში არცერთი პრომპტი არ მიდის მოდელამდე პირდაპირ — ის გაივლის Presidio/Gateway ფილტრს, სადაც PII იცვლება უსაფრთხო ტოკენებით, ხოლო პასუხში ხდება უკუჩანაცვლება."
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
    filename = "2.2 - კონტექსტის მართვა, PII შეზღუდვები და Prefill ტექნიკა.pptx"
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
    build_lecture_m7_2_2()
