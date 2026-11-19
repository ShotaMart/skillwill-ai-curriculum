import os
import sys
import shutil
import pptx
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN
from pptx.enum.shapes import MSO_SHAPE

def build_lecture_m7_1_2():
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
            p.text = "1.2 - მკაფიო ინსტრუქციის არქიტექტურა და ანტიპატერნები"
            p.alignment = PP_ALIGN.CENTER
            p.font.name = FONT_ARIAL
            p.font.size = Pt(32)
            p.font.bold = True
            p.font.color.rgb = COLOR_RED

            p2 = tf.add_paragraph()
            p2.text = "დირექტივები, შეზღუდვები, ტიპური შეცდომები და პრომპტინგის ინდუსტრიული ანტიპატერნები"
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
    def add_header(slide, title, category="ᲛᲝᲓᲣᲚᲘ 7 • PROMPT ENGINEERING • 1.2"):
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
    # SLIDE 2: Paradigm Shift (Vague Language vs Engineering Precision)
    # -------------------------------------------------------------------------
    s2 = created_slides[0]
    add_header(s2, "ორაზროვნება ბუნებრივ ენაში vs საინჟინრო სიცხადე (Clarity)", "ᲞᲐᲠᲐᲓᲘᲒᲛᲐ • VAGUE LANGUAGE VS DETERMINISTIC DIRECTIVES")

    cards_s2 = [
        ("🗣️ HUMAN CONVERSATION", "ადამიანური ორაზროვნება", "კონტექსტი და ინტუიცია", [
            ("ნაგულისხმევი წესები", "ადამიანები ბუნდოვან ფრაზებს („ცოტა შეამოკლე“, „კარგად დაწერე“) ავსებენ საერთო გამოცდილებით."),
            ("ჟესტები და ინტონაცია", "კომუნიკაციის უდიდესი ნაწილი ეყრდნობა არავერბალურ მინიშნებებსა და სოციალურ კონტექსტს."),
            ("ტოლერანტობა შეცდომებზე", "ადამიანი ადვილად ხვდება მთქმელის განზრახვას (Intent) არასრული ინფორმაციის დროსაც.")
        ]),
        ("⚙️ LLM INTERPRETATION", "მოდელის ალბათური რისკი", "სტატისტიკური გაურკვევლობა", [
            ("ინტუიციის არარსებობა", "LLM-ს არ აქვს ცნობიერება — ბუნდოვანი ფრაზა აფართოებს შესაძლო ალბათობების სპექტრს."),
            ("ზედაპირული შედეგები", "დაუზუსტებელ მოთხოვნაზე მოდელი ირჩევს ყველაზე ზოგად, ბანალურ ინტერნეტ-კლიშეებს."),
            ("არაპროგნოზირებადობა", "ერთი და იგივე ბუნდოვანი პრომპტი ყოველ გაშვებაზე განსხვავებულ და არასაიმედო პასუხს იძლევა.")
        ]),
        ("🎯 INSTRUCTION PRECISION", "საინჟინრო სიცხადე", "დეტერმინისტული დირექტივები", [
            ("ალბათობის შევიწროება", "ზუსტი დირექტივები მკვეთრად ავიწროებს ალბათობათა განაწილებას კონკრეტულ მიზანზე."),
            ("განმეორებადობა (Reproducibility)", "კარგად სტრუქტურირებული პრომპტი იძლევა 100%-ით სტაბილურ და პროგნოზირებად პასუხს."),
            ("ავტომატიზაციის გარანტია", "მკაფიო ინსტრუქცია საშუალებას გაძლევთ AI უსაფრთხოდ ჩართოთ პროგრამულ მილსადენში (Pipeline).")
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
    p.text = "ბუნებრივი ენა თავისი ბუნებით ორაზროვანია, ხოლო ნეირონული ქსელი — ალბათური მანქანა. Prompt Engineering-ის მიზანია ბუნებრივი ენის ისეთივე მკაცრი დისციპლინით გამოყენება, როგორითაც პროგრამული კოდი იწერება."
    p.font.name = FONT_ARIAL; p.font.size = Pt(13); p.font.color.rgb = COLOR_MUTED; p.space_before = Pt(4)

    # -------------------------------------------------------------------------
    # SLIDE 3: Anatomy: 4 Pillars of Clear Instruction
    # -------------------------------------------------------------------------
    s3 = created_slides[1]
    add_header(s3, "მკაფიო ინსტრუქციის 4 საყრდენი: ზმნა, აუდიტორია, მასშტაბი და შეზღუდვა", "ᲐᲜᲐᲢᲝᲛᲘᲐ • ინსტრუქციის საკვანძო ელემენტები")

    anatomy_s3 = [
        ("Action Verbs (ზმნები)", "არაინტერპრეტირებადი მოქმედება", [
            ("კონკრეტული დირექტივა", "გამოიყენეთ მკაფიო ზმნები: „დააკლასიფიცირე“, „ამოიღე“, „გადაიყვანე JSON-ში“."),
            ("ბუნდოვანების აკრძალვა", "მოერიდეთ ისეთ ზმნებს, როგორიცაა „მიხედე“, „რამე მოიფიქრე“, „დაამუშავე“."),
            ("ფოკუსირებული მიზანი", "თითოეულ წინადადებას უნდა ჰქონდეს ერთი კონკრეტული ამოცანა.")
        ]),
        ("Target Audience (აუდიტორია)", "აღქმის დონე და ენა", [
            ("პერსპექტივის მორგება", "განსაზღვრეთ ვისთვის იქმნება პასუხი (მაგ. „C-Level მენეჯერებისთვის“, „დამწყები სტუდენტებისთვის“)."),
            ("სირთულის დონე", "აუდიტორია კარნახობს მოდელს ტერმინოლოგიის სიღრმესა და აბსტრაქციის ხარისხს."),
            ("კონტექსტური ტონი", "მეგობრული, აკადემიური, კორპორაციული თუ ტექნიკური სტილი.")
        ]),
        ("Scope & Length (მასშტაბი)", "მოცულობის ზუსტი კონტროლი", [
            ("მკაცრი მეტრიკა", "მოითხოვეთ ზუსტი მოცულობა („მაქსიმუმ 3 ბულეტი“, „ზუსტად 200 სიტყვა“, „1 აბზაცი“)."),
            ("ტოკენების ეკონომია", "სიგრძის ლიმიტი იცავს API ხარჯებს და თავიდან აცილებს ზედმეტ „წყალს“."),
            ("ინფორმაციის სიმჭიდროვე", "აიძულებს მოდელს კონცენტრირდეს მხოლოდ ყველაზე მნიშვნელოვან ფაქტებზე.")
        ]),
        ("Negative Constraints", "უარყოფითი შეზღუდვები", [
            ("საზღვრების გავლება", "მკაფიოდ გაწერეთ რა არ უნდა გააკეთოს მოდელმა („არ ახსენო კონკურენტები“)."),
            ("ფორმატის აკრძალვა", "„არ გამოიყენო შესავალი მისალმება და დასკვნითი ფრაზები“."),
            ("უსაფრთხოების ფილტრი", "ჰალუცინაციების პრევენცია („არ გამოიგონო არარსებული თარიღები“).")
        ])
    ]

    for i, (title, subtitle, items) in enumerate(anatomy_s3):
        bg_col = COLOR_CARD_RED if i == 0 else COLOR_BG_LIGHT
        bor_col = COLOR_BORDER_RED if i == 0 else COLOR_BORDER
        c = s3.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, xs_4[i], step_y, step_w, step_h)
        c.fill.solid(); c.fill.fore_color.rgb = bg_col
        c.line.color.rgb = bor_col; c.line.width = Pt(1.5)

        tb = s3.shapes.add_textbox(xs_4[i] + Inches(0.25), step_y + Inches(0.35), step_w - Inches(0.5), step_h - Inches(0.7))
        tf = tb.text_frame; tf.word_wrap = True; tf.margin_left = tf.margin_top = 0

        p = tf.paragraphs[0]; p.text = f"ᲡᲐᲧᲠᲓᲔᲜᲘ {i+1}".upper()
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
    p = tfb.paragraphs[0]; p.text = "💡 საინჟინრო ფორმულა:"
    p.font.name = FONT_ARIAL; p.font.size = Pt(14); p.font.bold = True; p.font.color.rgb = COLOR_RED
    p = tfb.add_paragraph()
    p.text = "მკაფიო საინჟინრო ინსტრუქცია = [ზუსტი მოქმედების ზმნა] + [სამიზნე აუდიტორია] + [მოცულობის ზღვარი] + [მკაცრი უარყოფითი შეზღუდვა]. ამ ოთხი კომპონენტის ერთობლიობა მინიმუმამდე ამცირებს ინტერპრეტაციის რისკს."
    p.font.name = FONT_ARIAL; p.font.size = Pt(13); p.font.color.rgb = COLOR_MUTED; p.space_before = Pt(4)

    # -------------------------------------------------------------------------
    # SLIDE 4: Mechanism: Negation Handling & Attention
    # -------------------------------------------------------------------------
    s4 = created_slides[2]
    add_header(s4, "მექანიზმი: Negation Handling (უარყოფის პრობლემა) და Attention", "ᲛᲔᲥᲐᲜᲘᲖᲛᲘ • ყურადღების განაწილება და უარყოფითი წესები")

    # Left: Pink Elephant Problem
    b1 = s4.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(1.0), box_y, box_w, box_h)
    b1.fill.solid(); b1.fill.fore_color.rgb = COLOR_BG_LIGHT
    b1.line.color.rgb = COLOR_BORDER; b1.line.width = Pt(1.5)

    tb1 = s4.shapes.add_textbox(Inches(1.4), box_y + Inches(0.4), box_w - Inches(0.8), box_h - Inches(0.8))
    tf1 = tb1.text_frame; tf1.word_wrap = True; tf1.margin_left = tf1.margin_top = 0
    p = tf1.paragraphs[0]; p.text = "THE \"PINK ELEPHANT\" PROBLEM"
    p.font.name = FONT_ARIAL; p.font.size = Pt(12); p.font.bold = True; p.font.color.rgb = COLOR_MUTED

    p = tf1.add_paragraph(); p.text = "რატომ უჭირს LLM-ს უარყოფა?"
    p.font.name = FONT_ARIAL; p.font.size = Pt(21); p.font.bold = True; p.font.color.rgb = COLOR_DARK; p.space_before = Pt(4)

    neg_points = [
        ("Attention-ის აქტივაცია", "როდესაც წერთ „არ ახსენო ვარდისფერი სპილო“, სიტყვა „სპილო“ იღებს უდიდეს ყურადღების (Attention) წონას."),
        ("ტოკენების ალბათობის ზრდა", "აკრძალული სიტყვის პრომპტში შეტანა ავტომატურად ზრდის მასთან ასოცირებული ტოკენების ალბათობას."),
        ("არაპირდაპირი დარღვევა", "მოდელი ხშირად მაინც ახსენებს იმას, რაც აუკრძალეთ, ოღონდ უარყოფის კონტექსტში („მე არ ვსაუბრობ სპილოზე...“)."),
        ("ლოგიკური დაბნევა", "რთულ მრავალდონიან უარყოფებში („ნუ გამოიყენებ არაოფიციალურ...“) შეცდომის ალბათობა ექსპონენციალურად იზრდება.")
    ]
    for t, d in neg_points:
        p = tf1.add_paragraph(); p.text = f"• {t}: {d}"
        p.font.name = FONT_ARIAL; p.font.size = Pt(13); p.font.color.rgb = COLOR_TEXT_MUTED; p.space_before = Pt(8)

    # Right: Positive Framing
    b2 = s4.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(9.65), box_y, box_w, box_h)
    b2.fill.solid(); b2.fill.fore_color.rgb = COLOR_CARD_RED
    b2.line.color.rgb = COLOR_BORDER_RED; b2.line.width = Pt(1.5)

    tb2 = s4.shapes.add_textbox(Inches(10.05), box_y + Inches(0.4), box_w - Inches(0.8), box_h - Inches(0.8))
    tf2 = tb2.text_frame; tf2.word_wrap = True; tf2.margin_left = tf2.margin_top = 0
    p = tf2.paragraphs[0]; p.text = "POSITIVE FRAMING TECHNIQUE"
    p.font.name = FONT_ARIAL; p.font.size = Pt(12); p.font.bold = True; p.font.color.rgb = COLOR_RED

    p = tf2.add_paragraph(); p.text = "პოზიტიური ჩანაცვლების წესი"
    p.font.name = FONT_ARIAL; p.font.size = Pt(21); p.font.bold = True; p.font.color.rgb = COLOR_DARK; p.space_before = Pt(4)

    pos_points = [
        ("„რა არ გააკეთოს“ ➔ „რა გააკეთოს“", "აკრძალვა გარდაქმენით სასურველ ქცევად: ნაცვლად „ნუ იქნები უხეში“ ➔ „იყავი თავაზიანი და დიპლომატიური“."),
        ("სიგრძის პოზიტიური ლიმიტი", "ნაცვლად „არ დაწერო გრძლად“ ➔ „შეზღუდე პასუხი ზუსტად 3 მოკლე წინადადებით“."),
        ("სტრუქტურული ჩარჩო", "აკრძალვის ნაცვლად მიეცით სასურველი ველების მკაცრი ჩამონათვალი (მაგ. JSON Schema ან ცხრილი)."),
        ("მაღალი სტაბილურობა", "პოზიტიური დირექტივები მიმართავს მოდელის ყურადღების ვექტორს პირდაპირ სასურველი შედეგისკენ.")
    ]
    for t, d in pos_points:
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
    p = tfb.paragraphs[0]; p.text = "💡 ოქროს წესი:"
    p.font.name = FONT_ARIAL; p.font.size = Pt(14); p.font.bold = True; p.font.color.rgb = COLOR_RED
    p = tfb.add_paragraph()
    p.text = "უთხარით მოდელს რა გააკეთოს, ნაცვლად იმისა, თუ რა არ გააკეთოს. პოზიტიურად ფორმულირებული დირექტივა ყოველთვის უფრო მაღალ ალბათურ სტაბილურობასა და ხარისხს უზრუნველყოფს."
    p.font.name = FONT_ARIAL; p.font.size = Pt(13); p.font.color.rgb = COLOR_MUTED; p.space_before = Pt(4)

    # -------------------------------------------------------------------------
    # SLIDE 5: Ecosystem: 3 Major Anti-Patterns in Prompting
    # -------------------------------------------------------------------------
    s5 = created_slides[3]
    add_header(s5, "პრომპტინგის ანტიპატერნები: Bloat, Conflicts და Format Ambiguity", "ᲔᲙᲝᲡᲘᲡᲢᲔᲛᲐ • ტიპური შეცდომების კლასიფიკაცია ინდუსტრიაში")

    anti_s5 = [
        ("Prompt Bloat (გადატვირთვა)", "ზედმეტი ქაოსური ტექსტი", [
            ("სიმპტომი", "პრომპტში ჩაყრილია 300+ სიტყვა განმეორებადი განმარტებებით, არარელევანტური ისტორიითა და „წყლით“."),
            ("მექანიზმის ეფექტი", "Attention მატრიცა იფანტება (Prompt Dilution); მოდელი კარგავს ფოკუსს მთავარ ამოცანაზე."),
            ("გამოსავალი", "პრომპტის შეკუმშვა (Pruning), მკაფიო ბულეტები და მხოლოდ არსებითი კონტექსტის დატოვება.")
        ]),
        ("Conflicting Directives", "წინააღმდეგობრივი წესები", [
            ("სიმპტომი", "ურთიერთგამომრიცხავი მოთხოვნები (მაგ. „დაწერე დეტალური ანალიზი, მაგრამ იყავი ულტრა-მოკლე 1 აბზაცში“)."),
            ("მექანიზმის ეფექტი", "მოდელი ვარდება ლოგიკურ ჩიხში (Cognitive Dissonance); პასუხი ხდება ზედაპირული და გაურკვეველი."),
            ("გამოსავალი", "წესების პრიორიტეტიზაცია და კონფლიქტური შეზღუდვების მკაფიო იერარქიაში მოქცევა.")
        ]),
        ("Format Ambiguity", "ფორმატის დეფიციტი", [
            ("სიმპტომი", "გამომავალი სტრუქტურის მიუთითებლობა (მაგ. მოთხოვნა „გააანალიზე მონაცემები“ ფორმატის გარეშე)."),
            ("მექანიზმის ეფექტი", "მოდელი თავად ირჩევს ფორმატს (ტექსტი, ესე, ბულეტები), რაც გამორიცხავს ავტომატურ პარსინგს."),
            ("გამოსავალი", "გამოსავლის ფორმატის ხისტი განსაზღვრა (Markdown ცხრილი, JSON Schema, CSV).")
        ])
    ]

    for i, (title, subtitle, items) in enumerate(anti_s5):
        bg_col = COLOR_CARD_RED if i == 1 else COLOR_BG_LIGHT
        bor_col = COLOR_BORDER_RED if i == 1 else COLOR_BORDER
        c = s5.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, xs_3[i], card_y, card_w, card_h)
        c.fill.solid(); c.fill.fore_color.rgb = bg_col
        c.line.color.rgb = bor_col; c.line.width = Pt(1.5)

        tb = s5.shapes.add_textbox(xs_3[i] + Inches(0.35), card_y + Inches(0.35), card_w - Inches(0.7), card_h - Inches(0.7))
        tf = tb.text_frame; tf.word_wrap = True; tf.margin_left = tf.margin_top = 0

        p = tf.paragraphs[0]; p.text = f"ᲐᲜᲢᲘᲞᲐᲢᲔᲠᲜᲘ {i+1}".upper()
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
    bb = s5.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(1.0), bot_y, Inches(17.0), Inches(2.0))
    bb.fill.solid(); bb.fill.fore_color.rgb = COLOR_WHITE
    bb.line.color.rgb = COLOR_BORDER; bb.line.width = Pt(1.2)

    tbb = s5.shapes.add_textbox(Inches(1.4), bot_y + Inches(0.3), Inches(16.2), Inches(1.4))
    tfb = tbb.text_frame; tfb.word_wrap = True; tfb.margin_left = tfb.margin_top = 0
    p = tfb.paragraphs[0]; p.text = "💡 ანტიპატერნის შედეგი:"
    p.font.name = FONT_ARIAL; p.font.size = Pt(14); p.font.bold = True; p.font.color.rgb = COLOR_RED
    p = tfb.add_paragraph()
    p.text = "გადატვირთვა იწვევს ყურადღების გაფანტვას, წინააღმდეგობრივი წესები — ლოგიკურ ჩავარდნას, ხოლო ფორმატის უქონლობა კლავს პროგრამულ ინტეგრაციასა და ავტომატიზაციას."
    p.font.name = FONT_ARIAL; p.font.size = Pt(13); p.font.color.rgb = COLOR_MUTED; p.space_before = Pt(4)

    # -------------------------------------------------------------------------
    # SLIDE 6: Rules: 5 Golden Standards of Instruction Architecture
    # -------------------------------------------------------------------------
    s6 = created_slides[4]
    add_header(s6, "ინსტრუქციის საინჟინრო არქიტექტურის 5 ოქროს წესი", "ᲬᲔᲡᲔᲑᲘ ᲓᲐ ᲠᲘᲡᲙᲔᲑᲘ • სტანდარტები საიმედო შედეგის გარანტიისთვის")

    rules_s6 = [
        ("📐 STRUCTURAL DELIMITERS", "სტრუქტურული გამყოფები", "ინექციისგან დაცვა", [
            ("XML ტეგები და Markdown", "გამოიყენეთ <instructions>, <context>, <rules> ან სამმაგი ბრჭყალები (\"\"\")."),
            ("როლების გამიჯვნა", "მოდელი მკაფიოდ არჩევს სად მთავრდება სისტემური წესი და სად იწყება მომხმარებლის ტექსტი."),
            ("ინექციის პრევენცია", "იცავს სისტემას იმ შემთხვევებისგან, როდესაც შემავალი მონაცემი ცდილობს წესების გადაწერას.")
        ]),
        ("🎯 EXPLICIT CONSTRAINTS", "ცხადი შეზღუდვები", "საზღვრების კონტროლი", [
            ("ინტერპრეტაციის გამორიცხვა", "წინასწარ გაწერეთ დასაშვები ენები, ტერმინოლოგია და რიცხვების სიზუსტე (მაგ. „2 ათწილადი“)."),
            ("სტილი და პერსონა", "განსაზღვრეთ ტონი: „იყავი ობიექტური, მკაცრი, ფაქტებზე ორიენტირებული“."),
            ("მოცულობის ჩარჩო", "მკაფიო ლიმიტი: „მაქსიმუმ 5 პუნქტი, თითო პუნქტი არაუმეტეს 15 სიტყვისა“.")
        ]),
        ("🛡️ FAIL-SAFE / FALLBACK", "ავარიული წესები", "ჰალუცინაციის პრევენცია", [
            ("უცნობი მონაცემის მართვა", "მიუთითეთ: „თუ ინფორმაცია დოკუმენტში არ არის, დააბრუნე: NOT_FOUND. არ გამოიცნო“."),
            ("არასრული შეკითხვა", "დაავალეთ მოდელს დაუსვას მომხმარებელს დამაზუსტებელი კითხვა, ნაცვლად თვითნებური ვარაუდისა."),
            ("სანდოობის გარანტია", "გამორიცხავს მოგონილ ფაქტებს კრიტიკულ ბიზნეს-პროცესებში.")
        ])
    ]

    for i, (badge, title, subtitle, items) in enumerate(rules_s6):
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
    p = tfb.paragraphs[0]; p.text = "💡 საიმედოობის ფორმულა:"
    p.font.name = FONT_ARIAL; p.font.size = Pt(14); p.font.bold = True; p.font.color.rgb = COLOR_RED
    p = tfb.add_paragraph()
    p.text = "Delimiters (გამყოფები) იცავს ინექციისგან, Explicit Constraints (შეზღუდვები) იცავს ქაოსისგან, ხოლო Fallback Rules იცავს ჰალუცინაციისგან. ეს არის პროფესიული პრომპტინგის სამკუთხედი."
    p.font.name = FONT_ARIAL; p.font.size = Pt(13); p.font.color.rgb = COLOR_MUTED; p.space_before = Pt(4)

    # -------------------------------------------------------------------------
    # SLIDE 7: Workspaces: Platform Approaches (Anthropic, OpenAI, Google)
    # -------------------------------------------------------------------------
    s7 = created_slides[5]
    add_header(s7, "ინსტრუქციების გამართვა: Anthropic, OpenAI და Google მიდგომები", "ᲡᲐᲛᲣᲨᲐᲝ ᲒᲐᲠᲔᲛᲝ • საუკეთესო დოკუმენტირებული პრაქტიკა 2026")

    env_s7 = [
        ("Anthropic \"Be Clear & Direct\"", "docs.anthropic.com", [
            ("პირდაპირობის პრინციპი", "Claude საუკეთესოდ რეაგირებს პირდაპირ, მკაფიო მითითებებზე და არა თავაზიან ფრაზებზე."),
            ("XML იერარქია", "ტეგების გამოყენება (<context>, <instructions>, <formatting>) ზრდის წესების შესრულების სიზუსტეს 40%-ით."),
            ("Prefill Assistant", "პასუხის დასაწყისის კარნახი უზრუნველყოფს სასურველი ფორმატის უპირობო დაცვას.")
        ]),
        ("OpenAI Instruction Hierarchy", "platform.openai.com", [
            ("მკაცრი იერარქია", "System/Developer შეტყობინება ფლობს უმაღლეს პრიორიტეტს User შეტყობინებასთან კონფლიქტისას."),
            ("Structured Outputs", "Strict Mode (JSON Schema) გარანტიას იძლევა, რომ მოდელი არ გადაუხვევს მონაცემთა სტრუქტურას."),
            ("Thinking Effort კონტროლი", "რთული ლოგიკური წესების დროს მოდელი ხარჯავს დამატებით მსჯელობის ტოკენებს.")
        ]),
        ("Google AI Studio Rules", "ai.google.dev", [
            ("System Instructions", "გლობალური როლი და უარყოფითი შეზღუდვები ცალკე სისტემურ პანელშია გამოყოფილი."),
            ("Search Grounding", "სისტემური ინსტრუქციის დაკავშირება Google-ის ცოცხალ ძიებასთან ფაქტების გადასამოწმებლად."),
            ("Zero Temperature", "Temp=0.0 და Top-P=0.95 კომბინაცია ინსტრუქციების მაქსიმალურად მკაცრი შესრულებისთვის.")
        ])
    ]

    for i, (title, subtitle, items) in enumerate(env_s7):
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
    p = tfb.paragraphs[0]; p.text = "💡 ინდუსტრიული სტანდარტი:"
    p.font.name = FONT_ARIAL; p.font.size = Pt(14); p.font.bold = True; p.font.color.rgb = COLOR_RED
    p = tfb.add_paragraph()
    p.text = "სამივე წამყვანი ლაბორატორია (Anthropic, OpenAI, Google) ერთხმად აღიარებს: რაც უფრო ნაკლებია „ლიტერატურული წყალი“ პრომპტში და მეტია სტრუქტურა, მით უფრო მაღალია სისტემის საიმედოობა."
    p.font.name = FONT_ARIAL; p.font.size = Pt(13); p.font.color.rgb = COLOR_MUTED; p.space_before = Pt(4)

    # -------------------------------------------------------------------------
    # SLIDE 8: Experiment: Anti-Pattern Refactoring (Before vs After)
    # -------------------------------------------------------------------------
    s8 = created_slides[6]
    add_header(s8, "ანტიპატერნების გამოსწორება: Before vs After რეალურ ქეისებზე", "ᲔᲥᲡᲞᲔᲠᲘᲛᲔᲜᲢᲘ • შეცდომების ტრანსფორმაცია საინჟინრო პრომპტად")

    # Top container: Side-by-side comparison box
    top_box_y = Inches(2.35)
    top_box_h = Inches(3.35)

    # Before box
    b_before = s8.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(1.0), top_box_y, Inches(8.35), top_box_h)
    b_before.fill.solid(); b_before.fill.fore_color.rgb = COLOR_BG_LIGHT
    b_before.line.color.rgb = COLOR_BORDER; b_before.line.width = Pt(1.5)

    tb_b = s8.shapes.add_textbox(Inches(1.35), top_box_y + Inches(0.25), Inches(7.65), top_box_h - Inches(0.5))
    tf_b = tb_b.text_frame; tf_b.word_wrap = True; tf_b.margin_left = tf_b.margin_top = 0
    p = tf_b.paragraphs[0]; p.text = "❌ BEFORE: ᲢᲘᲞᲣᲠᲘ ᲐᲜᲢᲘᲞᲐᲢᲔᲠᲜᲔᲑᲘ"
    p.font.name = FONT_ARIAL; p.font.size = Pt(11); p.font.bold = True; p.font.color.rgb = COLOR_MUTED

    p = tf_b.add_paragraph(); p.text = "ქაოსური, უარყოფითი და ბუნდოვანი მოთხოვნები"
    p.font.name = FONT_ARIAL; p.font.size = Pt(15); p.font.bold = True; p.font.color.rgb = COLOR_DARK; p.space_before = Pt(4)

    before_cases = [
        ("ქეისი 1 (Bloat)", "„მჭირდება ანალიზი, ძალიან გთხოვ კარგად წაიკითხე, მომეცი საუკეთესო რჩევები და არაფერი გამოგრჩეს...“"),
        ("ქეისი 2 (Negation)", "„არ ახსენო ტექნიკური დეტალები, არ დაწერო გრძლად და არაფერი მოიგონო!“"),
        ("ქეისი 3 (Ambiguity)", "„გადახედე ამ ფაილს და მითხარი რა ხდება.“")
    ]
    for h, d in before_cases:
        p = tf_b.add_paragraph(); p.text = f"• {h}: {d}"
        p.font.name = FONT_ARIAL; p.font.size = Pt(12); p.font.color.rgb = COLOR_TEXT_MUTED; p.space_before = Pt(4)

    # After box
    b_after = s8.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(9.65), top_box_y, Inches(8.35), top_box_h)
    b_after.fill.solid(); b_after.fill.fore_color.rgb = COLOR_CARD_RED
    b_after.line.color.rgb = COLOR_BORDER_RED; b_after.line.width = Pt(1.5)

    tb_a = s8.shapes.add_textbox(Inches(10.0), top_box_y + Inches(0.25), Inches(7.65), top_box_h - Inches(0.5))
    tf_a = tb_a.text_frame; tf_a.word_wrap = True; tf_a.margin_left = tf_a.margin_top = 0
    p = tf_a.paragraphs[0]; p.text = "✅ AFTER: ᲡᲐᲘᲜᲟᲘᲜᲠᲝ ᲠᲔᲤᲐᲥᲢᲝᲠᲘᲜᲒᲘ"
    p.font.name = FONT_ARIAL; p.font.size = Pt(11); p.font.bold = True; p.font.color.rgb = COLOR_RED

    p = tf_a.add_paragraph(); p.text = "სტრუქტურირებული დირექტივები და ჩარჩო"
    p.font.name = FONT_ARIAL; p.font.size = Pt(15); p.font.bold = True; p.font.color.rgb = COLOR_DARK; p.space_before = Pt(4)

    after_cases = [
        ("რეფაქტორინგი 1", "<instructions> გამოყავი 3 საკვანძო რისკი. ფორმატი: Markdown ბულეტები. </instructions>"),
        ("რეფაქტორინგი 2", "„აუდიტორია: დამწყები. ტონი: მარტივი და ცხადი. სიგრძე: ზუსტად 2 წინადადება. Fallback: თუ უცნობია, დაწერე: NOT_FOUND“."),
        ("რეფაქტორინგი 3", "„როლი: Cloud Security Architect. ამოიღე CVE ნომრები და სერვერების IP-ები Markdown ცხრილში“.")
    ]
    for h, d in after_cases:
        p = tf_a.add_paragraph(); p.text = f"⭐ {h}: {d}"
        p.font.name = FONT_ARIAL; p.font.size = Pt(12); p.font.bold = (h == "რეფაქტორინგი 3"); p.font.color.rgb = COLOR_DARK; p.space_before = Pt(4)

    # Middle Model Cards (3 cards)
    mid_y = Inches(5.95)
    mid_h = Inches(1.8)
    models_exp = [
        ("OpenAI GPT-6", "openai.png", "Before: ჭარბი ტექსტი და ბუნდოვანი რჩევები; After: 100%-ით ზუსტი ცხრილი და მკაცრი დაცვა."),
        ("Claude 5.5 Sonnet", "claude.png", "Before: უარყოფის ნაწილობრივი უგულებელყოფა; After: XML ტეგების უნაკლო დისციპლინა."),
        ("Gemini 2.5 Pro", "gemini.png", "Before: ჰალუცინაციები უცნობ დეტალებზე; After: ზუსტი და საიმედო Fallback პასუხი (NOT_FOUND).")
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
    p.text = "პრომპტის სწორი რეფაქტორინგი პასუხის ხარისხს ზრდის არა მოდელის შეცვლით, არამედ ინსტრუქციის სიცხადით, უარყოფის პოზიტიური ჩანაცვლებითა და საზღვრების დაწესებით."
    p.font.name = FONT_ARIAL; p.font.size = Pt(13); p.font.color.rgb = COLOR_MUTED; p.space_before = Pt(4)

    # -------------------------------------------------------------------------
    # SLIDE 9: Ecosystem: Prompt Linting & Testing Tools (2026)
    # -------------------------------------------------------------------------
    s9 = created_slides[7]
    add_header(s9, "პრომპტების ხარისხის კონტროლი: Prompt Linting & Testing Tools", "ᲘᲜᲡᲢᲠᲣᲛᲔᲜᲢᲔᲑᲘ • ავტომატური ანალიზი, ვალიდაცია და CI/CD")

    tools_s9 = [
        ("🔍 PROMPT LINTING", "სტატიკური ანალიზი", "შეცდომების პრევენცია", [
            ("Promptfoo CLI", "ავტომატური CLI ხელსაწყო პრომპტების ხარისხის, სიცხადისა და უსაფრთხოების შესამოწმებლად."),
            ("LangLint & Static Checks", "ავლენს ბუნდოვან ზმნებს, წინააღმდეგობრივ შეზღუდვებსა და ტოკენების ჭარბ ხარჯვას."),
            ("CI/CD ინტეგრაცია", "პრომპტის Git-ში დაფუშვამდე ავტომატური შემოწმება, როგორც კოდის Linting-ისას.")
        ]),
        ("🧪 AUTOMATED ASSERTIONS", "ავტომატური Unit ტესტები", "ხარისხის ვალიდაცია", [
            ("DeepEval", "პრომპტების Unit ტესტირება: პასუხის რელევანტურობა, ტოქსიკურობა და ჰალუცინაციის დონე."),
            ("DSPy Assertions", "პროგრამული შემოწმება: თუ მოდელი დაარღვევს ფორმატს, DSPy ავტომატურად გაუშვებს კორექციის ციკლს."),
            ("JSON Schema Compliance", "მოდელის გამოსავლის 100%-იანი სინტაქსური და სემანტიკური ვალიდაცია.")
        ]),
        ("📊 REGRESSION & LOGGING", "რეგრესიული მონიტორინგი", "ვერსიების კონტროლი", [
            ("LangSmith", "პრომპტის ყველა გაშვების დეტალური ტრეისინგი, დაყოვნების (Latency) და ხარჯების მონიტორინგი."),
            ("Arize Phoenix", "ღია კოდის პლატფორმა ევალუაციისა და RAG პრომპტების დიაგნოსტიკისთვის."),
            ("Weights & Biases Prompts", "პრომპტის ვერსიების (v1.1 ➔ v1.2) შედარება სატესტო Golden Dataset-ზე.")
        ])
    ]

    for i, (badge, title, subtitle, items) in enumerate(tools_s9):
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
    p = tfb.paragraphs[0]; p.text = "💡 ინდუსტრიული მიდგომა:"
    p.font.name = FONT_ARIAL; p.font.size = Pt(14); p.font.bold = True; p.font.color.rgb = COLOR_RED
    p = tfb.add_paragraph()
    p.text = "2026 წელს პრომპტები აღარ იწერება „თვალზომით“ — ისინი გადიან სტატიკურ Linting-ს, Unit ტესტებსა და ავტომატურ ვალიდაციას, ისევე როგორც ნებისმიერი კრიტიკული საწარმოო კოდი."
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
    filename = "1.2 - მკაფიო ინსტრუქციის არქიტექტურა და ანტიპატერნები.pptx"
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
    build_lecture_m7_1_2()
