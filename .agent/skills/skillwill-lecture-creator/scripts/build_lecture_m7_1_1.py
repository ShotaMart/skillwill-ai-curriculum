import os
import sys
import shutil
import pptx
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN
from pptx.enum.shapes import MSO_SHAPE

def build_lecture_m7_1_1():
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
            p.text = "1.1 - რა არის Prompt Engineering და რატომ სჭირდება დეველოპერს ანალიტიკოსს"
            p.alignment = PP_ALIGN.CENTER
            p.font.name = FONT_ARIAL
            p.font.size = Pt(32)
            p.font.bold = True
            p.font.color.rgb = COLOR_RED

            p2 = tf.add_paragraph()
            p2.text = "მოდელის მენტალური მოდელი, Context Window, ტოკენიზაცია და Software 2.0 პარადიგმა"
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
    def add_header(slide, title, category="ᲛᲝᲓᲣᲚᲘ 7 • PROMPT ENGINEERING • 1.1"):
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
    # SLIDE 2: Paradigm Shift (Software 1.0 vs Software 2.0 / 3.0)
    # -------------------------------------------------------------------------
    s2 = created_slides[0]
    add_header(s2, "პარადიგმის ცვლილება: ტრადიციული პროგრამირება vs Prompt Engineering", "ᲞᲐᲠᲐᲓᲘᲒᲛᲐ • SOFTWARE 1.0 VS SOFTWARE 2.0 / 3.0")

    cards_s2 = [
        ("💻 TRADITIONAL", "Software 1.0 (კლასიკური კოდი)", "ექსპლიციტური ალგორითმები", [
            ("დეტერმინისტული ლოგიკა", "მკაცრად განსაზღვრული წესები (If/Else, Loops, Switch) — ყოველი ეტაპი წინასწარ გაწერილია."),
            ("კოდის მოცულობა და ხარჯი", "თითოეული ახალი ქეისისა თუ გამონაკლისისთვის საჭიროა ახალი ხელით დაწერილი ლოგიკა."),
            ("სისუსტე ბუნებრივ ენასთან", "კლასიკური წესებით ურთულესია არასტრუქტურირებული ტექსტის, განწყობის ან კონტექსტის დამუშავება.")
        ]),
        ("🧠 PROMPT ENGINEERING", "Software 2.0 / 3.0 (AI მოდელები)", "ბუნებრივი ენა როგორც კოდი", [
            ("პრობაბილისტური გამოთვლა", "მოდელი ითვლის სტატისტიკურ ალბათობებსა და პატერნებს, ნაცვლად ხისტი წესებისა."),
            ("დეკლარაციული მიდგომა", "თქვენ აღწერთ „რა გსურთ რომ მიიღოთ“ (სასურველი შედეგი) და არა „როგორ უნდა გაკეთდეს“ ნაბიჯ-ნაბიჯ."),
            ("ადაპტურობა", "ერთი და იგივე პრომპტი ადვილად უმკლავდება ათასობით სხვადასხვა ენობრივ ვარიაციასა და სინონიმს.")
        ]),
        ("🚀 PROFESSIONAL VALUE", "საინჟინრო უპირატესობა", "დეველოპერი და ანალიტიკოსი", [
            ("NLP ამოცანების რევოლუცია", "ტექსტის კლასიფიკაცია, ექსტრაქცია და სინთეზი იწერება რამდენიმე წამში მარტივი API-ით."),
            ("პროტოტიპირების სისწრაფე", "ახალი იდეებისა და ფუნქციონალის შემოწმება ხდება საათებში კვირების ნაცვლად."),
            ("ჰიბრიდული სისტემები", "კლასიკური პროგრამირება (1.0) და AI პრომპტინგი (2.0) ერთიანდება მოქნილ აპლიკაციებში.")
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
    p.text = "Prompt Engineering არ არის უბრალოდ „ჩატში საუბარი“ — ეს არის მაღალი დონის ინტერფეისი მილიარდობით პარამეტრიან ნეირონულ ქსელთან, სადაც თქვენი ინგლისური თუ ქართული ინსტრუქცია კომპილირდება ალბათურ ფუნქციად."
    p.font.name = FONT_ARIAL; p.font.size = Pt(13); p.font.color.rgb = COLOR_MUTED; p.space_before = Pt(4)

    # -------------------------------------------------------------------------
    # SLIDE 3: Model Mental Model: What LLM Actually Sees
    # -------------------------------------------------------------------------
    s3 = created_slides[1]
    add_header(s3, "მოდელის მენტალური მოდელი: რას „ხედავს“ LLM სინამდვილეში?", "ᲐᲜᲐᲢᲝᲛᲘᲐ • მოდელის აღქმა და მონაცემთა შიდა რეპრეზენტაცია")

    anatomy_s3 = [
        ("Tokenization (ტოკენიზაცია)", "სიტყვები როგორც რიცხვები", [
            ("Sub-word Chunks", "ტექსტი იყოფა ქვესიტყვებად (1 ტოკენი ≈ 4 სიმბოლოს ინგლისურად)."),
            ("ქართული დამწერლობა", "ქართული ანბანი მოითხოვს მეტ ტოკენს (1 სიტყვა = 2-5 ტოკენი)."),
            ("Token ID", "თითოეული ნაწილი გარდაიქმნება უნიკალურ რიცხვით ინდექსად.")
        ]),
        ("Vector Embeddings", "ვექტორული ჩადგმა", [
            ("High-Dim Space", "ტოკენები გადადის მრავალგანზომილებიან ვექტორებში."),
            ("სემანტიკური მანძილი", "მსგავსი მნიშვნელობის ცნებები გეომეტრიულად ახლოსაა (Cosine Similarity)."),
            ("კონტექსტურობა", "სიტყვის მნიშვნელობა იცვლება გარემომცველი ტოკენების მიხედვით.")
        ]),
        ("Context Window", "კონტექსტის ფანჯარა", [
            ("ოპერატიული მეხსიერება", "2026 წლის სტანდარტები: GPT-6 (128k-1M), Claude 5 (200k+), Gemini 4 (2M+)."),
            ("მკაცრი საზღვრები", "რაც ფანჯრის მიღმა რჩება, მოდელისთვის უბრალოდ არ არსებობს."),
            ("First-In / First-Out", "გრძელ დიალოგში ძველი კონტექსტი ავტომატურად იჭრება.")
        ]),
        ("Next-Token Prediction", "ალბათური პროგნოზირება", [
            ("ალბათობის გამოთვლა", "მოდელი ითვლის ლექსიკონის ყველა სიტყვის მომდევნო ალბათობას: P(W_n | W_1...n-1)."),
            ("სტატისტიკური ლოგიკა", "LLM-ს არ გააჩნია ცნობიერება — ის ეძებს ყველაზე მოსალოდნელ გაგრძელებას."),
            ("სტოქასტური შედეგი", "ერთსა და იმავე პრომპტზე შესაძლებელია ალტერნატიული გენერაცია.")
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

        p = tf.paragraphs[0]; p.text = f"ᲡᲕᲔᲢᲘ {i+1}".upper()
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
    p = tfb.paragraphs[0]; p.text = "💡 მოდელის მენტალური მოდელი:"
    p.font.name = FONT_ARIAL; p.font.size = Pt(14); p.font.bold = True; p.font.color.rgb = COLOR_RED
    p = tfb.add_paragraph()
    p.text = "LLM არ კითხულობს ტექსტს ადამიანის მსგავსად — ის ხედავს რიცხვების ნაკადს და ითვლის ყურადღების (Attention) მატრიცებს, რათა განსაზღვროს, რომელი სიტყვა უნდა მოჰყვეს წინას კონკრეტულ კონტექსტში."
    p.font.name = FONT_ARIAL; p.font.size = Pt(13); p.font.color.rgb = COLOR_MUTED; p.space_before = Pt(4)

    # -------------------------------------------------------------------------
    # SLIDE 4: Processing Cycle & Attention Mechanism
    # -------------------------------------------------------------------------
    s4 = created_slides[2]
    add_header(s4, "Prompt-ის დამუშავების ციკლი და ყურადღების (Attention) ნაკადი", "ᲛᲔᲥᲐᲜᲘᲖᲛᲘ • PROMPT PIPELINE და ტოკენების გენერაცია")

    # Left: Process Pipeline
    b1 = s4.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(1.0), box_y, box_w, box_h)
    b1.fill.solid(); b1.fill.fore_color.rgb = COLOR_BG_LIGHT
    b1.line.color.rgb = COLOR_BORDER; b1.line.width = Pt(1.5)

    tb1 = s4.shapes.add_textbox(Inches(1.4), box_y + Inches(0.4), box_w - Inches(0.8), box_h - Inches(0.8))
    tf1 = tb1.text_frame; tf1.word_wrap = True; tf1.margin_left = tf1.margin_top = 0
    p = tf1.paragraphs[0]; p.text = "PROMPT INGESTION & ATTENTION"
    p.font.name = FONT_ARIAL; p.font.size = Pt(12); p.font.bold = True; p.font.color.rgb = COLOR_MUTED

    p = tf1.add_paragraph(); p.text = "დამუშავების 4 ეტაპი"
    p.font.name = FONT_ARIAL; p.font.size = Pt(21); p.font.bold = True; p.font.color.rgb = COLOR_DARK; p.space_before = Pt(4)

    proc_points = [
        ("1. Token Ingestion", "მომხმარებლის პრომპტი იყოფა ტოკენებად და ერთიანდება სისტემურ ინსტრუქციასთან (System Prompt)."),
        ("2. Self-Attention Calculation", "თვით-ყურადღების მექანიზმი ითვლის კავშირებს (Q, K, V) პრომპტში არსებულ ყველა სიტყვას შორის."),
        ("3. Deep Layer Transformation", "ტრანსფორმერის ათეულობით ფენაში ყალიბდება სემანტიკური აზრი, ლოგიკა და კონტექსტი."),
        ("4. Auto-regressive Generation", "გენერირდება თითო ახალი ტოკენი, რომელიც ემატება კონტექსტს და პროცესი ციკლურად გრძელდება.")
    ]
    for t, d in proc_points:
        p = tf1.add_paragraph(); p.text = f"• {t}: {d}"
        p.font.name = FONT_ARIAL; p.font.size = Pt(13); p.font.color.rgb = COLOR_TEXT_MUTED; p.space_before = Pt(8)

    # Right: Stochasticity & Sampling
    b2 = s4.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(9.65), box_y, box_w, box_h)
    b2.fill.solid(); b2.fill.fore_color.rgb = COLOR_CARD_RED
    b2.line.color.rgb = COLOR_BORDER_RED; b2.line.width = Pt(1.5)

    tb2 = s4.shapes.add_textbox(Inches(10.05), box_y + Inches(0.4), box_w - Inches(0.8), box_h - Inches(0.8))
    tf2 = tb2.text_frame; tf2.word_wrap = True; tf2.margin_left = tf2.margin_top = 0
    p = tf2.paragraphs[0]; p.text = "DETERMINISM VS STOCHASTICITY"
    p.font.name = FONT_ARIAL; p.font.size = Pt(12); p.font.bold = True; p.font.color.rgb = COLOR_RED

    p = tf2.add_paragraph(); p.text = "ალბათური შერჩევა & პარამეტრები"
    p.font.name = FONT_ARIAL; p.font.size = Pt(21); p.font.bold = True; p.font.color.rgb = COLOR_DARK; p.space_before = Pt(4)

    stoch_points = [
        ("Logits & Softmax", "მოდელი აგენერირებს ლექსიკონის ყველა სიტყვის ალბათობას (Softmax გარდაქმნით)."),
        ("Temperature (ტემპერატურა)", "0.0-ზე მოდელი ყოველთვის ირჩევს უმაღლეს ალბათობას (Argmax); მაღალზე — არჩევანი მრავალფეროვანია."),
        ("Top-P & Top-K Truncation", "ალბათური ფილტრები ჭრის ნაკლებად სავარაუდო სიტყვების კუდს, რაც იცავს ტექსტს ქაოსისგან."),
        ("Attention Dilution", "ზედმეტად გრძელ და ბუნდოვან პრომპტში ყურადღების წონები იფანტება (Needle-in-a-Haystack ეფექტი).")
    ]
    for t, d in stoch_points:
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
    p = tfb.paragraphs[0]; p.text = "💡 საინჟინრო წესი:"
    p.font.name = FONT_ARIAL; p.font.size = Pt(14); p.font.bold = True; p.font.color.rgb = COLOR_RED
    p = tfb.add_paragraph()
    p.text = "პრომპტის თითოეული სიტყვა ცვლის ყურადღების (Attention) მთელ მათემატიკურ მატრიცას — არაზუსტი ან ზედმეტი სიტყვა მოდელის ყურადღებას მცდარი ალბათობისკენ მიმართავს."
    p.font.name = FONT_ARIAL; p.font.size = Pt(13); p.font.color.rgb = COLOR_MUTED; p.space_before = Pt(4)

    # -------------------------------------------------------------------------
    # SLIDE 5: Ecosystem: Frontier Models Comparison (2026)
    # -------------------------------------------------------------------------
    s5 = created_slides[3]
    add_header(s5, "მოდელების ხასიათი: OpenAI, Anthropic და Google 2026 წლის რეალობაში", "ᲔᲙᲝᲡᲘᲡᲢᲔᲛᲐ • წამყვანი LLM არქიტექტურების რეაქცია ინსტრუქციებზე")

    eco_s5 = [
        ("OpenAI GPT-6 (Sol & Astra)", "OpenAI Ecosystem", [
            ("ქცევის სტილი", "მკაცრად მიჰყვება სისტემურ დირექტივებს, ძლიერია კოდირებასა და ლოგიკურ მსჯელობაში."),
            ("პრომპტინგის სტილი", "მკაფიო როლი, ნაბიჯ-ნაბიჯ შეზღუდვები და Structured Outputs (JSON Schema)."),
            ("Integrated Reasoning", "ჩაშენებული რეგულირებადი მსჯელობის ბიუჯეტი (Thinking Effort) კომპლექსური ამოცანებისთვის.")
        ]),
        ("Anthropic Claude 5.x", "Anthropic Ecosystem", [
            ("ქცევის სტილი", "გამორჩეული ლიტერატურული ენა, მაღალი ინტელექტი, შესანიშნავად იცავს კონტექსტს."),
            ("პრომპტინგის სტილი", "XML ტეგების იერარქია (<context>, <instructions>) და Prefill Assistant Response-ის მხარდაჭერა."),
            ("Needle-in-a-Haystack", "საუკეთესო სიზუსტე გრძელ კონტექსტში (200,000+ ტოკენი) დამალული დეტალების პოვნაში.")
        ]),
        ("Google Gemini (Gemini 4 / 3.8 Flash)", "Google Ecosystem", [
            ("ქცევის სტილი", "Native Multimodal ბუნება — ტექსტს, ვიდეოკადრებს, აუდიოსა და კოდს აღიქვამს ერთიან ენაზე."),
            ("პრომპტინგის სტილი", "სისტემური ინსტრუქციები AI Studio-ში, Grounding with Google Search ფაქტების გადასამოწმებლად."),
            ("Massive Context (2M+)", "ერთდროულად ამუშავებს მთლიან საპროგრამო რეპოზიტორიას, წიგნებს ან საათობრივ ვიდეოს.")
        ])
    ]

    for i, (title, subtitle, items) in enumerate(eco_s5):
        bg_col = COLOR_CARD_RED if i == 1 else COLOR_BG_LIGHT
        bor_col = COLOR_BORDER_RED if i == 1 else COLOR_BORDER
        c = s5.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, xs_3[i], card_y, card_w, card_h)
        c.fill.solid(); c.fill.fore_color.rgb = bg_col
        c.line.color.rgb = bor_col; c.line.width = Pt(1.5)

        tb = s5.shapes.add_textbox(xs_3[i] + Inches(0.35), card_y + Inches(0.35), card_w - Inches(0.7), card_h - Inches(0.7))
        tf = tb.text_frame; tf.word_wrap = True; tf.margin_left = tf.margin_top = 0

        p = tf.paragraphs[0]; p.text = f"ᲝᲯᲐᲮᲘ {i+1}".upper()
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
    p = tfb.paragraphs[0]; p.text = "💡 მოდელების დივერსიფიკაცია:"
    p.font.name = FONT_ARIAL; p.font.size = Pt(14); p.font.bold = True; p.font.color.rgb = COLOR_RED
    p = tfb.add_paragraph()
    p.text = "არ არსებობს ერთი უნივერსალური პრომპტი — Claude საუკეთესოდ მუშაობს XML ტეგებით, GPT-6 ითხოვს მკაცრ ლოგიკურ დირექტივებს, ხოლო Gemini ბრწყინავს გრძელ მულტიმოდალურ კონტექსტთან."
    p.font.name = FONT_ARIAL; p.font.size = Pt(13); p.font.color.rgb = COLOR_MUTED; p.space_before = Pt(4)

    # -------------------------------------------------------------------------
    # SLIDE 6: Rules & Failure Modes: Best Practices & Risks
    # -------------------------------------------------------------------------
    s6 = created_slides[4]
    add_header(s6, "ეფექტური პრომპტინგის ოქროს წესები და კრიტიკული რისკები", "ᲬᲔᲡᲔᲑᲘ ᲓᲐ ᲠᲘᲡᲙᲔᲑᲘ • სიზუსტის სტანდარტები, ჰალუცინაციები და უსაფრთხოება")

    rules_s6 = [
        ("🎯 BEST PRACTICES", "ოქროს წესები", "საინჟინრო დისციპლინა", [
            ("აქტიური მოქმედების ზმნები", "გამოიყენეთ მკაფიო დირექტივები („დააკლასიფიცირე“, „ამოიღე მონაცემები“, „შეადარე“)."),
            ("ფორმატის კონტროლი", "ყოველთვის წინასწარ განსაზღვრეთ სტრუქტურა (Markdown ცხრილი, JSON Schema, ბულეტები)."),
            ("უარყოფითი შეზღუდვები", "მკაფიოდ გაწერეთ რა არ უნდა გააკეთოს მოდელმა (მაგ. „არ გამოიყენო შესავალი ტექსტი“).")
        ]),
        ("⚠️ FAILURE MODES", "კრიტიკული რისკები", "როდის ცდება მოდელი", [
            ("Hallucination (ჰალუცინაცია)", "მოდელი იგონებს დამაჯერებელ, მაგრამ არარსებულ ფაქტებს, ციტატებსა და თარიღებს."),
            ("Prompt Drift (გადახრა)", "გრძელ მრავალრაუნდიან დიალოგში მოდელი თანდათან კარგავს საწყისი ინსტრუქციის წესებს."),
            ("Over-constraining", "ურთიერთგამომრიცხავი წესების სიმრავლე, რაც იწვევს მოდელის ლოგიკურ დაბნევას.")
        ]),
        ("🔒 DATA PRIVACY", "მონაცემთა დაცვა", "კორპორაციული უსაფრთხოება", [
            ("PII და კონფიდენციალობა", "არასოდეს გაგზავნოთ პაროლები, საკრედიტო ბარათები ან კომპანიის საიდუმლო ინფორმაცია."),
            ("Prompt Leakage", "მოდელის მიერ შიდა სისტემური ინსტრუქციის (System Prompt) მომხმარებელზე გაცემის რისკი."),
            ("Prompt Injection", "მესამე მხარის მავნე ტექსტის მიერ მოდელის კონტროლის ხელში ჩაგდება.")
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
    p.text = "რაც უფრო ბუნდოვანია პრომპტი, მით მეტ თავისუფლებას იღებს ალბათური მოდელი — საინჟინრო სიზუსტე იწყება მკაცრი საზღვრებისა და ფორმატის წინასწარი დაწესებით."
    p.font.name = FONT_ARIAL; p.font.size = Pt(13); p.font.color.rgb = COLOR_MUTED; p.space_before = Pt(4)

    # -------------------------------------------------------------------------
    # SLIDE 7: Developer Workspaces: Playgrounds & Web IDEs
    # -------------------------------------------------------------------------
    s7 = created_slides[5]
    add_header(s7, "სად მუშაობს პრომპტ ინჟინერი? Developer Playgrounds & Web IDEs", "ᲡᲐᲛᲣᲨᲐᲝ ᲒᲐᲠᲔᲛᲝ • GOOGLE AI STUDIO, OPENAI WORKBENCH და ANTHROPIC CONSOLE")

    env_s7 = [
        ("Google AI Studio", "ai.google.dev", [
            ("სამუშაო რეჟიმები", "Chat Prompt (დიალოგი), Freeform Prompt (თავისუფალი ტექსტი), Structured Prompt (ცხრილური I/O)."),
            ("პარამეტრების მართვა", "Temperature, Top-P, Top-K, Safety Thresholds და System Instructions."),
            ("Get Code ექსპორტი", "ერთი დაწკაპუნებით მზა კოდი Python, JavaScript, cURL და Swift ენებზე.")
        ]),
        ("OpenAI Playground / Workbench", "platform.openai.com", [
            ("მოდელების მართვა", "GPT-6 (Sol, Astra), Thinking Effort კონტროლი და სისტემური როლის განსაზღვრა."),
            ("Structured Outputs", "JSON Schema-ს მკაცრი ვალიდაცია, რათა მოდელმა გამოიტანოს 100%-ით ვალიდური მონაცემები."),
            ("პარამეტრები", "Stop Sequences, Frequency და Presence Penalty განმეორებების შესამცირებლად.")
        ]),
        ("Anthropic Console", "console.anthropic.com", [
            ("Prompt Generator", "ბუნებრივ ენაზე აღწერილი ამოცანის ავტომატური გარდაქმნა მაღალი დონის Claude პრომპტად."),
            ("Prefill Assistant Response", "მოდელის პასუხის პირველი სიმბოლოების კარნახი (მაგ. {\"status\":), რაც ფორმატის გარანტიაა."),
            ("Prompt Evaluator", "სხვადასხვა ვერსიების ავტომატური შედარება სატესტო ნაკრებზე (Test Suites).")
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
    p = tfb.paragraphs[0]; p.text = "💡 პროფესიული მიდგომა:"
    p.font.name = FONT_ARIAL; p.font.size = Pt(14); p.font.bold = True; p.font.color.rgb = COLOR_RED
    p = tfb.add_paragraph()
    p.text = "პრომპტ ინჟინრები არ მუშაობენ სამომხმარებლო ჩატში (ChatGPT/Claude App) — ნამდვილი ტესტირება და ოპტიმიზაცია ხდება Developer Playground-ში, სადაც კონტროლდება პარამეტრები და System Message."
    p.font.name = FONT_ARIAL; p.font.size = Pt(13); p.font.color.rgb = COLOR_MUTED; p.space_before = Pt(4)

    # -------------------------------------------------------------------------
    # SLIDE 8: Experiment: Vague Prompt vs Engineered Prompt
    # -------------------------------------------------------------------------
    s8 = created_slides[6]
    add_header(s8, "პრაქტიკული შედარება: ბუნდოვანი vs საინჟინრო პრომპტი", "ᲔᲥᲡᲞᲔᲠᲘᲛᲔᲜᲢᲘ • INPUT / OUTPUT ანალიზი და მოდელების რეაქცია")

    # Top container: Side-by-side comparison box
    top_box_y = Inches(2.35)
    top_box_h = Inches(3.35)

    # Vague prompt box
    b_vague = s8.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(1.0), top_box_y, Inches(8.35), top_box_h)
    b_vague.fill.solid(); b_vague.fill.fore_color.rgb = COLOR_BG_LIGHT
    b_vague.line.color.rgb = COLOR_BORDER; b_vague.line.width = Pt(1.5)

    tb_v = s8.shapes.add_textbox(Inches(1.35), top_box_y + Inches(0.25), Inches(7.65), top_box_h - Inches(0.5))
    tf_v = tb_v.text_frame; tf_v.word_wrap = True; tf_v.margin_left = tf_v.margin_top = 0
    p = tf_v.paragraphs[0]; p.text = "❌ ᲑᲣᲜᲓᲝᲕᲐᲜᲘ ᲛᲝᲗᲮᲝᲕᲜᲐ (VAGUE PROMPT)"
    p.font.name = FONT_ARIAL; p.font.size = Pt(11); p.font.bold = True; p.font.color.rgb = COLOR_MUTED

    p = tf_v.add_paragraph(); p.text = "„დამიწერე შეჯამება ამ ტექსტის და მითხარი რა გავაკეთო.“"
    p.font.name = FONT_ARIAL; p.font.size = Pt(15); p.font.bold = True; p.font.color.rgb = COLOR_DARK; p.space_before = Pt(4)

    vague_flaws = [
        ("როლის არარსებობა", "მოდელი არ იცის, ვის პოზიციიდან უნდა უპასუხოს (იურისტი, ინჟინერი, სტუდენტი)."),
        ("კონტექსტის დეფიციტი", "არ არის დაკონკრეტებული აუდიტორია, მიზანი და ტექსტის პრიორიტეტები."),
        ("არაპროგნოზირებადი ფორმატი", "შედეგი: ზოგადი, ზედაპირული ესე, გაურკვეველი ნაბიჯები და ჭარბი შესავალი/დასკვნა.")
    ]
    for h, d in vague_flaws:
        p = tf_v.add_paragraph(); p.text = f"• {h}: {d}"
        p.font.name = FONT_ARIAL; p.font.size = Pt(12); p.font.color.rgb = COLOR_TEXT_MUTED; p.space_before = Pt(4)

    # Engineered prompt box
    b_eng = s8.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(9.65), top_box_y, Inches(8.35), top_box_h)
    b_eng.fill.solid(); b_eng.fill.fore_color.rgb = COLOR_CARD_RED
    b_eng.line.color.rgb = COLOR_BORDER_RED; b_eng.line.width = Pt(1.5)

    tb_e = s8.shapes.add_textbox(Inches(10.0), top_box_y + Inches(0.25), Inches(7.65), top_box_h - Inches(0.5))
    tf_e = tb_e.text_frame; tf_e.word_wrap = True; tf_e.margin_left = tf_e.margin_top = 0
    p = tf_e.paragraphs[0]; p.text = "✅ ᲡᲐᲘᲜᲟᲘᲜᲠᲝ ᲞᲠᲝᲛᲞᲢᲘ (ENGINEERED PROMPT)"
    p.font.name = FONT_ARIAL; p.font.size = Pt(11); p.font.bold = True; p.font.color.rgb = COLOR_RED

    p = tf_e.add_paragraph(); p.text = "„როლი: Senior Data Architect. ამოცანა: გააანალიზე სერვერის ლოგები.“"
    p.font.name = FONT_ARIAL; p.font.size = Pt(15); p.font.bold = True; p.font.color.rgb = COLOR_DARK; p.space_before = Pt(4)

    eng_strengths = [
        ("მკაფიო სტრუქტურა", "ფორმატი: Markdown ცხრილი (Timestamp, Error Type, Severity)."),
        ("ხისტი შეზღუდვა", "გამოიტანე მხოლოდ Critical და Error სტატუსები. არ გამოიყენო შესავალი."),
        ("დეტერმინისტული შედეგი", "100%-ით ზუსტი ცხრილი, ნულოვანი ზედმეტი ტექსტი, მზად მონაცემთა ბაზაში ჩასაწერად.")
    ]
    for h, d in eng_strengths:
        p = tf_e.add_paragraph(); p.text = f"⭐ {h}: {d}"
        p.font.name = FONT_ARIAL; p.font.size = Pt(12); p.font.bold = (h == "დეტერმინისტული შედეგი"); p.font.color.rgb = COLOR_DARK; p.space_before = Pt(4)

    # Middle Model Cards (3 cards)
    mid_y = Inches(5.95)
    mid_h = Inches(1.8)
    models_exp = [
        ("OpenAI GPT-6", "openai.png", "100% სიზუსტე ცხრილის დაცვაში, ნულოვანი ჰალუცინაცია, მყისიერი სტრუქტურირება."),
        ("Claude 5.5 Sonnet", "claude.png", "სრულყოფილი სინტაქსი, Markdown ცხრილი იდეალურად ფორმატირებული და წაკითხვადი."),
        ("Gemini 2.5 Pro", "gemini.png", "მყისიერი დამუშავება, ლოგების ანომალიების ზუსტი გამოყოფა და კორელაცია.")
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
    p.text = "როლი + მკაფიო ამოცანა + ფორმატის შეზღუდვა = დეტერმინისტული და საიმედო გამოსავალი ნებისმიერ წამყვან მოდელში. რაც უფრო ზუსტია შეზღუდვა, მით უფრო დაცულია სისტემა ჰალუცინაციისგან."
    p.font.name = FONT_ARIAL; p.font.size = Pt(13); p.font.color.rgb = COLOR_MUTED; p.space_before = Pt(4)

    # -------------------------------------------------------------------------
    # SLIDE 9: Ecosystem: Prompt Engineering & LLMOps Map (2026)
    # -------------------------------------------------------------------------
    s9 = created_slides[7]
    add_header(s9, "Prompt Engineering და LLMOps ინსტრუმენტების რუკა (2026)", "ᲘᲜᲡᲢᲠᲣᲛᲔᲜᲢᲔᲑᲘ • პროტოტიპირება, ვერსიონირება და ევალუაცია")

    ops_s9 = [
        ("🛠️ PROTOTYPING & PLAYGROUNDS", "პროტოტიპირების სტეკი", "პრომპტის შედგენა და API", [
            ("Google AI Studio", "Gemini მოდელების პროტოტიპირება, მულტიმოდალური ტესტირება და API კოდის გენერაცია."),
            ("Anthropic Console", "Claude მოდელების პრომპტინგი, XML სტრუქტურის ვალიდაცია და Prompt Generator."),
            ("OpenAI Workbench", "GPT-6 ტესტირება, Structured Outputs (JSON Schema) და პარამეტრების მართვა.")
        ]),
        ("📊 EVALUATION & TESTING", "ევალუაცია და ტესტირება", "ავტომატური ხარისხის კონტროლი", [
            ("Promptfoo", "CLI და CI/CD ინსტრუმენტი პრომპტების ავტომატური ტესტირებისა და უსაფრთხოების აუდიტისთვის."),
            ("LangSmith / LangChain", "პრომპტების გაშვების ტრეისინგი, დებაგინგი და მონიტორინგი რეალურ დროში."),
            ("Weights & Biases Prompts", "ექსპერიმენტების ვერსიონირება, მოდელების შედარება და მეტრიკების ვიზუალიზაცია.")
        ]),
        ("⚙️ PRODUCTION & GOVERNANCE", "საწარმოო მართვა", "უსაფრთხოება და ორკესტრაცია", [
            ("Portkey / AI Gateway", "პრომპტების ქეშირება, ტრაფიკის როუტინგი და საიმედოობის უზრუნველყოფა (Failover)."),
            ("DSPy", "პრომპტების ავტომატური ოპტიმიზაცია და კომპილაცია პროგრამული ალგორითმებით."),
            ("Lakera Gandalf", "Red-Teaming და Prompt Injection შეტევების სიმულაცია უსაფრთხოების შესამოწმებლად.")
        ])
    ]

    for i, (badge, title, subtitle, items) in enumerate(ops_s9):
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
    p.text = "Prompt Engineering 2026 წელს გასცდა „ხელოვნების“ ჩარჩოებს და იქცა პროგრამულ ინჟინერიად (LLMOps) — ტესტირების, ვერსიონირებისა და ავტომატური ოპტიმიზაციის სრული ინსტრუმენტული ჯაჭვით."
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
    filename = "1.1 - რა არის Prompt Engineering და რატომ სჭირდება დეველოპერს ანალიტიკოსს.pptx"
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
    build_lecture_m7_1_1()
