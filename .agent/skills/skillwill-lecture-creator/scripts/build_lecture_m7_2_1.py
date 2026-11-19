import os
import sys
import shutil
import pptx
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN
from pptx.enum.shapes import MSO_SHAPE

def build_lecture_m7_2_1():
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
            p.text = "2.1 - ეფექტური პრომპტის 6 საკვანძო კომპონენტი"
            p.alignment = PP_ALIGN.CENTER
            p.font.name = FONT_ARIAL
            p.font.size = Pt(32)
            p.font.bold = True
            p.font.color.rgb = COLOR_RED

            p2 = tf.add_paragraph()
            p2.text = "პრომპტის სრული არქიტექტურა, ელემენტების იერარქია და აბლაციური ტესტირება"
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
    def add_header(slide, title, category="ᲛᲝᲓᲣᲚᲘ 7 • PROMPT ENGINEERING • 2.1"):
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
    # SLIDE 2: Paradigm Shift (Monolithic Prompt vs Modular Architecture)
    # -------------------------------------------------------------------------
    s2 = created_slides[0]
    add_header(s2, "მონოლითური პრომპტი vs 6-კომპონენტიანი საინჟინრო არქიტექტურა", "ᲞᲐᲠᲐᲓᲘᲒᲛᲐ • MONOLITHIC PROMPT VS MODULAR ARCHITECTURE")

    cards_s2 = [
        ("🧱 MONOLITHIC PROMPT", "მონოლითური ქაოსი", "გაუმართავი სტრუქტურა", [
            ("შერეული შრეები", "ამოცანა, კონტექსტი, წესები და მონაცემები ჩაწერილია ერთ უწყვეტ აბზაცში გამყოფების გარეშე."),
            ("Attention Dilution", "მოდელს უჭირს გაარჩიოს, რა არის ფონური ინფორმაცია და რა — მკაცრად შესასრულებელი ბრძანება."),
            ("რთული დებაგინგი", "როდესაც პასუხი მცდარია, თითქმის შეუძლებელია იმის დადგენა, თუ რომელმა ფრაზამ გამოიწვია შეცდომა.")
        ]),
        ("🧩 MODULAR ARCHITECTURE", "6-კომპონენტიანი ჩარჩო", "ინჟინერიული დისციპლინა", [
            ("დამოუკიდებელი მოდულები", "პრომპტი იყოფა 6 ფუნქციურ ბლოკად: როლი, მიზანი, კონტექსტი, ინსტრუქცია, შეზღუდვა, ფორმატი."),
            ("მკაფიო საზღვრები", "სტრუქტურული გამყოფები (XML/Markdown) მოდელს აძლევს მკაფიო სიგნალს თითოეული ნაწილის როლზე."),
            ("ხელახალი გამოყენებადობა", "ერთი და იგივე როლის ან შეზღუდვების ბლოკი მარტივად ინტეგრირდება სხვადასხვა ამოცანაში.")
        ]),
        ("📈 ENGINEERING IMPACT", "საინჟინრო ეფექტი", "ტესტირება და სტანდარტები", [
            ("Ablation Testing", "თითოეული კომპონენტის ამოღება/დამატება იძლევა მისი რეალური ზეგავლენის გაზომვის საშუალებას."),
            ("გუნდური სტანდარტიზაცია", "დეველოპერები და ანალიტიკოსები იყენებენ ერთიან სტრუქტურულ შაბლონს პროექტის მასშტაბით."),
            ("მზადყოფნა API-სთვის", "მოდულური პრომპტი წამებში გადადის საწარმოო კოდში (Python, TypeScript, cURL).")
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
    p.text = "6-ელემენტიანი პრომპტი წააგავს ობიექტზე ორიენტირებულ კლასს ან ფუნქციას პროგრამირებაში — მკაფიო პარამეტრებით, ტიპიზაციით, ლოგიკითა და გარანტირებული გამომავალი ტიპით (Return Type)."
    p.font.name = FONT_ARIAL; p.font.size = Pt(13); p.font.color.rgb = COLOR_MUTED; p.space_before = Pt(4)

    # -------------------------------------------------------------------------
    # SLIDE 3: Anatomy: The 6-Element Framework
    # -------------------------------------------------------------------------
    s3 = created_slides[1]
    add_header(s3, "პრომპტის 6 საკვანძო კომპონენტი (The 6-Element Framework)", "ᲐᲜᲐᲢᲝᲛᲘᲐ • პრომპტის არქიტექტურული ჩონჩხი")

    anatomy_s3 = [
        ("🎯 GOAL & ROLE", "მიზანი და როლი", "პერსონა და ფოკუსი", [
            ("1. Task / Goal (ამოცანა)", "პირდაპირი მოქმედება (Root Action) — რა უნდა გაკეთდეს პირველ რიგში („დააკლასიფიცირე“, „გააანალიზე“)."),
            ("2. Persona / Role (როლი)", "ვინ არის მოდელი (Senior Data Architect, FinOps Lead, იურისტი) — განსაზღვრავს ექსპერტიზასა და ლექსიკონს."),
            ("კოგნიტური ფოკუსი", "როლი მოდელის ცოდნის გიგანტური ბაზიდან ააქტიურებს კონკრეტულ დარგობრივ ქვესივრცეს.")
        ]),
        ("📚 CONTEXT & INSTRUCTIONS", "კონტექსტი და ინსტრუქციები", "მონაცემები და პროცესი", [
            ("3. Context (კონტექსტი)", "ფონური ინფორმაცია, კომპანიის სპეციფიკა, შემავალი დოკუმენტები და ბიზნეს-გარემო."),
            ("4. Instructions (ინსტრუქციები)", "ნაბიჯ-ნაბიჯ ალგორითმი: რა თანმიმდევრობით უნდა დამუშავდეს მოცემული ინფორმაცია."),
            ("ლოგიკური ჯაჭვი", "ინსტრუქციები მოდელს უთითებს მსჯელობის ეტაპებს (Step-by-step Execution).")
        ]),
        ("⛔ CONSTRAINTS & FORMAT", "შეზღუდვები და ფორმატი", "საზღვრები და სტრუქტურა", [
            ("5. Constraints (შეზღუდვები)", "უარყოფითი წესები, საზღვრები, აკრძალვები და Fallback ქცევა („თუ უცნობია, დაწერე NOT_FOUND“)."),
            ("6. Output Format (ფორმატი)", "გამოსავლის ხისტი სტრუქტურა (Markdown ცხრილი, JSON Schema, სიგრძის ზღვარი)."),
            ("ინტეგრაციის გარანტია", "ფორმატის კონტროლი უზრუნველყოფს პასუხის მყისიერ წაკითხვას სხვა პროგრამების მიერ.")
        ])
    ]

    for i, (badge, title, subtitle, items) in enumerate(anatomy_s3):
        bg_col = COLOR_CARD_RED if i == 0 else COLOR_BG_LIGHT
        bor_col = COLOR_BORDER_RED if i == 0 else COLOR_BORDER
        c = s3.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, xs_3[i], card_y, card_w, card_h)
        c.fill.solid(); c.fill.fore_color.rgb = bg_col
        c.line.color.rgb = bor_col; c.line.width = Pt(1.5)

        tb = s3.shapes.add_textbox(xs_3[i] + Inches(0.35), card_y + Inches(0.35), card_w - Inches(0.7), card_h - Inches(0.7))
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
    bb = s3.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(1.0), bot_y, Inches(17.0), Inches(2.0))
    bb.fill.solid(); bb.fill.fore_color.rgb = COLOR_WHITE
    bb.line.color.rgb = COLOR_BORDER; bb.line.width = Pt(1.2)

    tbb = s3.shapes.add_textbox(Inches(1.4), bot_y + Inches(0.3), Inches(16.2), Inches(1.4))
    tfb = tbb.text_frame; tfb.word_wrap = True; tfb.margin_left = tfb.margin_top = 0
    p = tfb.paragraphs[0]; p.text = "💡 არქიტექტურული წესი:"
    p.font.name = FONT_ARIAL; p.font.size = Pt(14); p.font.bold = True; p.font.color.rgb = COLOR_RED
    p = tfb.add_paragraph()
    p.text = "პრომპტიდან ნებისმიერი ერთი ელემენტის ამოღება აჩენს ხარვეზს: უროლოდ — პასუხი ზოგადია, უკონტექსტოდ — მცდარია, უინსტრუქციოდ — ქაოსურია, უშეზღუდვოდ — მოუწესრიგებელია, ხოლო უფორმატოდ — გამოუსადეგარია ავტომატიზაციაში."
    p.font.name = FONT_ARIAL; p.font.size = Pt(13); p.font.color.rgb = COLOR_MUTED; p.space_before = Pt(4)

    # -------------------------------------------------------------------------
    # SLIDE 4: Mechanism: Ablation Testing (What Breaks vs What Restores)
    # -------------------------------------------------------------------------
    s4 = created_slides[2]
    add_header(s4, "აბლაციური ტესტირება: რა ხდება, როდესაც ელემენტს ვაკლებთ?", "ᲛᲔᲥᲐᲜᲘᲖᲛᲘ • კომპონენტების ამოღებისა და დაბრუნების ეფექტი")

    # Left: What Breaks?
    b1 = s4.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(1.0), box_y, box_w, box_h)
    b1.fill.solid(); b1.fill.fore_color.rgb = COLOR_BG_LIGHT
    b1.line.color.rgb = COLOR_BORDER; b1.line.width = Pt(1.5)

    tb1 = s4.shapes.add_textbox(Inches(1.4), box_y + Inches(0.4), box_w - Inches(0.8), box_h - Inches(0.8))
    tf1 = tb1.text_frame; tf1.word_wrap = True; tf1.margin_left = tf1.margin_top = 0
    p = tf1.paragraphs[0]; p.text = "ABLATION BREAKDOWN (რა ზიანდება?)"
    p.font.name = FONT_ARIAL; p.font.size = Pt(12); p.font.bold = True; p.font.color.rgb = COLOR_MUTED

    p = tf1.add_paragraph(); p.text = "ელემენტის ამოღების ეფექტი"
    p.font.name = FONT_ARIAL; p.font.size = Pt(21); p.font.bold = True; p.font.color.rgb = COLOR_DARK; p.space_before = Pt(4)

    ab_break_points = [
        ("ამოღებულია როლი (No Role)", "მოდელი ირჩევს ზოგადსაკაცობრიო ენას, აკლია პროფესიული ტერმინოლოგია და ექსპერტული სიღრმე."),
        ("ამოღებულია კონტექსტი (No Context)", "მოდელი იძულებულია თავად გამოიცნოს ბიზნეს-რეალობა, რაც იწვევს ჰალუცინაციებს."),
        ("ამოღებულია ინსტრუქცია (No Instructions)", "მოდელი ირჩევს თვითნებურ გზას, ტოვებს კრიტიკულ ეტაპებს ან აკეთებს არასწორ აქცენტებს."),
        ("ამოღებულია შეზღუდვები (No Constraints)", "პასუხი ივსება ჭარბი შესავლით, დაუზუსტებელი ფაქტებითა და ზედმეტი ტექსტით.")
    ]
    for t, d in ab_break_points:
        p = tf1.add_paragraph(); p.text = f"• {t}: {d}"
        p.font.name = FONT_ARIAL; p.font.size = Pt(13); p.font.color.rgb = COLOR_TEXT_MUTED; p.space_before = Pt(8)

    # Right: What Restores?
    b2 = s4.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(9.65), box_y, box_w, box_h)
    b2.fill.solid(); b2.fill.fore_color.rgb = COLOR_CARD_RED
    b2.line.color.rgb = COLOR_BORDER_RED; b2.line.width = Pt(1.5)

    tb2 = s4.shapes.add_textbox(Inches(10.05), box_y + Inches(0.4), box_w - Inches(0.8), box_h - Inches(0.8))
    tf2 = tb2.text_frame; tf2.word_wrap = True; tf2.margin_left = tf2.margin_top = 0
    p = tf2.paragraphs[0]; p.text = "RESTORED STABILITY (რა აღდგება?)"
    p.font.name = FONT_ARIAL; p.font.size = Pt(12); p.font.bold = True; p.font.color.rgb = COLOR_RED

    p = tf2.add_paragraph(); p.text = "ელემენტის დაბრუნების ეფექტი"
    p.font.name = FONT_ARIAL; p.font.size = Pt(21); p.font.bold = True; p.font.color.rgb = COLOR_DARK; p.space_before = Pt(4)

    ab_restore_points = [
        ("დაბრუნებული ფორმატი (Format Added)", "მყისიერად ჩნდება მკაცრად სტრუქტურირებული Markdown ცხრილი ან ვალიდური JSON სქემა."),
        ("დაბრუნებული ინსტრუქცია (Instructions Added)", "მსჯელობის ნაბიჯ-ნაბიჯ ლოგიკა ხდება გამჭვირვალე, აუდიტირებადი და პროგნოზირებადი."),
        ("დაბრუნებული შეზღუდვა (Constraints Added)", "წყდება ჰალუცინაციები, იფილტრება ზედმეტი შესავალი და პასუხი ხდება ლაკონიური."),
        ("დაბრუნებული კონტექსტი (Context Added)", "პასუხი ზუსტად ეფუძნება კომპანიის რეალურ მონაცემებსა და ფაქტებს.")
    ]
    for t, d in ab_restore_points:
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
    p = tfb.paragraphs[0]; p.text = "💡 საინჟინრო პრინციპი:"
    p.font.name = FONT_ARIAL; p.font.size = Pt(14); p.font.bold = True; p.font.color.rgb = COLOR_RED
    p = tfb.add_paragraph()
    p.text = "აბლაციური ტესტირება (Ablation Study) საუკეთესო გზაა პრომპტის დიაგნოსტიკისთვის: თუ ელემენტის ამოღება შედეგს არ ცვლის — ის ზედმეტია; ხოლო თუ ელემენტის ამოღებით სისტემა ინგრევა — ის კრიტიკული არქიტექტურული საყრდენია."
    p.font.name = FONT_ARIAL; p.font.size = Pt(13); p.font.color.rgb = COLOR_MUTED; p.space_before = Pt(4)

    # -------------------------------------------------------------------------
    # SLIDE 5: Ecosystem: 6 Elements Support Across 2026 Frontier Models
    # -------------------------------------------------------------------------
    s5 = created_slides[3]
    add_header(s5, "6 ელემენტის მართვა წამყვან მოდელებში: OpenAI, Claude და Gemini", "ᲔᲙᲝᲡᲘᲡᲢᲔᲛᲐ • პლატფორმების მიდგომა კომპონენტების მიმართ")

    eco_s5 = [
        ("OpenAI GPT-6 (Sol & Astra)", "OpenAI Ecosystem", [
            ("System Message განაწილება", "Role და Constraints გადადის Developer/System შეტყობინებაში მაქსიმალური პრიორიტეტით."),
            ("Strict Structured Outputs", "Output Format გარანტირებულია JSON Schema-ს მეშვეობით (ნულოვანი ვალიდაციის ერორი)."),
            ("Reasoning ბიუჯეტი", "Instructions ბლოკის სირთულეს მოდელი თავად არეგულირებს Thinking Effort-ის მიხედვით.")
        ]),
        ("Anthropic Claude 5.x", "Anthropic Ecosystem", [
            ("XML სემანტიკური ტეგები", "თითოეული ელემენტი თავსდება შესაბამის ტეგში (<role>, <task>, <context>, <rules>)."),
            ("Prefill Assistant Response", "Output Format-ის უპირობო დაცვა ასისტენტის პასუხის დასაწყისის კარნახით."),
            ("System Prompt დისციპლინა", "გამორჩეული მდგრადობა უარყოფითი შეზღუდვებისა და ეთიკური ჩარჩოების დაცვაში.")
        ]),
        ("Google Gemini (Gemini 4 / 3.8 Flash)", "Google Ecosystem", [
            ("System Instructions პანელი", "როლი და გლობალური ქცევის წესები გამოყოფილია ცალკე სისტემურ არხში."),
            ("Massive Context (2M+)", "Context-ის ბლოკში თავისუფლად თავსდება 100+ გვერდიანი დოკუმენტი ან საათობრივი ვიდეო."),
            ("Search Grounding", "კონტექსტის ავტომატური შევსება და ფაქტების გადამოწმება Google-ის ცოცხალი საძიებო ინდექსით.")
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

        p = tf.paragraphs[0]; p.text = f"ᲞᲚᲐᲢᲤᲝᲠᲛᲐ {i+1}".upper()
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
    p = tfb.paragraphs[0]; p.text = "💡 პლატფორმული სინერგია:"
    p.font.name = FONT_ARIAL; p.font.size = Pt(14); p.font.bold = True; p.font.color.rgb = COLOR_RED
    p = tfb.add_paragraph()
    p.text = "მიუხედავად იმისა, რომ სინტაქსი მოდელებს შორის განსხვავდება (XML vs JSON Schema vs System Instructions), 6-ელემენტიანი არქიტექტურული ბირთვი უცვლელია და უზრუნველყოფს პორტაბელურობას ნებისმიერ წამყვან LLM-ში."
    p.font.name = FONT_ARIAL; p.font.size = Pt(13); p.font.color.rgb = COLOR_MUTED; p.space_before = Pt(4)

    # -------------------------------------------------------------------------
    # SLIDE 6: Rules: Component Hierarchy & Conflict Resolution
    # -------------------------------------------------------------------------
    s6 = created_slides[4]
    add_header(s6, "კომპონენტების იერარქია და კონფლიქტის გადაჭრის წესები", "ᲬᲔᲡᲔᲑᲘ ᲓᲐ ᲠᲘᲡᲙᲔᲑᲘ • იერარქიული პრიორიტეტები და შეზღუდვების წესრიგი")

    hierarchy_s6 = [
        ("👑 HIGHEST PRIORITY", "უმაღლესი დონე", "უსაფრთხოება და ფორმატი", [
            ("Constraints & Security", "უსაფრთხოების წესები, ეთიკური ჩარჩო და მკაცრი შეზღუდვები გადაწონის ნებისმიერ სხვა ინსტრუქციას."),
            ("Output Format", "ფორმატის მოთხოვნა (მაგ. JSON Schema ან Markdown ცხრილი) უპირატესია თავისუფალ ტექსტზე."),
            ("კონფლიქტის პრიორიტეტი", "თუ მომხმარებელი ითხოვს ფორმატის დარღვევას, მოდელი უარს ამბობს უმაღლესი წესის სასარგებლოდ.")
        ]),
        ("⚖️ MEDIUM PRIORITY", "საშუალო დონე", "ამოცანა და როლი", [
            ("Task & Instructions", "ძირითადი ამოცანა და მისი გადაჭრის ალგორითმული ნაბიჯები წარმოადგენს შესრულების ბირთვს."),
            ("Role & Persona", "როლი ემსახურება ამოცანის შესრულებას — ის არ უნდა ეწინააღმდეგებოდეს უსაფრთხოების შეზღუდვებს."),
            ("ტონის ადაპტაცია", "როლი კარნახობს პროფესიულ ლექსიკონს, მაგრამ არ ცვლის ფაქტობრივი სიზუსტის ვალდებულებას.")
        ]),
        ("📖 BASELINE CONTEXT", "საბაზისო დონე", "მონაცემები და იზოლაცია", [
            ("Context როგორც მასალა", "კონტექსტი წარმოადგენს მასალას დასამუშავებლად და არა მმართველ დირექტივას."),
            ("ინექციისგან იზოლაცია", "კონტექსტში არსებული ტექსტი (მაგ. დოკუმენტში ეწერა „დაივიწყე წინა წესები“) არ უნდა აღიქმებოდეს ბრძანებად."),
            ("Data Boundary", "მკაცრი გამიჯვნა გამყოფებით (Delimiters) გამორიცხავს კონტექსტის მიერ ინსტრუქციის გადაწერას.")
        ])
    ]

    for i, (badge, title, subtitle, items) in enumerate(hierarchy_s6):
        bg_col = COLOR_CARD_RED if i == 0 else COLOR_BG_LIGHT
        bor_col = COLOR_BORDER_RED if i == 0 else COLOR_BORDER
        c = s6.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, xs_3[i], card_y, card_w, card_h)
        c.fill.solid(); c.fill.fore_color.rgb = bg_col
        c.line.color.rgb = bor_col; c.line.width = Pt(1.5)

        tb = s6.shapes.add_textbox(xs_3[i] + Inches(0.35), card_y + Inches(0.35), card_w - Inches(0.7), card_h - Inches(0.7))
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
    bb = s6.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(1.0), bot_y, Inches(17.0), Inches(2.0))
    bb.fill.solid(); bb.fill.fore_color.rgb = COLOR_WHITE
    bb.line.color.rgb = COLOR_BORDER; bb.line.width = Pt(1.2)

    tbb = s6.shapes.add_textbox(Inches(1.4), bot_y + Inches(0.3), Inches(16.2), Inches(1.4))
    tfb = tbb.text_frame; tfb.word_wrap = True; tfb.margin_left = tfb.margin_top = 0
    p = tfb.paragraphs[0]; p.text = "💡 იერარქიის ოქროს წესი:"
    p.font.name = FONT_ARIAL; p.font.size = Pt(14); p.font.bold = True; p.font.color.rgb = COLOR_RED
    p = tfb.add_paragraph()
    p.text = "როდესაც კომპონენტები ერთმანეთს ეჯახება, მოდელი ყოველთვის უნდა მიჰყვეს პრიორიტეტს: უსაფრთხოების შეზღუდვები > გამომავალი ფორმატი > ძირითადი ამოცანა > როლი > კონტექსტური მონაცემები."
    p.font.name = FONT_ARIAL; p.font.size = Pt(13); p.font.color.rgb = COLOR_MUTED; p.space_before = Pt(4)

    # -------------------------------------------------------------------------
    # SLIDE 7: Workspaces: Component Layout in Developer IDEs
    # -------------------------------------------------------------------------
    s7 = created_slides[5]
    add_header(s7, "6 ელემენტის მართვა Developer IDE-ებში: AI Studio, Console & Playground", "ᲡᲐᲛᲣᲨᲐᲝ ᲒᲐᲠᲔᲛᲝ • როგორ ნაწილდება კომპონენტები ინტერფეისში")

    ide_s7 = [
        ("System Instructions პანელი", "გლობალური გარემო", [
            ("Role & Persona", "ვინ არის მოდელი და რა დონის ექსპერტიზა გააჩნია (მაგ. Senior Security Engineer)."),
            ("Global Constraints", "შეზღუდვები, რომლებიც უცვლელად უნდა შესრულდეს ყველა რაუნდში."),
            ("Persistent Rules", "წესები, რომლებიც არ უნდა გადაიწეროს მომხმარებლის შემთხვევითი ფრაზებით.")
        ]),
        ("User & Context პანელი", "დინამიკური მონაცემები", [
            ("Task (ძირითადი ამოცანა)", "კონკრეტული სესიის მიზანი, რომელიც მომხმარებლის მიერ იგზავნება."),
            ("Context & Payload", "შემავალი დოკუმენტები, ცხრილები, JSON ობიექტები ან საუბრის ისტორია."),
            ("Step-by-step Execution", "დამუშავების კონკრეტული ნაბიჯები ამ კონკრეტული მონაცემებისთვის.")
        ]),
        ("Parameters & Schema პანელი", "მათემატიკური დაცვა", [
            ("Output Format / JSON Schema", "Strict Structured Outputs რეჟიმი სინტაქსური ვალიდურობის გარანტიისთვის."),
            ("Temperature & Top-P", "ალბათობების განაწილების რეგულირება (0.0 დეტერმინისტული შედეგისთვის)."),
            ("Stop Sequences", "სიმბოლოები, რომელთა დანახვისას მოდელი მყისიერად წყვეტს გენერაციას.")
        ])
    ]

    for i, (title, subtitle, items) in enumerate(ide_s7):
        bg_col = COLOR_CARD_RED if i == 0 else COLOR_BG_LIGHT
        bor_col = COLOR_BORDER_RED if i == 0 else COLOR_BORDER
        c = s7.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, xs_3[i], card_y, card_w, card_h)
        c.fill.solid(); c.fill.fore_color.rgb = bg_col
        c.line.color.rgb = bor_col; c.line.width = Pt(1.5)

        tb = s7.shapes.add_textbox(xs_3[i] + Inches(0.35), card_y + Inches(0.35), card_w - Inches(0.7), card_h - Inches(0.7))
        tf = tb.text_frame; tf.word_wrap = True; tf.margin_left = tf.margin_top = 0

        p = tf.paragraphs[0]; p.text = f"ᲖᲝᲜᲐ {i+1}".upper()
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
    p = tfb.paragraphs[0]; p.text = "💡 დეველოპერული არქიტექტურა:"
    p.font.name = FONT_ARIAL; p.font.size = Pt(14); p.font.bold = True; p.font.color.rgb = COLOR_RED
    p = tfb.add_paragraph()
    p.text = "Developer IDE-ები სპეციალურად არის დაყოფილი ფუნქციურ ზონებად, რათა 6 ელემენტი ერთმანეთში არ აირიოს — System ზონა წესებისთვისაა, User ზონა დინამიკური მონაცემებისთვის, ხოლო Parameters ზონა მათემატიკური ფორმატისთვის."
    p.font.name = FONT_ARIAL; p.font.size = Pt(13); p.font.color.rgb = COLOR_MUTED; p.space_before = Pt(4)

    # -------------------------------------------------------------------------
    # SLIDE 8: Experiment: 1-Element to 6-Element Progressive Building
    # -------------------------------------------------------------------------
    s8 = created_slides[6]
    add_header(s8, "პრაქტიკული ექსპერიმენტი: 1-ელემენტიანიდან 6-ელემენტიან პრომპტამდე", "ᲔᲥᲡᲞᲔᲠᲘᲛᲔᲜᲢᲘ • ხარისხის ევოლუცია კომპონენტების დამატებით")

    # Top container: Evolution stages
    top_box_y = Inches(2.35)
    top_box_h = Inches(3.35)

    # Left: v1 & v2 box
    b_v12 = s8.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(1.0), top_box_y, Inches(8.35), top_box_h)
    b_v12.fill.solid(); b_v12.fill.fore_color.rgb = COLOR_BG_LIGHT
    b_v12.line.color.rgb = COLOR_BORDER; b_v12.line.width = Pt(1.5)

    tb_12 = s8.shapes.add_textbox(Inches(1.35), top_box_y + Inches(0.25), Inches(7.65), top_box_h - Inches(0.5))
    tf_12 = tb_12.text_frame; tf_12.word_wrap = True; tf_12.margin_left = tf_12.margin_top = 0
    p = tf_12.paragraphs[0]; p.text = "⚠️ V1 & V2: ᲜᲐᲬᲘᲚᲝᲑᲠᲘᲕᲘ ᲞᲠᲝᲛᲞᲢᲔᲑᲘ"
    p.font.name = FONT_ARIAL; p.font.size = Pt(11); p.font.bold = True; p.font.color.rgb = COLOR_MUTED

    p = tf_12.add_paragraph(); p.text = "ევოლუციის საწყისი ეტაპები"
    p.font.name = FONT_ARIAL; p.font.size = Pt(15); p.font.bold = True; p.font.color.rgb = COLOR_DARK; p.space_before = Pt(4)

    stages_v12 = [
        ("v1 (მხოლოდ Task)", "„დამიწერე რეპორტი გაყიდვებზე.“ ➔ შედეგი: ზოგადი ესე, გაურკვეველი ციფრები, გამოუსადეგარი მონაცემები."),
        ("v2 (+ Role & Context)", "„როლი: Data Analyst. კონტექსტი: Q3 შემოსავალი 20%-ით შემცირდა. გააანალიზე.“ ➔ შედეგი: პროფესიული ტონი, მაგრამ არასტრუქტურირებული დასკვნები."),
        ("ხარვეზის მიზეზი", "ფორმატისა და შეზღუდვების გარეშე მოდელი ირჩევს თვითნებურ სიგრძესა და სტრუქტურას.")
    ]
    for h, d in stages_v12:
        p = tf_12.add_paragraph(); p.text = f"• {h}: {d}"
        p.font.name = FONT_ARIAL; p.font.size = Pt(12); p.font.color.rgb = COLOR_TEXT_MUTED; p.space_before = Pt(4)

    # Right: v3 Full 6-Element box
    b_v3 = s8.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(9.65), top_box_y, Inches(8.35), top_box_h)
    b_v3.fill.solid(); b_v3.fill.fore_color.rgb = COLOR_CARD_RED
    b_v3.line.color.rgb = COLOR_BORDER_RED; b_v3.line.width = Pt(1.5)

    tb_v3 = s8.shapes.add_textbox(Inches(10.0), top_box_y + Inches(0.25), Inches(7.65), top_box_h - Inches(0.5))
    tf_v3 = tb_v3.text_frame; tf_v3.word_wrap = True; tf_v3.margin_left = tf_v3.margin_top = 0
    p = tf_v3.paragraphs[0]; p.text = "✅ V3: ᲡᲠᲣᲚᲘ 6-ᲔᲚᲔᲛᲔᲜᲢᲘᲐᲜᲘ ᲞᲠᲝᲛᲞᲢᲘ"
    p.font.name = FONT_ARIAL; p.font.size = Pt(11); p.font.bold = True; p.font.color.rgb = COLOR_RED

    p = tf_v3.add_paragraph(); p.text = "სრული საინჟინრო არქიტექტურა"
    p.font.name = FONT_ARIAL; p.font.size = Pt(15); p.font.bold = True; p.font.color.rgb = COLOR_DARK; p.space_before = Pt(4)

    stages_v3 = [
        ("კომპონენტების სრული სტეკი", "როლი + მიზანი + კონტექსტი + ეტაპობრივი ინსტრუქცია + შეზღუდვები + გამოსავლის ფორმატი."),
        ("სტრუქტურული შეზღუდვა", "გამოიტანე მხოლოდ Markdown ცხრილი: Region, Drop%, Primary Cause. შესავლის გარეშე."),
        ("დეტერმინისტული შედეგი", "100%-ით ზუსტი, მზა ბიზნეს-რეპორტი კონკრეტული მეტრიკებითა და მოქმედების გეგმით.")
    ]
    for h, d in stages_v3:
        p = tf_v3.add_paragraph(); p.text = f"⭐ {h}: {d}"
        p.font.name = FONT_ARIAL; p.font.size = Pt(12); p.font.bold = (h == "დეტერმინისტული შედეგი"); p.font.color.rgb = COLOR_DARK; p.space_before = Pt(4)

    # Middle Model Cards (3 cards)
    mid_y = Inches(5.95)
    mid_h = Inches(1.8)
    models_exp = [
        ("OpenAI GPT-6", "openai.png", "v1: ბუნდოვანი ესე; v3: სრული სტრუქტურული სიზუსტე, მკაცრად დაცული რიცხვითი დელტები."),
        ("Claude 5.5 Sonnet", "claude.png", "v1: ჭარბი შესავალი; v3: უმაღლესი ხარისხის ანალიტიკური ტონი, იდეალური Markdown ცხრილი."),
        ("Gemini 2.5 Pro", "gemini.png", "v1: გამოგონილი ციფრები; v3: მყისიერი დამუშავება და ანომალიების ზუსტი იდენტიფიცირება.")
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
    p.text = "თითოეული დამატებული ელემენტი პროგნოზირებადობის ხარისხს საშუალოდ 25-30%-ით ზრდის, ხოლო სრული 6 ელემენტი სისტემას საწარმოო საიმედოობამდე მიიყვანს."
    p.font.name = FONT_ARIAL; p.font.size = Pt(13); p.font.color.rgb = COLOR_MUTED; p.space_before = Pt(4)

    # -------------------------------------------------------------------------
    # SLIDE 9: Ecosystem: Prompt Template Engines & Registries (2026)
    # -------------------------------------------------------------------------
    s9 = created_slides[7]
    add_header(s9, "პრომპტის შაბლონები და მენეჯმენტი: Prompt Template Engines 2026", "ᲘᲜᲡᲢᲠᲣᲛᲔᲜᲢᲔᲑᲘ • შაბლონიზაცია, ცვლადები და ვერსიონირება")

    registries_s9 = [
        ("🗂️ TEMPLATE ENGINES", "შაბლონიზატორები", "მოდულური აწყობა", [
            ("LangChain PromptTemplates", "6 ელემენტის მოდულური აწყობა დინამიკური ცვლადებით ({role}, {context}, {task})."),
            ("Jinja2 for LLMs", "პირობითი ლოგიკა (If/Else) და ციკლები პრომპტის გენერაციის დროს სხვადასხვა მომხმარებლისთვის."),
            ("LlamaIndex Prompts", "RAG სისტემებისთვის კონტექსტის ავტომატური ინექცია წინასწარ გამზადებულ ჩარჩოში.")
        ]),
        ("📚 PROMPT REGISTRIES", "პრომპტების რეესტრი", "გუნდური ვერსიონირება", [
            ("LangSmith Hub", "ცენტრალიზებული რეესტრი პრომპტების ვერსიონირებისთვის (Git-like tags: v1.2-prod, v2.0-staging)."),
            ("Weights & Biases Prompts", "პრომპტის ცვლილებების ისტორია, ავტორების კონტროლი და ერთი დაწკაპუნებით Rollback."),
            ("Humanloop Registry", "პრომპტების გაზიარება და თანამშრომლობა ინჟინრებსა და პროდუქტის მენეჯერებს შორის.")
        ]),
        ("🔒 COMPILATION & OPTIMIZATION", "ავტომატური ოპტიმიზაცია", "კოდით მართვა", [
            ("DSPy Teleprompter", "პრომპტების ავტომატური ოპტიმიზაცია და კომპილაცია პროგრამული ალგორითმებით."),
            ("Agenta Platform", "ღია კოდის LLMOps გარემო პრომპტების სწრაფი ევალუაციისა და A/B ტესტირებისთვის."),
            ("Promptfoo Integration", "პრომპტ-შაბლონების ავტომატური ვალიდაცია CI/CD მილსადენში ყოველი ცვლილებისას.")
        ])
    ]

    for i, (badge, title, subtitle, items) in enumerate(registries_s9):
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
    p.text = "საწარმოო AI სისტემებში პრომპტები არ ინახება „მძიმე კოდში“ (Hardcoded) — ისინი ცოცხლობენ ცენტრალიზებულ Prompt Registry-ში, სადაც 6-ვე ელემენტი ცალკე მართვადი, ვერსიონირებადი და ტესტირებადი კომპონენტია."
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
    filename = "2.1 - ეფექტური პრომპტის 6 საკვანძო კომპონენტი.pptx"
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
    build_lecture_m7_2_1()
