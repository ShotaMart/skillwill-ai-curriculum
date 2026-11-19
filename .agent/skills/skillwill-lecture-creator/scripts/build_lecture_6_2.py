import os
import sys
import shutil
import pptx
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN
from pptx.enum.shapes import MSO_SHAPE

def build_lecture_6_2():
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
            p.text = "6.2 - ChatGPT სრულად II: Custom GPTs, Projects, Actions და Enterprise AI"
            p.alignment = PP_ALIGN.CENTER
            p.font.name = FONT_ARIAL
            p.font.size = Pt(34)
            p.font.bold = True
            p.font.color.rgb = COLOR_RED

            p2 = tf.add_paragraph()
            p2.text = "პერსონალიზებული AI ასისტენტები, გარე API ინტეგრაციები და კორპორაციული მართვა"
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
    def add_header(slide, title, category="ᲛᲝᲓᲣᲚᲘ 5 • OPENAI ეკოსისტემა • 6.2"):
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
    # SLIDE 2: What is Custom GPTs? (Specialized AI Assistants)
    # -------------------------------------------------------------------------
    s2 = created_slides[0]
    add_header(s2, "რა არის Custom GPTs? (მიზნობრივი ასისტენტები კოდირების გარეშე)", "პერსონალიზაცია • ზოგადი ჩატბოტიდან ვიწრო პროფესიულ აგენტამდე")

    # Left: Standard General ChatGPT
    b1 = s2.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(1.0), box_y, box_w, box_h)
    b1.fill.solid(); b1.fill.fore_color.rgb = COLOR_BG_LIGHT
    b1.line.color.rgb = COLOR_BORDER; b1.line.width = Pt(1.5)

    tb1 = s2.shapes.add_textbox(Inches(1.4), box_y + Inches(0.4), box_w - Inches(0.8), box_h - Inches(0.8))
    tf1 = tb1.text_frame; tf1.word_wrap = True; tf1.margin_left = tf1.margin_top = 0
    p = tf1.paragraphs[0]; p.text = "ᲩᲕᲔᲣᲚᲔᲑᲠᲘᲕᲘ CHATGPT"
    p.font.name = FONT_ARIAL; p.font.size = Pt(12); p.font.bold = True; p.font.color.rgb = COLOR_MUTED

    p = tf1.add_paragraph(); p.text = "General-Purpose Assistant (ზოგადი ჩატი)"
    p.font.name = FONT_ARIAL; p.font.size = Pt(21); p.font.bold = True; p.font.color.rgb = COLOR_DARK; p.space_before = Pt(4)

    gen_points = [
        ("ზოგადი ცოდნა", "მოდელმა იცის ყველაფერი ცოტ-ცოტა, მაგრამ არ ფლობს თქვენი ორგანიზაციის სპეციფიკურ მეთოდოლოგიას."),
        ("პრომპტების განმეორება", "ყოველ ახალ საუბარში მომხმარებელს უწევს როლის, ტონის, შეზღუდვებისა და ფაილების თავიდან ახსნა."),
        ("კონტექსტის გაბნევა", "რთულ ამოცანებში ზოგადი მოდელი ხშირად გადადის არასასურველ ფორმატებზე ან ზოგად რჩევებზე."),
        ("გუნდური სტანდარტის არარსებობა", "თითოეული თანამშრომელი წერს განსხვავებულ პრომპტს და იღებს არაერთგვაროვან ხარისხს.")
    ]
    for t, d in gen_points:
        p = tf1.add_paragraph(); p.text = f"• {t}: {d}"
        p.font.name = FONT_ARIAL; p.font.size = Pt(13); p.font.color.rgb = COLOR_TEXT_MUTED; p.space_before = Pt(8)

    # Right: Custom GPTs
    b2 = s2.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(9.65), box_y, box_w, box_h)
    b2.fill.solid(); b2.fill.fore_color.rgb = COLOR_CARD_RED
    b2.line.color.rgb = COLOR_BORDER_RED; b2.line.width = Pt(1.5)

    tb2 = s2.shapes.add_textbox(Inches(10.05), box_y + Inches(0.4), box_w - Inches(0.8), box_h - Inches(0.8))
    tf2 = tb2.text_frame; tf2.word_wrap = True; tf2.margin_left = tf2.margin_top = 0
    p = tf2.paragraphs[0]; p.text = "CUSTOM GPTS (2026)"
    p.font.name = FONT_ARIAL; p.font.size = Pt(12); p.font.bold = True; p.font.color.rgb = COLOR_RED

    p = tf2.add_paragraph(); p.text = "Specialized Agent (მიზნობრივი ასისტენტი)"
    p.font.name = FONT_ARIAL; p.font.size = Pt(21); p.font.bold = True; p.font.color.rgb = COLOR_DARK; p.space_before = Pt(4)

    spec_points = [
        ("Pre-configured Persona (მუდმივი როლი)", "წინასწარ გაწერილი სისტემური ინსტრუქცია, რომელიც არასოდეს ავიწყდება მოდელს."),
        ("Embedded Knowledge Base (ჩაშენებული ცოდნა)", "ატვირთული შიდა დოკუმენტები, გაიდლაინები და ცხრილები, რომლებსაც მოდელი ყოველთვის ეყრდნობა."),
        ("Actions & API Execution (მოქმედების უნარი)", "შეუძლია გარე სისტემებთან (CRM, Slack, Google Calendar) უსაფრთხო დაკავშირება."),
        ("No-Code Creation (მარტივი აწყობა)", "ასისტენტის შექმნა შეუძლია ნებისმიერ თანამშრომელს ბუნებრივი ენით (Natural Language).")
    ]
    for t, d in spec_points:
        p = tf2.add_paragraph(); p.text = f"⭐ {t}"
        p.font.name = FONT_ARIAL; p.font.size = Pt(13); p.font.bold = True; p.font.color.rgb = COLOR_DARK; p.space_before = Pt(8)
        p = tf2.add_paragraph(); p.text = d
        p.font.name = FONT_ARIAL; p.font.size = Pt(12); p.font.color.rgb = COLOR_TEXT_MUTED; p.space_before = Pt(2)

    # Bottom takeaway
    bb = s2.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(1.0), bot_y, Inches(17.0), Inches(2.0))
    bb.fill.solid(); bb.fill.fore_color.rgb = COLOR_WHITE
    bb.line.color.rgb = COLOR_BORDER; bb.line.width = Pt(1.2)

    tbb = s2.shapes.add_textbox(Inches(1.4), bot_y + Inches(0.3), Inches(16.2), Inches(1.4))
    tfb = tbb.text_frame; tfb.word_wrap = True; tfb.margin_left = tfb.margin_top = 0
    p = tfb.paragraphs[0]; p.text = "💡 მთავარი არსი (მარტივი ანალოგია):"
    p.font.name = FONT_ARIAL; p.font.size = Pt(14); p.font.bold = True; p.font.color.rgb = COLOR_RED
    p = tfb.add_paragraph()
    p.text = "Custom GPT არის თქვენი ორგანიზაციის შიდა გაიდლაინებითა და დოკუმენტებით წინასწარ მომზადებული ვირტუალური თანამშრომელი, რომელიც ყოველ ჯერზე იდეალურად და სტანდარტულად ასრულებს ერთ კონკრეტულ ბიზნეს-ამოცანას."
    p.font.name = FONT_ARIAL; p.font.size = Pt(13); p.font.color.rgb = COLOR_MUTED; p.space_before = Pt(4)

    # -------------------------------------------------------------------------
    # SLIDE 3: GPT Builder Architecture (Instructions, Knowledge, Capabilities)
    # -------------------------------------------------------------------------
    s3 = created_slides[1]
    add_header(s3, "GPT Builder-ის შიდა არქიტექტურა (Instructions, Knowledge და Capabilities)", "არქიტექტურა • 3 საყრდენი სვეტი ეფექტური ასისტენტის შესაქმნელად")

    cards_s3 = [
        ("📝 INSTRUCTIONS", "Instructions (სისტემური წესები)", "ქცევის ჩარჩო", [
            ("როლი და პერსონა", "ზუსტად ვინ არის ასისტენტი (მაგ. „Skillwill-ის სადიპლომო ნაშრომების ტექნიკური რეცენზენტი“)."),
            ("მკაცრი შეზღუდვები (Negative Constraints)", "რა არის დაუშვებელი (მაგ. „არ გამოიყენო გაურკვეველი ტერმინები“, „არ მოიგონო ფაქტები“)."),
            ("შედეგის ფორმატი", "სტრუქტურირებული ცხრილები, Markdown ბლოკები ან ნაბიჯ-ნაბიჯ ჩამონათვალი.")
        ]),
        ("📚 KNOWLEDGE", "Knowledge (ცოდნის ბაზა)", "დოკუმენტური საფუძველი", [
            ("შიდა დოკუმენტების ატვირთვა", "PDF, DOCX, TXT, CSV, JSON და პროგრამული კოდის რეპოზიტორიები."),
            ("RAG და ვექტორული ძებნა", "მომხმარებლის კითხვაზე მოდელი პოულობს შესაბამის აბზაცებს და პასუხობს ზუსტი ციტატებით."),
            ("დაცული კონტექსტი", "ცოდნის ფაილები არ არის საჯარო და გარე მომხმარებელი ვერ გადმოწერს მათ ორიგინალ სახით.")
        ]),
        ("⚡ CAPABILITIES", "Capabilities (ჩაშენებული ხელსაწყოები)", "ფუნქციური უნარები", [
            ("Web Search", "საჭიროების შემთხვევაში უახლესი ინფორმაციის ვებიდან მოძიება და გადამოწმება."),
            ("Canvas & Code Interpreter", "Python კოდის გაშვება, მონაცემთა დამუშავება და ინტერაქტიული Canvas-ში ჩაწერა."),
            ("DALL-E 3 Image Generation", "ილუსტრაციების, სქემების, ლოგოებისა და სურათების შექმნა ბრიფის მიხედვით.")
        ])
    ]

    for i, (badge, title, subtitle, items) in enumerate(cards_s3):
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
    p = tfb.paragraphs[0]; p.text = "💡 ოქროს წესი (Best Practice):"
    p.font.name = FONT_ARIAL; p.font.size = Pt(14); p.font.bold = True; p.font.color.rgb = COLOR_RED
    p = tfb.add_paragraph()
    p.text = "კარგი GPT-ის საიდუმლო მკაფიო შეზღუდვებშია: ნაცვლად ზოგადი ფრაზებისა, მიეცით ასისტენტს მკაცრი ინსტრუქცია — რა ეტაპებით იმოქმედოს, რა წყაროებს დაეყრდნოს და რა შემთხვევაში თქვას „არ ვიცი“."
    p.font.name = FONT_ARIAL; p.font.size = Pt(13); p.font.color.rgb = COLOR_MUTED; p.space_before = Pt(4)

    # -------------------------------------------------------------------------
    # SLIDE 4: GPT Actions & External API Integrations
    # -------------------------------------------------------------------------
    s4 = created_slides[2]
    add_header(s4, "GPT Actions და გარე API ინტეგრაციები (OpenAPI სტანდარტი)", "ინტეგრაცია • საუბრიდან გარე სისტემებში რეალური მოქმედებების შესრულებამდე")

    # Left: Actions Architecture
    b1 = s4.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(1.0), box_y, box_w, box_h)
    b1.fill.solid(); b1.fill.fore_color.rgb = COLOR_BG_LIGHT
    b1.line.color.rgb = COLOR_BORDER; b1.line.width = Pt(1.5)

    tb1 = s4.shapes.add_textbox(Inches(1.4), box_y + Inches(0.4), box_w - Inches(0.8), box_h - Inches(0.8))
    tf1 = tb1.text_frame; tf1.word_wrap = True; tf1.margin_left = tf1.margin_top = 0
    p = tf1.paragraphs[0]; p.text = "ACTIONS ᲐᲠᲥᲘᲢᲔᲥᲢᲣᲠᲐ"
    p.font.name = FONT_ARIAL; p.font.size = Pt(12); p.font.bold = True; p.font.color.rgb = COLOR_MUTED

    p = tf1.add_paragraph(); p.text = "OpenAPI 3.0 სპეციფიკაცია"
    p.font.name = FONT_ARIAL; p.font.size = Pt(21); p.font.bold = True; p.font.color.rgb = COLOR_DARK; p.space_before = Pt(4)

    actions_points = [
        ("OpenAPI Schema (სქემა)", "JSON ან YAML ფორმატში გაწერილი API სტრუქტურა, რომელიც აღწერს ხელმისაწვდომ Endpoints-სა და პარამეტრებს."),
        ("Function Calling (ფუნქციების გამოძახება)", "მოდელი თავად ხვდება, როდის სჭირდება გარე ბაზასთან დაკავშირება და აგენერირებს სწორ JSON მოთხოვნას."),
        ("მონაცემთა ორმხრივი გაცვლა", "მონაცემების წაკითხვა (GET) და ახალი ჩანაწერების შექმნა ან განახლება (POST/PUT/DELETE)."),
        ("უსაფრთხო ავტორიზაცია", "API Key, OAuth 2.0 ან Bearer Token მომხმარებლის უსაფრთხო იდენტიფიკაციისთვის.")
    ]
    for t, d in actions_points:
        p = tf1.add_paragraph(); p.text = f"• {t}: {d}"
        p.font.name = FONT_ARIAL; p.font.size = Pt(13); p.font.color.rgb = COLOR_TEXT_MUTED; p.space_before = Pt(8)

    # Right: Practical Business Scenarios
    b2 = s4.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(9.65), box_y, box_w, box_h)
    b2.fill.solid(); b2.fill.fore_color.rgb = COLOR_CARD_RED
    b2.line.color.rgb = COLOR_BORDER_RED; b2.line.width = Pt(1.5)

    tb2 = s4.shapes.add_textbox(Inches(10.05), box_y + Inches(0.4), box_w - Inches(0.8), box_h - Inches(0.8))
    tf2 = tb2.text_frame; tf2.word_wrap = True; tf2.margin_left = tf2.margin_top = 0
    p = tf2.paragraphs[0]; p.text = "ᲞᲠᲐᲥᲢᲘᲙᲣᲚᲘ ᲡᲪᲔᲜᲐᲠᲔᲑᲘ"
    p.font.name = FONT_ARIAL; p.font.size = Pt(12); p.font.bold = True; p.font.color.rgb = COLOR_RED

    p = tf2.add_paragraph(); p.text = "ბიზნეს-პროცესების ავტომატიზაცია"
    p.font.name = FONT_ARIAL; p.font.size = Pt(21); p.font.bold = True; p.font.color.rgb = COLOR_DARK; p.space_before = Pt(4)

    scen_points = [
        ("CRM ინტეგრაცია (HubSpot / Salesforce)", "კლიენტის სტატუსის მყისიერი შემოწმება და საუბრის შედეგად ახალი ლიდის ავტომატური რეგისტრაცია."),
        ("კალენდარი და თასქები (Calendar & Jira)", "შეხვედრის ჩანიშვნა თავისუფალ სლოტში ან ახალი დეველოპერული ამოცანის გახსნა Jira/Notion-ში."),
        ("მონაცემთა ბაზების გამოკითხვა (SQL / DB)", "PostgreSQL ან BigQuery ბაზიდან ფინანსური მონაცემების უსაფრთხო ამოღება და რეზიუმირება."),
        ("შეტყობინებების გაგზავნა (Slack / Teams)", "ანალიტიკური რეპორტის პირდაპირი გაგზავნა კომპანიის შესაბამის სამუშაო არხში.")
    ]
    for t, d in scen_points:
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
    p = tfb.paragraphs[0]; p.text = "💡 უსაფრთხოების ბარიერი (Confirmation Step):"
    p.font.name = FONT_ARIAL; p.font.size = Pt(14); p.font.bold = True; p.font.color.rgb = COLOR_RED
    p = tfb.add_paragraph()
    p.text = "Actions-ის გაშვებამდე ChatGPT ყოველთვის სთხოვს მომხმარებელს დადასტურებას (Allow / Always Allow), რათა სრულად გამოირიცხოს უნებართვო ფინანსური ტრანზაქციები ან სენსიტიური მონაცემების გარე სერვერზე გადაცემა."
    p.font.name = FONT_ARIAL; p.font.size = Pt(13); p.font.color.rgb = COLOR_MUTED; p.space_before = Pt(4)

    # -------------------------------------------------------------------------
    # SLIDE 5: ChatGPT Projects & Workspaces
    # -------------------------------------------------------------------------
    s5 = created_slides[3]
    add_header(s5, "ChatGPT Projects და Workspaces (გუნდური სამუშაო გარემო)", "სამუშაო სივრცეები • როგორ ვმართავთ საერთო კონტექსტსა და გუნდურ ცოდნას")

    cards_s5 = [
        ("📁 PROJECTS", "Projects სტრუქტურა", "სამუშაო სივრცეები", [
            ("გამოყოფილი გარემო", "ცალკე საქაღალდეები კონკრეტული პროექტებისთვის (მაგ. „მარკეტინგული კამპანია Q4“, „პროდუქტის აუდიტი“)."),
            ("Project Knowledge", "ატვირთული ფაილები და ცხრილები ხელმისაწვდომია პროექტის შიგნით გახსნილი ყველა ჩატისთვის."),
            ("Project Instructions", "პროექტის სპეციფიკური ქცევის წესები და სტანდარტები, რომელიც ყველა შიდა ჩატზე ვრცელდება.")
        ]),
        ("⚖️ COMPARISON", "Custom GPT vs Project", "სწორი გამიჯვნა", [
            ("Custom GPT (ინსტრუმენტი)", "ვიწრო როლის მქონე სპეციალიზებული ასისტენტი კონკრეტული ამოცანის შესასრულებლად (Reusable Tool)."),
            ("ChatGPT Project (გარემო)", "კონტეინერი, რომელიც აერთიანებს დოკუმენტებს, ჩატებსა და GPT-ებს საერთო მიზნის ირგვლივ (Workspace)."),
            ("კომბინირებული ძალა", "პროექტის შიგნით შეგიძლიათ გამოიყენოთ როგორც ზოგადი მოდელი (Sol/Astra), ისე Custom GPT-ები.")
        ]),
        ("👥 COLLABORATION", "გუნდური თანამშრომლობა", "ცოდნის უწყვეტობა", [
            ("Shared Context", "გუნდის წევრები მუშაობენ ერთსა და იმავე დოკუმენტურ ბაზაზე ინფორმაციის დუბლირების გარეშე."),
            ("ჩატების იზოლაცია", "საუბრები რჩება პირადული, სანამ მომხმარებელი არ გადაწყვეტს მის პროექტში გაზიარებას."),
            ("Fast Onboarding", "ახალი თანამშრომელი მყისიერად იღებს წვდომას პროექტის მთელ ისტორიასა და მასალებზე.")
        ])
    ]

    for i, (badge, title, subtitle, items) in enumerate(cards_s5):
        bg_col = COLOR_CARD_RED if i == 1 else COLOR_BG_LIGHT
        bor_col = COLOR_BORDER_RED if i == 1 else COLOR_BORDER
        c = s5.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, xs_3[i], card_y, card_w, card_h)
        c.fill.solid(); c.fill.fore_color.rgb = bg_col
        c.line.color.rgb = bor_col; c.line.width = Pt(1.5)

        tb = s5.shapes.add_textbox(xs_3[i] + Inches(0.35), card_y + Inches(0.35), card_w - Inches(0.7), card_h - Inches(0.7))
        tf = tb.text_frame; tf.word_wrap = True; tf.margin_left = tf.margin_top = 0

        p = tf.paragraphs[0]; p.text = badge.upper()
        p.font.name = FONT_ARIAL; p.font.size = Pt(11); p.font.bold = True
        p.font.color.rgb = COLOR_RED if i == 1 else COLOR_MUTED

        p = tf.add_paragraph(); p.text = title
        p.font.name = FONT_ARIAL; p.font.size = Pt(18); p.font.bold = True; p.font.color.rgb = COLOR_DARK; p.space_before = Pt(4)

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
    p = tfb.paragraphs[0]; p.text = "💡 ორგანიზაციული წესი:"
    p.font.name = FONT_ARIAL; p.font.size = Pt(14); p.font.bold = True; p.font.color.rgb = COLOR_RED
    p = tfb.add_paragraph()
    p.text = "თუ გჭირდებათ რუტინული პროცესის ავტომატიზაცია (მაგ. ინვოისის შემოწმება) — შექმენით Custom GPT; ხოლო თუ მუშაობთ კომპლექსურ პროექტზე (მაგ. ახალი პროდუქტის გაშვება) — გახსენით ChatGPT Project."
    p.font.name = FONT_ARIAL; p.font.size = Pt(13); p.font.color.rgb = COLOR_MUTED; p.space_before = Pt(4)

    # -------------------------------------------------------------------------
    # SLIDE 6: GPT Store & Sharing Hierarchy
    # -------------------------------------------------------------------------
    s6 = created_slides[4]
    add_header(s6, "GPT Store და გაზიარების იერარქია (Distribution & Access)", "დისტრიბუცია • წვდომის 3 დონე და კორპორაციული უსაფრთხოება")

    cards_s6 = [
        ("🔒 LEVEL 1", "1. Only Me (პირადი ასისტენტი)", "სრული იზოლაცია", [
            ("ინდივიდუალური წვდომა", "ასისტენტი ხელმისაწვდომია მხოლოდ მისი უშუალო შემქმნელისთვის."),
            ("სატესტო გარემო", "იდეალურია ექსპერიმენტებისთვის, პირადი პრომპტებისა და შიდა დრაფტების დასახვეწად."),
            ("ნულოვანი რისკი", "მონაცემები, ფაილები და ინსტრუქციები არასოდეს ტოვებს პირად ანგარიშს.")
        ]),
        ("🔗 LEVEL 2", "2. Anyone with a Link", "შერჩევითი გაზიარება", [
            ("პირდაპირი ბმული", "წვდომა აქვთ მხოლოდ იმ პირებს, ვისაც გაუგზავნით პირდაპირ უნიკალურ ბმულს (URL)."),
            ("კოლაბორაცია პარტნიორებთან", "მოსახერხებელია გარე კონტრაქტორებთან, ფრილანსერებთან და კლიენტებთან მუშაობისას."),
            ("დამალვა კატალოგში", "ასისტენტი არ იძებნება საჯარო GPT Store-ში და არ ჩანს მესამე პირებისთვის.")
        ]),
        ("🏢 LEVEL 3", "3. Organization / Public Store", "მასშტაბური დისტრიბუცია", [
            ("Enterprise Workspace Store", "ასისტენტი ხელმისაწვდომია მხოლოდ კომპანიის ავტორიზებული თანამშრომლებისთვის."),
            ("Public GPT Store", "გლობალური კატალოგი მილიონობით მომხმარებლისთვის რეიტინგებითა და მონეტიზაციით."),
            ("ადმინისტრატორის კონტროლი", "IT ადმინისტრატორს შეუძლია აკრძალოს საჯარო GPT-ების გამოყენება და დაუშვას მხოლოდ შიდა ასისტენტები.")
        ])
    ]

    for i, (badge, title, subtitle, items) in enumerate(cards_s6):
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
    p = tfb.paragraphs[0]; p.text = "💡 კორპორაციული უსაფრთხოების წესი:"
    p.font.name = FONT_ARIAL; p.font.size = Pt(14); p.font.bold = True; p.font.color.rgb = COLOR_RED
    p = tfb.add_paragraph()
    p.text = "კომპანიის შიდა დოკუმენტებზე შექმნილი GPT-ები არასოდეს არ უნდა გამოქვეყნდეს Public სტატუსით — ისინი უნდა განთავსდეს მკაცრად Organization Workspace-ის შიგნით, რათა შიდა ნოუ-ჰაუ არ გახდეს საჯარო."
    p.font.name = FONT_ARIAL; p.font.size = Pt(13); p.font.color.rgb = COLOR_MUTED; p.space_before = Pt(4)

    # -------------------------------------------------------------------------
    # SLIDE 7: OpenAI Data Policy & Zero Data Retention
    # -------------------------------------------------------------------------
    s7 = created_slides[5]
    add_header(s7, "OpenAI Data Policy და Zero Data Retention (ZDR)", "კონფიდენციალურობა • სად მიდის და როგორ არის დაცული თქვენი მონაცემები")

    cards_s7 = [
        ("🚫 ZERO TRAINING", "Zero Training on Business Data", "მოდელის წვრთნის აკრძალვა", [
            ("მკაცრი აკრძალვა", "OpenAI არ იყენებს Business, Enterprise და API მომხმარებლების ჩატებსა და ფაილებს მოდელების საწვრთნელად."),
            ("IP Protection", "თქვენი კომერციული საიდუმლოება, კოდი და ანალიტიკა რჩება 100%-ით თქვენს მფლობელობაში."),
            ("Consumer Opt-Out", "ინდივიდუალურ (Free/Plus) მომხმარებლებსაც შეუძლიათ პარამეტრებიდან გათიშონ მოდელის გაუმჯობესების ფუნქცია.")
        ]),
        ("🗑️ ZERO RETENTION", "Zero Data Retention (ZDR)", "მყისიერი წაშლა", [
            ("ნულოვანი შენახვა", "მოთხოვნის დამუშავებისთანავე სერვერებიდან იშლება დროებითი ტოკენები და ფაილების ასლები."),
            ("ლოგირების გარეშე", "მოთხოვნები არ ინახება მუდმივ მონაცემთა ბაზებში ან სერვერის ლოგებში."),
            ("მარეგულირებელი სტანდარტები", "აკმაყოფილებს ყველაზე მკაცრ საბანკო, სამედიცინო და სამთავრობო მოთხოვნებს.")
        ]),
        ("🔐 ENCRYPTION", "მონაცემთა დაშიფვრა", "უსაფრთხოების სტანდარტები", [
            ("Encryption in Transit", "ყველა მონაცემი იშიფრება გადაცემის დროს თანამედროვე TLS 1.3 პროტოკოლით."),
            ("Encryption at Rest", "შენახული მონაცემები დაცულია სამხედრო დონის AES-256 დაშიფვრით."),
            ("Customer-Managed Keys (CMK)", "Enterprise მომხმარებლებს შეუძლიათ გამოიყენონ საკუთარი დაშიფვრის გასაღებები.")
        ])
    ]

    for i, (badge, title, subtitle, items) in enumerate(cards_s7):
        bg_col = COLOR_CARD_RED if i == 0 else COLOR_BG_LIGHT
        bor_col = COLOR_BORDER_RED if i == 0 else COLOR_BORDER
        c = s7.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, xs_3[i], card_y, card_w, card_h)
        c.fill.solid(); c.fill.fore_color.rgb = bg_col
        c.line.color.rgb = bor_col; c.line.width = Pt(1.5)

        tb = s7.shapes.add_textbox(xs_3[i] + Inches(0.35), card_y + Inches(0.35), card_w - Inches(0.7), card_h - Inches(0.7))
        tf = tb.text_frame; tf.word_wrap = True; tf.margin_left = tf.margin_top = 0

        p = tf.paragraphs[0]; p.text = badge.upper()
        p.font.name = FONT_ARIAL; p.font.size = Pt(11); p.font.bold = True
        p.font.color.rgb = COLOR_RED if i == 0 else COLOR_MUTED

        p = tf.add_paragraph(); p.text = title
        p.font.name = FONT_ARIAL; p.font.size = Pt(18); p.font.bold = True; p.font.color.rgb = COLOR_DARK; p.space_before = Pt(4)

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
    p = tfb.paragraphs[0]; p.text = "💡 კორპორაციული ნდობის საფუძველი:"
    p.font.name = FONT_ARIAL; p.font.size = Pt(14); p.font.bold = True; p.font.color.rgb = COLOR_RED
    p = tfb.add_paragraph()
    p.text = "OpenAI-ს თანამედროვე პოლიტიკა მკაფიოდ მიჯნავს უფასო სამომხმარებლო სერვისს კორპორაციული გარემოსგან: ბიზნეს მონაცემები არასოდეს ხდება საჯარო და არ გამოიყენება AI-ს სწავლებისთვის."
    p.font.name = FONT_ARIAL; p.font.size = Pt(13); p.font.color.rgb = COLOR_MUTED; p.space_before = Pt(4)

    # -------------------------------------------------------------------------
    # SLIDE 8: ChatGPT Enterprise & Business Architecture
    # -------------------------------------------------------------------------
    s8 = created_slides[6]
    add_header(s8, "ChatGPT Enterprise და Business არქიტექტურა", "ინფრასტრუქტურა • კორპორაციული კონტროლი, ადმინისტრირება და აუდიტი")

    pillars_s8 = [
        ("1. SSO & SCIM Directory", "მომხმარებელთა მართვა", [
            ("Single Sign-On (SSO)", "ერთიანი უსაფრთხო ავტორიზაცია Okta, Microsoft Entra ID და Google Workspace-ით."),
            ("SCIM Provisioning", "თანამშრომლის დამატების ან წასვლისას ანგარიშებისა და წვდომების ავტომატური მართვა.")
        ]),
        ("2. Admin Console & Analytics", "ცენტრალიზებული კონტროლი", [
            ("ერთიანი ადმინ-პანელი", "მომხმარებლების, ლიცენზიებისა და შიდა Custom GPT-ების ცენტრალიზებული მართვა."),
            ("Usage Analytics", "დეტალური სტატისტიკა — რომელი დეპარტამენტი რა ინტენსივობით იყენებს AI-ს.")
        ]),
        ("3. Domain Verification & Security", "დომენის კონტროლი", [
            ("დომენის ვერიფიკაცია", "კომპანიის კორპორაციული ელფოსტით დარეგისტრირებული ყველა ანგარიშის მართვა."),
            ("Actions Whitelisting", "ადმინისტრატორი განსაზღვრავს, რომელ გარე API-ებთან აქვთ Custom GPT-ებს დაკავშირების უფლება.")
        ]),
        ("4. Compliance Standards", "საერთაშორისო აუდიტი", [
            ("SOC 2 Type II აუდიტი", "დამოუკიდებელი აუდიტორული შემოწმება უსაფრთხოებასა და კონფიდენციალურობაზე."),
            ("HIPAA & GDPR", "ჯანდაცვისა და ევროპული პერსონალური მონაცემების დაცვის კანონმდებლობასთან სრული თანხვედრა.")
        ])
    ]

    for i, (title, subtitle, items) in enumerate(pillars_s8):
        bg_col = COLOR_CARD_RED if i == 0 else COLOR_BG_LIGHT
        bor_col = COLOR_BORDER_RED if i == 0 else COLOR_BORDER
        c = s8.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, xs_4[i], step_y, step_w, step_h)
        c.fill.solid(); c.fill.fore_color.rgb = bg_col
        c.line.color.rgb = bor_col; c.line.width = Pt(1.5)

        tb = s8.shapes.add_textbox(xs_4[i] + Inches(0.25), step_y + Inches(0.35), step_w - Inches(0.5), step_h - Inches(0.7))
        tf = tb.text_frame; tf.word_wrap = True; tf.margin_left = tf.margin_top = 0

        p = tf.paragraphs[0]; p.text = f"სვეტი {i+1}".upper()
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
    p = tfb.paragraphs[0]; p.text = "🎯 Enterprise შეჯამება:"
    p.font.name = FONT_ARIAL; p.font.size = Pt(14); p.font.bold = True; p.font.color.rgb = COLOR_RED
    p = tfb.add_paragraph()
    p.text = "ChatGPT Enterprise უზრუნველყოფს უსაფრთხოების, აუდიტისა და მართვის უმაღლეს დონეს, რაც საშუალებას აძლევს მსხვილ ორგანიზაციებს უსაფრთხოდ დანერგონ AI მასშტაბურად."
    p.font.name = FONT_ARIAL; p.font.size = Pt(13); p.font.color.rgb = COLOR_MUTED; p.space_before = Pt(4)

    # -------------------------------------------------------------------------
    # SLIDE 9: ChatGPT vs Claude Comparative Scorecard (2026 Leaders)
    # -------------------------------------------------------------------------
    s9 = created_slides[7]
    add_header(s9, "ChatGPT vs Claude შედარებითი Scorecard (2026 წლის ლიდერები)", "შედარებითი ანალიზი • ორი წამყვანი პლატფორმის ობიექტური შეფასება")

    cards_s9 = [
        ("🏆 ANTHROPIC", "სად ლიდერობს Claude?", "ანალიტიკა და ინჟინერია", [
            ("Massive Context & Retrieval", "200K+ ტოკენის უზადო აღქმა, წიგნებისა და რთული კოდის რეპოზიტორიების ანალიზი."),
            ("Logic & Coding Depth", "Claude 5.5 Sonnet და Opus ლიდერობს რთულ ალგორითმულ კოდირებასა და მათემატიკაში."),
            ("MCP არქიტექტურა", "ღია Model Context Protocol და ავტონომიური ტერმინალის აგენტი Claude Code."),
            ("Artifacts Ecosystem", "ინტერაქტიული დოკუმენტებისა და აპლიკაციების დამოუკიდებელი მართვა.")
        ]),
        ("🏆 OPENAI", "სად ლიდერობს ChatGPT?", "მულტიმოდალობა და კვლევა", [
            ("Advanced Voice Mode", "მშობლიური აუდიო-სტრიმინგი ემოციებით, ტემპითა და მყისიერი შეწყვეტის ფუნქციით."),
            ("SearchGPT & Deep Research", "უმაღლესი დონის ცოცხალი ვებ-ძიება და ავტონომიური სიღრმისეული კვლევა."),
            ("OpenAI Canvas", "ტექსტისა და კოდის ორმხრივი, შერჩევითი რედაქტირების უმაღლესი კომფორტი."),
            ("Custom GPTs & Store", "ყველაზე დიდი გლობალური ეკოსისტემა მზა ასისტენტებითა და API Actions-ით.")
        ]),
        ("🎯 FRAMEWORK", "სწორი შერჩევის წესი", "Decision Guide", [
            ("აირჩიეთ Claude", "სიღრმისეული ანალიტიკისთვის, დიდი დოკუმენტების დამუშავებისთვის, სისტემური პროგრამირებისთვის."),
            ("აირჩიეთ ChatGPT", "ხმოვანი ინტერაქციისთვის, ცოცხალი ვებ-კვლევისთვის, Canvas-ით მუშაობისთვის და Custom GPT-ებისთვის."),
            ("სინერგიული ეფექტი", "თანამედროვე კომპანიები წარმატებით იყენებენ ორივე სისტემას მათი ძლიერი მხარეების მიხედვით.")
        ])
    ]

    for i, (badge, title, subtitle, items) in enumerate(cards_s9):
        bg_col = COLOR_CARD_RED if i == 2 else COLOR_BG_LIGHT
        bor_col = COLOR_BORDER_RED if i == 2 else COLOR_BORDER
        c = s9.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, xs_3[i], card_y, card_w, card_h)
        c.fill.solid(); c.fill.fore_color.rgb = bg_col
        c.line.color.rgb = bor_col; c.line.width = Pt(1.5)

        tb = s9.shapes.add_textbox(xs_3[i] + Inches(0.35), card_y + Inches(0.35), card_w - Inches(0.7), card_h - Inches(0.7))
        tf = tb.text_frame; tf.word_wrap = True; tf.margin_left = tf.margin_top = 0

        p = tf.paragraphs[0]; p.text = badge.upper()
        p.font.name = FONT_ARIAL; p.font.size = Pt(11); p.font.bold = True
        p.font.color.rgb = COLOR_RED if i == 2 else COLOR_MUTED

        p = tf.add_paragraph(); p.text = title
        p.font.name = FONT_ARIAL; p.font.size = Pt(18); p.font.bold = True; p.font.color.rgb = COLOR_DARK; p.space_before = Pt(4)

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
    p.text = "არ არსებობს ერთი აბსოლუტური გამარჯვებული — 2026 წელს Claude და ChatGPT სრულყოფილად ავსებენ ერთმანეთს: Claude არის თქვენი ღრმა ანალიტიკოსი და ინჟინერი, ხოლო ChatGPT — თქვენი კვლევითი, ხმოვანი და მულტიმოდალური პარტნიორი."
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
    filename = "6.2 - ChatGPT სრულად II - Custom GPTs, Projects, Actions და Enterprise AI.pptx"
    out_master = os.path.join(r'd:\ხელოვნური ინტელექტის ინსტრუმენტები', filename)
    prs.save(out_master)
    print(f"Master PPTX for Lecture 6.2 generated successfully: {out_master}")

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
    build_lecture_6_2()
