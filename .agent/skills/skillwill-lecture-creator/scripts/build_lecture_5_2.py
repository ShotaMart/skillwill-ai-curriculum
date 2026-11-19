import os
import sys
import shutil
import pptx
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN
from pptx.enum.shapes import MSO_SHAPE

def build_lecture_5_2():
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
            p.text = "5.2 - Claude სრულად III: MCP არქიტექტურა, Claude Code, Cowork და Enterprise AI"
            p.alignment = PP_ALIGN.CENTER
            p.font.name = FONT_ARIAL
            p.font.size = Pt(36)
            p.font.bold = True
            p.font.color.rgb = COLOR_RED

            p2 = tf.add_paragraph()
            p2.text = "ღია პროტოკოლები, ავტონომიური აგენტები, CLI ხელსაწყოები და კორპორაციული უსაფრთხოება"
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
    def add_header(slide, title, category="ᲛᲝᲓᲣᲚᲘ 5 • CLAUDE ეკოსისტემა • 5.2"):
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
    # SLIDE 2: What is MCP (Model Context Protocol)?
    # -------------------------------------------------------------------------
    s2 = created_slides[0]
    add_header(s2, "რა არის MCP (Model Context Protocol)? AI-ს უნივერსალური USB-C", "სტანდარტიზაცია • როგორ აკავშირებს ღია პროტოკოლი AI-ს ნებისმიერ სისტემასთან")

    box_w = Inches(8.35); box_h = Inches(5.4); box_y = Inches(2.35)

    # Left: Chaos before MCP
    b1 = s2.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(1.0), box_y, box_w, box_h)
    b1.fill.solid(); b1.fill.fore_color.rgb = COLOR_BG_LIGHT
    b1.line.color.rgb = COLOR_BORDER; b1.line.width = Pt(1.5)

    tb1 = s2.shapes.add_textbox(Inches(1.4), box_y + Inches(0.4), box_w - Inches(0.8), box_h - Inches(0.8))
    tf1 = tb1.text_frame; tf1.word_wrap = True; tf1.margin_left = tf1.margin_top = 0
    p = tf1.paragraphs[0]; p.text = "MCP-ᲛᲓᲔ ᲐᲠᲡᲔᲑᲣᲚᲘ ᲥᲐᲝᲡᲘ"
    p.font.name = FONT_ARIAL; p.font.size = Pt(12); p.font.bold = True; p.font.color.rgb = COLOR_MUTED

    p = tf1.add_paragraph(); p.text = "ფრაგმენტული ინტეგრაციები (N x M პრობლემა)"
    p.font.name = FONT_ARIAL; p.font.size = Pt(21); p.font.bold = True; p.font.color.rgb = COLOR_DARK; p.space_before = Pt(4)

    mcp_old_points = [
        ("ცალკეული ადაპტერები", "ყოველ ახალ AI მოდელს სჭირდებოდა ინდივიდუალური პლაგინები GitHub-ისთვის, Slack-ისთვის, მონაცემთა ბაზებისთვის."),
        ("მომწოდებელზე მიბმულობა (Vendor Lock-in)", "ერთი მოდელისთვის დაწერილი ხელსაწყო არ მუშაობდა სხვა სისტემაში."),
        ("მაღალი საინჟინრო ხარჯი", "დეველოპერები დროს ხარჯავდნენ API-ების გადაწერაზე, ნაცვლად რეალური ფუნქციონალის შექმნისა."),
        ("უსაფრთხოების რისკები", "ყოველ ინტეგრაციას ჰქონდა ავტორიზაციისა და მონაცემთა გაჟონვის განსხვავებული დაუცველი წერტილები.")
    ]
    for t, d in mcp_old_points:
        p = tf1.add_paragraph(); p.text = f"• {t}: {d}"
        p.font.name = FONT_ARIAL; p.font.size = Pt(13); p.font.color.rgb = COLOR_TEXT_MUTED; p.space_before = Pt(8)

    # Right: MCP Open Standard
    b2 = s2.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(9.65), box_y, box_w, box_h)
    b2.fill.solid(); b2.fill.fore_color.rgb = COLOR_CARD_RED
    b2.line.color.rgb = COLOR_BORDER_RED; b2.line.width = Pt(1.5)

    tb2 = s2.shapes.add_textbox(Inches(10.05), box_y + Inches(0.4), box_w - Inches(0.8), box_h - Inches(0.8))
    tf2 = tb2.text_frame; tf2.word_wrap = True; tf2.margin_left = tf2.margin_top = 0
    p = tf2.paragraphs[0]; p.text = "MCP ᲡᲢᲐᲜᲓᲐᲠᲢᲘ (2026)"
    p.font.name = FONT_ARIAL; p.font.size = Pt(12); p.font.bold = True; p.font.color.rgb = COLOR_RED

    p = tf2.add_paragraph(); p.text = "Model Context Protocol (ღია სტანდარტი)"
    p.font.name = FONT_ARIAL; p.font.size = Pt(21); p.font.bold = True; p.font.color.rgb = COLOR_DARK; p.space_before = Pt(4)

    mcp_new_points = [
        ("Open Standard (ღია ინდუსტრიული სტანდარტი)", "Anthropic-ის მიერ შექმნილი პროტოკოლი, რომელიც AI-ს აძლევს სტანდარტულ წვდომას გარე მონაცემებთან."),
        ("უნივერსალური თავსებადობა", "ერთი MCP სერვერი (მაგ. PostgreSQL) ერთნაირად ემსახურება Claude-ს, Claude Code-ს და ნებისმიერ სხვა AI ასისტენტს."),
        ("უსაფრთხო ლოკალური არხი", "კომუნიკაცია ხდება უსაფრთხო JSON-RPC არხით ლოკალურად ან დაშიფრულ ქსელში."),
        ("მზარდი ეკოსისტემა", "ათასობით მზა სერვერი (GitHub, Postgres, Slack, Linear, Google Workspace, Brave Search).")
    ]
    for t, d in mcp_new_points:
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
    p.text = "MCP არის ხელოვნური ინტელექტის „USB-C პორტი“ — როგორც ერთმა უნივერსალურმა კაბელმა ჩაანაცვლა ათობით განსხვავებული დამტენი, ისე MCP-მ გააერთიანა AI-ს კავშირი ფაილებთან, მონაცემთა ბაზებთან და სერვისებთან."
    p.font.name = FONT_ARIAL; p.font.size = Pt(13); p.font.color.rgb = COLOR_MUTED; p.space_before = Pt(4)

    # -------------------------------------------------------------------------
    # SLIDE 3: MCP 3-tier Architecture
    # -------------------------------------------------------------------------
    s3 = created_slides[1]
    add_header(s3, "MCP სამდონიანი არქიტექტურა (Host, Client და Server)", "არქიტექტურა • როგორ მოძრაობს მონაცემები უსაფრთხო პროტოკოლში")

    card_w = Inches(5.4); card_h = Inches(5.4); card_y = Inches(2.35)
    xs = [Inches(1.0), Inches(6.8), Inches(12.6)]

    cards_s3 = [
        ("💻 HOST", "MCP Host (მასპინძელი)", "სამუშაო გარემო", [
            ("სამუშაო ინტერფეისი", "აპლიკაცია, სადაც მომხმარებელი მუშაობს (Claude Desktop, Claude Code, IDE, ბრაუზერი)."),
            ("მომხმარებლის მოთხოვნის მიღება", "აფიქსირებს მომხმარებლის პრომპტს და ინიცირებს სამუშაო პროცესს."),
            ("უსაფრთხოების კონტექსტი", "აკონტროლებს წვდომის უფლებებს და მომხმარებლის თანხმობას მოქმედებების შესრულებაზე.")
        ]),
        ("🔌 CLIENT", "MCP Client (კლიენტი)", "დამაკავშირებელი ფენა", [
            ("პროტოკოლის ლოგიკა", "შუამავალი ფენა, რომელიც ინახავს 1:1 კავშირებს ხელმისაწვდომ MCP სერვერებთან."),
            ("Tool Discovery (ხელსაწყოების აღმოჩენა)", "ამოწმებს რა შესაძლებლობები აქვს თითოეულ სერვერს (Tools, Resources, Prompts)."),
            ("მოთხოვნის მარშრუტიზაცია", "მოდელის გადაწყვეტილების მიხედვით აგზავნის ბრძანებას სწორ სერვერთან.")
        ]),
        ("🗄️ SERVER", "MCP Server (სპეციალიზებული სერვერი)", "რესურსების მიმწოდებელი", [
            ("რესურსის იზოლაცია", "მსუბუქი სერვისი, რომელიც პასუხისმგებელია კონკრეტულ რესურსზე (Filesystem, PostgreSQL, GitHub, Jira)."),
            ("3 ტიპის შესაძლებლობა", "Prompts (შაბლონები), Resources (მონაცემები/ფაილები) და Tools (შემსრულებელი ფუნქციები)."),
            ("ლოკალური გაშვება", "ხშირად ეშვება პირდაპირ მომხმარებლის კომპიუტერზე, მონაცემთა გარეთ გაუსვლელად.")
        ])
    ]

    for i, (badge, title, subtitle, items) in enumerate(cards_s3):
        bg_col = COLOR_CARD_RED if i == 1 else COLOR_BG_LIGHT
        bor_col = COLOR_BORDER_RED if i == 1 else COLOR_BORDER
        c = s3.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, xs[i], card_y, card_w, card_h)
        c.fill.solid(); c.fill.fore_color.rgb = bg_col
        c.line.color.rgb = bor_col; c.line.width = Pt(1.5)

        tb = s3.shapes.add_textbox(xs[i] + Inches(0.35), card_y + Inches(0.35), card_w - Inches(0.7), card_h - Inches(0.7))
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
    bb = s3.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(1.0), bot_y, Inches(17.0), Inches(2.0))
    bb.fill.solid(); bb.fill.fore_color.rgb = COLOR_WHITE
    bb.line.color.rgb = COLOR_BORDER; bb.line.width = Pt(1.2)

    tbb = s3.shapes.add_textbox(Inches(1.4), bot_y + Inches(0.3), Inches(16.2), Inches(1.4))
    tfb = tbb.text_frame; tfb.word_wrap = True; tfb.margin_left = tfb.margin_top = 0
    p = tfb.paragraphs[0]; p.text = "💡 უსაფრთხოების გარანტია:"
    p.font.name = FONT_ARIAL; p.font.size = Pt(14); p.font.bold = True; p.font.color.rgb = COLOR_RED
    p = tfb.add_paragraph()
    p.text = "MCP არქიტექტურაში მონაცემთა ნაკადი მკაცრად კონტროლირებადია: AI მოდელს არ აქვს პირდაპირი წვდომა თქვენს კომპიუტერთან — ის ურთიერთობს მხოლოდ MCP სერვერთან, რომელიც თითოეულ ბრძანებას წინასწარ განსაზღვრული წესებით ამოწმებს."
    p.font.name = FONT_ARIAL; p.font.size = Pt(13); p.font.color.rgb = COLOR_MUTED; p.space_before = Pt(4)

    # -------------------------------------------------------------------------
    # SLIDE 4: From Chatbot to Autonomous Agent (Agentic Loop)
    # -------------------------------------------------------------------------
    s4 = created_slides[2]
    add_header(s4, "ჩატბოტიდან ავტონომიურ აგენტამდე (Agentic Architecture)", "მექანიზმი • როგორ იღებს AI გადაწყვეტილებებს და იყენებს ხელსაწყოებს")

    step_w = Inches(3.95); step_h = Inches(5.4); step_y = Inches(2.35)
    xs_s4 = [Inches(1.0), Inches(5.35), Inches(9.7), Inches(14.05)]

    steps_s4 = [
        ("1. Goal (მიზანი)", "დავალების ფორმულირება", [
            ("მაღალი დონის ბრძანება", "მომხმარებელი აძლევს საბოლოო მიზანს (მაგ. „მოაგვარე ბაგი #412 და გაუშვი ტესტები“)."),
            ("კონტექსტის გააზრება", "მოდელი აანალიზებს გარემოს, ხელმისაწვდომ ხელსაწყოებსა და შეზღუდვებს.")
        ]),
        ("2. Planning (გეგმა)", "დევიზია და ნაბიჯები", [
            ("ნაბიჯების გაწერა", "ამოცანის დაყოფა ქვეამოცანებად (ფაილების მოძიება ➔ ანალიზი ➔ კოდის ჩასწორება)."),
            ("პრიორიტეტიზაცია", "ლოგიკური თანმიმდევრობის განსაზღვრა და ალტერნატივების შეფასება.")
        ]),
        ("3. Tool Use (მოქმედება)", "ხელსაწყოების გამოძახება", [
            ("API და CLI ბრძანებები", "MCP-ით ფაილების კითხვა, ბრძანებების გაშვება ტერმინალში, მონაცემთა ბაზების ჩაწერა."),
            ("რეალური მოქმედება", "აგენტი აღარ ელოდება მომხმარებელს, თავად ასრულებს ნაბიჯ-ნაბიჯ ოპერაციებს.")
        ]),
        ("4. Self-Correction (კორექცია)", "შედეგის გადამოწმება", [
            ("შედეგის შეფასება", "გაეშვა თუ არა ტესტი? მივიღეთ თუ არა შეცდომა (Error Output)?"),
            ("უკუკავშირის ციკლი", "თუ ტესტი ჩავარდა, აგენტი აანალიზებს ერორს, ასწორებს კოდს და ცდის ხელახლა დასრულებამდე.")
        ])
    ]

    for i, (title, subtitle, items) in enumerate(steps_s4):
        bg_col = COLOR_CARD_RED if i == 2 else COLOR_BG_LIGHT
        bor_col = COLOR_BORDER_RED if i == 2 else COLOR_BORDER
        c = s4.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, xs_s4[i], step_y, step_w, step_h)
        c.fill.solid(); c.fill.fore_color.rgb = bg_col
        c.line.color.rgb = bor_col; c.line.width = Pt(1.5)

        tb = s4.shapes.add_textbox(xs_s4[i] + Inches(0.25), step_y + Inches(0.35), step_w - Inches(0.5), step_h - Inches(0.7))
        tf = tb.text_frame; tf.word_wrap = True; tf.margin_left = tf.margin_top = 0

        p = tf.paragraphs[0]; p.text = f"ᲜᲐᲑᲘᲯᲘ {i+1}"
        p.font.name = FONT_ARIAL; p.font.size = Pt(11); p.font.bold = True
        p.font.color.rgb = COLOR_RED if i == 2 else COLOR_MUTED

        p = tf.add_paragraph(); p.text = title
        p.font.name = FONT_ARIAL; p.font.size = Pt(16); p.font.bold = True; p.font.color.rgb = COLOR_DARK; p.space_before = Pt(4)

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
    p = tfb.paragraphs[0]; p.text = "🎯 აგენტური პარადიგმა:"
    p.font.name = FONT_ARIAL; p.font.size = Pt(14); p.font.bold = True; p.font.color.rgb = COLOR_RED
    p = tfb.add_paragraph()
    p.text = "ტრადიციული ჩატბოტი არის „მრჩეველი“ (მხოლოდ გეუბნებათ რა უნდა გააკეთოთ), ხოლო ავტონომიური აგენტი არის „შემსრულებელი“ — ის იღებს მიზანს, იყენებს ხელსაწყოებს, ამოწმებს საკუთარ თავს და თავად აგვარებს პრობლემას."
    p.font.name = FONT_ARIAL; p.font.size = Pt(13); p.font.color.rgb = COLOR_MUTED; p.space_before = Pt(4)

    # -------------------------------------------------------------------------
    # SLIDE 5: Claude Code - ტერმინალის ავტონომიური აგენტი
    # -------------------------------------------------------------------------
    s5 = created_slides[3]
    add_header(s5, "Claude Code (ტერმინალის ავტონომიური დეველოპერი)", "ინსტრუმენტი • CLI აგენტი რთული პროგრამული პროექტების სამართავად")

    cards_s5 = [
        ("💻 CODEBASE", "Full Codebase Understanding", "პროექტის სრული აღქმა", [
            ("რეპოზიტორიის მასშტაბი", "კითხულობს და იაზრებს ათასობით ფაილს, არქიტექტურულ დამოკიდებულებებსა და კონფიგურაციებს."),
            ("ჭკვიანი ინდექსაცია", "სწრაფად პოულობს შესაბამის ფუნქციებს, კლასებსა და ლოგიკურ კავშირებს მთელ პროექტში."),
            ("Git კონტექსტი", "აანალიზებს Git ისტორიას, ბოლო Commit-ებსა და განშტოებების (Branches) სტრუქტურას.")
        ]),
        ("⚙️ EXECUTION", "Autonomous Coding & Execution", "ავტონომიური რედაქტირება", [
            ("ფაილების პირდაპირი მოდიფიკაცია", "თავად ხსნის, არედაქტირებს და ქმნის ფაილებს თქვენს ლოკალურ დისკზე."),
            ("ტერმინალის ბრძანებების გაშვება", "უშვებს Build პროცესებს, Linters, Type-checkers და ავტომატურ Unit ტესტებს."),
            ("ბაგების ავტო-შესწორება", "თუ ტესტმა შეცდომა აჩვენა, აგენტი პირდაპირ ტერმინალშივე ასწორებს კოდს და ამოწმებს ხელახლა.")
        ]),
        ("🐙 WORKFLOW", "Git Workflow & Pull Requests", "ვერსიათა კონტროლი", [
            ("განშტოებების მართვა", "ქმნის ახალ Branch-ს კონკრეტული ამოცანისთვის სუფთა იზოლაციაში."),
            ("აზრიანი Commit-ები", "აჯამებს შესრულებულ სამუშაოს და აკეთებს მკაფიო, ინფორმაციულ Commit შეტყობინებებს."),
            ("Pull Request-ის მომზადება", "GitHub/GitLab-ზე ამზადებს სრულყოფილ PR დოკუმენტაციას კოლეგების განხილვისთვის.")
        ])
    ]

    for i, (badge, title, subtitle, items) in enumerate(cards_s5):
        bg_col = COLOR_CARD_RED if i == 1 else COLOR_BG_LIGHT
        bor_col = COLOR_BORDER_RED if i == 1 else COLOR_BORDER
        c = s5.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, xs[i], card_y, card_w, card_h)
        c.fill.solid(); c.fill.fore_color.rgb = bg_col
        c.line.color.rgb = bor_col; c.line.width = Pt(1.5)

        tb = s5.shapes.add_textbox(xs[i] + Inches(0.35), card_y + Inches(0.35), card_w - Inches(0.7), card_h - Inches(0.7))
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
    p = tfb.paragraphs[0]; p.text = "💡 2026 წლის საინჟინრო რეალობა:"
    p.font.name = FONT_ARIAL; p.font.size = Pt(14); p.font.bold = True; p.font.color.rgb = COLOR_RED
    p = tfb.add_paragraph()
    p.text = "Claude Code აქცევს დეველოპერს არქიტექტორად: ნაცვლად რუტინული კოდის წერისა, ინჟინერი აძლევს Claude Code-ს მაღალი დონის ამოცანებს, აკონტროლებს მის მუშაობას და ამოწმებს საბოლოო შედეგს."
    p.font.name = FONT_ARIAL; p.font.size = Pt(13); p.font.color.rgb = COLOR_MUTED; p.space_before = Pt(4)

    # -------------------------------------------------------------------------
    # SLIDE 6: Claude Cowork - ადამიანისა და AI-ს თანამშრომლობა
    # -------------------------------------------------------------------------
    s6 = created_slides[4]
    add_header(s6, "Claude Cowork (ადამიანისა და AI-ს ერთობლივი სამუშაო სივრცე)", "კოლაბორაცია • მულტი-აგენტური დავალებები და ინტერაქტიული სამუშაო მაგიდა")

    cards_s6 = [
        ("🖥️ CANVAS", "Shared Context Canvas", "ინტერაქტიული მაგიდა", [
            ("სინქრონული სამუშაო სივრცე", "ადამიანი და AI ერთდროულად ხედავენ და ამუშავებენ პროექტის მიმდინარეობას."),
            ("ვიზუალური დაფა", "დოკუმენტების მონახაზები, დიაგრამები და სტრუქტურები იქმნება რეალურ დროში."),
            ("უწყვეტი ინტერაქცია", "მომხმარებელს შეუძლია ნებისმიერ მომენტში შეასწოროს დეტალი ან მიმართულება მისცეს აგენტს.")
        ]),
        ("⏳ BACKGROUND", "Background Task Execution", "ფონური პროცესები", [
            ("ხანგრძლივი ამოცანების დელეგირება", "აგენტს შეგიძლიათ დაავალოთ 2-საათიანი კვლევა ან დიდი ბაზის ანალიზი."),
            ("პარალელური მუშაობა", "სანამ აგენტი ამუშავებს მონაცემებს ფონში, თქვენ აგრძელებთ სხვა დავალებების შესრულებას."),
            ("შეტყობინება დასრულებაზე", "სამუშაოს დასრულებისთანავე Claude გაწვდით მზა, სტრუქტურირებულ რეზიუმეს.")
        ]),
        ("🛡️ CONTROL", "Human-in-the-Loop Safeguard", "ადამიანის კონტროლი", [
            ("კრიტიკული მოქმედებების დადასტურება", "სანამ აგენტი წაშლის ფაილს, გააგზავნის წერილს ან განახორციელებს გადახდას, ითხოვს ადამიანის თანხმობას."),
            ("გადაწყვეტილების გამჭვირვალობა", "აგენტი ხსნის, რატომ ირჩევს ამ ნაბიჯს და რა რისკები ახლავს მას."),
            ("სრული მართვადობა", "კონტროლი ყოველთვის რჩება ადამიანის ხელში.")
        ])
    ]

    for i, (badge, title, subtitle, items) in enumerate(cards_s6):
        bg_col = COLOR_CARD_RED if i == 0 else COLOR_BG_LIGHT
        bor_col = COLOR_BORDER_RED if i == 0 else COLOR_BORDER
        c = s6.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, xs[i], card_y, card_w, card_h)
        c.fill.solid(); c.fill.fore_color.rgb = bg_col
        c.line.color.rgb = bor_col; c.line.width = Pt(1.5)

        tb = s6.shapes.add_textbox(xs[i] + Inches(0.35), card_y + Inches(0.35), card_w - Inches(0.7), card_h - Inches(0.7))
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
    bb = s6.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(1.0), bot_y, Inches(17.0), Inches(2.0))
    bb.fill.solid(); bb.fill.fore_color.rgb = COLOR_WHITE
    bb.line.color.rgb = COLOR_BORDER; bb.line.width = Pt(1.2)

    tbb = s6.shapes.add_textbox(Inches(1.4), bot_y + Inches(0.3), Inches(16.2), Inches(1.4))
    tfb = tbb.text_frame; tfb.word_wrap = True; tfb.margin_left = tfb.margin_top = 0
    p = tfb.paragraphs[0]; p.text = "💡 პარადიგმის ცვლა:"
    p.font.name = FONT_ARIAL; p.font.size = Pt(14); p.font.bold = True; p.font.color.rgb = COLOR_RED
    p = tfb.add_paragraph()
    p.text = "Claude Cowork-ით AI გამოდის მარტივი კითხვა-პასუხის რეჟიმიდან და ხდება თქვენი სრულფასოვანი ციფრული კოლეგა, რომელთანაც იყოფთ სამუშაო დატვირთვას."
    p.font.name = FONT_ARIAL; p.font.size = Pt(13); p.font.color.rgb = COLOR_MUTED; p.space_before = Pt(4)

    # -------------------------------------------------------------------------
    # SLIDE 7: Constitutional AI და უსაფრთხოების მექანიზმები
    # -------------------------------------------------------------------------
    s7 = created_slides[5]
    add_header(s7, "Constitutional AI და მოდელის უსაფრთხოების ჩარჩო", "უსაფრთხოება • როგორ აკონტროლებს Anthropic მოდელის ეთიკასა და სანდოობას")

    # Left: Standard RLHF
    b1 = s7.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(1.0), box_y, box_w, box_h)
    b1.fill.solid(); b1.fill.fore_color.rgb = COLOR_BG_LIGHT
    b1.line.color.rgb = COLOR_BORDER; b1.line.width = Pt(1.5)

    tb1 = s7.shapes.add_textbox(Inches(1.4), box_y + Inches(0.4), box_w - Inches(0.8), box_h - Inches(0.8))
    tf1 = tb1.text_frame; tf1.word_wrap = True; tf1.margin_left = tf1.margin_top = 0
    p = tf1.paragraphs[0]; p.text = "ᲢᲠᲐᲓᲘᲪᲘᲣᲚᲘ ᲛᲘᲓᲒᲝᲛᲐ"
    p.font.name = FONT_ARIAL; p.font.size = Pt(12); p.font.bold = True; p.font.color.rgb = COLOR_MUTED

    p = tf1.add_paragraph(); p.text = "Standard RLHF (ადამიანის გამოხმაურება)"
    p.font.name = FONT_ARIAL; p.font.size = Pt(21); p.font.bold = True; p.font.color.rgb = COLOR_DARK; p.space_before = Pt(4)

    rlhf_points = [
        ("ადამიანური შეფასების ლიმიტი", "მოდელის გაწვრთნა ეყრდნობა ათასობით შემფასებლის სუბიექტურ მოსაზრებას."),
        ("შეუსაბამობა და არაერთგვაროვნება", "სხვადასხვა ადამიანს განსხვავებული ეთიკური და ლოგიკური სტანდარტები აქვს."),
        ("მაღალი ხარჯი და დაბალი ტემპი", "რთული სამეცნიერო და უსაფრთხოების ქცევების მასშტაბირება ადამიანებით თითქმის შეუძლებელია."),
        ("სიყალბის რისკი (Sycophancy)", "მოდელი ცდილობს ასიამოვნოს მომხმარებელს, მაშინაც კი როცა მომხმარებელი ცდება.")
    ]
    for t, d in rlhf_points:
        p = tf1.add_paragraph(); p.text = f"• {t}: {d}"
        p.font.name = FONT_ARIAL; p.font.size = Pt(13); p.font.color.rgb = COLOR_TEXT_MUTED; p.space_before = Pt(8)

    # Right: Constitutional AI
    b2 = s7.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(9.65), box_y, box_w, box_h)
    b2.fill.solid(); b2.fill.fore_color.rgb = COLOR_CARD_RED
    b2.line.color.rgb = COLOR_BORDER_RED; b2.line.width = Pt(1.5)

    tb2 = s7.shapes.add_textbox(Inches(10.05), box_y + Inches(0.4), box_w - Inches(0.8), box_h - Inches(0.8))
    tf2 = tb2.text_frame; tf2.word_wrap = True; tf2.margin_left = tf2.margin_top = 0
    p = tf2.paragraphs[0]; p.text = "ANTHROPIC-ᲘᲡ ᲘᲜᲝᲕᲐᲪᲘᲐ"
    p.font.name = FONT_ARIAL; p.font.size = Pt(12); p.font.bold = True; p.font.color.rgb = COLOR_RED

    p = tf2.add_paragraph(); p.text = "Constitutional AI (RLAIF - AI-ს თვითკრიტიკა)"
    p.font.name = FONT_ARIAL; p.font.size = Pt(21); p.font.bold = True; p.font.color.rgb = COLOR_DARK; p.space_before = Pt(4)

    cai_points = [
        ("მკაფიო კონსტიტუცია", "გაწერილი წესები, რომლებიც ეფუძნება გაეროს ადამიანის უფლებათა დეკლარაციასა და AI უსაფრთხოების პრინციპებს."),
        ("ორეტაპიანი თვითკრიტიკა (Self-Critique)", "მოდელი თავად აფასებს საკუთარ პასუხებს კონსტიტუციის მიხედვით და ასწორებს შეცდომებს."),
        ("მავნე კონტენტის ბლოკირება", "კიბერთავდასხმების, ბიოლოგიური საფრთხეებისა და დეზინფორმაციის მყარი დაცვა."),
        ("პატიოსნება და ობიექტურობა", "მოდელი არ მალავს ფაქტებს და გულწრფელად აღიარებს უცნობ დეტალებს.")
    ]
    for t, d in cai_points:
        p = tf2.add_paragraph(); p.text = f"⭐ {t}"
        p.font.name = FONT_ARIAL; p.font.size = Pt(13); p.font.bold = True; p.font.color.rgb = COLOR_DARK; p.space_before = Pt(8)
        p = tf2.add_paragraph(); p.text = d
        p.font.name = FONT_ARIAL; p.font.size = Pt(12); p.font.color.rgb = COLOR_TEXT_MUTED; p.space_before = Pt(2)

    # Bottom takeaway
    bb = s7.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(1.0), bot_y, Inches(17.0), Inches(2.0))
    bb.fill.solid(); bb.fill.fore_color.rgb = COLOR_WHITE
    bb.line.color.rgb = COLOR_BORDER; bb.line.width = Pt(1.2)

    tbb = s7.shapes.add_textbox(Inches(1.4), bot_y + Inches(0.3), Inches(16.2), Inches(1.4))
    tfb = tbb.text_frame; tfb.word_wrap = True; tfb.margin_left = tfb.margin_top = 0
    p = tfb.paragraphs[0]; p.text = "💡 უსაფრთხოების შედეგი:"
    p.font.name = FONT_ARIAL; p.font.size = Pt(14); p.font.bold = True; p.font.color.rgb = COLOR_RED
    p = tfb.add_paragraph()
    p.text = "Constitutional AI-ს წყალობით Claude ითვლება მსოფლიოში ყველაზე მდგრად, ეთიკურ და პროგნოზირებად მოდელად, რაც კრიტიკულად მნიშვნელოვანია მსხვილი ორგანიზაციებისთვის."
    p.font.name = FONT_ARIAL; p.font.size = Pt(13); p.font.color.rgb = COLOR_MUTED; p.space_before = Pt(4)

    # -------------------------------------------------------------------------
    # SLIDE 8: Enterprise Data Policy და Zero Data Retention
    # -------------------------------------------------------------------------
    s8 = created_slides[6]
    add_header(s8, "Enterprise მონაცემთა პოლიტიკა და Zero Data Retention", "კონფიდენციალურობა • სად მიდის და როგორ არის დაცული თქვენი მონაცემები")

    cards_s8 = [
        ("🚫 NO TRAINING", "Zero Training on Customer Data", "მონაცემთა დაცვა", [
            ("მკაცრი აკრძალვა", "Anthropic არ იყენებს ბიზნეს მომხმარებლების მოთხოვნებს, ფაილებსა და კოდს მომავალი AI მოდელების საწვრთნელად."),
            ("ინტელექტუალური საკუთრების გარანტია", "თქვენი საავტორო უფლებები, შიდა კოდი და ანალიტიკა რჩება 100%-ით თქვენს მფლობელობაში."),
            ("კომერციული საიდუმლოება", "გამორიცხულია კონფიდენციალური მონაცემების სხვა მომხმარებლის პასუხში გაჟონვა.")
        ]),
        ("🗑️ ZERO RETENTION", "Zero Data Retention (ZDR)", "ნულოვანი შენახვა", [
            ("მყისიერი წაშლა", "მოთხოვნის დამუშავებისა და პასუხის დაბრუნებისთანავე სერვერებიდან იშლება დროებითი ტოკენები."),
            ("ლოგების გარეშე", "მოთხოვნის შინაარსი არ ინახება მუდმივ მონაცემთა ბაზებში ან სერვერის ლოგებში."),
            ("მარეგულირებელი სტანდარტები", "აკმაყოფილებს ყველაზე მკაცრ საბანკო და სამედიცინო მოთხოვნებს.")
        ]),
        ("🔐 ENCRYPTION", "Enterprise-Grade Encryption", "დაშიფვრა ყველა დონეზე", [
            ("Encryption in Transit", "ყველა მონაცემი იშიფრება გადაცემის დროს თანამედროვე TLS 1.3 პროტოკოლით."),
            ("Encryption at Rest", "შენახული მონაცემები დაცულია სამხედრო დონის AES-256 დაშიფვრით."),
            ("გასაღებების მართვა (KMS)", "შესაძლებელია კომპანიის საკუთარი დაშიფვრის გასაღებების (Customer-Managed Keys) გამოყენება.")
        ])
    ]

    for i, (badge, title, subtitle, items) in enumerate(cards_s8):
        bg_col = COLOR_CARD_RED if i == 0 else COLOR_BG_LIGHT
        bor_col = COLOR_BORDER_RED if i == 0 else COLOR_BORDER
        c = s8.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, xs[i], card_y, card_w, card_h)
        c.fill.solid(); c.fill.fore_color.rgb = bg_col
        c.line.color.rgb = bor_col; c.line.width = Pt(1.5)

        tb = s8.shapes.add_textbox(xs[i] + Inches(0.35), card_y + Inches(0.35), card_w - Inches(0.7), card_h - Inches(0.7))
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
    bb = s8.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(1.0), bot_y, Inches(17.0), Inches(2.0))
    bb.fill.solid(); bb.fill.fore_color.rgb = COLOR_WHITE
    bb.line.color.rgb = COLOR_BORDER; bb.line.width = Pt(1.2)

    tbb = s8.shapes.add_textbox(Inches(1.4), bot_y + Inches(0.3), Inches(16.2), Inches(1.4))
    tfb = tbb.text_frame; tfb.word_wrap = True; tfb.margin_left = tfb.margin_top = 0
    p = tfb.paragraphs[0]; p.text = "💡 კორპორაციული ნდობის საფუძველი:"
    p.font.name = FONT_ARIAL; p.font.size = Pt(14); p.font.bold = True; p.font.color.rgb = COLOR_RED
    p = tfb.add_paragraph()
    p.text = "ბანკებს, საავადმყოფოებს, სახელმწიფო უწყებებსა და იურიდიულ ფირმებს შეუძლიათ მშვიდად ენდონ Claude-ს, რადგან მათი მონაცემები არასოდეს ხდება საჯარო და არ გამოიყენება AI-ს სწავლებისთვის."
    p.font.name = FONT_ARIAL; p.font.size = Pt(13); p.font.color.rgb = COLOR_MUTED; p.space_before = Pt(4)

    # -------------------------------------------------------------------------
    # SLIDE 9: Claude Enterprise არქიტექტურა
    # -------------------------------------------------------------------------
    s9 = created_slides[7]
    add_header(s9, "Claude Enterprise არქიტექტურა (კორპორაციული კონტროლი)", "ინფრასტრუქტურა • ადმინისტრირება, შესაბამისობა და გლობალური სტანდარტები")

    pillars_s9 = [
        ("1. SSO & SCIM Directory", "მომხმარებელთა მართვა", [
            ("Single Sign-On (SSO)", "ერთიანი უსაფრთხო შესვლა Okta, Google Workspace, Microsoft Entra ID-ით."),
            ("SCIM Provisioning", "თანამშრომლის დამატების ან წასვლისას ანგარიშების ავტომატური მართვა და წვდომების ჩამოხსნა.")
        ]),
        ("2. Audit Logs & SIEM", "უსაფრთხოების მონიტორინგი", [
            ("დეტალური აუდიტის ლოგები", "ვინ, როდის, რომელი IP მისამართიდან და რა ფაილით ისარგებლა Claude-ში."),
            ("SIEM ინტეგრაცია", "ლოგების პირდაპირი ექსპორტი Splunk, Datadog ან Azure Sentinel-ში კიბერმონიტორინგისთვის.")
        ]),
        ("3. Role-Based Access (RBAC)", "იერარქიული წვდომები", [
            ("დეპარტამენტების იზოლაცია", "ფინანსურ პროექტებზე წვდომა აქვთ მხოლოდ ფინანსისტებს, დეველოპერულზე — ინჟინრებს."),
            ("ადმინისტრატორის კონტროლი", "ცენტრალიზებული პოლიტიკა Connectors-ისა და MCP სერვერების დასაშვებად.")
        ]),
        ("4. Compliance Standards", "საერთაშორისო აუდიტი", [
            ("SOC 2 Type II", "დამოუკიდებელი აუდიტი უსაფრთხოებასა და კონფიდენციალურობაზე."),
            ("HIPAA & GDPR Compliance", "ევროპული და ამერიკული პერსონალური მონაცემების დაცვის კანონმდებლობასთან სრული თანხვედრა.")
        ])
    ]

    for i, (title, subtitle, items) in enumerate(pillars_s9):
        bg_col = COLOR_CARD_RED if i == 0 else COLOR_BG_LIGHT
        bor_col = COLOR_BORDER_RED if i == 0 else COLOR_BORDER
        c = s9.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, xs_s4[i], step_y, step_w, step_h)
        c.fill.solid(); c.fill.fore_color.rgb = bg_col
        c.line.color.rgb = bor_col; c.line.width = Pt(1.5)

        tb = s9.shapes.add_textbox(xs_s4[i] + Inches(0.25), step_y + Inches(0.35), step_w - Inches(0.5), step_h - Inches(0.7))
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
    bb = s9.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(1.0), bot_y, Inches(17.0), Inches(2.0))
    bb.fill.solid(); bb.fill.fore_color.rgb = COLOR_WHITE
    bb.line.color.rgb = COLOR_BORDER; bb.line.width = Pt(1.2)

    tbb = s9.shapes.add_textbox(Inches(1.4), bot_y + Inches(0.3), Inches(16.2), Inches(1.4))
    tfb = tbb.text_frame; tfb.word_wrap = True; tfb.margin_left = tfb.margin_top = 0
    p = tfb.paragraphs[0]; p.text = "🎯 Enterprise შეჯამება:"
    p.font.name = FONT_ARIAL; p.font.size = Pt(14); p.font.bold = True; p.font.color.rgb = COLOR_RED
    p = tfb.add_paragraph()
    p.text = "Claude Enterprise აძლევს მსხვილ ორგანიზაციებს სრულ კონტროლს, გამჭვირვალობასა და უსაფრთხოებას, რაც საშუალებას აძლევს AI დანერგონ ორგანიზაციის ყველა დეპარტამენტში რისკების გარეშე."
    p.font.name = FONT_ARIAL; p.font.size = Pt(13); p.font.color.rgb = COLOR_MUTED; p.space_before = Pt(4)

    # -------------------------------------------------------------------------
    # FINAL POST-PROCESSING: Strict 100% Arial & Clean Notes
    # -------------------------------------------------------------------------
    for slide in prs.slides:
        # Clear notes
        if slide.has_notes_slide:
            notes_tf = slide.notes_slide.notes_text_frame
            notes_tf.text = ""

        # Enforce Arial everywhere
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
    filename = "5.2 - Claude სრულად III - MCP არქიტექტურა, Claude Code, Cowork და Enterprise AI.pptx"
    out_master = os.path.join(r'd:\ხელოვნური ინტელექტის ინსტრუმენტები', filename)
    prs.save(out_master)
    print(f"Master PPTX for Lecture 5.2 generated successfully: {out_master}")

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
    build_lecture_5_2()
