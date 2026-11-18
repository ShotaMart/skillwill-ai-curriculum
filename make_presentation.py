import os
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN
from pptx.enum.shapes import MSO_SHAPE

def build_visual_presentation():
    template_path = 'd:/ხელოვნური ინტელექტის ინსტრუმენტები/Skillwill Presentation Example .pptx'
    prs = Presentation(template_path)

    # Assets
    asset_dir = 'd:/ხელოვნური ინტელექტის ინსტრუმენტები/.template_assets'
    logo_dir = os.path.join(asset_dir, 'logos')

    img_corner_bl = os.path.join(asset_dir, 'image2.png')
    img_corner_tr = os.path.join(asset_dir, 'image3.png')
    img_icon_title = os.path.join(asset_dir, 'image5.png')

    logo_chatgpt = os.path.join(logo_dir, 'chatgpt.png')
    logo_claude = os.path.join(logo_dir, 'claude.png')
    logo_gemini = os.path.join(logo_dir, 'gemini.png')
    logo_copilot = os.path.join(logo_dir, 'copilot.png')

    # Brand Colors
    COLOR_RED = RGBColor(213, 43, 30)       # #D52B1E
    COLOR_DARK = RGBColor(17, 24, 39)        # #111827
    COLOR_MUTED = RGBColor(107, 114, 128)    # #6B7280
    COLOR_BORDER = RGBColor(226, 232, 240)   # #E2E8F0
    COLOR_WHITE = RGBColor(255, 255, 255)
    COLOR_CARD_RED = RGBColor(254, 242, 242)
    COLOR_BORDER_RED = RGBColor(254, 202, 202)

    FONT_TITLE = 'TBC Contractica CAPS Bold'
    FONT_BODY = 'Segoe UI'

    # Slide 1: Cover Slide
    s1 = prs.slides[0]
    for shape in s1.shapes:
        if shape.has_text_frame and shape.text_frame.text.strip():
            tf = shape.text_frame
            tf.clear()
            p = tf.paragraphs[0]
            p.text = "რა არის AI და პლატფორმების 4 კატეგორია"
            p.alignment = PP_ALIGN.CENTER
            p.font.name = FONT_TITLE
            p.font.size = Pt(46)
            p.font.bold = True
            p.font.color.rgb = COLOR_RED

            p2 = tf.add_paragraph()
            p2.text = "მოდული 5: შესავალი ხელოვნურ ინტელექტში • ლექცია #1\nშოთა მარტყოფლიშვილი | Skillwill 2026"
            p2.alignment = PP_ALIGN.CENTER
            p2.font.name = FONT_BODY
            p2.font.size = Pt(22)
            p2.font.color.rgb = COLOR_MUTED
            p2.space_before = Pt(18)
            break

    # Remove template slides 2 & 3
    rId2 = prs.slides._sldIdLst[1].rId
    prs.part.drop_rel(rId2)
    del prs.slides._sldIdLst[1]

    rId3 = prs.slides._sldIdLst[1].rId
    prs.part.drop_rel(rId3)
    del prs.slides._sldIdLst[1]

    blank_layout = prs.slide_layouts[6]

    def add_header(slide, title, category="ᲛᲝᲓᲣᲚᲘ 5 • ᲨᲔᲡᲐᲕᲐᲚᲘ AI"):
        if os.path.exists(img_corner_tr):
            slide.shapes.add_picture(img_corner_tr, Inches(18.3), Inches(0), width=Inches(1.7))
        if os.path.exists(img_icon_title):
            slide.shapes.add_picture(img_icon_title, Inches(1.2), Inches(0.92), width=Inches(0.48))

        tb = slide.shapes.add_textbox(Inches(1.8), Inches(0.65), Inches(16.0), Inches(1.2))
        tf = tb.text_frame
        tf.word_wrap = True
        tf.margin_left = tf.margin_top = tf.margin_right = tf.margin_bottom = 0

        p_cat = tf.paragraphs[0]
        p_cat.text = category.upper()
        p_cat.font.name = FONT_TITLE
        p_cat.font.size = Pt(13)
        p_cat.font.bold = True
        p_cat.font.color.rgb = COLOR_RED

        p_t = tf.add_paragraph()
        p_t.text = title
        p_t.font.name = FONT_TITLE
        p_t.font.size = Pt(34)
        p_t.font.bold = True
        p_t.font.color.rgb = COLOR_DARK
        p_t.space_before = Pt(4)

        line = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(1.2), Inches(2.0), Inches(17.6), Inches(0.02))
        line.fill.solid()
        line.fill.fore_color.rgb = COLOR_BORDER
        line.line.color.rgb = COLOR_BORDER

        foot = slide.shapes.add_textbox(Inches(1.2), Inches(10.3), Inches(17.6), Inches(0.4))
        ftf = foot.text_frame
        ftf.word_wrap = True
        p_foot = ftf.paragraphs[0]
        p_foot.text = "SKILLWILL • ხელოვნური ინტელექტის ინსტრუმენტები | ლექტორი: შოთა მარტყოფლიშვილი"
        p_foot.font.name = FONT_BODY
        p_foot.font.size = Pt(13)
        p_foot.font.color.rgb = COLOR_MUTED

    # -------------------------------------------------------------------------
    # SLIDE 2: რა არის AI? (ცნებების იერარქია)
    # -------------------------------------------------------------------------
    s2 = prs.slides.add_slide(blank_layout)
    add_header(s2, "რა არის AI? — ცნებების იერარქია და ევოლუცია")

    concepts = [
        {
            "tag": "ქოლგა ცნება",
            "name": "Artificial Intelligence (AI)",
            "sub": "მანქანის უნარი, მოახდინოს ადამიანური ინტელექტუალური ქცევის სიმულაცია — აღქმა, სწავლა და ამოცანების ავტონომიური გადაწყვეტა.",
            "color_border": RGBColor(203, 213, 225),
            "color_bg": RGBColor(248, 250, 252),
            "accent": RGBColor(71, 85, 105)
        },
        {
            "tag": "ქვედარგი",
            "name": "Machine Learning (ML)",
            "sub": "ალგორითმები და სტატისტიკური მოდელები, რომლებიც სწავლობენ მონაცემებიდან მკაცრი წესების ხელით გაწერის გარეშე.",
            "color_border": RGBColor(147, 197, 253),
            "color_bg": RGBColor(239, 246, 255),
            "accent": RGBColor(37, 99, 235)
        },
        {
            "tag": "დღევანდელი რევოლუცია",
            "name": "Generative AI & LLM",
            "sub": "ტრანსფორმერული ნეიროქსელები, რომლებიც აგენერირებენ ახალ შინაარსს: ტექსტს, პროგრამულ კოდს, გამოსახულებასა და აგენტურ ლოგიკას.",
            "color_border": RGBColor(252, 165, 165),
            "color_bg": RGBColor(254, 242, 242),
            "accent": COLOR_RED
        }
    ]

    card_w = Inches(5.6)
    gap = Inches(0.4)
    start_x = Inches(1.2)
    card_y = Inches(2.35)
    card_h = Inches(4.3)

    for i, c in enumerate(concepts):
        cx = start_x + i * (card_w + gap)
        card = s2.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, cx, card_y, card_w, card_h)
        card.fill.solid()
        card.fill.fore_color.rgb = c['color_bg']
        card.line.color.rgb = c['color_border']
        card.line.width = Pt(1.5)

        tb = s2.shapes.add_textbox(cx + Inches(0.4), card_y + Inches(0.4), card_w - Inches(0.8), card_h - Inches(0.8))
        tf = tb.text_frame
        tf.word_wrap = True

        p_tag = tf.paragraphs[0]
        p_tag.text = c['tag'].upper()
        p_tag.font.name = FONT_TITLE
        p_tag.font.size = Pt(13)
        p_tag.font.bold = True
        p_tag.font.color.rgb = c['accent']

        p_name = tf.add_paragraph()
        p_name.text = c['name']
        p_name.font.name = FONT_TITLE
        p_name.font.size = Pt(24)
        p_name.font.bold = True
        p_name.font.color.rgb = COLOR_DARK
        p_name.space_before = Pt(8)

        p_sub = tf.add_paragraph()
        p_sub.text = c['sub']
        p_sub.font.name = FONT_BODY
        p_sub.font.size = Pt(17)
        p_sub.font.color.rgb = RGBColor(55, 65, 81)
        p_sub.space_before = Pt(14)

    # 5 Signs of Intelligence
    banner = s2.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(1.2), Inches(6.9), Inches(17.6), Inches(3.0))
    banner.fill.solid()
    banner.fill.fore_color.rgb = COLOR_WHITE
    banner.line.color.rgb = COLOR_BORDER
    banner.line.width = Pt(1.5)

    btb = s2.shapes.add_textbox(Inches(1.6), Inches(7.1), Inches(16.8), Inches(2.6))
    btf = btb.text_frame
    btf.word_wrap = True

    p_bt = btf.paragraphs[0]
    p_bt.text = "ინტელექტუალური ქცევის 5 საკვანძო ნიშანი (რას აკეთებს AI):"
    p_bt.font.name = FONT_TITLE
    p_bt.font.size = Pt(20)
    p_bt.font.bold = True
    p_bt.font.color.rgb = COLOR_DARK

    signs = [
        ("1. დასწავლა (Learning)", "გამოცდილებიდან და მონაცემებიდან შაბლონებისა და წესების ათვისება"),
        ("2. მსჯელობა (Reasoning)", "ლოგიკური დასკვნების გამოტანა არსებული ფაქტებისა და კონტექსტის საფუძველზე"),
        ("3. პრობლემის გადაჭრა", "მიზნის მისაღწევად ოპტიმალური სტრატეგიისა და ალგორითმის შერჩევა"),
        ("4. აღქმა (Perception)", "მრავალმოდალური ინფორმაციის (ტექსტი, გამოსახულება, ხმა, გარემო) გაანალიზება"),
        ("5. ენობრივი კომუნიკაცია", "ბუნებრივი ადამიანური ენის (ქართული, ინგლისური) გაგება, ინტერპრეტაცია და გენერაცია")
    ]

    for title_s, desc_s in signs:
        p_s = btf.add_paragraph()
        p_s.text = f"•  {title_s}: {desc_s}"
        p_s.font.name = FONT_BODY
        p_s.font.size = Pt(16.5)
        p_s.font.color.rgb = COLOR_DARK
        p_s.space_before = Pt(4)

    # -------------------------------------------------------------------------
    # SLIDE 3: AI პლატფორმების 4 კატეგორია
    # -------------------------------------------------------------------------
    s3 = prs.slides.add_slide(blank_layout)
    add_header(s3, "AI პლატფორმების 4 კატეგორია — სად და როგორ ვიყენებთ?", "პლატფორმების კლასიფიკაცია")

    categories = [
        {
            "num": "01",
            "title": "Consumer AI ჩატბოტები",
            "target": "ყოველდღიური ინდივიდუალური ამოცანები",
            "desc": "მზა ვებ-აპლიკაციები სწრაფი კითხვა-პასუხისთვის, ტექსტების რედაქტირებისთვის, კვლევისა და იდეების გენერაციისთვის.",
            "examples": "ChatGPT Free/Plus, Claude.ai, Gemini Web",
            "color_border": RGBColor(147, 197, 253),
            "badge_color": RGBColor(37, 99, 235)
        },
        {
            "num": "02",
            "title": "Enterprise Copilots",
            "target": "კორპორატიული სამუშაო გარემო",
            "desc": "საოფისე პროგრამებში ჩაშენებული, უსაფრთხო ასისტენტები კომპანიის შიდა დოკუმენტებთან წვდომით (SSO, Audit, Data Privacy).",
            "examples": "Microsoft 365 Copilot, Google Workspace Gemini, ChatGPT Team/Enterprise",
            "color_border": RGBColor(167, 243, 208),
            "badge_color": RGBColor(5, 150, 105)
        },
        {
            "num": "03",
            "title": "Cloud Developer Platforms",
            "target": "ინჟინრები, ანალიტიკოსები და API",
            "desc": "დეველოპერული გარემო, სადაც სრულად კონტროლდება პარამეტრები (System Prompt, Temperature, JSON Mode, Function Calling).",
            "examples": "Google AI Studio, Anthropic Console, OpenAI Platform, Azure AI Foundry",
            "color_border": RGBColor(253, 230, 138),
            "badge_color": RGBColor(217, 119, 6)
        },
        {
            "num": "04",
            "title": "Open-Weights / Local AI",
            "target": "სრული კონფიდენციალურობა და იზოლაცია",
            "desc": "ღია მოდელები, რომლებიც ეშვება საკუთარ კომპიუტერზე ან ლოკალურ სერვერზე ინტერნეტის გარეშე, მონაცემთა 100% დაცვით.",
            "examples": "Meta LLaMA 3, DeepSeek-R1, Mistral, Ollama, LM Studio",
            "color_border": RGBColor(252, 165, 165),
            "badge_color": COLOR_RED
        }
    ]

    card_h = Inches(1.72)
    gap_y = Inches(0.2)
    start_y = Inches(2.35)

    for i, cat in enumerate(categories):
        cy = start_y + i * (card_h + gap_y)

        card = s3.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(1.2), cy, Inches(17.6), card_h)
        card.fill.solid()
        card.fill.fore_color.rgb = COLOR_WHITE
        card.line.color.rgb = cat['color_border']
        card.line.width = Pt(1.5)

        # Left Number Badge
        nb = s3.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(1.5), cy + Inches(0.32), Inches(1.08), Inches(1.08))
        nb.fill.solid()
        nb.fill.fore_color.rgb = cat['badge_color']
        nb.line.fill.background()
        ntf = nb.text_frame
        p_n = ntf.paragraphs[0]
        p_n.text = cat['num']
        p_n.font.name = FONT_TITLE
        p_n.font.size = Pt(26)
        p_n.font.bold = True
        p_n.font.color.rgb = COLOR_WHITE
        p_n.alignment = PP_ALIGN.CENTER

        # Content Box
        tb = s3.shapes.add_textbox(Inches(2.85), cy + Inches(0.12), Inches(15.6), Inches(1.5))
        tf = tb.text_frame
        tf.word_wrap = True

        p_t = tf.paragraphs[0]
        p_t.text = f"{cat['title']}   —   [{cat['target']}]"
        p_t.font.name = FONT_TITLE
        p_t.font.size = Pt(22)
        p_t.font.bold = True
        p_t.font.color.rgb = COLOR_DARK

        p_desc = tf.add_paragraph()
        p_desc.text = cat['desc']
        p_desc.font.name = FONT_BODY
        p_desc.font.size = Pt(16.5)
        p_desc.font.color.rgb = RGBColor(75, 85, 99)
        p_desc.space_before = Pt(3)

        p_ex = tf.add_paragraph()
        p_ex.text = "მაგალითები: " + cat['examples']
        p_ex.font.name = FONT_BODY
        p_ex.font.size = Pt(16)
        p_ex.font.bold = True
        p_ex.font.color.rgb = cat['badge_color']
        p_ex.space_before = Pt(3)

    # -------------------------------------------------------------------------
    # SLIDE 4: 4 წამყვანი ეკოსისტემა (ლოგოების ბარათები)
    # -------------------------------------------------------------------------
    s4 = prs.slides.add_slide(blank_layout)
    add_header(s4, "გლობალური AI ეკოსისტემა — 4 წამყვანი მოთამაშე", "ეკოსისტემის რუკა")

    platforms = [
        {
            "name": "ChatGPT",
            "company": "OpenAI",
            "logo": logo_chatgpt,
            "badge": "ყველაზე პოპულარული",
            "features": ["GPT-4o მულტიმოდალური მოდელი", "Reasoning o-სერია (o1, o3-mini)", "Canvas & მონაცემთა ანალიზი"],
            "logo_h": Inches(1.2)
        },
        {
            "name": "Claude",
            "company": "Anthropic",
            "logo": logo_claude,
            "badge": "სიღრმისეული ანალიტიკა & კოდი",
            "features": ["Claude 3.5 Sonnet & Haiku", "Artifacts ინტერაქტიული გარემო", "200,000 ტოკენიანი კონტექსტი"],
            "logo_h": Inches(0.7)
        },
        {
            "name": "Gemini",
            "company": "Google DeepMind",
            "logo": logo_gemini,
            "badge": "ექსტრემალური კონტექსტი",
            "features": ["1-2 მილიონი ტოკენის კონტექსტი", "Workspace & AI Studio ინტეგრაცია", "NotebookLM წყაროზე დაფუძნებული AI"],
            "logo_h": Inches(0.8)
        },
        {
            "name": "Microsoft Copilot",
            "company": "Microsoft",
            "logo": logo_copilot,
            "badge": "Enterprise ინტეგრაცია",
            "features": ["M365 ინტეგრაცია (Word, Excel, PPT)", "Commercial Data Protection", "GitHub Copilot დეველოპერებისთვის"],
            "logo_h": Inches(0.7)
        }
    ]

    col_w = Inches(4.15)
    gap = Inches(0.33)
    start_x = Inches(1.2)
    card_y = Inches(2.35)
    card_h = Inches(6.5)

    for i, p in enumerate(platforms):
        cx = start_x + i * (col_w + gap)

        card = s4.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, cx, card_y, col_w, card_h)
        card.fill.solid()
        card.fill.fore_color.rgb = COLOR_WHITE
        card.line.color.rgb = COLOR_BORDER
        card.line.width = Pt(1.5)

        if os.path.exists(p['logo']):
            logo_top = card_y + Inches(0.35)
            s4.shapes.add_picture(p['logo'], cx + Inches(0.45), logo_top, height=p['logo_h'])

        tb = s4.shapes.add_textbox(cx + Inches(0.4), card_y + Inches(1.75), col_w - Inches(0.8), card_h - Inches(1.9))
        tf = tb.text_frame
        tf.word_wrap = True

        p_name = tf.paragraphs[0]
        p_name.text = p['name']
        p_name.font.name = FONT_TITLE
        p_name.font.size = Pt(26)
        p_name.font.bold = True
        p_name.font.color.rgb = COLOR_DARK

        p_comp = tf.add_paragraph()
        p_comp.text = p['company']
        p_comp.font.name = FONT_BODY
        p_comp.font.size = Pt(15)
        p_comp.font.color.rgb = COLOR_MUTED
        p_comp.space_before = Pt(2)

        p_b = tf.add_paragraph()
        p_b.text = p['badge'].upper()
        p_b.font.name = FONT_TITLE
        p_b.font.size = Pt(13)
        p_b.font.bold = True
        p_b.font.color.rgb = COLOR_RED
        p_b.space_before = Pt(8)

        for feat in p['features']:
            pf = tf.add_paragraph()
            pf.text = "✓ " + feat
            pf.font.name = FONT_BODY
            pf.font.size = Pt(16.5)
            pf.font.color.rgb = COLOR_DARK
            pf.space_before = Pt(10)

    # Bottom strip: Alternative open models
    alt_banner = s4.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(1.2), Inches(9.0), Inches(17.6), Inches(1.05))
    alt_banner.fill.solid()
    alt_banner.fill.fore_color.rgb = RGBColor(241, 245, 249)
    alt_banner.line.color.rgb = COLOR_BORDER
    alt_banner.line.width = Pt(1)

    alt_tb = s4.shapes.add_textbox(Inches(1.5), Inches(9.12), Inches(17.0), Inches(0.8))
    alt_tf = alt_tb.text_frame
    alt_tf.word_wrap = True
    p_alt = alt_tf.paragraphs[0]
    p_alt.text = "⚡ ღია & ალტერნატიული ეკოსისტემები: DeepSeek (R1 / V3 Reasoning), Meta LLaMA 3 (Open-Weights), Mistral AI, Qwen."
    p_alt.font.name = FONT_BODY
    p_alt.font.size = Pt(16)
    p_alt.font.bold = True
    p_alt.font.color.rgb = COLOR_DARK

    # -------------------------------------------------------------------------
    # SLIDE 5: მონაცემთა უსაფრთხოების „არასდროს“ წესები
    # -------------------------------------------------------------------------
    s5 = prs.slides.add_slide(blank_layout)
    add_header(s5, "მონაცემთა დაცვა — 3 „არასდროს“ წესი ორგანიზაციაში", "უსაფრთხოება & რისკების კონტროლი")

    rules = [
        {
            "num": "🛑 წესი 1",
            "title": "არასდროს გააგზავნო PII და საიდუმლოებები საჯარო ჩატში",
            "points": [
                "პერსონალური მონაცემები: პირადი ნომრები, პასპორტის სკანები, სამედიცინო ჩანაწერები.",
                "ფინანსური ინფორმაცია: საბანკო რეკვიზიტები, კომპანიის შიდა ბიუჯეტი და ინვოისები.",
                "ტექნიკური საიდუმლოება: კომპანიის დახურული საწყისი კოდი (Source Code) და API გასაღებები."
            ]
        },
        {
            "num": "🛑 წესი 2",
            "title": "არასდროს გამოიყენო AI კორპორატიულად Opt-Out-ის გარეშე",
            "points": [
                "უფასო/Consumer ჩატბოტები ნაგულისხმევად იყენებენ საუბრებს მოდელის შემდეგი წვრთნისთვის.",
                "სავალდებულოა პარამეტრებში Model Training-ის გათიშვა (Data Opt-Out) პირველივე დღეს.",
                "სამუშაო მონაცემებისთვის გამოიყენება მხოლოდ Enterprise / Zero-Data Retention პაკეტები."
            ]
        },
        {
            "num": "🛑 წესი 3",
            "title": "არასდროს ენდო ფაქტებს ბრმად (სავალდებულო გადამოწმება)",
            "points": [
                "გენერაციული AI აგენერირებს ალბათურ ტექსტს და ხშირად დამაჯერებლად ჰალუცინირებს.",
                "თარიღები, ციტატები, კანონის მუხლები და ციფრები ყოველთვის მოითხოვს პირველწყაროსთან შემოწმებას.",
                "ადამიანი რჩება 100%-ით პასუხისმგებელი საბოლოო შედეგზე (Human-in-the-Loop პრინციპი)."
            ]
        }
    ]

    card_w = Inches(5.6)
    gap = Inches(0.4)
    start_x = Inches(1.2)
    card_y = Inches(2.35)
    card_h = Inches(7.5)

    for i, r in enumerate(rules):
        cx = start_x + i * (card_w + gap)

        card = s5.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, cx, card_y, card_w, card_h)
        card.fill.solid()
        card.fill.fore_color.rgb = COLOR_CARD_RED
        card.line.color.rgb = COLOR_BORDER_RED
        card.line.width = Pt(1.5)

        tb = s5.shapes.add_textbox(cx + Inches(0.4), card_y + Inches(0.4), card_w - Inches(0.8), card_h - Inches(0.8))
        tf = tb.text_frame
        tf.word_wrap = True

        p_n = tf.paragraphs[0]
        p_n.text = r['num']
        p_n.font.name = FONT_TITLE
        p_n.font.size = Pt(20)
        p_n.font.bold = True
        p_n.font.color.rgb = COLOR_RED

        p_t = tf.add_paragraph()
        p_t.text = r['title']
        p_t.font.name = FONT_TITLE
        p_t.font.size = Pt(22)
        p_t.font.bold = True
        p_t.font.color.rgb = COLOR_DARK
        p_t.space_before = Pt(8)

        line = s5.shapes.add_shape(MSO_SHAPE.RECTANGLE, cx + Inches(0.4), card_y + Inches(2.3), card_w - Inches(0.8), Inches(0.01))
        line.fill.solid()
        line.fill.fore_color.rgb = COLOR_BORDER_RED
        line.line.color.rgb = COLOR_BORDER_RED

        tb_body = s5.shapes.add_textbox(cx + Inches(0.4), card_y + Inches(2.45), card_w - Inches(0.8), card_h - Inches(2.6))
        tf_body = tb_body.text_frame
        tf_body.word_wrap = True
        first = True
        for pt in r['points']:
            p_pt = tf_body.paragraphs[0] if first else tf_body.add_paragraph()
            first = False
            p_pt.text = "⚠️ " + pt
            p_pt.font.name = FONT_BODY
            p_pt.font.size = Pt(17.5)
            p_pt.font.color.rgb = COLOR_DARK
            p_pt.space_before = Pt(16)

    # -------------------------------------------------------------------------
    # SLIDE 6: Closing Slide
    # -------------------------------------------------------------------------
    s6 = prs.slides.add_slide(blank_layout)
    if os.path.exists(img_corner_bl):
        s6.shapes.add_picture(img_corner_bl, Inches(0), Inches(8.75), width=Inches(2.52))
    if os.path.exists(img_corner_tr):
        s6.shapes.add_picture(img_corner_tr, Inches(17.53), Inches(0), width=Inches(2.47))

    close_box = s6.shapes.add_textbox(Inches(4.0), Inches(4.5), Inches(12.0), Inches(2.8))
    ctf = close_box.text_frame
    ctf.word_wrap = True

    p_close = ctf.paragraphs[0]
    p_close.text = "მადლობა ყურადღებისთვის!"
    p_close.alignment = PP_ALIGN.CENTER
    p_close.font.name = FONT_TITLE
    p_close.font.size = Pt(48)
    p_close.font.bold = True
    p_close.font.color.rgb = COLOR_DARK

    p_subclose = ctf.add_paragraph()
    p_subclose.text = "კითხვა-პასუხი & დისკუსია\nSkillwill • 2026"
    p_subclose.alignment = PP_ALIGN.CENTER
    p_subclose.font.name = FONT_BODY
    p_subclose.font.size = Pt(24)
    p_subclose.font.color.rgb = COLOR_RED
    p_subclose.space_before = Pt(18)

    # Save
    out_file = 'd:/ხელოვნური ინტელექტის ინსტრუმენტები/Skillwill_მოდული5_ლექცია1_პრეზენტაცია_ნიმუში.pptx'
    prs.save(out_file)
    print("SUCCESS: Regenerated presentation with larger, easily readable typography:", out_file)

if __name__ == '__main__':
    build_visual_presentation()
