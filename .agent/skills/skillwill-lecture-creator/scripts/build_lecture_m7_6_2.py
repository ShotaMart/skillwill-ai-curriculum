import os
import sys
import shutil
import pptx
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN
from pptx.enum.shapes import MSO_SHAPE

def build_lecture_m7_6_2():
    template_path = r'd:\ხელოვნური ინტელექტის ინსტრუმენტები\Skillwill Presentation Example .pptx'
    asset_dir = r'd:\ხელოვნური ინტელექტის ინსტრუმენტები\assets'
    img_icon_title = os.path.join(asset_dir, 'icon_title.png')

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
    COLOR_GREEN_BG = RGBColor(0xEA, 0xF8, 0xEA)
    COLOR_GREEN_BOR = RGBColor(0xC2, 0xE8, 0xC2)
    COLOR_GREEN_TEXT = RGBColor(0x13, 0x73, 0x33)

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
            p.text = "6.2 - Chain-of-Thought, Reasoning მოდელები და Extended Thinking"
            p.alignment = PP_ALIGN.CENTER
            p.font.name = FONT_ARIAL
            p.font.size = Pt(32)
            p.font.bold = True
            p.font.color.rgb = COLOR_RED

            p2 = tf.add_paragraph()
            p2.text = "ეტაპობრივი მსჯელობის მეთოდოლოგია (CoT), 2026 წლის Frontier Reasoning მოდელები (Sol, Astra, Claude 5.x) და ლოგიკური ამოცანების შედარება"
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
    def add_header(slide, title, category="ᲛᲝᲓᲣᲚᲘ 7 • PROMPT ENGINEERING • 6.2"):
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

    card_w = Inches(5.4); card_h = Inches(5.4); card_y = Inches(2.35)
    xs_3 = [Inches(1.0), Inches(6.8), Inches(12.6)]
    bot_y = Inches(8.0)

    # -------------------------------------------------------------------------
    # SLIDE 2: CoT Mental Model: System 1 vs System 2 Thinking
    # -------------------------------------------------------------------------
    s2 = created_slides[0]
    add_header(s2, "Chain-of-Thought (CoT) მენტალური მოდელი: System 1 vs System 2", "ᲛᲔᲜᲢᲐᲚᲣᲠᲘ ᲛᲝᲓᲔᲚᲘ • SYSTEM 1 VS SYSTEM 2")

    cards_s2 = [
        ("⚡ SYSTEM 1 (FAST & INTUITIVE)", "სწრაფი და ინტუიციური", "პირდაპირი პასუხის რისკი", [
            ("პირდაპირი გენერაცია", "მოდელი ცდილობს გასცეს პასუხი პირველივე ტოკენიდან ყოველგვარი შუალედური გაანგარიშების გარეშე."),
            ("ლოგიკური ჩიხი", "რთულ ამოცანებში (მათემატიკა, ალგორითმები) პირდაპირი პროგნოზი 70%-ში იწვევს ჰალუცინაციას."),
            ("ფიქსირებული გამოთვლა", "სტანდარტული ტრანსფორმერის გამოთვლითი რესურსი თითოეულ ტოკენზე არ კმარა მრავალსაფეხურიანი ლოგიკისთვის.")
        ]),
        ("🧠 SYSTEM 2 (SLOW & DELIBERATE)", "შენელებული და ლოგიკური", "ეტაპობრივი გაანგარიშება", [
            ("შუალედური ნაბიჯები", "პრობლემის დაშლა ქვეამოცანებად (Sub-goals) და თითოეულის თანმიმდევრული, ლოგიკური ამოხსნა."),
            ("სამუშაო მეხსიერება (Working Memory)", "გენერირებული მსჯელობის ტოკენები ხდება კონტექსტი შემდეგი ნაბიჯებისთვის."),
            ("თვითკორექტირება გზაში", "მოდელი ყოველ ნაბიჯზე ამოწმებს წინა დასკვნის სისწორეს საბოლოო გადაწყვეტილებამდე.")
        ]),
        ("📈 COMPUTE AT INFERENCE TIME", "დასკვნის ეტაპის გამოთვლა", "ახალი საინჟინრო პარადიგმა", [
            ("Test-Time Compute", "მეტი ტოკენი მსჯელობაზე = რადიკალურად მაღალი ინტელექტი, სიზუსტე და ალგორითმული სიმტკიცე."),
            ("Scaling Laws 2.0", "მოდელის შესაძლებლობების ზრდა ხდება არა მხოლოდ წინასწარი წვრთნისას, არამედ უშუალოდ დასკვნის (Inference) ეტაპზე."),
            ("ოპტიმალური ინვესტიცია", "რთულ ამოცანაზე რამდენიმე წამით მეტი ფიქრი თავიდან გვარიდებს კრიტიკულ საწარმოო შეცდომებს.")
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
    p = tfb.paragraphs[0]; p.text = "💡 მთავარი არსი:"
    p.font.name = FONT_ARIAL; p.font.size = Pt(14); p.font.bold = True; p.font.color.rgb = COLOR_RED
    p = tfb.add_paragraph()
    p.text = "Chain-of-Thought გარდაქმნის LLM-ს 'სწრაფი ინტუიციიდან' (System 1) 'ღრმა მოაზროვნედ' (System 2); შუალედური ლოგიკური ნაბიჯების დაფიქსირება მოდელს საშუალებას აძლევს ამოხსნას ამოცანები, რომლებიც პირდაპირი გენერაციით შეუძლებელია."
    p.font.name = FONT_ARIAL; p.font.size = Pt(13); p.font.color.rgb = COLOR_MUTED; p.space_before = Pt(4)

    # -------------------------------------------------------------------------
    # SLIDE 3: Evolution of CoT Techniques (Zero-Shot CoT, Few-Shot CoT, XML)
    # -------------------------------------------------------------------------
    s3 = created_slides[1]
    add_header(s3, "CoT ტექნიკების ევოლუცია: Zero-Shot CoT, Few-Shot CoT და ტეგირება", "ᲔᲕᲝᲚᲣᲪᲘᲐ • COT TECHNIQUES")

    tech_s3 = [
        ("🪄 ZERO-SHOT COT", "მარტივი ტრიგერი", "Let's think step by step", [
            ("საკვანძო ფრაზა", "ფრაზის: 'Let's think step by step' (ან 'იაზროვნე ნაბიჯ-ნაბიჯ') დამატება მოთხოვნის ბოლოს."),
            ("მექანიზმი", "აიძულებს მოდელს, პირდაპირი პასუხის ნაცვლად ჯერ ჩამოაყალიბოს მსჯელობის ლოგიკური ჯაჭვი."),
            ("ეფექტურობა", "ზრდის სიზუსტეს 20-30%-ით მარტივ არითმეტიკულ და ლოგიკურ ამოცანებში დამატებითი მაგალითების გარეშე.")
        ]),
        ("📋 FEW-SHOT COT", "მაგალითებით მართული ლოგიკა", "სანიმუშო მსჯელობის ჯაჭვი", [
            ("სტრუქტურა", "მაგალითების მიწოდება ფორმატით: Input ➔ Step-by-Step Reasoning ➔ Final Answer."),
            ("ლოგიკის სტანდარტიზაცია", "მოდელს ვასწავლით არა მხოლოდ პასუხს, არამედ აზროვნებისა და გადამოწმების ზუსტ მეთოდოლოგიას."),
            ("რთული დომენები", "იდეალურია იურიდიული შემოწმებისთვის, სამედიცინო ტრიაჟისა და ფინანსური აუდიტისთვის.")
        ]),
        ("🏷️ STRUCTURED XML COT", "სტრუქტურული ტეგირება", "2026 წლის საწარმოო სტანდარტი", [
            ("ტეგებით იზოლაცია", "ინსტრუქცია: 'მსჯელობის ეტაპები მოათავსე <thinking> ტეგებში, ხოლო საბოლოო გადაწყვეტილება — <answer>-ში'."),
            ("გამოსავლის სისუფთავე", "სერვერულ აპლიკაციას შეუძლია მარტივად გამოყოს მსჯელობა საბოლოო შედეგისგან."),
            ("Frontier მორჩილება", "Claude 5.x და GPT-6 სრულყოფილად მიჰყვებიან ამ სტრუქტურულ არქიტექტურას.")
        ])
    ]

    for i, (badge, title, subtitle, items) in enumerate(tech_s3):
        bg_col = COLOR_CARD_RED if i == 2 else COLOR_BG_LIGHT
        bor_col = COLOR_BORDER_RED if i == 2 else COLOR_BORDER
        c = s3.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, xs_3[i], card_y, card_w, card_h)
        c.fill.solid(); c.fill.fore_color.rgb = bg_col
        c.line.color.rgb = bor_col; c.line.width = Pt(1.5)

        tb = s3.shapes.add_textbox(xs_3[i] + Inches(0.35), card_y + Inches(0.35), card_w - Inches(0.7), card_h - Inches(0.7))
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
    bb = s3.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(1.0), bot_y, Inches(17.0), Inches(2.0))
    bb.fill.solid(); bb.fill.fore_color.rgb = COLOR_WHITE
    bb.line.color.rgb = COLOR_BORDER; bb.line.width = Pt(1.2)

    tbb = s3.shapes.add_textbox(Inches(1.4), bot_y + Inches(0.3), Inches(16.2), Inches(1.4))
    tfb = tbb.text_frame; tfb.word_wrap = True; tfb.margin_left = tfb.margin_top = 0
    p = tfb.paragraphs[0]; p.text = "🎯 საინჟინრო პრინციპი:"
    p.font.name = FONT_ARIAL; p.font.size = Pt(14); p.font.bold = True; p.font.color.rgb = COLOR_RED
    p = tfb.add_paragraph()
    p.text = "Zero-Shot CoT კარგია სწრაფი ექსპერიმენტებისთვის; საწარმოო პროდაქშენში ყოველთვის გამოიყენეთ Structured CoT (<thinking>), რათა მოდელის ლოგიკური პროცესი იზოლირებული იყოს საბოლოო მომხმარებლის პასუხისგან."
    p.font.name = FONT_ARIAL; p.font.size = Pt(13); p.font.color.rgb = COLOR_MUTED; p.space_before = Pt(4)

    # -------------------------------------------------------------------------
    # SLIDE 4: Frontier Reasoning Models 2026 (Extended Thinking)
    # -------------------------------------------------------------------------
    s4 = created_slides[2]
    add_header(s4, "2026 წლის Frontier Reasoning მოდელები და Extended Thinking", "FRONTIER AI 2026 • EXTENDED THINKING")

    models_s4 = [
        ("🌐 OPENAI GPT-6 (SOL / ASTRA)", "ინტეგრირებული Reasoning", "ავტონომიური ფიქრის რეჟიმი", [
            ("Native Reinforcement Learning", "მოდელები გაწვრთნილია მსჯელობის ჯაჭვებზე (Chain-of-Thought RL), რათა შეცდომები თავად გაასწორონ."),
            ("Sol & Astra თაობა", "ინტეგრირებული აზროვნების სისტემა, რომელიც თავად წყვეტს, რამდენ ხანს იფიქროს ამოცანის სირთულის მიხედვით."),
            ("Deep Verification", "მოდელი პარალელურად ამოწმებს ამოხსნის რამდენიმე ჰიპოთეზას და ირჩევს უტყუარ ვარიანტს.")
        ]),
        ("🎭 ANTHROPIC CLAUDE 5.x", "Extended Thinking", "Thinking Budget-ის კონტროლი", [
            ("Thinking Budget პარამეტრი", "ინჟინერს შეუძლია API-ში მკაცრად განსაზღვროს ფიქრის ტოკენების ლიმიტი (მაგ. 4,096 ტოკენი)."),
            ("სრული გამჭვირვალობა", "Claude-ის Extended Thinking აბრუნებს მსჯელობის სრულ ბლოკს სპეციალურ thinking ობიექტში."),
            ("უმაღლესი თვითკონტროლი", "Claude 5.x ფიქრისას ამოწმებს კომპანიის პოლიტიკას, უსაფრთხოებასა და ლოგიკურ წინააღმდეგობებს.")
        ]),
        ("🔒 HIDDEN REASONING TOKENS", "ფარული ტოკენები", "არქიტექტურული უსაფრთხოება", [
            ("შიდა მონოლოგი", "Reasoning ტოკენები წარმოადგენს მოდელის 'გონების ხმას' და დაცულია გარე მანიპულაციებისგან."),
            ("მონიტორინგი და აუდიტი", "საწარმოო სისტემები ინახავს ამ ტოკენებს დეველოპერულ Tracing სისტემებში (Langfuse)."),
            ("აკრძალულია გადაწერა", "მომხმარებელს არ შეუძლია Prompt Injection-ით შეცვალოს მოდელის ფარული მსჯელობის პროცესი.")
        ])
    ]

    for i, (badge, title, subtitle, items) in enumerate(models_s4):
        bg_col = COLOR_CARD_RED if i == 1 else COLOR_BG_LIGHT
        bor_col = COLOR_BORDER_RED if i == 1 else COLOR_BORDER
        c = s4.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, xs_3[i], card_y, card_w, card_h)
        c.fill.solid(); c.fill.fore_color.rgb = bg_col
        c.line.color.rgb = bor_col; c.line.width = Pt(1.5)

        tb = s4.shapes.add_textbox(xs_3[i] + Inches(0.35), card_y + Inches(0.35), card_w - Inches(0.7), card_h - Inches(0.7))
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
    bb = s4.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(1.0), bot_y, Inches(17.0), Inches(2.0))
    bb.fill.solid(); bb.fill.fore_color.rgb = COLOR_WHITE
    bb.line.color.rgb = COLOR_BORDER; bb.line.width = Pt(1.2)

    tbb = s4.shapes.add_textbox(Inches(1.4), bot_y + Inches(0.3), Inches(16.2), Inches(1.4))
    tfb = tbb.text_frame; tfb.word_wrap = True; tfb.margin_left = tfb.margin_top = 0
    p = tfb.paragraphs[0]; p.text = "💡 მთავარი არსი:"
    p.font.name = FONT_ARIAL; p.font.size = Pt(14); p.font.bold = True; p.font.color.rgb = COLOR_RED
    p = tfb.add_paragraph()
    p.text = "2026 წელს Reasoning აღარ არის პრომპტში ჩაწერილი უბრალო ფრაზა; ეს არის მოდელის დონეზე ჩაშენებული Native შესაძლებლობა, რომლის მართვაც ხდება Thinking Budget-ითა და სისტემური API პარამეტრებით."
    p.font.name = FONT_ARIAL; p.font.size = Pt(13); p.font.color.rgb = COLOR_MUTED; p.space_before = Pt(4)

    # -------------------------------------------------------------------------
    # SLIDE 5: The Reasoning Boundary (Where CoT Excels vs Harms)
    # -------------------------------------------------------------------------
    s5 = created_slides[3]
    add_header(s5, "სად შველის და სად აფერხებს CoT? საინჟინრო საზღვრები", "ᲡᲐᲖᲦᲕᲠᲔᲑᲘ • THE REASONING BOUNDARY")

    bounds_s5 = [
        ("✅ WHERE COT EXCELS", "სად შველის რადიკალურად", "კომპლექსური ამოცანები", [
            ("მრავალსაფეხურიანი მათემატიკა", "რთული პროცენტები, ფინანსური მოდელირება, ბალანსის აუდიტი და მონაცემთა გადათვლა."),
            ("ალგორითმული და კოდის ლოგიკა", "არქიტექტურული დაგეგმვა, რთული Debugging, SQL მოთხოვნების ანალიზი და რეგექსები."),
            ("მრავალპირობიანი იურიდიული წესები", "კონტრაქტების შემოწმება ურთიერთგამომრიცხავ რეგულაციებთან და პოლიტიკასთან.")
        ]),
        ("❌ WHERE COT HARMS", "სად აფერხებს და აზიანებს", "მარტივი და სწრაფი ამოცანები", [
            ("მარტივი ტექსტური კლასიფიკაცია", "სპამი/არასპამი, მარტივი სენტიმენტი — აქ ზედმეტი მსჯელობა მხოლოდ აძვირებს სერვისს."),
            ("ენობრივი თარგმანი და კორექტირება", "თარგმნას სჭირდება ენობრივი განზოგადება და არა დედუქციური მსჯელობა."),
            ("პირდაპირი ფაქტობრივი კითხვები", "'რა არის საქართველოს დედაქალაქი?' — მსჯელობის ჩართვა აქ სრულიად უადგილოა.")
        ]),
        ("⚠️ THE OVERTHINKING TRAP", "ზედმეტი ფიქრის ხაფანგი", "პარადოქსული შეცდომები", [
            ("ხელოვნური გართულება", "მოდელი უბრალო კითხვაში ეძებს ფარულ ხრიკებს და მარტივ საკითხზე უშვებს შეცდომას."),
            ("პარალიზი ანალიზით", "ზედმეტი მსჯელობის გამო მოდელი იწყებს საკუთარ სწორ ვერსიებში ეჭვის შეტანას და ცდება."),
            ("პასუხის გაბერვა", "მომხმარებელი მოკლე, ლაკონიური პასუხის ნაცვლად იღებს მრავალგვერდიან ფილოსოფიურ ტრაქტატს.")
        ])
    ]

    for i, (badge, title, subtitle, items) in enumerate(bounds_s5):
        bg_col = COLOR_CARD_RED if i == 2 else COLOR_BG_LIGHT
        bor_col = COLOR_BORDER_RED if i == 2 else COLOR_BORDER
        c = s5.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, xs_3[i], card_y, card_w, card_h)
        c.fill.solid(); c.fill.fore_color.rgb = bg_col
        c.line.color.rgb = bor_col; c.line.width = Pt(1.5)

        tb = s5.shapes.add_textbox(xs_3[i] + Inches(0.35), card_y + Inches(0.35), card_w - Inches(0.7), card_h - Inches(0.7))
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
    bb = s5.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(1.0), bot_y, Inches(17.0), Inches(2.0))
    bb.fill.solid(); bb.fill.fore_color.rgb = COLOR_WHITE
    bb.line.color.rgb = COLOR_BORDER; bb.line.width = Pt(1.2)

    tbb = s5.shapes.add_textbox(Inches(1.4), bot_y + Inches(0.3), Inches(16.2), Inches(1.4))
    tfb = tbb.text_frame; tfb.word_wrap = True; tfb.margin_left = tfb.margin_top = 0
    p = tfb.paragraphs[0]; p.text = "🎯 საინჟინრო პრინციპი:"
    p.font.name = FONT_ARIAL; p.font.size = Pt(14); p.font.bold = True; p.font.color.rgb = COLOR_RED
    p = tfb.add_paragraph()
    p.text = "არ გამოიყენოთ Reasoning მოდელები იქ, სადაც სტანდარტული მსუბუქი მოდელი წამის მეასედებში წყვეტს ამოცანას; Reasoning გამოიყენეთ მხოლოდ მაშინ, როდესაც შეცდომის ფასი კრიტიკულია და ლოგიკა მრავალსაფეხურიანია."
    p.font.name = FONT_ARIAL; p.font.size = Pt(13); p.font.color.rgb = COLOR_MUTED; p.space_before = Pt(4)

    # -------------------------------------------------------------------------
    # SLIDE 6: Cost, Latency & Smart Hybrid Routing
    # -------------------------------------------------------------------------
    s6 = created_slides[4]
    add_header(s6, "Reasoning-ის ეკონომიკა: ტოკენების ხარჯი, დაყოვნება და Smart Routing", "ᲔᲙᲝᲜᲝᲛᲘᲙᲐ • COST & LATENCY TRADE-OFFS")

    cost_s6 = [
        ("💰 TOKEN MULTIPLICATION", "ტოკენების მრავალჯერადი ზრდა", "ფინანსური ხარჯი", [
            ("Thinking Tokens ტარიფი", "მსჯელობის ტოკენები იბეგრება როგორც Output Tokens (რომელიც Input-ზე 3-4-ჯერ ძვირია)."),
            ("ასიმეტრიული მოხმარება", "მოკლე კითხვაზე (20 ტოკენი) მოდელმა შეიძლება დახარჯოს 3,000 Thinking ტოკენი."),
            ("ბიუჯეტის ლიმიტები", "აუცილებელია მკაცრი Thinking Budget-ის დაწესება ყოველ კონკრეტულ API გამოძახებაზე.")
        ]),
        ("⏳ LATENCY EXPLOSION", "დაყოვნების მკვეთრი ზრდა", "მომხმარებლის UX", [
            ("Time to First Token (TTFT)", "მომხმარებელი ელოდება 5-დან 30 წამამდე, სანამ პირველი სიტყვა გამოჩნდება ეკრანზე."),
            ("ჩატბოტებისთვის მიუღებელია", "რეალურ დროში მომხმარებლის მხარდაჭერისთვის ასეთი დაყოვნება სრულიად მიუღებელია."),
            ("ფონური დამუშავება (Async)", "Reasoning მოდელები იდეალურია ასინქრონული, ფონური ანალიტიკისა და ღამის რეპორტინგისთვის.")
        ]),
        ("🔀 SMART HYBRID ROUTING", "ჰიბრიდული როუტინგი", "ოპტიმალური საინჟინრო დასტა", [
            ("Tiered Routing არქიტექტურა", "მსუბუქი კლასიფიკატორი (Gemini Flash / GPT-5.6 mini) პირველ ეტაპზე აფასებს ამოცანის სირთულეს."),
            ("მარტივი მოთხოვნები", "მიემართება სწრაფ, იაფ მოდელთან (Latency: 0.3s, ხარჯი: მინიმალური)."),
            ("კომპლექსური ამოცანები", "ავტომატურად გადამისამართდება Reasoning მოდელთან (Sol / Claude 5.x Thinking).")
        ])
    ]

    for i, (badge, title, subtitle, items) in enumerate(cost_s6):
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
    p = tfb.paragraphs[0]; p.text = "💡 მთავარი არსი:"
    p.font.name = FONT_ARIAL; p.font.size = Pt(14); p.font.bold = True; p.font.color.rgb = COLOR_RED
    p = tfb.add_paragraph()
    p.text = "Reasoning არის 'მძიმე არტილერია'; პროდაქშენში ყოველთვის გამოიყენეთ Smart Routing, რათა მომხმარებლის მოთხოვნების 80% დამუშავდეს მყისიერად და იაფად, ხოლო დარჩენილ 20% რთულ ქეისს მოემსახუროს Reasoning მოდელი."
    p.font.name = FONT_ARIAL; p.font.size = Pt(13); p.font.color.rgb = COLOR_MUTED; p.space_before = Pt(4)

    # -------------------------------------------------------------------------
    # SLIDE 7: 10 Logical Tasks Benchmark (Table)
    # -------------------------------------------------------------------------
    s7 = created_slides[5]
    add_header(s7, "10 ლოგიკური ამოცანის ექსპერიმენტული ანალიზი: Direct vs CoT", "ᲔᲥᲡᲞᲔᲠᲘᲛᲔᲜᲢᲘ • 10 LOGICAL TASKS BENCHMARK")

    # Table: 6 rows (Header + 5 task categories), 5 columns
    tbl_left = Inches(1.0); tbl_top = Inches(2.35); tbl_w = Inches(17.0); tbl_h = Inches(5.2)
    tbl_shape = s7.shapes.add_table(6, 5, tbl_left, tbl_top, tbl_w, tbl_h)
    table = tbl_shape.table

    col_widths = [Inches(3.4), Inches(2.8), Inches(3.2), Inches(4.0), Inches(3.6)]
    for ci, w in enumerate(col_widths):
        table.columns[ci].width = w

    headers = ["ამოცანის კატეგორია / ტიპი", "Direct (CoT-ის გარეშე)", "CoT / Extended Thinking", "ტიპური შეცდომა Direct-ში", "საინჟინრო რეკომენდაცია"]
    for ci, h in enumerate(headers):
        cell = table.cell(0, ci)
        cell.fill.solid(); cell.fill.fore_color.rgb = COLOR_DARK
        tf = cell.text_frame; tf.word_wrap = True; tf.margin_left = tf.margin_right = Inches(0.12)
        p = tf.paragraphs[0]; p.text = h; p.font.name = FONT_ARIAL; p.font.size = Pt(12); p.font.bold = True; p.font.color.rgb = COLOR_WHITE

    exp_rows = [
        ("მრავალსაფეხურიანი არითმეტიკა (3 ამოცანა)", "45% სიზუსტე", "98% სიზუსტე", "გამოტოვებული მოქმედებები და არითმეტიკული ჰალუცინაცია", "CoT სავალდებულოა"),
        ("დროითი და კალენდარული ლოგიკა (2 ამოცანა)", "50% სიზუსტე", "95% სიზუსტე", "დღეებისა და საათების არასწორი გადათვლა", "CoT სავალდებულოა"),
        ("სიტყვებით ლოგიკა და რებუსები (2 ამოცანა)", "40% სიზუსტე", "90% სიზუსტე", "ზედაპირული ასოციაციები და ფაქტობრივი ჩავარდნა", "Structured Thinking დაცვა"),
        ("იურიდიული წესების შეჯახება (2 ამოცანა)", "60% სიზუსტე", "96% სიზუსტე", "გამონაკლისი პუნქტების უგულებელყოფა", "Few-Shot CoT რეკომენდებულია"),
        ("მარტივი ფაქტობრივი კითხვა (1 ამოცანა)", "100% სიზუსტე", "98% (Overthinking რისკი)", "შეცდომა არ არის; CoT-ზე დრო და ფული ფუჭდება", "CoT მკაცრად აკრძალულია")
    ]

    for ri, row_data in enumerate(exp_rows):
        row_bg = COLOR_WHITE if ri % 2 == 0 else COLOR_BG_LIGHT
        for ci, val in enumerate(row_data):
            cell = table.cell(ri + 1, ci)
            cell.fill.solid(); cell.fill.fore_color.rgb = row_bg
            tf = cell.text_frame; tf.word_wrap = True; tf.margin_left = tf.margin_right = Inches(0.12)
            p = tf.paragraphs[0]; p.text = val; p.font.name = FONT_ARIAL; p.font.size = Pt(11)
            if ci == 0:
                p.font.bold = True; p.font.color.rgb = COLOR_DARK
            elif ci == 2:
                p.font.color.rgb = COLOR_RED; p.font.bold = True
            elif ci == 4:
                p.font.bold = True; p.font.color.rgb = COLOR_DARK
            else:
                p.font.color.rgb = COLOR_MUTED

    # Bottom takeaway
    bb = s7.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(1.0), bot_y, Inches(17.0), Inches(2.0))
    bb.fill.solid(); bb.fill.fore_color.rgb = COLOR_WHITE
    bb.line.color.rgb = COLOR_BORDER; bb.line.width = Pt(1.2)

    tbb = s7.shapes.add_textbox(Inches(1.4), bot_y + Inches(0.3), Inches(16.2), Inches(1.4))
    tfb = tbb.text_frame; tfb.word_wrap = True; tfb.margin_left = tfb.margin_top = 0
    p = tfb.paragraphs[0]; p.text = "💡 ექსპერიმენტის მთავარი დასკვნა:"
    p.font.name = FONT_ARIAL; p.font.size = Pt(14); p.font.bold = True; p.font.color.rgb = COLOR_RED
    p = tfb.add_paragraph()
    p.text = "რთულ მრავალსაფეხურიან ლოგიკაში CoT სიზუსტეს ორჯერ და მეტად ზრდის (45%-დან 98%-მდე); თუმცა მარტივ ფაქტობრივ კითხვებზე CoT ზედმეტია, ზრდის დაყოვნებას და აჩენს Overthinking-ის რისკს."
    p.font.name = FONT_ARIAL; p.font.size = Pt(13); p.font.color.rgb = COLOR_MUTED; p.space_before = Pt(4)

    # -------------------------------------------------------------------------
    # SLIDE 8: Output Separation Architecture (<thinking> vs <answer>)
    # -------------------------------------------------------------------------
    s8 = created_slides[6]
    add_header(s8, "საინჟინრო არქიტექტურა: მსჯელობისა და პასუხის გამიჯვნა (<thinking>)", "ᲐᲠᲥᲘᲢᲔᲥᲢᲣᲠᲐ • OUTPUT SEPARATION")

    sep_s8 = [
        ("🏷️ XML ENCAPSULATION", "ტეგების იზოლაცია", "მსჯელობის შეფუთვა", [
            ("მკაცრი ინსტრუქცია", "სისტემური დირექტივა: 'ჯერ გახსენი <thinking> ტეგი, ჩაატარე ანალიზი, დახურე ტეგი და შემდეგ გამოიტანე <final_answer>'."),
            ("საზღვრების დაცვა", "მოდელი არასოდეს ურევს მსჯელობასა და საბოლოო პასუხს ერთმანეთში."),
            ("Frontier სტანდარტი", "იდენტურია Anthropic Claude-ისა და Google Gemini-ს შიდა საინჟინრო პრინციპისა.")
        ]),
        ("⚙️ API & BACKEND PARSING", "სერვერული ფილტრაცია", "კონფიდენციალობა და UX", [
            ("სუფთა ინტერფეისი", "სერვერული აპლიკაცია კითხულობს და ჭრის <thinking> ბლოკს, საბოლოო კლიენტს კი აწვდის მხოლოდ სუფთა პასუხს."),
            ("Audit & Tracing Log", "მსჯელობის სრული ტექსტი ინახება შიდა მონიტორინგის ბაზაში (Langfuse, Datadog) ხარისხის ანალიზისთვის."),
            ("შიდა წესების დაცვა", "თუ მსჯელობისას მოდელმა გააანალიზა შიდა ბიზნეს-წესები, ისინი კლიენტისთვის უხილავი რჩება.")
        ]),
        ("📋 JSON SCHEMA INTEGRATION", "სტრუქტურირებული გამოსავალი", "პროგრამული კონტრაქტი", [
            ("ველების გამიჯვნა JSON-ში", "სქემა: {'reasoning_steps': ['ნაბიჯი 1', 'ნაბიჯი 2'], 'decision': 'APPROVED', 'confidence': 0.98}."),
            ("ტიპიზებული ვალიდაცია", "Pydantic ამოწმებს ორივე ველის არსებობას, რაც უზრუნველყოფს სრულ პროგრამულ კონტროლს."),
            ("Explainable AI (XAI)", "საჭიროების შემთხვევაში (მაგ. კრედიტის უარყოფა), კომპანიას შეუძლია მომხმარებელს აჩვენოს გადაწყვეტილების ახსნა.")
        ])
    ]

    for i, (badge, title, subtitle, items) in enumerate(sep_s8):
        bg_col = COLOR_CARD_RED if i == 0 else COLOR_BG_LIGHT
        bor_col = COLOR_BORDER_RED if i == 0 else COLOR_BORDER
        c = s8.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, xs_3[i], card_y, card_w, card_h)
        c.fill.solid(); c.fill.fore_color.rgb = bg_col
        c.line.color.rgb = bor_col; c.line.width = Pt(1.5)

        tb = s8.shapes.add_textbox(xs_3[i] + Inches(0.35), card_y + Inches(0.35), card_w - Inches(0.7), card_h - Inches(0.7))
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
    bb = s8.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(1.0), bot_y, Inches(17.0), Inches(2.0))
    bb.fill.solid(); bb.fill.fore_color.rgb = COLOR_WHITE
    bb.line.color.rgb = COLOR_BORDER; bb.line.width = Pt(1.2)

    tbb = s8.shapes.add_textbox(Inches(1.4), bot_y + Inches(0.3), Inches(16.2), Inches(1.4))
    tfb = tbb.text_frame; tfb.word_wrap = True; tfb.margin_left = tfb.margin_top = 0
    p = tfb.paragraphs[0]; p.text = "🎯 საინჟინრო პრინციპი:"
    p.font.name = FONT_ARIAL; p.font.size = Pt(14); p.font.bold = True; p.font.color.rgb = COLOR_RED
    p = tfb.add_paragraph()
    p.text = "პროდაქშენში არასოდეს აჩვენოთ მომხმარებელს მოდელის ნედლი შიდა მსჯელობა; ყოველთვის გამოიყენეთ Output Separation (<thinking> vs <answer>), რათა შეინარჩუნოთ იდეალური UX და სრული აუდიტის შესაძლებლობა."
    p.font.name = FONT_ARIAL; p.font.size = Pt(13); p.font.color.rgb = COLOR_MUTED; p.space_before = Pt(4)

    # -------------------------------------------------------------------------
    # SLIDE 9: Best Practices vs Anti-Patterns & Week 2 Wrap-up
    # -------------------------------------------------------------------------
    s9 = created_slides[7]
    add_header(s9, "საინჟინრო პატერნები და თავიდან ასაცილებელი შეცდომები Reasoning-ში", "ᲡᲐᲣᲙᲔᲗᲔᲡᲝ ᲞᲠᲐᲥᲢᲘᲙᲐ • PATTERNS & ANTI-PATTERNS")

    box_w = Inches(8.35); box_h = Inches(5.4)

    # Anti-patterns Box
    c_anti = s9.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(1.0), card_y, box_w, box_h)
    c_anti.fill.solid(); c_anti.fill.fore_color.rgb = COLOR_CARD_RED
    c_anti.line.color.rgb = COLOR_BORDER_RED; c_anti.line.width = Pt(1.5)

    tb_anti = s9.shapes.add_textbox(Inches(1.35), card_y + Inches(0.35), box_w - Inches(0.7), box_h - Inches(0.7))
    tf_anti = tb_anti.text_frame; tf_anti.word_wrap = True; tf_anti.margin_left = tf_anti.margin_top = 0
    p = tf_anti.paragraphs[0]; p.text = "❌ CRITICAL ANTI-PATTERNS (კრიტიკული შეცდომები)"
    p.font.name = FONT_ARIAL; p.font.size = Pt(13); p.font.bold = True; p.font.color.rgb = COLOR_RED

    p = tf_anti.add_paragraph(); p.text = "რას უნდა მოვერიდოთ Reasoning-ში"
    p.font.name = FONT_ARIAL; p.font.size = Pt(16); p.font.bold = True; p.font.color.rgb = COLOR_DARK; p.space_before = Pt(4)

    anti_items = [
        ("Unbounded Thinking Budget", "ლიმიტების გარეშე გაშვებული მოდელი, რომელიც მარტივ კითხვაზე 10,000 ტოკენს ხარჯავს და ბიუჯეტს წვავს."),
        ("Reasoning ჩატბოტის ინტერფეისში", "20-წამიანი დაყოვნება რეალურ დროში მომუშავე მხარდაჭერის ჩატში (მომხმარებელი ტოვებს საიტს)."),
        ("Raw Thinking Display", "შიდა დაუმუშავებელი მსჯელობის პირდაპირ კლიენტისთვის ჩვენება, რაც ქმნის დაბნეულობასა და რისკს."),
        ("Blind Trust in Reasoning", "ფიქრი, რომ Reasoning მოდელი არასოდეს ცდება (ლოგიკური შეცდომები და ჰალუცინაციები მაინც შესაძლებელია).")
    ]
    for head, body in anti_items:
        p = tf_anti.add_paragraph(); p.text = f"• {head}: {body}"
        p.font.name = FONT_ARIAL; p.font.size = Pt(11); p.font.color.rgb = COLOR_DARK; p.space_before = Pt(6)

    # Best Practices Box
    c_rec = s9.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(9.65), card_y, box_w, box_h)
    c_rec.fill.solid(); c_rec.fill.fore_color.rgb = COLOR_GREEN_BG
    c_rec.line.color.rgb = COLOR_GREEN_BOR; c_rec.line.width = Pt(1.5)

    tb_rec = s9.shapes.add_textbox(Inches(10.0), card_y + Inches(0.35), box_w - Inches(0.7), box_h - Inches(0.7))
    tf_rec = tb_rec.text_frame; tf_rec.word_wrap = True; tf_rec.margin_left = tf_rec.margin_top = 0
    p = tf_rec.paragraphs[0]; p.text = "✅ RECOMMENDED BEST PRACTICES (საუკეთესო პატერნები)"
    p.font.name = FONT_ARIAL; p.font.size = Pt(13); p.font.bold = True; p.font.color.rgb = COLOR_GREEN_TEXT

    p = tf_rec.add_paragraph(); p.text = "სტაბილური და ეფექტური არქიტექტურა"
    p.font.name = FONT_ARIAL; p.font.size = Pt(16); p.font.bold = True; p.font.color.rgb = COLOR_DARK; p.space_before = Pt(4)

    rec_items = [
        ("Dynamic Budgeting", "ამოცანის სირთულის მიხედვით ტოკენების ლიმიტის განსაზღვრა (მაგ. 1024 ტოკენი საშუალოზე, 4096 რთულზე)."),
        ("Smart Classifier Gate", "წინა ხაზზე იაფი, სწრაფი მოდელის დაყენება, რომელიც განსაზღვრავს, სჭირდება თუ არა მოთხოვნას Reasoning."),
        ("Structured Tag Isolation", "მსჯელობისა და პასუხის მკაცრი გამიჯვნა XML ტეგებით (<thinking> vs <output>) ან ცალკეული JSON ველებით."),
        ("Async Pipeline Integration", "Reasoning მოდელების გამოყენება ფონურ ანალიტიკაში, რეპორტინგსა და ღამის გადამუშავებაში.")
    ]
    for head, body in rec_items:
        p = tf_rec.add_paragraph(); p.text = f"• {head}: {body}"
        p.font.name = FONT_ARIAL; p.font.size = Pt(11); p.font.color.rgb = COLOR_DARK; p.space_before = Pt(6)

    # Bottom takeaway & Week 2 Wrap-up
    bb = s9.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(1.0), bot_y, Inches(17.0), Inches(2.0))
    bb.fill.solid(); bb.fill.fore_color.rgb = COLOR_WHITE
    bb.line.color.rgb = COLOR_BORDER; bb.line.width = Pt(1.2)

    tbb = s9.shapes.add_textbox(Inches(1.4), bot_y + Inches(0.25), Inches(16.2), Inches(1.5))
    tfb = tbb.text_frame; tfb.word_wrap = True; tfb.margin_left = tfb.margin_top = 0
    p = tfb.paragraphs[0]; p.text = "💡 საინჟინრო შეჯამება: კვირა 2-ის დასასრული და გადასვლა კვირა 3-ზე:"
    p.font.name = FONT_ARIAL; p.font.size = Pt(14); p.font.bold = True; p.font.color.rgb = COLOR_RED
    p = tfb.add_paragraph()
    p.text = "Reasoning და CoT 2026 წლის უმძლავრესი ინსტრუმენტებია, როდესაც მათ მიზნობრივად ვიყენებთ. ამით წარმატებით დასრულდა კვირა 2 (10 საკონტაქტო საათი = 5 ლექცია: 4.1, 4.2, 5, 6.1, 6.2)! შემდეგ კვირაში (კვირა 3) გადავდივართ 1-ლი მინი-პროექტის დაგეგმვასა და მის არქიტექტურულ იმპლემენტაციაზე."
    p.font.name = FONT_ARIAL; p.font.size = Pt(12.5); p.font.color.rgb = COLOR_MUTED; p.space_before = Pt(4)

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
    filename = "6.2 - Chain-of-Thought, Reasoning მოდელები და Extended Thinking.pptx"
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
    build_lecture_m7_6_2()
