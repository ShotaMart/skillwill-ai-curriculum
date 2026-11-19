import os
import sys
import shutil
import pptx
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN
from pptx.enum.shapes import MSO_SHAPE

def build_lecture_6_1():
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
            p.text = "6.1 - ChatGPT სრულად I: ინტერფეისი, მოდელები (GPT-6), Voice, Canvas და Deep Research"
            p.alignment = PP_ALIGN.CENTER
            p.font.name = FONT_ARIAL
            p.font.size = Pt(34)
            p.font.bold = True
            p.font.color.rgb = COLOR_RED

            p2 = tf.add_paragraph()
            p2.text = "OpenAI-ს უახლესი ეკოსისტემა, მულტიმოდალური ინტერაქცია და სიღრმისეული კვლევა"
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
    def add_header(slide, title, category="ᲛᲝᲓᲣᲚᲘ 5 • OPENAI ეკოსისტემა • 6.1"):
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
    # SLIDE 2: OpenAI Models Evolution & GPT-6 Architecture
    # -------------------------------------------------------------------------
    s2 = created_slides[0]
    add_header(s2, "OpenAI მოდელების ევოლუცია და GPT-6 არქიტექტურა", "არქიტექტურა • ინტეგრირებული მსჯელობის ახალი ეპოქა (2026)")

    # Left: Legacy Milestones
    b1 = s2.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(1.0), box_y, box_w, box_h)
    b1.fill.solid(); b1.fill.fore_color.rgb = COLOR_BG_LIGHT
    b1.line.color.rgb = COLOR_BORDER; b1.line.width = Pt(1.5)

    tb1 = s2.shapes.add_textbox(Inches(1.4), box_y + Inches(0.4), box_w - Inches(0.8), box_h - Inches(0.8))
    tf1 = tb1.text_frame; tf1.word_wrap = True; tf1.margin_left = tf1.margin_top = 0
    p = tf1.paragraphs[0]; p.text = "ᲔᲕᲝᲚᲣᲪᲘᲣᲠᲘ ᲡᲐᲤᲔᲮᲣᲠᲔᲑᲘ (2024-2025)"
    p.font.name = FONT_ARIAL; p.font.size = Pt(12); p.font.bold = True; p.font.color.rgb = COLOR_MUTED

    p = tf1.add_paragraph(); p.text = "ფრაგმენტული მოდელები (GPT-4o და o1/o3)"
    p.font.name = FONT_ARIAL; p.font.size = Pt(21); p.font.bold = True; p.font.color.rgb = COLOR_DARK; p.space_before = Pt(4)

    legacy_points = [
        ("ცალკეული მოდელების პრობლემა", "მომხმარებელს უწევდა არჩევანის გაკეთება: სწრაფი ჩატბოტი (GPT-4o) თუ ნელი მოაზროვნე (o1/o3)."),
        ("Reasoning-ის იზოლაცია", "o-სერიის მოდელები მოითხოვდა ცალკე რეჟიმს, არ ჰქონდათ სრულყოფილი ხელსაწყოების მხარდაჭერა და ჰქონდათ მაღალი დაყოვნება."),
        ("კონტექსტური ბარიერი", "რთულ ამოცანებში მოდელის ცვლა საუბრის შუაგულში არღვევდა სამუშაო პროცესს."),
        ("მოძველებული სუფიქსები", "ტერმინები „Turbo“, „Mini“ და მოდელების მრავალრიცხოვანი ვერსიები აბნევდა ბიზნესს.")
    ]
    for t, d in legacy_points:
        p = tf1.add_paragraph(); p.text = f"• {t}: {d}"
        p.font.name = FONT_ARIAL; p.font.size = Pt(13); p.font.color.rgb = COLOR_TEXT_MUTED; p.space_before = Pt(8)

    # Right: Unified GPT-6 Generation
    b2 = s2.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(9.65), box_y, box_w, box_h)
    b2.fill.solid(); b2.fill.fore_color.rgb = COLOR_CARD_RED
    b2.line.color.rgb = COLOR_BORDER_RED; b2.line.width = Pt(1.5)

    tb2 = s2.shapes.add_textbox(Inches(10.05), box_y + Inches(0.4), box_w - Inches(0.8), box_h - Inches(0.8))
    tf2 = tb2.text_frame; tf2.word_wrap = True; tf2.margin_left = tf2.margin_top = 0
    p = tf2.paragraphs[0]; p.text = "ᲗᲐᲜᲐᲛᲔᲓᲠᲝᲕᲔ ᲡᲢᲐᲜᲓᲐᲠᲢᲘ (2026)"
    p.font.name = FONT_ARIAL; p.font.size = Pt(12); p.font.bold = True; p.font.color.rgb = COLOR_RED

    p = tf2.add_paragraph(); p.text = "ერთიანი GPT-6 და ინტეგრირებული მსჯელობა"
    p.font.name = FONT_ARIAL; p.font.size = Pt(21); p.font.bold = True; p.font.color.rgb = COLOR_DARK; p.space_before = Pt(4)

    unified_points = [
        ("Integrated Reasoning (ჩაშენებული ლოგიკა)", "მსჯელობის უნარი (Chain-of-Thought) სრულად ინტეგრირებულია GPT-6-ის ბირთვულ არქიტექტურაში."),
        ("Scalable Thinking Effort (მსჯელობის რეგულირება)", "მომხმარებელს ან სისტემას შეუძლია დინამიკურად აირჩიოს ფიქრის სიღრმე: Low, Medium, High, Max."),
        ("უნივერსალური მულტიმოდალობა", "ხედვა, ხმა, ტექსტი, კოდი და ხელსაწყოების გამოყენება ერთიანად ხელმისაწვდომია ნებისმიერი დონის ამოცანაში."),
        ("Test-Time Compute (დრო გამოთვლებზე)", "რთულ ამოცანებზე მოდელი ხარჯავს დამატებით გამოთვლით დროს, რაც რადიკალურად ამცირებს შეცდომებს.")
    ]
    for t, d in unified_points:
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
    p.text = "თუ ადრე ცალ-ცალკე გვყავდა სწრაფი მთხრობელი (GPT-4) და ღრმად მოაზროვნე მათემატიკოსი (o1), GPT-6 თაობაში ისინი გაერთიანდნენ ერთ ტვინში, სადაც თავად ირჩევთ, რამდენად ღრმად უნდა დაფიქრდეს ასისტენტი პასუხის გაცემამდე."
    p.font.name = FONT_ARIAL; p.font.size = Pt(13); p.font.color.rgb = COLOR_MUTED; p.space_before = Pt(4)

    # -------------------------------------------------------------------------
    # SLIDE 3: OpenAI Capability Tiers (Astra, Sol, Terra, Luna)
    # -------------------------------------------------------------------------
    s3 = created_slides[1]
    add_header(s3, "OpenAI Capability Tiers (Astra, Sol, Terra და Luna)", "მოდელების ხაზი • კოსმოსური სხეულების სახელდების ახალი სისტემა (2026)")

    tiers_s3 = [
        ("Astra (Frontier Flagship)", "უმაღლესი ინტელექტი", [
            ("დანიშნულება", "OpenAI-ს ყველაზე ძლიერი მოდელი რთული კვლევებისთვის, საინჟინრო ამოცანებისა და ავტონომიური აგენტებისთვის."),
            ("ხელმისაწვდომობა", "ChatGPT Pro ტარიფში და საწარმოო API-ში მაქსიმალური მსჯელობის სიღრმით (Max Thinking Effort).")
        ]),
        ("Sol (Workhorse - GPT-6.2 Sol)", "პროფესიული სტანდარტი", [
            ("დანიშნულება", "ძირითადი მოდელი ბიზნესისთვის, დოკუმენტების მრავალმხრივი ანალიზისთვის, კოდირებისა და რთული სამუშაოებისთვის."),
            ("ბალანსი", "Astra-ს დონის ინტელექტი ოპტიმალური სისწრაფითა და ხელსაყრელი ფასით (Plus/Team სტანდარტი).")
        ]),
        ("Terra (Balanced Everyday)", "ვიზუალი და ავტომატიზაცია", [
            ("დანიშნულება", "ოპტიმიზებულია სწრაფი ვიზუალური ანალიზისთვის, ბუნებრივი წერისა და კომპიუტერის/ბრაუზერის ავტომატიზაციისთვის."),
            ("Computer Use", "იდეალურია UI ელემენტების ამოცნობისა და ინტერაქტიული სამუშაო მაგიდის მართვისთვის.")
        ]),
        ("Luna (High-Volume Economy)", "სისწრაფე და ეკონომია", [
            ("დანიშნულება", "ულტრა-სწრაფი, დაბალხარჯიანი მოდელი რუტინული ოპერაციებისთვის, მონაცემთა ექსტრაქციისა და კლასიფიკაციისთვის."),
            ("მასშტაბი", "უზრუნველყოფს Free ტარიფის სტაბილურ მუშაობას და მილიონობით მოთხოვნის მყისიერ დამუშავებას.")
        ])
    ]

    for i, (title, subtitle, items) in enumerate(tiers_s3):
        bg_col = COLOR_CARD_RED if i == 1 else COLOR_BG_LIGHT
        bor_col = COLOR_BORDER_RED if i == 1 else COLOR_BORDER
        c = s3.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, xs_4[i], step_y, step_w, step_h)
        c.fill.solid(); c.fill.fore_color.rgb = bg_col
        c.line.color.rgb = bor_col; c.line.width = Pt(1.5)

        tb = s3.shapes.add_textbox(xs_4[i] + Inches(0.25), step_y + Inches(0.35), step_w - Inches(0.5), step_h - Inches(0.7))
        tf = tb.text_frame; tf.word_wrap = True; tf.margin_left = tf.margin_top = 0

        p = tf.paragraphs[0]; p.text = f"TIER {i+1}".upper()
        p.font.name = FONT_ARIAL; p.font.size = Pt(11); p.font.bold = True
        p.font.color.rgb = COLOR_RED if i == 1 else COLOR_MUTED

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
    p = tfb.paragraphs[0]; p.text = "💡 არქიტექტურული პრინციპი:"
    p.font.name = FONT_ARIAL; p.font.size = Pt(14); p.font.bold = True; p.font.color.rgb = COLOR_RED
    p = tfb.add_paragraph()
    p.text = "OpenAI-ს სახელდების ახალი სისტემა (Astra, Sol, Terra, Luna) მომხმარებლებსა და ინჟინრებს აძლევს მკაფიო ორიენტირს: ყოველ ამოცანას ერგება ზუსტად შერჩეული დონე ინტელექტის, სისწრაფისა და ბიუჯეტის მიხედვით."
    p.font.name = FONT_ARIAL; p.font.size = Pt(13); p.font.color.rgb = COLOR_MUTED; p.space_before = Pt(4)

    # -------------------------------------------------------------------------
    # SLIDE 4: Pricing & Subscription Tiers (2026)
    # -------------------------------------------------------------------------
    s4 = created_slides[2]
    add_header(s4, "სატარიფო გეგმები (Free, Go, Plus, Pro, Business და Enterprise)", "ეკონომიკა • რომელი პაკეტი რა საჭიროებას პასუხობს (2026)")

    plans_s4 = [
        ("Free & Go ($0 - $8)", "საბაზისო წვდომა", [
            ("Free ($0/თვე)", "ყოველდღიური საუბრები GPT-6 Luna-თი, საბაზისო ძებნა და ფაილების შეზღუდული ატვირთვა."),
            ("Go (~$8/თვე)", "მსუბუქი მომხმარებლებისთვის: გაზრდილი ლიმიტები, Projects და გაფართოებული მეხსიერება.")
        ]),
        ("Plus ($20/თვე)", "პროფესიონალის არჩევანი", [
            ("ძირითადი მოდელი", "GPT-6 Sol სრული ინტეგრირებული მსჯელობით (Reasoning) და მაღალი ლიმიტებით."),
            ("ინსტრუმენტები", "OpenAI Canvas, Deep Research, Custom GPT-ების შექმნა და Advanced Voice Mode.")
        ]),
        ("Pro ($100 - $200/თვე)", "Power Users & მკვლევრები", [
            ("Frontier Intelligence", "შეუზღუდავი წვდომა GPT-6 Astra-ზე მაქსიმალური გამოთვლითი დროით (Max Thinking)."),
            ("Codex & Agent Suite", "ავტონომიური დეველოპერული აგენტები, CLI ხელსაწყოები და ულტრა-სწრაფი გამოთვლები.")
        ]),
        ("Business & Enterprise", "კორპორაციული მართვა", [
            ("Business ($20-25/user)", "მინიმუმ 2 მომხმარებელი, ადმინ-პანელი, საერთო პროექტები, მონაცემთა წვრთნის აკრძალვა."),
            ("Enterprise", "მორგებული კონტრაქტი, ZDR (Zero Data Retention), SSO/SCIM, HIPAA და შეუზღუდავი სისწრაფე.")
        ])
    ]

    for i, (title, subtitle, items) in enumerate(plans_s4):
        bg_col = COLOR_CARD_RED if i == 1 else COLOR_BG_LIGHT
        bor_col = COLOR_BORDER_RED if i == 1 else COLOR_BORDER
        c = s4.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, xs_4[i], step_y, step_w, step_h)
        c.fill.solid(); c.fill.fore_color.rgb = bg_col
        c.line.color.rgb = bor_col; c.line.width = Pt(1.5)

        tb = s4.shapes.add_textbox(xs_4[i] + Inches(0.25), step_y + Inches(0.35), step_w - Inches(0.5), step_h - Inches(0.7))
        tf = tb.text_frame; tf.word_wrap = True; tf.margin_left = tf.margin_top = 0

        p = tf.paragraphs[0]; p.text = f"ᲞᲐᲙᲔᲢᲘ {i+1}".upper()
        p.font.name = FONT_ARIAL; p.font.size = Pt(11); p.font.bold = True
        p.font.color.rgb = COLOR_RED if i == 1 else COLOR_MUTED

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
    p = tfb.paragraphs[0]; p.text = "💡 შერჩევის რეკომენდაცია:"
    p.font.name = FONT_ARIAL; p.font.size = Pt(14); p.font.bold = True; p.font.color.rgb = COLOR_RED
    p = tfb.add_paragraph()
    p.text = "ინდივიდუალური პროფესიონალისთვის ოპტიმალურია Plus ($20) ტარიფი GPT-6 Sol-ით; კომპანიებისთვის აუცილებელია Business/Enterprise მონაცემთა კონფიდენციალურობისა და მოდელის წვრთნის აკრძალვის (Zero Training) გამო."
    p.font.name = FONT_ARIAL; p.font.size = Pt(13); p.font.color.rgb = COLOR_MUTED; p.space_before = Pt(4)

    # -------------------------------------------------------------------------
    # SLIDE 5: Multimodal Interface (Files, Data Analysis & Vision)
    # -------------------------------------------------------------------------
    s5 = created_slides[3]
    add_header(s5, "მულტიმოდალური ინტერფეისი (ფაილები, ცხრილები და Advanced Vision)", "ინსტრუმენტები • როგორ ამუშავებს მოდელი შერეულ მონაცემებს")

    cards_s5 = [
        ("📄 დოკუმენტები", "დოკუმენტების დამუშავება", "PDF, Word, Codebases", [
            ("მასიური ფაილების კითხვა", "PDF, DOCX, TXT, JSON და პროგრამული კოდის რეპოზიტორიების ანალიზი."),
            ("კონტექსტური სინთეზი", "მრავალგვერდიანი ხელშეკრულებებისა და კვლევების სწრაფი შეჯამება და კითხვა-პასუხი."),
            ("Cross-Referencing", "რამდენიმე ფაილის ერთდროული შედარება და შეუსაბამობების იდენტიფიცირება.")
        ]),
        ("📊 ანალიტიკა", "Advanced Data Analysis", "Python Sandbox ღრუბელში", [
            ("ავტონომიური Python გარემო", "ChatGPT უშვებს Python კოდს იზოლირებულ ვირტუალურ მანქანაში რეალურ დროში."),
            ("მონაცემთა გაწმენდა და გაანგარიშება", "Excel/CSV ცხრილების სტატისტიკური ანალიზი, რეგრესიები და პროგნოზები."),
            ("ვიზუალიზაცია", "ინტერაქტიული დიაგრამების, გრაფიკებისა და სითბური რუკების (Heatmaps) გენერაცია.")
        ]),
        ("👁️ VISION", "Advanced Vision & Layout", "გრაფიკული აღქმა", [
            ("ვიზუალური სტრუქტურის აღქმა", "სქემების, არქიტექტურული ნახაზების, ინვოისებისა და ჩეკების დეტალური ამოცნობა."),
            ("UI/UX ანალიტიკა", "ინტერფეისის მაკეტების (Figma/Web სქრინების) შეფასება და კოდში გადაყვანა."),
            ("Spatial Understanding", "ობიექტების ურთიერთგანლაგებისა და ვიზუალური იერარქიის ზუსტი გააზრება.")
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
    p = tfb.paragraphs[0]; p.text = "💡 პრაქტიკული ძალა:"
    p.font.name = FONT_ARIAL; p.font.size = Pt(14); p.font.bold = True; p.font.color.rgb = COLOR_RED
    p = tfb.add_paragraph()
    p.text = "მულტიმოდალობა ნიშნავს, რომ აღარ გჭირდებათ მონაცემების წინასწარი ხელით გარდაქმნა — შეგიძლიათ ატვირთოთ ფინანსური PDF, Excel-ის დაუმუშავებელი ცხრილი და დიაგრამის ფოტო ერთდროულად, და მოდელი მათ ერთიან კონტექსტში გააანალიზებს."
    p.font.name = FONT_ARIAL; p.font.size = Pt(13); p.font.color.rgb = COLOR_MUTED; p.space_before = Pt(4)

    # -------------------------------------------------------------------------
    # SLIDE 6: Advanced Voice Mode (Native Audio Streaming)
    # -------------------------------------------------------------------------
    s6 = created_slides[4]
    add_header(s6, "Advanced Voice Mode (მშობლიური აუდიო-სტრიმინგი)", "ინტერაქცია • ხმოვანი კომუნიკაცია შუალედური ტექსტის გარეშე")

    # Left: Old TTS/STT vs Native Audio
    b1 = s6.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(1.0), box_y, box_w, box_h)
    b1.fill.solid(); b1.fill.fore_color.rgb = COLOR_BG_LIGHT
    b1.line.color.rgb = COLOR_BORDER; b1.line.width = Pt(1.5)

    tb1 = s6.shapes.add_textbox(Inches(1.4), box_y + Inches(0.4), box_w - Inches(0.8), box_h - Inches(0.8))
    tf1 = tb1.text_frame; tf1.word_wrap = True; tf1.margin_left = tf1.margin_top = 0
    p = tf1.paragraphs[0]; p.text = "ᲫᲕᲔᲚᲘ 3-ᲔᲢᲐᲞᲘᲐᲜᲘ ᲛᲝᲓᲔᲚᲘ (LEGACY)"
    p.font.name = FONT_ARIAL; p.font.size = Pt(12); p.font.bold = True; p.font.color.rgb = COLOR_MUTED

    p = tf1.add_paragraph(); p.text = "STT ➔ LLM ➔ TTS მილსადენი"
    p.font.name = FONT_ARIAL; p.font.size = Pt(21); p.font.bold = True; p.font.color.rgb = COLOR_DARK; p.space_before = Pt(4)

    voice_old_points = [
        ("მაღალი დაყოვნება (Latency)", "ხმა ჯერ გარდაიქმნებოდა ტექსტად (Whisper), მუშავდებოდა LLM-ში და შემდეგ სხვა მოდელი ახმოვანებდა (2-3 წამიანი პაუზა)."),
        ("ინტონაციის დაკარგვა", "მოდელი ვერ აღიქვამდა მომხმარებლის ემოციას, სარკაზმს, ჩურჩულს, სუნთქვასა და ტემპს."),
        ("შეწყვეტის შეუძლებლობა", "მომხმარებელს უწევდა მოდელის მონოლოგის ბოლომდე მოსმენა საუბარში ჩარევის გარეშე."),
        ("რობოტული ხმა", "სინთეზირებული ხმა იყო მონოტონური და არაბუნებრივი.")
    ]
    for t, d in voice_old_points:
        p = tf1.add_paragraph(); p.text = f"• {t}: {d}"
        p.font.name = FONT_ARIAL; p.font.size = Pt(13); p.font.color.rgb = COLOR_TEXT_MUTED; p.space_before = Pt(8)

    # Right: Native Audio-to-Audio
    b2 = s6.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(9.65), box_y, box_w, box_h)
    b2.fill.solid(); b2.fill.fore_color.rgb = COLOR_CARD_RED
    b2.line.color.rgb = COLOR_BORDER_RED; b2.line.width = Pt(1.5)

    tb2 = s6.shapes.add_textbox(Inches(10.05), box_y + Inches(0.4), box_w - Inches(0.8), box_h - Inches(0.8))
    tf2 = tb2.text_frame; tf2.word_wrap = True; tf2.margin_left = tf2.margin_top = 0
    p = tf2.paragraphs[0]; p.text = "ADVANCED VOICE MODE (2026)"
    p.font.name = FONT_ARIAL; p.font.size = Pt(12); p.font.bold = True; p.font.color.rgb = COLOR_RED

    p = tf2.add_paragraph(); p.text = "Native Audio-to-Audio Multimodal"
    p.font.name = FONT_ARIAL; p.font.size = Pt(21); p.font.bold = True; p.font.color.rgb = COLOR_DARK; p.space_before = Pt(4)

    voice_new_points = [
        ("მყისიერი რეაქცია (~300ms)", "ხმა მუშავდება პირდაპირ აუდიო ტოკენებად, რაც ადამიანური საუბრის სისწრაფის იდენტურია."),
        ("ემოციური ინტელექტი", "აღიქვამს ხმის ტემბრს, სიცილს, ეჭვს და თავადაც პასუხობს შესაბამისი ემოციური ინტონაციით."),
        ("Natural Interruption (მყისიერი შეწყვეტა)", "შეგიძლიათ ნებისმიერ სიტყვაზე შეაწყვეტინოთ საუბარი და მოდელი მყისიერად გაჩერდება და მოგისმენთ."),
        ("ვიდეო/კამერა ინტეგრაცია", "საუბრის პარალელურად მოდელს შეუძლია რეალურ დროში დაინახოს თქვენი კამერის გამოსახულება.")
    ]
    for t, d in voice_new_points:
        p = tf2.add_paragraph(); p.text = f"⭐ {t}"
        p.font.name = FONT_ARIAL; p.font.size = Pt(13); p.font.bold = True; p.font.color.rgb = COLOR_DARK; p.space_before = Pt(8)
        p = tf2.add_paragraph(); p.text = d
        p.font.name = FONT_ARIAL; p.font.size = Pt(12); p.font.color.rgb = COLOR_TEXT_MUTED; p.space_before = Pt(2)

    # Bottom takeaway
    bb = s6.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(1.0), bot_y, Inches(17.0), Inches(2.0))
    bb.fill.solid(); bb.fill.fore_color.rgb = COLOR_WHITE
    bb.line.color.rgb = COLOR_BORDER; bb.line.width = Pt(1.2)

    tbb = s6.shapes.add_textbox(Inches(1.4), bot_y + Inches(0.3), Inches(16.2), Inches(1.4))
    tfb = tbb.text_frame; tfb.word_wrap = True; tfb.margin_left = tfb.margin_top = 0
    p = tfb.paragraphs[0]; p.text = "💡 მთავარი არსი (მარტივი ანალოგია):"
    p.font.name = FONT_ARIAL; p.font.size = Pt(14); p.font.bold = True; p.font.color.rgb = COLOR_RED
    p = tfb.add_paragraph()
    p.text = "Advanced Voice Mode არის ნამდვილი ცოცხალი დიალოგი: AI საუბრობს არა როგორც რობოტი მთარგმნელი, არამედ როგორც ადამიანი ტელეფონზე — შეუძლია იხუმროს, შეცვალოს ტონი და გაჩერდეს როგორც კი სიტყვას შეაშველებთ."
    p.font.name = FONT_ARIAL; p.font.size = Pt(13); p.font.color.rgb = COLOR_MUTED; p.space_before = Pt(4)

    # -------------------------------------------------------------------------
    # SLIDE 7: OpenAI Canvas (Interactive Workspace)
    # -------------------------------------------------------------------------
    s7 = created_slides[5]
    add_header(s7, "OpenAI Canvas (ინტერაქტიული სამუშაო სივრცე)", "პროდუქტიულობა • ტექსტისა და კოდის პარალელური რედაქტირება")

    cards_s7 = [
        ("🖥️ INTERFACE", "Side-by-Side Canvas", "ინტერაქტიული დაფა", [
            ("ორპანელიანი სამუშაო გარემო", "მარცხნივ რჩება საუბრის ჩატი, მარჯვნივ იხსნება სრული დოკუმენტი ან კოდი."),
            ("კონტექსტის შენარჩუნება", "ჩატში მოთხოვნის გაგზავნისას მთელი ტექსტი თავიდან არ გენერირდება — იცვლება მხოლოდ საჭირო ნაწილი."),
            ("ვერსიების ისტორია", "ნებისმიერ დროს შეგიძლიათ წინა ვერსიაზე დაბრუნება (Undo/History Slider).")
        ]),
        ("✍️ WRITING", "ტექსტის დამუშავება", "რედაქტორის ინსტრუმენტები", [
            ("Inline Editing", "მონიშნეთ კონკრეტული აბზაცი და უთხარით: „ეს წინადადება გახადე უფრო დამაჯერებელი“."),
            ("სტილისა და სიგრძის სლაიდერი", "ტექსტის სიგრძის (მოკლე / ვრცელი) და კითხვის დონის (საბავშვო / აკადემიური) ერთი დაწკაპუნებით რეგულირება."),
            ("ფინალური გაპრიალება", "გრამატიკული შემოწმება, ემოციური ტონის ცვლილება და ფორმატირება.")
        ]),
        ("💻 CODING", "კოდის ედიტორი", "დეველოპერის დამხმარე", [
            ("Inline Code Review", "Claude Artifacts-ის ანალოგურად, კოდი ჩანს ცალკე ედიტორში სინტაქსის განათებით."),
            ("ავტომატური ბაგ-ფიქსი", "მონიშნულ ფუნქციაზე შეცდომის პოვნა, ლოგების დამატება და კომენტარების ჩაწერა."),
            ("ენების პორტირება", "კოდის ერთი ენიდან მეორეზე (მაგ. Python-იდან TypeScript-ზე) სწრაფი გადაყვანა.")
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
    p = tfb.paragraphs[0]; p.text = "💡 პარადიგმის ცვლა:"
    p.font.name = FONT_ARIAL; p.font.size = Pt(14); p.font.bold = True; p.font.color.rgb = COLOR_RED
    p = tfb.add_paragraph()
    p.text = "Canvas აქცევს ChatGPT-ს სრულფასოვან ტექსტურ და პროგრამულ პროცესორად: ნაცვლად იმისა, რომ ყოველ შესწორებაზე მოდელმა 10 გვერდი თავიდან გამოგიგზავნოთ, თქვენ და AI ერთად ასწორებთ კონკრეტულ წინადადებებსა და კოდის ხაზებს."
    p.font.name = FONT_ARIAL; p.font.size = Pt(13); p.font.color.rgb = COLOR_MUTED; p.space_before = Pt(4)

    # -------------------------------------------------------------------------
    # SLIDE 8: SearchGPT & Deep Research (Autonomous Web Intelligence)
    # -------------------------------------------------------------------------
    s8 = created_slides[6]
    add_header(s8, "SearchGPT და Deep Research (ავტონომიური კვლევა)", "ძიება • ცოცხალი ვებ-ინფორმაციიდან მრავალეტაპიან აკადემიურ ანალიზამდე")

    # Left: SearchGPT
    b1 = s8.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(1.0), box_y, box_w, box_h)
    b1.fill.solid(); b1.fill.fore_color.rgb = COLOR_BG_LIGHT
    b1.line.color.rgb = COLOR_BORDER; b1.line.width = Pt(1.5)

    tb1 = s8.shapes.add_textbox(Inches(1.4), box_y + Inches(0.4), box_w - Inches(0.8), box_h - Inches(0.8))
    tf1 = tb1.text_frame; tf1.word_wrap = True; tf1.margin_left = tf1.margin_top = 0
    p = tf1.paragraphs[0]; p.text = "ᲡᲬᲠᲐᲤᲘ ᲕᲔᲑ-ᲫᲘᲔᲑᲐ (REAL-TIME)"
    p.font.name = FONT_ARIAL; p.font.size = Pt(12); p.font.bold = True; p.font.color.rgb = COLOR_MUTED

    p = tf1.add_paragraph(); p.text = "SearchGPT (ცოცხალი წყაროები)"
    p.font.name = FONT_ARIAL; p.font.size = Pt(21); p.font.bold = True; p.font.color.rgb = COLOR_DARK; p.space_before = Pt(4)

    search_points = [
        ("მყისიერი აქტუალური ინფორმაცია", "ამინდი, აქციების ფასები, უახლესი ამბები და სპორტული შედეგები პირდაპირ ვებიდან."),
        ("ინტერაქტიული ვიჯეტები", "ვიზუალური რუკები, საფონდო გრაფიკები და ცხრილები პასუხის შიგნით."),
        ("Inline ციტირება და გამჭვირვალობა", "ყველა ფაქტს ახლავს პირდაპირი ბმული პირველწყაროზე, რაც გამორიცხავს მოძველებულ მონაცემებს."),
        ("ტრადიციული ძებნის ალტერნატივა", "საძიებო სისტემებში 10 ლურჯი ბმულის ნაცვლად, იღებთ პირდაპირ სინთეზირებულ პასუხს.")
    ]
    for t, d in search_points:
        p = tf1.add_paragraph(); p.text = f"• {t}: {d}"
        p.font.name = FONT_ARIAL; p.font.size = Pt(13); p.font.color.rgb = COLOR_TEXT_MUTED; p.space_before = Pt(8)

    # Right: Deep Research
    b2 = s8.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(9.65), box_y, box_w, box_h)
    b2.fill.solid(); b2.fill.fore_color.rgb = COLOR_CARD_RED
    b2.line.color.rgb = COLOR_BORDER_RED; b2.line.width = Pt(1.5)

    tb2 = s8.shapes.add_textbox(Inches(10.05), box_y + Inches(0.4), box_w - Inches(0.8), box_h - Inches(0.8))
    tf2 = tb2.text_frame; tf2.word_wrap = True; tf2.margin_left = tf2.margin_top = 0
    p = tf2.paragraphs[0]; p.text = "ᲐᲕᲢᲝᲜᲝᲛᲘᲣᲠᲘ ᲐᲒᲔᲜᲢᲣᲠᲘ ᲙᲕᲚᲔᲕᲐ (2026)"
    p.font.name = FONT_ARIAL; p.font.size = Pt(12); p.font.bold = True; p.font.color.rgb = COLOR_RED

    p = tf2.add_paragraph(); p.text = "Deep Research (მრავალსაათიანი ანალიზი)"
    p.font.name = FONT_ARIAL; p.font.size = Pt(21); p.font.bold = True; p.font.color.rgb = COLOR_DARK; p.space_before = Pt(4)

    deep_points = [
        ("ავტონომიური მრავალეტაპიანი ძიება", "მოდელი თავად აგენერირებს ათობით საძიებო მოთხოვნას, კითხულობს ასობით სტატიას და PDF კვლევას."),
        ("წყაროების კრიტიკული შეჯერება", "ამოწმებს ურთიერთგამომრიცხავ მონაცემებს და ახდენს ფაქტების კროს-ვალიდაციას."),
        ("ვრცელი ანალიტიკური რეპორტი", "ქმნის 10-20 გვერდიან აკადემიური დონის დოკუმენტს სტრუქტურირებული თავებითა და ციტატებით."),
        ("Background Task Execution", "მუშაობს ფონურ რეჟიმში 15-40 წუთის განმავლობაში, სანამ მომხმარებელი დაკავებულია სხვა საქმით.")
    ]
    for t, d in deep_points:
        p = tf2.add_paragraph(); p.text = f"⭐ {t}"
        p.font.name = FONT_ARIAL; p.font.size = Pt(13); p.font.bold = True; p.font.color.rgb = COLOR_DARK; p.space_before = Pt(8)
        p = tf2.add_paragraph(); p.text = d
        p.font.name = FONT_ARIAL; p.font.size = Pt(12); p.font.color.rgb = COLOR_TEXT_MUTED; p.space_before = Pt(2)

    # Bottom takeaway
    bb = s8.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(1.0), bot_y, Inches(17.0), Inches(2.0))
    bb.fill.solid(); bb.fill.fore_color.rgb = COLOR_WHITE
    bb.line.color.rgb = COLOR_BORDER; bb.line.width = Pt(1.2)

    tbb = s8.shapes.add_textbox(Inches(1.4), bot_y + Inches(0.3), Inches(16.2), Inches(1.4))
    tfb = tbb.text_frame; tfb.word_wrap = True; tfb.margin_left = tfb.margin_top = 0
    p = tfb.paragraphs[0]; p.text = "💡 ძიების ახალი სტანდარტი:"
    p.font.name = FONT_ARIAL; p.font.size = Pt(14); p.font.bold = True; p.font.color.rgb = COLOR_RED
    p = tfb.add_paragraph()
    p.text = "SearchGPT პასუხობს კითხვას: „რა ხდება ახლა?“, ხოლო Deep Research პასუხობს კითხვას: „გამიკეთე ბაზრის სრული კვლევა, შეადარე 20 კონკურენტი და დამიწერე საინვესტიციო მემორანდუმი“."
    p.font.name = FONT_ARIAL; p.font.size = Pt(13); p.font.color.rgb = COLOR_MUTED; p.space_before = Pt(4)

    # -------------------------------------------------------------------------
    # SLIDE 9: Memory, Personalization & User Controls
    # -------------------------------------------------------------------------
    s9 = created_slides[7]
    add_header(s9, "მეხსიერება და პერსონალიზაცია (Memory & User Controls)", "პერსონალიზაცია • როგორ იმახსოვრებს ChatGPT მომხმარებლის პროფილს")

    cards_s9 = [
        ("🧠 MEMORY", "Cross-Chat Memory", "ფაქტების დამახსოვრება", [
            ("ავტომატური დასწავლა", "მოდელი საუბრისას იმახსოვრებს თქვენს პროფესიას, ოჯახურ სტატუსს, პროექტებსა და ჩვევებს."),
            ("პირდაპირი მითითებები", "შეგიძლიათ პირდაპირ უთხრათ: „დაიმახსოვრე, რომ ჩემი კომპანია იყენებს Python/Django სტეკს“."),
            ("მუდმივი კონტექსტი", "ყოველი ახალი საუბარი იწყება თქვენი კონტექსტის გათვალისწინებით.")
        ]),
        ("⚙️ CONTROLS", "Custom Instructions", "ქცევის გლობალური წესები", [
            ("როლის განსაზღვრა", "„ვინ ხართ თქვენ?“ — ეხმარება მოდელს პასუხების თქვენს ცოდნის დონეზე მორგებაში."),
            ("პასუხის სტილი", "„როგორ გსურთ რომ გიპასუხოთ?“ — ლაკონიური, აკადემიური, იუმორით, კოდის ფორმატში."),
            ("მკაცრი შეზღუდვები", "აკრძალვების გაწერა (მაგ. „არასოდეს გამოიყენო emoji-ები“, „მიპასუხე მხოლოდ ქართულად“).")
        ]),
        ("🔒 PRIVACY", "კონფიდენციალურობა", "მონაცემთა მართვა", [
            ("Memory Management", "პარამეტრებში შეგიძლიათ სათითაოდ ნახოთ, შეცვალოთ ან წაშალოთ დამახსოვრებული ფაქტები."),
            ("Temporary Chats", "დროებითი ჩატები არ ინახავს ისტორიას და არ ამატებს ახალ ფაქტებს მეხსიერებაში."),
            ("Training Opt-Out", "შეგიძლიათ გამორთოთ თქვენი ჩატების გამოყენება OpenAI-ს მომავალი მოდელების საწვრთნელად.")
        ])
    ]

    for i, (badge, title, subtitle, items) in enumerate(cards_s9):
        bg_col = COLOR_CARD_RED if i == 0 else COLOR_BG_LIGHT
        bor_col = COLOR_BORDER_RED if i == 0 else COLOR_BORDER
        c = s9.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, xs_3[i], card_y, card_w, card_h)
        c.fill.solid(); c.fill.fore_color.rgb = bg_col
        c.line.color.rgb = bor_col; c.line.width = Pt(1.5)

        tb = s9.shapes.add_textbox(xs_3[i] + Inches(0.35), card_y + Inches(0.35), card_w - Inches(0.7), card_h - Inches(0.7))
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
    bb = s9.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(1.0), bot_y, Inches(17.0), Inches(2.0))
    bb.fill.solid(); bb.fill.fore_color.rgb = COLOR_WHITE
    bb.line.color.rgb = COLOR_BORDER; bb.line.width = Pt(1.2)

    tbb = s9.shapes.add_textbox(Inches(1.4), bot_y + Inches(0.3), Inches(16.2), Inches(1.4))
    tfb = tbb.text_frame; tfb.word_wrap = True; tfb.margin_left = tfb.margin_top = 0
    p = tfb.paragraphs[0]; p.text = "🎯 შემაჯამებელი დასკვნა:"
    p.font.name = FONT_ARIAL; p.font.size = Pt(14); p.font.bold = True; p.font.color.rgb = COLOR_RED
    p = tfb.add_paragraph()
    p.text = "ChatGPT აღარ არის უბრალოდ ჩატბოტი: GPT-6 ინტეგრირებული მსჯელობით, Advanced Voice-ით, Canvas-ით, Deep Research-ითა და მეხსიერებით ის წარმოადგენს უნივერსალურ პერსონალურ ოპერაციულ სისტემას ინტელექტუალური შრომისთვის."
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
    filename = "6.1 - ChatGPT სრულად I - ინტერფეისი, მოდელები, Voice, Canvas და Deep Research.pptx"
    out_master = os.path.join(r'd:\ხელოვნური ინტელექტის ინსტრუმენტები', filename)
    prs.save(out_master)
    print(f"Master PPTX for Lecture 6.1 generated successfully: {out_master}")

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
    build_lecture_6_1()
