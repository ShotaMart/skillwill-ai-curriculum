import os
import sys
import shutil
import pptx
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN
from pptx.enum.shapes import MSO_SHAPE

def build_lecture_m7_7_2():
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
            p.text = "7.2 - რთული ამოცანების დეკომპოზიცია და Prompt Chaining არქიტექტურა"
            p.alignment = PP_ALIGN.CENTER
            p.font.name = FONT_ARIAL
            p.font.size = Pt(30)
            p.font.bold = True
            p.font.color.rgb = COLOR_RED

            p2 = tf.add_paragraph()
            p2.text = "მონოლითური პრომპტის დეკომპოზიცია, State Management და 3-საფეხურიანი საწარმოო მილსადენი"
            p2.alignment = PP_ALIGN.CENTER
            p2.font.name = FONT_ARIAL
            p2.font.size = Pt(17)
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
    def add_header(slide, title, category="ᲛᲝᲓᲣᲚᲘ 7 • PROMPT ENGINEERING • 7.2"):
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
    # SLIDE 2: Limits of Monolithic Prompts
    # -------------------------------------------------------------------------
    s2 = created_slides[0]
    add_header(s2, "რატომ მარცხდება მონოლითური პრომპტი? ყურადღების დეფიციტი და ქაოსი", "ᲞᲐᲠᲐᲓᲘᲒᲛᲐ • LIMITS OF MONOLITHIC PROMPTS")

    cards_s2 = [
        ("🧠 COGNITIVE OVERLOAD", "კოგნიტური გადატვირთვა", "ყველაფერი ერთ პრომპტში", [
            ("კონფლიქტური მიზნები", "როდესაც ერთ პრომპტს ერთდროულად ვთხოვთ მონაცემთა ამოღებას, ლოგიკურ ანალიზს, წერასა და PII უსაფრთხოებას."),
            ("Attention დეგრადაცია", "მოდელის Self-Attention შრეები ვერ ახერხებს ათობით განსხვავებული ინსტრუქციის თანაბარი წონით დამუშავებას."),
            ("ხარისხის კომპრომისი", "მოდელი კარგად აკეთებს ერთ ნაწილს, მაგრამ უხეშ შეცდომებს უშვებს მეორე, თანაბრად მნიშვნელოვან ნაწილში.")
        ]),
        ("📉 LOST IN THE MIDDLE", "შუაში დაკარგვა", "გრძელი კონტექსტის მახე", [
            ("პოზიციური მიკერძოება", "მოდელი საუკეთესოდ ამჩნევს ტექსტის დასაწყისს და დასასრულს, ხოლო შუა ნაწილში არსებული წესები იგნორირებულია."),
            ("შეზღუდვების დარღვევა", "გრძელ მონოლითურ პრომპტებში (2,000+ ტოკენი) ფორმატისა და ტონის შეზღუდვები ხშირად ირღვევა."),
            ("არასტაბილური ქცევა", "ყოველ გამოძახებაზე განსხვავებული წესი 'ავიწყდება' მოდელს, რაც იწვევს არაპროგნოზირებად შედეგებს.")
        ]),
        ("🔍 DEBUGGING NIGHTMARE", "დიაგნოსტიკის ჩიხი", "გაურკვეველი შეცდომები", [
            ("გაუმჭვირვალე პროცესი", "თუ მონოლითურმა სისტემამ დააბრუნა არასწორი ანგარიში, შეუძლებელია იმის გაგება, სად დაუშვა შეცდომა — ამოღებისას თუ სინთეზისას."),
            ("Unintended Regression", "ერთი წესის ფორმულირების ჩასწორება მოულოდნელად აფუჭებს სხვა წესების მუშაობას."),
            ("Unit ტესტების არარსებობა", "შეუძლებელია ცალკეული ქვეგადაწყვეტილების დამოუკიდებლად იზოლირება და გატესტვა.")
        ])
    ]

    for i, (badge, title, subtitle, items) in enumerate(cards_s2):
        bg_col = COLOR_CARD_RED if i == 0 else COLOR_BG_LIGHT
        bor_col = COLOR_BORDER_RED if i == 0 else COLOR_BORDER
        c = s2.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, xs_3[i], card_y, card_w, card_h)
        c.fill.solid(); c.fill.fore_color.rgb = bg_col
        c.line.color.rgb = bor_col; c.line.width = Pt(1.5)

        tb = s2.shapes.add_textbox(xs_3[i] + Inches(0.35), card_y + Inches(0.35), card_w - Inches(0.7), card_h - Inches(0.7))
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
    bb = s2.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(1.0), bot_y, Inches(17.0), Inches(2.0))
    bb.fill.solid(); bb.fill.fore_color.rgb = COLOR_WHITE
    bb.line.color.rgb = COLOR_BORDER; bb.line.width = Pt(1.2)

    tbb = s2.shapes.add_textbox(Inches(1.4), bot_y + Inches(0.3), Inches(16.2), Inches(1.4))
    tfb = tbb.text_frame; tfb.word_wrap = True; tfb.margin_left = tfb.margin_top = 0
    p = tfb.paragraphs[0]; p.text = "💡 მთავარი არსი:"
    p.font.name = FONT_ARIAL; p.font.size = Pt(14); p.font.bold = True; p.font.color.rgb = COLOR_RED
    p = tfb.add_paragraph()
    p.text = "მონოლითური პრომპტი არის AI ინჟინერიის 'სპაგეტი კოდი'; რთული ამოცანების ერთიან ტექსტში შეკუმშვა გარდაუვლად იწვევს ყურადღების დეფიციტს, ჰალუცინაციებსა და სისტემის მყიფეობას."
    p.font.name = FONT_ARIAL; p.font.size = Pt(13); p.font.color.rgb = COLOR_MUTED; p.space_before = Pt(4)

    # -------------------------------------------------------------------------
    # SLIDE 3: Principles of Prompt Chaining (Divide & Conquer)
    # -------------------------------------------------------------------------
    s3 = created_slides[1]
    add_header(s3, "Prompt Chaining-ის მენტალური მოდელი: დაყავი და იბატონე AI-ში", "ᲛᲔᲜᲢᲐᲚᲣᲠᲘ ᲛᲝᲓᲔᲚᲘ • PROMPT CHAINING PRINCIPLES")

    chain_prin_s3 = [
        ("🎯 SINGLE RESPONSIBILITY", "ერთიანი პასუხისმგებლობა", "ერთი ნაბიჯი = ერთი ამოცანა", [
            ("იზოლირებული როლი", "თითოეული ნაბიჯი ფოკუსირებულია მკაცრად ერთ საინჟინრო ოპერაციაზე (მაგ. მხოლოდ მონაცემთა ექსტრაქცია)."),
            ("მინიმალური ინსტრუქცია", "მოკლე, ლაკონიური და კონკრეტული პრომპტი, რომელსაც მოდელი 100%-იანი სიზუსტით ასრულებს."),
            ("ჰეტეროგენული მოდელები", "შესაძლებელია თითოეულ ნაბიჯს მოვარგოთ ოპტიმალური მოდელი (სწრაფი ექსტრაქციისთვის, მძლავრი ანალიზისთვის).")
        ]),
        ("📦 ISOLATED CONTEXT", "კონტექსტის იზოლაცია", "მონაცემთა სისუფთავე", [
            ("ხმაურის გაფილტვრა", "ყოველი მომდევნო ნაბიჯი იღებს მხოლოდ იმ სუფთა მონაცემს, რაც მისთვის კრიტიკულად აუცილებელია."),
            ("ტოკენების ეკონომია", "წინა ნაბიჯების შუალედური მსჯელობა არ იგზავნება შემდეგ ეტაპზე, რაც რადიკალურად ზოგავს ბიუჯეტს."),
            ("კონტექსტის ჰიგიენა", "მოდელი მუშაობს კომპაქტურ, სუფთა გარემოში, სადაც არ არსებობს ყურადღების გაფანტვის რისკი.")
        ]),
        ("🛡️ INTERMEDIATE VALIDATION", "შუალედური ვალიდაცია", "ხარისხის კარიბჭეები", [
            ("პროგრამული შემოწმება", "ნაბიჯებს შორის მონაცემი მოწმდება ტრადიციული კოდით (Pydantic, Regex, JSON Schema)."),
            ("შეცდომის შეჩერება", "თუ პირველმა ნაბიჯმა მონაცემი არასწორად ამოიღო, პროცესი ჩერდება და არ გადადის გენერაციის ეტაპზე."),
            ("მოდულური Unit ტესტები", "თითოეული პრომპტი ტესტირდება, იზომება და ოპტიმიზდება სხვებისგან სრულიად დამოუკიდებლად.")
        ])
    ]

    for i, (badge, title, subtitle, items) in enumerate(chain_prin_s3):
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
    p = tfb.paragraphs[0]; p.text = "🎯 საინჟინრო პრინციპი:"
    p.font.name = FONT_ARIAL; p.font.size = Pt(14); p.font.bold = True; p.font.color.rgb = COLOR_RED
    p = tfb.add_paragraph()
    p.text = "Prompt Chaining გარდაქმნის AI-ს გაურკვეველი მონოლითიდან საიმედო პროგრამულ მილსადენად; თითოეული ნაბიჯი არის დამოუკიდებელი, ტესტირებადი ფუნქცია მკაცრი შემავალი და გამომავალი კონტრაქტით."
    p.font.name = FONT_ARIAL; p.font.size = Pt(13); p.font.color.rgb = COLOR_MUTED; p.space_before = Pt(4)

    # -------------------------------------------------------------------------
    # SLIDE 4: 3-Step Production Pipeline (Extraction -> Draft -> Verify)
    # -------------------------------------------------------------------------
    s4 = created_slides[2]
    add_header(s4, "3-საფეხურიანი საწარმოო ჯაჭვი: Extraction ➔ Draft ➔ Verification", "ᲡᲐᲬᲐᲠᲛᲝᲝ ᲛᲘᲚᲡᲐᲓᲔᲜᲘ • 3-STEP PIPELINE")

    pipe_s4 = [
        ("📊 STEP 1: EXTRACTION", "ფაქტების ამოღება", "მშრალი მონაცემების იზოლაცია", [
            ("დანიშნულება", "დოკუმენტიდან მხოლოდ ციფრების, სახელების, თარიღებისა და ფაქტების ამოღება JSON ფორმატში."),
            ("კონფიგურაცია", "ტემპერატურა T=0.0, მკაცრი Pydantic სქემა, ყოველგვარი შემოქმედებითი ტექსტის აკრძალვა."),
            ("გამოსავალი", "სუფთა, ვალიდური JSON ობიექტი მტკიცებულებებითა და წყაროს ციტატებით (Grounded Data).")
        ]),
        ("✍️ STEP 2: SYNTHESIS & DRAFT", "სინთეზი და დრაფტი", "ანალიტიკური ტექსტის შექმნა", [
            ("დანიშნულება", "ამოღებული JSON მონაცემების საფუძველზე სრულფასოვანი ბიზნეს-ანგარიშის ან პასუხის დაწერა."),
            ("კონფიგურაცია", "ტემპერატურა T=0.3 – 0.5, ბრენდის ტონი (Tone of Voice), სტრუქტურული Markdown სათაურები."),
            ("კონტექსტის უპირატესობა", "მოდელი ხედავს მხოლოდ 1-ლი ნაბიჯის JSON-ს და არა 100-გვერდიან ნედლ დოკუმენტს!")
        ]),
        ("🛡️ STEP 3: QUALITY GATE", "ვერიფიკაცია და აუდიტი", "პოლიტიკისა და PII-ს შემოწმება", [
            ("დანიშნულება", "დაწერილი დრაფტის შემოწმება კომპანიის პოლიტიკასთან, უსაფრთხოებასა და ფაქტობრივ სიზუსტესთან."),
            ("კონფიგურაცია", "Evaluator პრომპტი, რომელიც აბრუნებს status: APPROVED / REJECTED და ხარვეზების სიას."),
            ("საბოლოო დაცვა", "პირადი მონაცემების (PII) და სისტემური წესების გაჟონვის საბოლოო ბლოკირება გამოტანამდე.")
        ])
    ]

    for i, (badge, title, subtitle, items) in enumerate(pipe_s4):
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
    p.text = "ეს 3-საფეხურიანი არქიტექტურა არის საწარმოო ოქროს სტანდარტი; მე-2 ნაბიჯს მიეწოდება მხოლოდ სუფთა JSON, რაც 95%-ით გამორიცხავს ჰალუცინაციებს, ხოლო მე-3 ნაბიჯი მოქმედებს როგორც დამოუკიდებელი უსაფრთხოების კარიბჭე."
    p.font.name = FONT_ARIAL; p.font.size = Pt(13); p.font.color.rgb = COLOR_MUTED; p.space_before = Pt(4)

    # -------------------------------------------------------------------------
    # SLIDE 5: State Management & Typed Contracts
    # -------------------------------------------------------------------------
    s5 = created_slides[3]
    add_header(s5, "მონაცემთა გადაცემა და State Management ნაბიჯებს შორის", "ᲐᲠᲥᲘᲢᲔᲥᲢᲣᲠᲐ • STATE MANAGEMENT & CONTRACTS")

    state_s5 = [
        ("📋 TYPED CONTRACTS", "ტიპიზებული კონტრაქტები", "JSON Schema & Pydantic", [
            ("მკაცრი ინტერფეისი", "ყოველი ნაბიჯის გამოსავალი აღწერილია Pydantic მოდელით, რაც უზრუნველყოფს ტიპების სისწორეს."),
            ("გარანტირებული ვალიდაცია", "თუ Step 1-მა დააბრუნა არასწორი JSON, ის არ გადაეცემა Step 2-ს; აპლიკაცია იჭერს შეცდომას."),
            ("დოკუმენტირებადობა", "დეველოპერთა გუნდმა ზუსტად იცის თითოეული ნაბიჯის შემავალი და გამომავალი მონაცემთა სტრუქტურა.")
        ]),
        ("💾 STATEFUL ORCHESTRATION", "მდგომარეობის მართვა", "Stateful აპი vs Stateless LLM", [
            ("მოდელი რჩება Stateless", "LLM არ ინახავს შინაგან მეხსიერებას; ის მხოლოდ ამუშავებს კონკრეტულ მოწოდებულ შეტყობინებას."),
            ("გარე State Engine", "სერვერული აპლიკაცია (Python, Temporal, Redis) ინახავს სესიის მიმდინარეობას, ისტორიას და შედეგებს."),
            ("Pause & Resume", "ხანგრძლივი პროცესების შეჩერება და განახლება (მაგ. ადამიანის დასტურის მოლოდინი HITL რეჟიმში).")
        ]),
        ("🧹 CONTEXT HYGIENE", "კონტექსტის ჰიგიენა", "ხმაურის მინიმიზაცია", [
            ("Selective Passing", "შემდეგ ნაბიჯს გადაეცემა მხოლოდ მისი მუშაობისთვის კრიტიკული ველები და არა მთელი ისტორია."),
            ("ტოკენების ეკონომია", "კონტექსტის არასაჭირო გაბერვის თავიდან აცილება 70%-ით ამცირებს API გადასახადს."),
            ("ყურადღების კონცენტრაცია", "მოდელი ფოკუსირებულია მხოლოდ სუფთა, რელევანტურ მონაცემებზე ყოველგვარი ყურადღების გაფანტვის გარეშე.")
        ])
    ]

    for i, (badge, title, subtitle, items) in enumerate(state_s5):
        bg_col = COLOR_CARD_RED if i == 0 else COLOR_BG_LIGHT
        bor_col = COLOR_BORDER_RED if i == 0 else COLOR_BORDER
        c = s5.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, xs_3[i], card_y, card_w, card_h)
        c.fill.solid(); c.fill.fore_color.rgb = bg_col
        c.line.color.rgb = bor_col; c.line.width = Pt(1.5)

        tb = s5.shapes.add_textbox(xs_3[i] + Inches(0.35), card_y + Inches(0.35), card_w - Inches(0.7), card_h - Inches(0.7))
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
    bb = s5.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(1.0), bot_y, Inches(17.0), Inches(2.0))
    bb.fill.solid(); bb.fill.fore_color.rgb = COLOR_WHITE
    bb.line.color.rgb = COLOR_BORDER; bb.line.width = Pt(1.2)

    tbb = s5.shapes.add_textbox(Inches(1.4), bot_y + Inches(0.3), Inches(16.2), Inches(1.4))
    tfb = tbb.text_frame; tfb.word_wrap = True; tfb.margin_left = tfb.margin_top = 0
    p = tfb.paragraphs[0]; p.text = "🎯 საინჟინრო პრინციპი:"
    p.font.name = FONT_ARIAL; p.font.size = Pt(14); p.font.bold = True; p.font.color.rgb = COLOR_RED
    p = tfb.add_paragraph()
    p.text = "ჯაჭვში მონაცემთა მოძრაობა უნდა იმართებოდეს პროგრამული კოდით და არა მოდელის მეხსიერებით; ტიპიზებული კონტრაქტები (Pydantic) და კონტექსტის ჰიგიენა უზრუნველყოფს სისტემის პროგნოზირებადობას."
    p.font.name = FONT_ARIAL; p.font.size = Pt(13); p.font.color.rgb = COLOR_MUTED; p.space_before = Pt(4)

    # -------------------------------------------------------------------------
    # SLIDE 6: Error Handling, Retries & Circuit Breakers
    # -------------------------------------------------------------------------
    s6 = created_slides[4]
    add_header(s6, "შეცდომების მართვა და Fallback მექანიზმები: რა ვქნათ რგოლის გაწყვეტისას?", "საიმედოობა • ERROR HANDLING & CIRCUIT BREAKERS")

    err_s6 = [
        ("🔄 STEP-LEVEL RETRY", "მიზნობრივი Retry", "შეცდომის გამოსწორება", [
            ("ლოკალური აღდგენა", "ჩავარდნისას ხელახლა ეშვება მხოლოდ კონკრეტული ნაბიჯი და არა მთელი ჯაჭვი თავიდან!"),
            ("Error Feedback Injection", "მოდელს ეგზავნება ვალიდაციის ერორი: 'შენი პასუხი არ შეიცავს ველს price; გამოასწორე!'."),
            ("მაღალი წარმატება", "მოდელი შემთხვევათა 90%-ში წარმატებით ასწორებს შეცდომას მეორე განმეორებით ცდაზე.")
        ]),
        ("🛑 MAX RETRIES & CIRCUIT BREAKER", "ლიმიტები და გაჩერება", "უსასრულო ციკლების აღკვეთა", [
            ("მკაცრი ლიმიტი", "მაქსიმუმ 2 განმეორებითი ცდა (Retry Limit = 2); დაუსრულებელი ლუპების კატეგორიული აკრძალვა."),
            ("Circuit Breaker", "თუ API პასუხობს 500 ერორით ან მოდელი მუდმივად ცდება, ჯაჭვი მყისიერად წყდება."),
            ("ფინანსური დაცვა", "იცავს კომპანიას გაუთვალისწინებელი ფინანსური გადახარჯვისა და ტოკენების უაზრო დაწვისგან.")
        ]),
        ("🛡️ GRACEFUL FALLBACK", "გრაციოზული დეგრადაცია", "უსაფრთხო გასასვლელი", [
            ("Fallback შაბლონი", "კრიტიკული ხარვეზისას სისტემა გადადის გამარტივებულ, წინასწარ გამზადებულ უსაფრთხო პასუხზე."),
            ("Human Escalation", "ჩავარდნილი მოთხოვნის ავტომატური გადაგზავნა ოპერატორთან სრული დიაგნოსტიკური ლოგით."),
            ("მომხმარებლის ინფორმირება", "კლიენტი იღებს ზრდილობიან, გასაგებ შეტყობინებას და არა გაურკვეველ სისტემურ ერორს.")
        ])
    ]

    for i, (badge, title, subtitle, items) in enumerate(err_s6):
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
    p = tfb.paragraphs[0]; p.text = "💡 მთავარი არსი:"
    p.font.name = FONT_ARIAL; p.font.size = Pt(14); p.font.bold = True; p.font.color.rgb = COLOR_RED
    p = tfb.add_paragraph()
    p.text = "ჯაჭვის სიმტკიცე განისაზღვრება მისი ყველაზე სუსტი რგოლით; საწარმოო Chaining მოითხოვს მიზნობრივ Retry მექანიზმს შეცდომის უკუკავშირით და ავარიულ Circuit Breaker-ს ხარჯების დასაცავად."
    p.font.name = FONT_ARIAL; p.font.size = Pt(13); p.font.color.rgb = COLOR_MUTED; p.space_before = Pt(4)

    # -------------------------------------------------------------------------
    # SLIDE 7: Monolithic Prompt vs Prompt Chaining (Table)
    # -------------------------------------------------------------------------
    s7 = created_slides[5]
    add_header(s7, "Monolithic Prompt vs Prompt Chaining: არქიტექტურული შედარება", "ᲨᲔᲓᲐᲠᲔᲑᲘᲗᲘ ᲐᲜᲐᲚᲘᲖᲘ • MONOLITH VS CHAIN")

    # Table: 6 rows, 5 columns
    tbl_left = Inches(1.0); tbl_top = Inches(2.35); tbl_w = Inches(17.0); tbl_h = Inches(5.2)
    tbl_shape = s7.shapes.add_table(6, 5, tbl_left, tbl_top, tbl_w, tbl_h)
    table = tbl_shape.table

    col_widths = [Inches(3.2), Inches(3.4), Inches(3.4), Inches(3.5), Inches(3.5)]
    for ci, w in enumerate(col_widths):
        table.columns[ci].width = w

    headers = ["კრიტერიუმი / პარამეტრი", "Monolithic Prompt (მონოლითი)", "Prompt Chaining (ჯაჭვი)", "რისკი / ნეგატიური გავლენა", "საინჟინრო რეკომენდაცია"]
    for ci, h in enumerate(headers):
        cell = table.cell(0, ci)
        cell.fill.solid(); cell.fill.fore_color.rgb = COLOR_DARK
        tf = cell.text_frame; tf.word_wrap = True; tf.margin_left = tf.margin_right = Inches(0.12)
        p = tf.paragraphs[0]; p.text = h; p.font.name = FONT_ARIAL; p.font.size = Pt(12); p.font.bold = True; p.font.color.rgb = COLOR_WHITE

    chain_rows = [
        ("ფაქტობრივი სიზუსტე", "65% – 75% (ხშირი ჰალუცინაციები)", "95% – 98% (იზოლირებული ექსტრაქცია)", "მაღალი ბიზნეს-რისკი მონოლითში", "გამოიყენეთ Chaining რთულ ამოცანებზე"),
        ("მოდულური ტესტირებადობა", "შეუძლებელია ცალკეული ლოგიკის გატესტვა", "თითოეულ ნაბიჯს აქვს საკუთარი Unit ტესტი", "გაურკვეველი რეგრესიები მონოლითში", "Unit ტესტები და Golden Set ყოველ ეტაპზე"),
        ("შეცდომების დიაგნოსტიკა", "Debugging Nightmare — გაურკვეველია მიზეზი", "მყისიერი ლოკალიზაცია (ჩავარდა Step 2)", "დროისა და დეველოპერული რესურსის კარგვა", "დეტალური Step-level Tracing მონიტორინგი"),
        ("ტოკენების ხარჯი", "ერთი შეხედვით ნაკლები, თუმცა ბევრი Retry", "მეტი გამოძახება, მაგრამ იდეალური Caching", "ხარჯების მოულოდნელი ზრდა", "Prompt Caching ნაბიჯების სისტემურ პრეფიქსზე"),
        ("პარალელური დამუშავება", "მკაცრად თანმიმდევრული (Sequential)", "შესაძლებელია დატოტვა (Parallel Branching)", "სისტემის მაღალი დაყოვნება მონოლითში", "პარალელური ნაბიჯები დამოუკიდებელ ველებზე")
    ]

    for ri, row_data in enumerate(chain_rows):
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
    p = tfb.paragraphs[0]; p.text = "💡 შედარების მთავარი დასკვნა:"
    p.font.name = FONT_ARIAL; p.font.size = Pt(14); p.font.bold = True; p.font.color.rgb = COLOR_RED
    p = tfb.add_paragraph()
    p.text = "მონოლითური პრომპტი მისაღებია მხოლოდ უმარტივეს 1-საფეხურიან მოთხოვნებზე; ნებისმიერი საწარმოო ბიზნეს-პროცესი მოითხოვს Prompt Chaining-ს ხარისხის, ტესტირებადობისა და დიაგნოსტიკის უზრუნველსაყოფად."
    p.font.name = FONT_ARIAL; p.font.size = Pt(13); p.font.color.rgb = COLOR_MUTED; p.space_before = Pt(4)

    # -------------------------------------------------------------------------
    # SLIDE 8: Modern Orchestration & Observability 2026
    # -------------------------------------------------------------------------
    s8 = created_slides[6]
    add_header(s8, "თანამედროვე ტექნოლოგიური დასტა: Chaining ძრავები და Observability", "ᲢᲔᲥᲜᲝᲚᲝᲒᲘᲣᲠᲘ ᲡᲢᲔᲙᲘ • ORCHESTRATION 2026")

    stack_s8 = [
        ("💻 CODE-FIRST ENGINES", "კოდზე დაფუძნებული ძრავები", "Python & TypeScript", [
            ("LangChain LCEL", "დეკლარაციული მილსადენები (prompt | model | parser) პარალელიზაციითა და Fallback-ების მხარდაჭერით."),
            ("LlamaIndex Workflows", "მოვლენებზე დაფუძნებული (Event-Driven) მილსადენები მონაცემთა ექსტრაქციისა და RAG სისტემებისთვის."),
            ("Native Async Python", "უმარტივესი და ყველაზე კონტროლირებადი მიდგომა გარე მძიმე ბიბლიოთეკების გარეშე (asyncio & Pydantic).")
        ]),
        ("☁️ STATE ENGINES & GATEWAYS", "ღრუბლოვანი ძრავები", "საწარმოო საიმედოობა", [
            ("Google AI Studio Chaining", "მრავალეტაპიანი პრომპტების ვიზუალური აწყობა და Vertex AI მილსადენებთან ინტეგრაცია."),
            ("Temporal & Inngest", "საიმედო ორკესტრაცია ხანგრძლივი პროცესებისთვის (Durable Execution, ავტომატური Retries)."),
            ("Cloud Microservices", "AWS Step Functions / Cloud Run — მიკროსერვისული არქიტექტურა მასშტაბური საწარმოო ჯაჭვებისთვის.")
        ]),
        ("📊 OBSERVABILITY & TRACING", "მონიტორინგი და ტრასირება", "მილსადენის ხედვადობა", [
            ("Langfuse", "თითოეული ნაბიჯის შემავალი/გამომავალი მონაცემების, დაყოვნებისა და ხარჯის დეტალური ვიზუალიზაცია რეალურ დროში."),
            ("Arize Phoenix", "ღია კოდის Tracing სისტემა ჯაჭვში ხარვეზებისა და ხარისხის დეგრადაციის მყისიერად აღმოსაჩენად."),
            ("OpenTelemetry", "ინდუსტრიული სტანდარტის ტელემეტრია, რომელიც აერთიანებს AI ლოგებს საწარმოო მონიტორინგთან.")
        ])
    ]

    for i, (badge, title, subtitle, items) in enumerate(stack_s8):
        bg_col = COLOR_CARD_RED if i == 2 else COLOR_BG_LIGHT
        bor_col = COLOR_BORDER_RED if i == 2 else COLOR_BORDER
        c = s8.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, xs_3[i], card_y, card_w, card_h)
        c.fill.solid(); c.fill.fore_color.rgb = bg_col
        c.line.color.rgb = bor_col; c.line.width = Pt(1.5)

        tb = s8.shapes.add_textbox(xs_3[i] + Inches(0.35), card_y + Inches(0.35), card_w - Inches(0.7), card_h - Inches(0.7))
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
    bb = s8.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(1.0), bot_y, Inches(17.0), Inches(2.0))
    bb.fill.solid(); bb.fill.fore_color.rgb = COLOR_WHITE
    bb.line.color.rgb = COLOR_BORDER; bb.line.width = Pt(1.2)

    tbb = s8.shapes.add_textbox(Inches(1.4), bot_y + Inches(0.3), Inches(16.2), Inches(1.4))
    tfb = tbb.text_frame; tfb.word_wrap = True; tfb.margin_left = tfb.margin_top = 0
    p = tfb.paragraphs[0]; p.text = "🎯 საინჟინრო პრინციპი:"
    p.font.name = FONT_ARIAL; p.font.size = Pt(14); p.font.bold = True; p.font.color.rgb = COLOR_RED
    p = tfb.add_paragraph()
    p.text = "2026 წელს ჯაჭვების მართვა მოითხოვს სრულ Observability-ს (Langfuse, Phoenix); თქვენ რეალურ დროში უნდა ხედავდეთ თითოეული ნაბიჯის ტოკენების ხარჯს, ლატენტობასა და მონაცემთა ნაკადს."
    p.font.name = FONT_ARIAL; p.font.size = Pt(13); p.font.color.rgb = COLOR_MUTED; p.space_before = Pt(4)

    # -------------------------------------------------------------------------
    # SLIDE 9: Best Practices vs Anti-Patterns & Transition to Topic 8.1
    # -------------------------------------------------------------------------
    s9 = created_slides[7]
    add_header(s9, "საინჟინრო პატერნები და თავიდან ასაცილებელი შეცდომები Chaining-ში", "ᲡᲐᲣᲙᲔᲗᲔᲡᲝ ᲞᲠᲐᲥᲢᲘᲙᲐ • PATTERNS & ANTI-PATTERNS")

    box_w = Inches(8.35); box_h = Inches(5.4)

    # Anti-patterns Box
    c_anti = s9.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(1.0), card_y, box_w, box_h)
    c_anti.fill.solid(); c_anti.fill.fore_color.rgb = COLOR_CARD_RED
    c_anti.line.color.rgb = COLOR_BORDER_RED; c_anti.line.width = Pt(1.5)

    tb_anti = s9.shapes.add_textbox(Inches(1.35), card_y + Inches(0.35), box_w - Inches(0.7), box_h - Inches(0.7))
    tf_anti = tb_anti.text_frame; tf_anti.word_wrap = True; tf_anti.margin_left = tf_anti.margin_top = 0
    p = tf_anti.paragraphs[0]; p.text = "❌ CRITICAL ANTI-PATTERNS (კრიტიკული შეცდომები)"
    p.font.name = FONT_ARIAL; p.font.size = Pt(13); p.font.bold = True; p.font.color.rgb = COLOR_RED

    p = tf_anti.add_paragraph(); p.text = "რას უნდა მოვერიდოთ Chaining-ში"
    p.font.name = FONT_ARIAL; p.font.size = Pt(16); p.font.bold = True; p.font.color.rgb = COLOR_DARK; p.space_before = Pt(4)

    anti_items = [
        ("Fragile Unvalidated Chains", "მონაცემთა გადაცემა ნაბიჯებს შორის ვალიდაციის გარეშე (ერთი შეცდომა მთელ ჯაჭვს კრაშავს)."),
        ("Over-Chaining", "მარტივი ამოცანის დაშლა 10 უსარგებლო ნაბიჯად, რაც 10-ჯერ ზრდის სისტემის დაყოვნებასა და ხარჯებს."),
        ("Stateful Accumulation", "ყოველ ახალ ნაბიჯში წინა ყველა ნაბიჯის ნედლი ტექსტის დაგროვება და კონტექსტის დაბინძურება."),
        ("Missing Fallbacks", "ავარიული გეგმის არარსებობა API შეფერხების ან მოდელის განმეორებითი ჩავარდნის დროს.")
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

    p = tf_rec.add_paragraph(); p.text = "სტაბილური და დაცული Chaining არქიტექტურა"
    p.font.name = FONT_ARIAL; p.font.size = Pt(16); p.font.bold = True; p.font.color.rgb = COLOR_DARK; p.space_before = Pt(4)

    rec_items = [
        ("Strict Schema Validation", "ყოველი ნაბიჯის გამოსავალი მოწმდება Pydantic სქემით მომდევნო ეტაპზე გადაცემამდე."),
        ("Lean Step Context", "შემდეგ ნაბიჯს მიეწოდება მხოლოდ საჭირო ველები, რაც ინარჩუნებს კონტექსტის სისუფთავეს."),
        ("Parallel Branching", "დამოუკიდებელი ამოცანების პარალელური გაშვება (asyncio.gather), რაც რადიკალურად ამცირებს დაყოვნებას."),
        ("Circuit Breakers & Alerts", "კრიტიკული შეცდომისას ჯაჭვის მყისიერი გაჩერება და დეველოპერთა გუნდის ავტომატური გაფრთხილება.")
    ]
    for head, body in rec_items:
        p = tf_rec.add_paragraph(); p.text = f"• {head}: {body}"
        p.font.name = FONT_ARIAL; p.font.size = Pt(11); p.font.color.rgb = COLOR_DARK; p.space_before = Pt(6)

    # Bottom takeaway & transition to Lecture 8.1
    bb = s9.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(1.0), bot_y, Inches(17.0), Inches(2.0))
    bb.fill.solid(); bb.fill.fore_color.rgb = COLOR_WHITE
    bb.line.color.rgb = COLOR_BORDER; bb.line.width = Pt(1.2)

    tbb = s9.shapes.add_textbox(Inches(1.4), bot_y + Inches(0.25), Inches(16.2), Inches(1.5))
    tfb = tbb.text_frame; tfb.word_wrap = True; tfb.margin_left = tfb.margin_top = 0
    p = tfb.paragraphs[0]; p.text = "💡 საინჟინრო შეჯამება და გადასვლა ლექციაზე 8.1:"
    p.font.name = FONT_ARIAL; p.font.size = Pt(14); p.font.bold = True; p.font.color.rgb = COLOR_RED
    p = tfb.add_paragraph()
    p.text = "Prompt Chaining აქცევს AI-ს საიმედო საინჟინრო კომპონენტად. შემდეგ ლექციაში (8.1: მოწინავე მსჯელობის მეთოდები — Self-Consistency, ToT, CoVe და ტოკენების ეკონომიკა) შევისწავლით: როგორ ავაწყოთ მრავალტოტიანი აზროვნების ხეები (Tree of Thoughts) და როდის ღირს მათი გამოყენება საწარმოო ამოცანებში."
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
    filename = "7.2 - რთული ამოცანების დეკომპოზიცია და Prompt Chaining არქიტექტურა.pptx"
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
    build_lecture_m7_7_2()
