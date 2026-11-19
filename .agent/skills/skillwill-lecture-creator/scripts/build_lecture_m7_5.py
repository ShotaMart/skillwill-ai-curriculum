import os
import sys
import shutil
import pptx
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN
from pptx.enum.shapes import MSO_SHAPE

def build_lecture_m7_5():
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
            p.text = "5 - Few-Shot Prompting: მაგალითების შერჩევა და შედეგის სტაბილურობა"
            p.alignment = PP_ALIGN.CENTER
            p.font.name = FONT_ARIAL
            p.font.size = Pt(32)
            p.font.bold = True
            p.font.color.rgb = COLOR_RED

            p2 = tf.add_paragraph()
            p2.text = "In-Context Learning მექანიზმი, მაგალითების რაოდენობა (1, 3, 5), ფორმატის ერთგვაროვნება და Label Bias-ის პრევენცია"
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
    def add_header(slide, title, category="ᲛᲝᲓᲣᲚᲘ 7 • PROMPT ENGINEERING • 5"):
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
    # SLIDE 2: In-Context Learning (ICL) & Few-Shot Mental Model
    # -------------------------------------------------------------------------
    s2 = created_slides[0]
    add_header(s2, "Few-Shot მენტალური მოდელი: როგორ „სწავლობს“ LLM კონტექსტში?", "ᲛᲔᲜᲢᲐᲚᲣᲠᲘ ᲛᲝᲓᲔᲚᲘ • IN-CONTEXT LEARNING")

    cards_s2 = [
        ("🧠 IN-CONTEXT LEARNING (ICL)", "სწავლება კონტექსტში", "წონების განახლების გარეშე", [
            ("ყურადღების მექანიზმი (Self-Attention)", "მოდელი Attention შრეებით ავლენს სემანტიკურ და სტრუქტურულ კანონზომიერებებს მოცემულ მაგალითებს შორის."),
            ("დროებითი ადაპტაცია", "არ ხდება მოდელის შინაგანი წონების (Weights) შეცვლა; ცოდნა აქტიურდება მხოლოდ მიმდინარე კონტექსტის ფანჯარაში."),
            ("ანალოგიური მსჯელობა", "მოდელი ახდენს მაგალითებში ნაჩვენები პატერნის ზუსტ ექსტრაპოლაციას ახალ, შეუცნობელ შემავალ მონაცემზე.")
        ]),
        ("🎯 TASK CONDITIONING", "დავალებაზე კონდიცირება", "ოპერაციული რეჟიმის ჩაკეტვა", [
            ("ალბათური სივრცის შევიწროება", "მაგალითები რადიკალურად ამცირებს გაურკვევლობას (Entropy) და მოდელს გადაჰყავს ზუსტ სამუშაო რეჟიმში."),
            ("ორაზროვნების გაფანტვა", "როდესაც წესები რამდენიმენაირად შეიძლება გავიგოთ, მაგალითი აჩვენებს ერთადერთ სწორ საინჟინრო ინტერპრეტაციას."),
            ("პრიორიტეტების დემონსტრაცია", "მოდელი მკაფიოდ ხედავს, რა დეტალს მიაქციოს განსაკუთრებული ყურადღება და რა უგულებელყოს.")
        ]),
        ("📐 IMPLICIT STYLISTICS", "იმპლიციტური სტილი", "ფორმატი ახსნის გარეშე", [
            ("სიგრძისა და ტონის კოპირება", "მოდელი ინტუიციურად იმეორებს მაგალითის სიგრძეს (მაგ. მკაცრად 1 წინადადება) და ფორმალურობის ხარისხს."),
            ("სტრუქტურული მიმიკრია", "JSON სქემის გასაღებები, პუნქტუაცია და ტეგები ზუსტად მეორდება ყოველგვარი დამატებითი წესის გარეშეც."),
            ("ჟარგონისა და ენის ათვისება", "კონკრეტული დომენის ლექსიკა და ტერმინოლოგიური ნიუანსები ბუნებრივად გადადის გენერირებულ პასუხში.")
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
    p = tfb.paragraphs[0]; p.text = "💡 მთავარი არსი (საინჟინრო ანალოგია):"
    p.font.name = FONT_ARIAL; p.font.size = Pt(14); p.font.bold = True; p.font.color.rgb = COLOR_RED
    p = tfb.add_paragraph()
    p.text = "ერთი მაგალითი ათას ინსტრუქციად ღირს; In-Context Learning მოდელს აძლევს 'სამუშაო შაბლონს', რომლის მიხედვითაც ის ზუსტად ახდენს სასურველი ლოგიკისა და ფორმატის რეპროდუცირებას წესების გადატვირთვის გარეშე."
    p.font.name = FONT_ARIAL; p.font.size = Pt(13); p.font.color.rgb = COLOR_MUTED; p.space_before = Pt(4)

    # -------------------------------------------------------------------------
    # SLIDE 3: Shot Count Economics (1, 3 vs 5 Shots)
    # -------------------------------------------------------------------------
    s3 = created_slides[1]
    add_header(s3, "მაგალითების რაოდენობის მათემატიკა და ეკონომიკა: 1, 3 თუ 5 Shot?", "ᲡᲐᲘᲜᲟᲘᲜᲠᲝ ᲐᲠᲩᲔᲕᲐᲜᲘ • SHOT COUNT ECONOMICS")

    shots_s3 = [
        ("1️⃣ ONE-SHOT (1 მაგალითი)", "სტრუქტურული წამყვანი", "მინიმალური ტოკენების ხარჯი", [
            ("დანიშნულება", "გამოსავლის სასურველი ფორმატის (JSON/Markdown) დაფიქსირება და საბაზისო სტილის ჩვენება."),
            ("უპირატესობები", "მინიმალური Token Overhead და დაბალი დაყოვნება (Latency); იდეალურია მარტივი ტრანსფორმაციებისთვის."),
            ("საფრთხე & რისკი", "მაღალი მგრძნობელობა ერთ მაგალითზე — თუ მაგალითი უნიკალურია, მოდელი მას ზედმეტად მიემსგავსება.")
        ]),
        ("3️⃣ THREE-SHOT (3 მაგალითი)", "ინდუსტრიული Sweet Spot", "ოპტიმალური ბალანსი", [
            ("სანიმუშო არქიტექტურა", "2 ტიპური საბაზისო ქეისი + 1 ალტერნატიული ან სასაზღვრო ვარიანტი სხვადასხვა ლეიბლით."),
            ("უპირატესობები", "უზრუნველყოფს სტაბილურ სიზუსტეს (Accuracy) და გამორიცხავს ერთ კონკრეტულ ქეისზე გადახრას."),
            ("საწარმოო სტანდარტი", "2026 წლის პროდაქშენ სერვისების 70%-ისთვის წარმოადგენს ხარისხისა და ხარჯის საუკეთესო ოქროს შუალედს.")
        ]),
        ("5️⃣ FIVE-SHOT (5 მაგალითი)", "მაქსიმალური სტაბილურობა", "რთული კლასიფიკაცია", [
            ("დანიშნულება", "მრავალკლასიანი კლასიფიკაცია, რთული დომენური წესები და ზღვრული ქეისების დაფარვა."),
            ("უპირატესობები", "თითქმის 100%-იანი იმუნიტეტი ფორმატის დარღვევაზე და მინიმალური დისპერსია (Variance) პასუხებში."),
            ("საინჟინრო მოთხოვნა", "აუცილებელია Prompt Caching-ის გამოყენება, რათა 800+ ტოკენიანი მაგალითების ბაზა 90%-ით გაიაფდეს.")
        ])
    ]

    for i, (badge, title, subtitle, items) in enumerate(shots_s3):
        bg_col = COLOR_CARD_RED if i == 1 else COLOR_BG_LIGHT
        bor_col = COLOR_BORDER_RED if i == 1 else COLOR_BORDER
        c = s3.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, xs_3[i], card_y, card_w, card_h)
        c.fill.solid(); c.fill.fore_color.rgb = bg_col
        c.line.color.rgb = bor_col; c.line.width = Pt(1.5)

        tb = s3.shapes.add_textbox(xs_3[i] + Inches(0.35), card_y + Inches(0.35), card_w - Inches(0.7), card_h - Inches(0.7))
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
    bb = s3.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(1.0), bot_y, Inches(17.0), Inches(2.0))
    bb.fill.solid(); bb.fill.fore_color.rgb = COLOR_WHITE
    bb.line.color.rgb = COLOR_BORDER; bb.line.width = Pt(1.2)

    tbb = s3.shapes.add_textbox(Inches(1.4), bot_y + Inches(0.3), Inches(16.2), Inches(1.4))
    tfb = tbb.text_frame; tfb.word_wrap = True; tfb.margin_left = tfb.margin_top = 0
    p = tfb.paragraphs[0]; p.text = "🎯 საინჟინრო პრინციპი:"
    p.font.name = FONT_ARIAL; p.font.size = Pt(14); p.font.bold = True; p.font.color.rgb = COLOR_RED
    p = tfb.add_paragraph()
    p.text = "დაიწყეთ 1-shot-ით ფორმატის დასაფიქსირებლად; თუ მოდელი ცდება კლასიფიკაციაში, გადადით 3-shot-ზე; 5-shot გამოიყენეთ მხოლოდ რთულ, მრავალგანზომილებიან ამოცანებში და აუცილებლად ჩართეთ Prompt Caching."
    p.font.name = FONT_ARIAL; p.font.size = Pt(13); p.font.color.rgb = COLOR_MUTED; p.space_before = Pt(4)

    # -------------------------------------------------------------------------
    # SLIDE 4: Selection Strategy & Structural Isomorphism
    # -------------------------------------------------------------------------
    s4 = created_slides[2]
    add_header(s4, "მაგალითების შერჩევის სტრატეგია: ცუდი მაგალითი უარესია, ვიდრე არცერთი!", "ᲛᲝᲜᲐᲪᲔᲛᲗᲐ ᲮᲐᲠᲘᲡᲮᲘ • SELECTION STRATEGY & ISOMORPHISM")

    strat_s4 = [
        ("📐 STRUCTURAL ISOMORPHISM", "სტრუქტურული იზომორფიზმი", "ფორმატის სრული ერთგვაროვნება", [
            ("იდენტური ველები", "ყველა მაგალითში JSON გასაღებების სახელები, მონაცემთა ტიპები და თანმიმდევრობა უნდა ემთხვეოდეს ერთმანეთს."),
            ("სინტაქსური ერთგვაროვნება", "თუ ერთ მაგალითში JSON ფორმატია, მეორე მაგალითში არავითარ შემთხვევაში არ უნდა იყოს უბრალო ტექსტი."),
            ("XML ტეგების სტანდარტი", "თითოეული მაგალითის შეფუთვა მკაფიო კონტეინერში: <example id='1'><input>...</input><output>...</output></example>.")
        ]),
        ("🌈 DOMAIN DIVERSITY", "დომენური მრავალფეროვნება", "განზოგადების გაზრდა", [
            ("სცენარების ვარიაცია", "არ გამოიყენოთ 3 იდენტური მოკლე წინადადება; აჩვენეთ მოკლე, საშუალო და რთული სტრუქტურის შემავალი ტექსტები."),
            ("ლექსიკური სიმდიდრე", "გამოიყენეთ განსხვავებული სინონიმები და წინადადების განსხვავებული წყობა სხვადასხვა მაგალითში."),
            ("Overfitting-ის პრევენცია", "მრავალფეროვნება იცავს მოდელს კონკრეტულ სიტყვებზე ჩაციკვლისა და მცდარი ასოციაციების შექმნისგან.")
        ]),
        ("💎 GROUND TRUTH PURITY", "ეტალონური სიზუსტე", "ნულოვანი ტოლერანტობა შეცდომებზე", [
            ("ექსპერტული ვერიფიკაცია", "მაგალითების ლეიბლები და პასუხები გადამოწმებული უნდა იყოს ადამიანური დომენის ექსპერტის მიერ."),
            ("ხარვეზის გამრავლება", "თუ მაგალითში გაიპარა თუნდაც ერთი უზუსტობა, მოდელი მას განაზოგადებს და შეცდომების კასკადს შექმნის."),
            ("ცოდნის რეგულარული აუდიტი", "რეგულარულად გადაამოწმეთ, ხომ არ მოძველდა მაგალითებში მოყვანილი ფაქტები კომპანიის პოლიტიკის ცვლილების შემდეგ.")
        ])
    ]

    for i, (badge, title, subtitle, items) in enumerate(strat_s4):
        bg_col = COLOR_CARD_RED if i == 0 else COLOR_BG_LIGHT
        bor_col = COLOR_BORDER_RED if i == 0 else COLOR_BORDER
        c = s4.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, xs_3[i], card_y, card_w, card_h)
        c.fill.solid(); c.fill.fore_color.rgb = bg_col
        c.line.color.rgb = bor_col; c.line.width = Pt(1.5)

        tb = s4.shapes.add_textbox(xs_3[i] + Inches(0.35), card_y + Inches(0.35), card_w - Inches(0.7), card_h - Inches(0.7))
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
    bb = s4.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(1.0), bot_y, Inches(17.0), Inches(2.0))
    bb.fill.solid(); bb.fill.fore_color.rgb = COLOR_WHITE
    bb.line.color.rgb = COLOR_BORDER; bb.line.width = Pt(1.2)

    tbb = s4.shapes.add_textbox(Inches(1.4), bot_y + Inches(0.3), Inches(16.2), Inches(1.4))
    tfb = tbb.text_frame; tfb.word_wrap = True; tfb.margin_left = tfb.margin_top = 0
    p = tfb.paragraphs[0]; p.text = "💡 მთავარი არსი:"
    p.font.name = FONT_ARIAL; p.font.size = Pt(14); p.font.bold = True; p.font.color.rgb = COLOR_RED
    p = tfb.add_paragraph()
    p.text = "LLM-ები მაგალითებს აღიქვამენ როგორც აბსოლუტურ ჭეშმარიტებას (Ground Truth); ერთმა უზუსტო ან განსხვავებულად ფორმატირებულმა მაგალითმაც კი შეიძლება მთელი სისტემის საიმედოობა გაანახევროს."
    p.font.name = FONT_ARIAL; p.font.size = Pt(13); p.font.color.rgb = COLOR_MUTED; p.space_before = Pt(4)

    # -------------------------------------------------------------------------
    # SLIDE 5: Hidden Biases (Majority Label Bias, Recency Bias & Distribution)
    # -------------------------------------------------------------------------
    s5 = created_slides[3]
    add_header(s5, "ფარული ანომალიები Few-Shot-ში: Majority Label Bias და Recency Bias", "ᲐᲜᲝᲛᲐᲚᲘᲔᲑᲘᲡ ᲞᲠᲔᲕᲔᲜᲪᲘᲐ • BIAS & DISTRIBUTION DRIFT")

    bias_s5 = [
        ("⚖️ MAJORITY LABEL BIAS", "უმრავლესობის გადახრა", "ალბათური დისბალანსი", [
            ("პრობლემის არსი", "თუ 4 მაგალითიდან 3 არის 'დადებითი' და მხოლოდ 1 'უარყოფითი', მოდელი ახალ გაურკვეველ ქეისსაც ინერციით 'დადებითად' მონიშნავს."),
            ("ალბათური აღქმა", "მოდელი კონტექსტში არსებულ მაგალითების სიხშირეს აღიქვამს როგორც სამყაროში მოვლენის საბაზისო ალბათობას (Prior Probability)."),
            ("საინჟინრო გადაჭრა", "Balanced Few-Shot: ყველა შესაძლო საკლასიფიკაციო კლასს უნდა ჰქონდეს ზუსტად თანაბარი რაოდენობის წარმომადგენლობა.")
        ]),
        ("⏳ RECENCY BIAS", "ბოლო ელემენტის ეფექტი", "Attention-ის გადახრა", [
            ("პრობლემის არსი", "LLM-ები არაპროპორციულად დიდ ყურადღებას აქცევენ ბოლო, უშუალოდ მომხმარებლის კითხვასთან ყველაზე ახლოს მდებარე მაგალითს."),
            ("გავლენა გადაწყვეტილებაზე", "თუ ბოლო მაგალითის კლასია 'ტექნიკური შეცდომა', მოდელი მიდრეკილია ახალი შეკითხვაც ამ კატეგორიას მიაკუთვნოს."),
            ("საინჟინრო გადაჭრა", "მაგალითების თანმიმდევრობის შემთხვევითი შერევა (Shuffling) და ნეიტრალური გამყოფი ტეგების გამოყენება.")
        ]),
        ("📊 COMMON TOKEN BIAS", "გავრცელებული სიტყვების ეფექტი", "Pre-training ინერცია", [
            ("პრობლემის არსი", "მოდელს ქვეცნობიერად ურჩევნია გამოიყენოს ის ლეიბლები, რომლებიც უფრო ხშირად გვხვდებოდა Pre-training-ის მონაცემებში."),
            ("მაგალითი", "მოდელი უფრო ადვილად პროგნოზირებს 'True/False'-ს, ვიდრე კონკრეტულ შიდა სტატუსებს ('Passed/Rejected')."),
            ("საინჟინრო გადაჭრა", "მკაცრი Enum სქემის დაფიქსირება და თითოეული ლეიბლისთვის კონკრეტული სანიმუშო კონტექსტის მიბმა.")
        ])
    ]

    for i, (badge, title, subtitle, items) in enumerate(bias_s5):
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
    p.text = "Few-Shot მაგალითების კოლექცია უნდა იყოს მკაცრად დაბალანსებული: თანაბარი რაოდენობის კლასები, მოწესრიგებული თანმიმდევრობა და ნეიტრალური ტეგებით იზოლაცია გამორიცხავს მოდელის სტატისტიკურ გადახრას."
    p.font.name = FONT_ARIAL; p.font.size = Pt(13); p.font.color.rgb = COLOR_MUTED; p.space_before = Pt(4)

    # -------------------------------------------------------------------------
    # SLIDE 6: Edge Cases & Negative Examples
    # -------------------------------------------------------------------------
    s6 = created_slides[4]
    add_header(s6, "ზღვრული ქეისები და უარყოფითი მაგალითები: როგორ ვასწავლოთ საზღვრები?", "ᲡᲐᲖᲦᲕᲠᲔᲑᲘᲡ ᲓᲐᲪᲕᲐ • EDGE CASES & NEGATIVE EXAMPLES")

    edge_s6 = [
        ("🧗 BORDERLINE / EDGE CASES", "სასაზღვრო სიტუაციები", "კრიტიკული გადაწყვეტილებები", [
            ("კომპლექსური ქეისები", "სიტუაციები, სადაც მომხმარებლის ტექსტი ერთდროულად ორ კატეგორიას ეხება (მაგ. ფინანსური პრეტენზია ტექნიკური შეცდომის ფონზე)."),
            ("პრიორიტეტის ჩვენება", "მაგალითი თვალნათლივ უჩვენებს მოდელს, რომელი წესი იმარჯვებს ინტერესთა კონფლიქტის დროს."),
            ("Reasoning ველი", "რთულ მაგალითებში მოკლე explanation ველის დამატება მოდელს ასწავლის გადაწყვეტილების მიღების ლოგიკას.")
        ]),
        ("🚫 CONTRASTIVE EXAMPLES", "კონტრასტული მაგალითები", "Bad vs Good დემონსტრაცია", [
            ("შეცდომის პირდაპირი ჩვენება", "Wrong Output vs Correct Output — მოდელი ხედავს ტიპურ შეცდომას და მის საპირისპირო სწორ ვარიანტს."),
            ("ჰალუცინაციების ბლოკირება", "კონტრასტით ვაჩვენებთ, რომ დაუდასტურებელი ფაქტის თვითნებური დამატება მკაცრად დაუშვებელია."),
            ("ტონის კორექცია", "მშრალი, აგრესიული პასუხის შედარება ემპათიურ, პროფესიულ ფორმულირებასთან იმავე კონტექსტში.")
        ]),
        ("🛡️ HARD REFUSAL EXAMPLES", "უარის თქმის მაგალითები", "უსაფრთხოების ესკალაცია", [
            ("მავნე მოთხოვნები", "მაგალითი, სადაც მომხმარებელი ცდილობს სისტემური წესების გაგებას, ხოლო ასისტენტი ზრდილობიანად უარყოფს მას."),
            ("არაკომპეტენტური კითხვები", "ასისტენტი აღიარებს, რომ საკითხი სცდება მის უფლებამოსილებას და ამისამართებს ოპერატორთან."),
            ("Handoff ფორმატი", "ადამიანთან გადაცემის სტრუქტურის (Handoff payload) დემონსტრირება რეალურ მაგალითზე.")
        ])
    ]

    for i, (badge, title, subtitle, items) in enumerate(edge_s6):
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
    p.text = "მხოლოდ პოზიტიური მაგალითების ჩვენება მოდელს აქცევს გულუბრყვილოდ; სწორედ Edge Cases და Refusal Examples უყალიბებს სისტემას მკაცრ იმუნიტეტს შეცდომებისა და ჰაკერული ინექციების მიმართ."
    p.font.name = FONT_ARIAL; p.font.size = Pt(13); p.font.color.rgb = COLOR_MUTED; p.space_before = Pt(4)

    # -------------------------------------------------------------------------
    # SLIDE 7: Experimental Benchmarks: Zero-Shot vs 3-Shot vs 5-Shot (Table)
    # -------------------------------------------------------------------------
    s7 = created_slides[5]
    add_header(s7, "Zero-Shot vs 3-Shot vs 5-Shot კლასიფიკაციაში: ექსპერიმენტული შედარება", "ᲨᲔᲓᲐᲠᲔᲑᲘᲗᲘ ᲐᲜᲐᲚᲘᲖᲘ • EXPERIMENTAL BENCHMARKS")

    # Table: 6 rows, 5 columns
    tbl_left = Inches(1.0); tbl_top = Inches(2.35); tbl_w = Inches(17.0); tbl_h = Inches(5.2)
    tbl_shape = s7.shapes.add_table(6, 5, tbl_left, tbl_top, tbl_w, tbl_h)
    table = tbl_shape.table

    col_widths = [Inches(3.2), Inches(3.2), Inches(3.4), Inches(3.6), Inches(3.6)]
    for ci, w in enumerate(col_widths):
        table.columns[ci].width = w

    headers = ["პარამეტრი / მეტრიკა", "Zero-Shot (ინსტრუქცია)", "3-Shot (ოპტიმალური)", "5-Shot (მაქსიმალური)", "საინჟინრო რეკომენდაცია"]
    for ci, h in enumerate(headers):
        cell = table.cell(0, ci)
        cell.fill.solid(); cell.fill.fore_color.rgb = COLOR_DARK
        tf = cell.text_frame; tf.word_wrap = True; tf.margin_left = tf.margin_right = Inches(0.12)
        p = tf.paragraphs[0]; p.text = h; p.font.name = FONT_ARIAL; p.font.size = Pt(12); p.font.bold = True; p.font.color.rgb = COLOR_WHITE

    bench_rows = [
        ("კლასიფიკაციის სიზუსტე (Accuracy)", "72% – 78% (ზოგად ამოცანებზე)", "88% – 93% (სტაბილურად მაღალი)", "94% – 97% (ზღვრულ ქეისებზეც)", "3-Shot სავსებით საკმარისია ამოცანების 80%-ში"),
        ("ფორმატის დაცვა (JSON Validity)", "85% – 90% (ეპარება ზედმეტი ტექსტი)", "98.5% (მკაცრად ვალიდური სტრუქტურა)", "99.9% (აბსოლუტური ერთგვაროვნება)", "1-3 Shot სრულად აგვარებს ფორმატის პრობლემას"),
        ("ტოკენების ხარჯი (Token Overhead)", "0 დამატებითი ტოკენი", "~300 – 600 დამატებითი ტოკენი", "~800 – 1500 დამატებითი ტოკენი", "5-Shot-ზე აუცილებელია Prompt Caching"),
        ("პასუხის დაყოვნება (Latency TTFT)", "მინიმალური (< 300 ms)", "მცირედით მომატებული (~450 ms)", "შესამჩნევი (~700 ms ქეშის გარეშე)", "Prompt Caching ამცირებს დაყოვნებას 80%-ით"),
        ("მგრძნობელობა ხმაურზე (Variance)", "მაღალი (მგრძნობიარეა ფორმულირებაზე)", "დაბალი (მაგალითები იცავს გადახრისგან)", "უკიდურესად დაბალი (მყარი ინვარიანტულობა)", "Few-Shot მკვეთრად ზრდის შედეგის განმეორებადობას")
    ]

    for ri, row_data in enumerate(bench_rows):
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
    p.text = "3-Shot მიდგომა უზრუნველყოფს სიზუსტის 15-20%-იან ნახტომს Zero-Shot-თან შედარებით მინიმალური ხარჯით; 5-Shot გამართლებულია მხოლოდ მაშინ, როდესაც საქმე გვაქვს რთულ და მრავალკლასიან კლასიფიკაციასთან."
    p.font.name = FONT_ARIAL; p.font.size = Pt(13); p.font.color.rgb = COLOR_MUTED; p.space_before = Pt(4)

    # -------------------------------------------------------------------------
    # SLIDE 8: Dynamic Few-Shot & Semantic Selection (RAG + Vector DB 2026)
    # -------------------------------------------------------------------------
    s8 = created_slides[6]
    add_header(s8, "დინამიკური Few-Shot: სემანტიკური შერჩევა ვექტორული ბაზიდან (2026)", "ᲛᲝᲬᲘᲜᲐᲕᲔ ᲐᲠᲥᲘᲢᲔᲥᲢᲣᲠᲐ • DYNAMIC FEW-SHOT & RAG")

    dyn_s8 = [
        ("🔍 SEMANTIC SELECTION", "სემანტიკური ძიება", "მოთხოვნაზე მორგებული მაგალითები", [
            ("სტატიკურიდან დინამიკურზე", "პრომპტში მუდმივად ერთი და იგივე მაგალითების ჩაკერვის ნაცვლად, სისტემა დინამიკურად ირჩევს ყველაზე რელევანტურებს."),
            ("Vector DB ინტეგრაცია", "ათასობით ისტორიული, ადამიანის მიერ ვალიდირებული ქეისი ინახება ვექტორულ ბაზაში (Pinecone, Chroma, Qdrant)."),
            ("Cosine Similarity", "შემავალი შეკითხვის ემბედინგის მიხედვით 3 სემანტიკურად ყველაზე მსგავსი მაგალითის ამოღება რეალურ დროში.")
        ]),
        ("🛡️ QUALITY & RELEVANCE", "ხარისხის ფილტრაცია", "მხოლოდ ეტალონური ქეისები", [
            ("Dynamic Sanitization", "ვექტორული ბაზიდან ამოღებული მაგალითები მოწმდება, რომ არ შეიცავდეს PII-ს ან მოძველებულ ბიზნეს-პოლიტიკას."),
            ("Relevance Threshold", "თუ ვექტორული მსგავსება დაბალია (< 0.75), სისტემა იყენებს უსაფრთხო, სტანდარტულ ნაგულისხმევ მაგალითებს."),
            ("Balanced Selection", "ალგორითმი ირჩევს 2 პირდაპირ მსგავს ქეისს + 1 კონტრასტულ სასაზღვრო მაგალითს სრული სურათისთვის.")
        ]),
        ("⚡ FEW-SHOT CACHING 2026", "ქეშირება და მასშტაბირება", "ინფრასტრუქტურული ოპტიმიზაცია", [
            ("Static Prefix + Dynamic Suffix", "სისტემური ინსტრუქცია და ბაზისური მაგალითები მუდმივად ქეშირებულია (Prompt Caching)."),
            ("90%-იანი ეკონომია", "Anthropic, OpenAI და Gemini ქეშირების წყალობით დინამიკური Few-Shot-ის ფასი პრაქტიკულად უტოლდება Zero-Shot-ს."),
            ("საწარმოო მასშტაბი", "მილიონობით მოთხოვნის დამუშავება უმაღლესი სიზუსტით, უსწრაფესი რეაგირებითა და მინიმალური ბიუჯეტით.")
        ])
    ]

    for i, (badge, title, subtitle, items) in enumerate(dyn_s8):
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
    p.text = "2026 წელს წამყვანი საწარმოო არქიტექტურა იყენებს Dynamic Few-Shot-ს: ვექტორული ბაზიდან ამოღებული ზუსტი ანალოგები და Prompt Caching უზრუნველყოფს მაქსიმალურ სიზუსტეს მინიმალურ ფასად."
    p.font.name = FONT_ARIAL; p.font.size = Pt(13); p.font.color.rgb = COLOR_MUTED; p.space_before = Pt(4)

    # -------------------------------------------------------------------------
    # SLIDE 9: Engineering Patterns vs Anti-Patterns & Transition to Topic 6.1
    # -------------------------------------------------------------------------
    s9 = created_slides[7]
    add_header(s9, "საინჟინრო პატერნები და თავიდან ასაცილებელი შეცდომები Few-Shot-ში", "ᲡᲐᲣᲙᲔᲗᲔᲡᲝ ᲞᲠᲐᲥᲢᲘᲙᲐ • PATTERNS & ANTI-PATTERNS")

    box_w = Inches(8.35); box_h = Inches(5.4)

    # Anti-patterns Box
    c_anti = s9.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(1.0), card_y, box_w, box_h)
    c_anti.fill.solid(); c_anti.fill.fore_color.rgb = COLOR_CARD_RED
    c_anti.line.color.rgb = COLOR_BORDER_RED; c_anti.line.width = Pt(1.5)

    tb_anti = s9.shapes.add_textbox(Inches(1.35), card_y + Inches(0.35), box_w - Inches(0.7), box_h - Inches(0.7))
    tf_anti = tb_anti.text_frame; tf_anti.word_wrap = True; tf_anti.margin_left = tf_anti.margin_top = 0
    p = tf_anti.paragraphs[0]; p.text = "❌ CRITICAL ANTI-PATTERNS (კრიტიკული შეცდომები)"
    p.font.name = FONT_ARIAL; p.font.size = Pt(13); p.font.bold = True; p.font.color.rgb = COLOR_RED

    p = tf_anti.add_paragraph(); p.text = "რას უნდა მოვერიდოთ Few-Shot-ში"
    p.font.name = FONT_ARIAL; p.font.size = Pt(16); p.font.bold = True; p.font.color.rgb = COLOR_DARK; p.space_before = Pt(4)

    anti_items = [
        ("Label Imbalance", "4 დადებითი და 0 უარყოფითი მაგალითი, რაც იწვევს მოდელის სრულ სტატისტიკურ გადახრას და დაბრმავებას."),
        ("Inconsistent Formatting", "ერთ მაგალითში JSON, მეორეში გასაღები-მნიშვნელობა, მესამეში თავისუფალი ტექსტი — მოდელი იბნევა."),
        ("Hallucinated Data in Demos", "მაგალითში მცდარი ფაქტის მოყვანა, რაც მოდელს ჰალუცინაციისა და ცრუ დაპირებების ლიცენზიას აძლევს."),
        ("Context Stuffing", "10+ გრძელი მაგალითის ჩაყრა, რაც უაზროდ ავსებს კონტექსტს და იწვევს 'Lost in the Middle' დეგრადაციას.")
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

    p = tf_rec.add_paragraph(); p.text = "სტაბილური Few-Shot არქიტექტურა"
    p.font.name = FONT_ARIAL; p.font.size = Pt(16); p.font.bold = True; p.font.color.rgb = COLOR_DARK; p.space_before = Pt(4)

    rec_items = [
        ("Strict Balanced Classes", "თითოეული საკლასიფიკაციო კატეგორიის მკაცრად თანაბარი წარმომადგენლობა მაგალითებში."),
        ("XML Tag Wrapping", "თითოეული მაგალითის შეფუთვა მკაფიო კონტეინერში (<example id='1'>...</example>)."),
        ("Step-by-Step Reasoning in Demos", "რთულ მაგალითებში მოკლე ახსნის ჩასმა ('explanation': '...'), რაც მოდელს აზროვნებას ასწავლის."),
        ("Systematic Verification", "მაგალითების ბაზის რეგულარული აუდიტი და განახლება კომპანიის პოლიტიკის ცვლილებისთანავე.")
    ]
    for head, body in rec_items:
        p = tf_rec.add_paragraph(); p.text = f"• {head}: {body}"
        p.font.name = FONT_ARIAL; p.font.size = Pt(11); p.font.color.rgb = COLOR_DARK; p.space_before = Pt(6)

    # Bottom takeaway & transition to Lecture 6.1
    bb = s9.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(1.0), bot_y, Inches(17.0), Inches(2.0))
    bb.fill.solid(); bb.fill.fore_color.rgb = COLOR_WHITE
    bb.line.color.rgb = COLOR_BORDER; bb.line.width = Pt(1.2)

    tbb = s9.shapes.add_textbox(Inches(1.4), bot_y + Inches(0.25), Inches(16.2), Inches(1.5))
    tfb = tbb.text_frame; tfb.word_wrap = True; tfb.margin_left = tfb.margin_top = 0
    p = tfb.paragraphs[0]; p.text = "💡 საინჟინრო შეჯამება და გადასვლა ლექციაზე 6.1:"
    p.font.name = FONT_ARIAL; p.font.size = Pt(14); p.font.bold = True; p.font.color.rgb = COLOR_RED
    p = tfb.add_paragraph()
    p.text = "Few-Shot Prompting არის ყველაზე საიმედო ხიდი Zero-Shot-სა და Fine-Tuning-ს შორის. შემდეგ ლექციაში (6.1: ამოცანის ტიპების ბიბლიოთეკა და reasoning რეჟიმები) შევისწავლით: 4 ძირითადი ამოცანის ტიპს (Summarize, Infer, Transform, Expand), Chain-of-Thought მეთოდოლოგიას და თანამედროვე Reasoning მოდელების მუშაობის პრინციპებს."
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
    filename = "5 - Few-Shot Prompting - მაგალითების შერჩევა და შედეგის სტაბილურობა.pptx"
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
    build_lecture_m7_5()
