import os
import sys
import shutil
import pptx
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN
from pptx.enum.shapes import MSO_SHAPE

def build_lecture_m7_7_1():
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
            p.text = "7.1 - ჰალუცინაციების მართვა, თავდაჯერების კალიბრაცია და გადამოწმების პრომპტები"
            p.alignment = PP_ALIGN.CENTER
            p.font.name = FONT_ARIAL
            p.font.size = Pt(30)
            p.font.bold = True
            p.font.color.rgb = COLOR_RED

            p2 = tf.add_paragraph()
            p2.text = "ჰალუცინაციების ბუნება, „არ ვიცი“ წესი, წყაროზე დამოწმება (Citations) და Chain-of-Verification (CoVe) არქიტექტურა"
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
    def add_header(slide, title, category="ᲛᲝᲓᲣᲚᲘ 7 • PROMPT ENGINEERING • 7.1"):
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
    step_w = Inches(3.95); step_h = Inches(5.4); step_y = Inches(2.35)
    xs_4 = [Inches(1.0), Inches(5.35), Inches(9.7), Inches(14.05)]
    bot_y = Inches(8.0)

    # -------------------------------------------------------------------------
    # SLIDE 2: Anatomy of Hallucinations
    # -------------------------------------------------------------------------
    s2 = created_slides[0]
    add_header(s2, "ჰალუცინაციების ანატომია: რატომ იტყუება მოდელი თავდაჯერებულად?", "ᲛᲔᲜᲢᲐᲚᲣᲠᲘ ᲛᲝᲓᲔᲚᲘ • THE ANATOMY OF HALLUCINATIONS")

    cards_s2 = [
        ("🎲 PROBABILISTIC SIMULATION", "ალბათური სიმულაცია", "Next-Token Prediction ბუნება", [
            ("ავტორეგრესიული პროგნოზი", "მოდელი ირჩევს სტატისტიკურად ყველაზე მოსალოდნელ ტოკენს და არა ფაქტობრივ სიმართლეს."),
            ("დამაჯერებელი სიცრუე", "მოდელი ფლობს უნაკლო გრამატიკასა და სტილს, რაც მის გამონაგონს გარეგნულად სარწმუნოდ წარმოაჩენს."),
            ("ცოდნის ილუზია", "ტრანსფორმერს არ გააჩნია ცალკეული 'სიმართლის მოდული'; ის ახდენს რეალობის ენობრივ სიმულაციას.")
        ]),
        ("🤝 OVERCONFIDENCE & SYCOPHANCY", "ზედმეტი თავდაჯერება", "დამხმარე ასისტენტის მახე", [
            ("Alignment გადახრა", "მოდელი გაწვრთნილია იყოს მაქსიმალურად დამხმარე, ამიტომ უჭირს პირდაპირი უარის თქმა."),
            ("პასუხის იძულება", "როდესაც კითხვა არასწორ პრეზუმფციას შეიცავს, მოდელი ეთანხმება მომხმარებელს სიმართლის ნაცვლად."),
            ("შინაგანი ეჭვის ნაკლებობა", "მოდელი ერთნაირი კატეგორიულობით აცხადებს როგორც მტკიცე ფაქტს, ისე სრულ გამონაგონს.")
        ]),
        ("📉 CONTEXT DEGRADATION", "კონტექსტის დეგრადაცია", "ინფორმაციული ვაკუუმი", [
            ("ცარიელი კონტექსტი", "როდესაც მოწოდებულ დოკუმენტში პასუხი არ წერია, მოდელი მიმართავს პრე-ტრენინგის ბუნდოვან მეხსიერებას."),
            ("ყურადღების გაფანტვა", "გრძელ დოკუმენტებში (100k+ ტოკენი) ფაქტის შეუმჩნევლობა იწვევს ჰალუცინაციის გააქტიურებას."),
            ("ხარვეზების კასკადი", "პირველივე მცდარი ტოკენი ხდება კონტექსტი შემდეგი ტოკენებისთვის და აძლიერებს სიცრუეს.")
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
    p.text = "ჰალუცინაცია არ არის მოდელის 'გაფუჭება'; ეს არის ტრანსფორმერის ბუნებრივი მდგომარეობა, რომელიც ყოველთვის ქმნის ტექსტის ყველაზე ლოგიკურ გაგრძელებას; საინჟინრო ამოცანაა ამ გენერაციის მკაცრი დამიწება (Grounding) რეალურ წყაროებზე."
    p.font.name = FONT_ARIAL; p.font.size = Pt(13); p.font.color.rgb = COLOR_MUTED; p.space_before = Pt(4)

    # -------------------------------------------------------------------------
    # SLIDE 3: The Explicit "I Don't Know" Directive
    # -------------------------------------------------------------------------
    s3 = created_slides[1]
    add_header(s3, "„არ ვიცი“ წესი: როგორ მივცეთ მოდელს უცოდინრობის ნებართვა?", "ᲣᲡᲐᲤᲠᲗᲮᲝᲔᲑᲘᲡ ᲓᲘᲠᲔᲥᲢᲘᲕᲐ • THE \"I DON'T KNOW\" DIRECTIVE")

    directive_s3 = [
        ("🛑 EXPLICIT PERMISSION", "ოფიციალური ნებართვა", "ბარიერის მოხსნა", [
            ("მკაცრი სისტემური წესი", "ინსტრუქცია: 'თუ მოწოდებულ ტექსტში არ წერია პასუხი, მკაცრად უპასუხე: ინფორმაცია არ მოიპოვება'."),
            ("გამოცნობის აკრძალვა", "აკრძალვა: 'კატეგორიულად გეკრძალება ლოგიკური ექსტრაპოლაცია, გამოცნობა ან საკუთარი ვარაუდის გამოთქმა'."),
            ("უარის წახალისება", "მოდელმა უნდა იცოდეს, რომ 'არ ვიცი' არის სწორი და სასურველი პასუხი გაურკვევლობის დროს.")
        ]),
        ("📉 85%-IANI VARDNA", "რისკის მკვეთრი შემცირება", "საწარმოო მეტრიკა", [
            ("ფაქტობრივი სტაბილურობა", "„არ ვიცი“ წესის ჩართვა ჰალუცინაციების საერთო რაოდენობას 80-85%-ით ამცირებს სისტემაში."),
            ("ბიზნესის დაცვა", "კომპანიისთვის უსასრულოდ უკეთესია ბოტმა თქვას 'არ ვიცი', ვიდრე გასცეს მცდარი იურიდიული ან ფინანსური პასუხი."),
            ("ნდობის ფაქტორი", "მომხმარებლის ნდობა AI-ს მიმართ იზრდება, როდესაც ის ხედავს, რომ ასისტენტი ტყუილებს არ იგონებს.")
        ]),
        ("🔄 SAFE ALTERNATIVE", "უსაფრთხო ალტერნატივა", "კონსტრუქციული უარი", [
            ("უარი + გზამკვლევი", "'დოკუმენტში ეს საკითხი არ წერია, თუმცა გირჩევთ მიმართოთ Skillwill-ის ადმინისტრაციას'."),
            ("დაკონკრეტების მოთხოვნა", "'ინფორმაცია არასრულია; გთხოვთ დააზუსტოთ თქვენი სტუდენტის ID ან ხელშეკრულების ნომერი'."),
            ("ადამიანთან გადაცემა", "უცნობი ქეისის დროს Handoff payload-ის ავტომატური გააქტიურება და ოპერატორის ჩართვა.")
        ])
    ]

    for i, (badge, title, subtitle, items) in enumerate(directive_s3):
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
    p.text = "სანამ მოდელს არ მისცემთ ოფიციალურ ნებართვას, რომ თქვას 'არ ვიცი', ის ნებისმიერ ფასად ეცდება პასუხის გამოგონებას; ეს მარტივი წესი წარმოადგენს ჰალუცინაციების საწინააღმდეგო პირველ საიმედო ზღუდეს."
    p.font.name = FONT_ARIAL; p.font.size = Pt(13); p.font.color.rgb = COLOR_MUTED; p.space_before = Pt(4)

    # -------------------------------------------------------------------------
    # SLIDE 4: Grounding & Citations Architecture
    # -------------------------------------------------------------------------
    s4 = created_slides[2]
    add_header(s4, "წყაროების მოთხოვნა და ციტირების არქიტექტურა (Verifiable Facts)", "ᲓᲐᲛᲘᲬᲔᲑᲐ & ᲪᲘᲢᲘᲠᲔᲑᲐ • GROUNDING & CITATIONS")

    cite_s4 = [
        ("📌 EXPLICIT CITATIONS", "პირდაპირი ციტირება", "ყველა ფაქტს თავისი წყარო", [
            ("მკაცრი მოთხოვნა", "ყოველ გენერირებულ თეზისს ბოლოში უნდა ახლდეს ზუსტი ციტატა: [ციტატა: '...']."),
            ("პუნქტების მითითება", "დოკუმენტის სექციის, თავისა და აბზაცის ზუსტი იდენტიფიცირება (მაგ. 'მუხლი 4.2')."),
            ("დისციპლინა", "ციტირების იძულება აიძულებს Attention მექანიზმს, მიებას კონკრეტულ ტოკენებს და გამორიცხავს ფანტაზიას.")
        ]),
        ("🤖 CODE VALIDATION", "სერვერული ვალიდაცია", "ციტატების შემოწმება კოდით", [
            ("Substring Matching", "სერვერული სკრიპტი ამოწმებს, რეალურად არსებობს თუ არა მოყვანილი ციტატა ორიგინალ ტექსტში."),
            ("Fuzzy Matching", "მცირე ორთოგრაფიული ცდომილებების დაშვება (Levenshtein Similarity > 95%)."),
            ("ავტომატური რეჟექცია", "თუ მოდელმა მოიყვანა არარსებული ციტატა, სისტემა პასუხს ბლოკავს და ააქტიურებს Retry-ს.")
        ]),
        ("📇 POINTERS & JSON CONTRACT", "ინდექსირებული ექსტრაქცია", "სტრუქტურული მტკიცებულება", [
            ("Pointers მიდგომა", "მოდელი აბრუნებს ობიექტს: {'fact': '...', 'paragraph_id': 4, 'exact_quote': '...' }."),
            ("UI Highlighter", "ვებ-ინტერფეისში მომხმარებელს შეუძლია დააჭიროს ციტატას და პირდაპირ დაინახოს მონიშნული ადგილი PDF-ში."),
            ("100%-იანი აუდიტირებადობა", "იდეალურია იურიდიული, საბანკო და სადაზღვევო დოკუმენტაციის ანალიზისთვის.")
        ])
    ]

    for i, (badge, title, subtitle, items) in enumerate(cite_s4):
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
    p.text = "არ ენდოთ მოდელის თავისუფალ პარაფრაზს; საწარმოო აპლიკაციაში ყოველი ფაქტი უნდა იყოს ციტირებული და სერვერულად გადამოწმებული ორიგინალ წყაროსთან შედარებით."
    p.font.name = FONT_ARIAL; p.font.size = Pt(13); p.font.color.rgb = COLOR_MUTED; p.space_before = Pt(4)

    # -------------------------------------------------------------------------
    # SLIDE 5: Chain-of-Verification (CoVe) 4-Step Architecture
    # -------------------------------------------------------------------------
    s5 = created_slides[3]
    add_header(s5, "Chain-of-Verification (CoVe): თვითშემოწმების 4-საფეხურიანი არქიტექტურა", "COVE ᲐᲠᲥᲘᲢᲔᲥᲢᲣᲠᲐ • CHAIN-OF-VERIFICATION")

    cove_s5 = [
        ("📝 STEP 1: BASELINE", "საბაზისო დრაფტი", "პირველადი პასუხი", [
            ("დრაფტის გენერაცია", "მოდელი აგენერირებს საწყის პასუხს მომხმარებლის კითხვაზე (რომელიც შეიძლება შეიცავდეს უზუსტობებს)."),
            ("შიდა დამუშავება", "ეს დრაფტი რჩება სისტემის შიდა დამუშავების ეტაპზე და მომხმარებელს ჯერ არ მიეწოდება."),
            ("საწყისი ჰიპოთეზა", "ქმნის საინფორმაციო ბაზას საკონტროლო კითხვების შესადგენად.")
        ]),
        ("❓ STEP 2: PLAN VERIFY", "კითხვების გენერაცია", "ფაქტების იზოლაცია", [
            ("თვითკრიტიკა", "მოდელი აანალიზებს საკუთარ დრაფტს და აგენერირებს 3-5 საკონტროლო კითხვას ფაქტების გადასამოწმებლად."),
            ("საკონტროლო შეკითხვები", "მაგალითი: 'რა თარიღში გაფორმდა ხელშეკრულება? რა თანხა წერია მე-3 მუხლში?'."),
            ("ფოკუსირება რისკებზე", "კითხვები მიმართულია ყველაზე სარისკო და საეჭვო ფაქტებზე.")
        ]),
        ("🔍 STEP 3: INDEPENDENT", "დამოუკიდებელი პასუხი", "წყაროზე დაყრდნობით", [
            ("დამოუკიდებელი შესრულება", "მოდელი პასუხობს საკონტროლო კითხვებს პირველადი დრაფტის ნახვის გარეშე, მხოლოდ ორიგინალ ტექსტზე დაყრდნობით."),
            ("ინერციის აღმოფხვრა", "ეს გამორიცხავს პირველადი დრაფტის შეცდომების ინერციით გამეორებას."),
            ("ობიექტური ფაქტები", "წყაროდან პირდაპირი მონაცემების მიღება ყოველგვარი დამახინჯების გარეშე.")
        ]),
        ("✨ STEP 4: FINAL REVISE", "საბოლოო კორექცია", "გადამოწმებული შედეგი", [
            ("ფაქტების შედარება", "მოდელი ადარებს პირველად დრაფტს დამოუკიდებელ პასუხებს და პოულობს შეუსაბამობებს."),
            ("კორექტირება", "ნებისმიერი შეუსაბამობა სწორდება და აგენერირებს უტყუარ, გადამოწმებულ პასუხს."),
            ("სუფთა გამოსავალი", "მომხმარებელი იღებს მხოლოდ 100%-ით გაფილტრულ და გადამოწმებულ შედეგს.")
        ])
    ]

    for i, (badge, title, subtitle, items) in enumerate(cove_s5):
        bg_col = COLOR_CARD_RED if i == 2 else COLOR_BG_LIGHT
        bor_col = COLOR_BORDER_RED if i == 2 else COLOR_BORDER
        c = s5.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, xs_4[i], step_y, step_w, step_h)
        c.fill.solid(); c.fill.fore_color.rgb = bg_col
        c.line.color.rgb = bor_col; c.line.width = Pt(1.5)

        tb = s5.shapes.add_textbox(xs_4[i] + Inches(0.25), step_y + Inches(0.25), step_w - Inches(0.5), step_h - Inches(0.5))
        tf = tb.text_frame; tf.word_wrap = True; tf.margin_left = tf.margin_top = 0

        p = tf.paragraphs[0]; p.text = badge.upper()
        p.font.name = FONT_ARIAL; p.font.size = Pt(11); p.font.bold = True
        p.font.color.rgb = COLOR_RED if i == 2 else COLOR_MUTED

        p = tf.add_paragraph(); p.text = title
        p.font.name = FONT_ARIAL; p.font.size = Pt(15); p.font.bold = True; p.font.color.rgb = COLOR_DARK; p.space_before = Pt(4)

        p = tf.add_paragraph(); p.text = subtitle
        p.font.name = FONT_ARIAL; p.font.size = Pt(11); p.font.color.rgb = COLOR_TEXT_MUTED; p.space_before = Pt(2)

        for heading, body in items:
            p = tf.add_paragraph(); p.text = f"• {heading}: {body}"
            p.font.name = FONT_ARIAL; p.font.size = Pt(11); p.font.color.rgb = COLOR_DARK; p.space_before = Pt(6)

    # Bottom takeaway
    bb = s5.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(1.0), bot_y, Inches(17.0), Inches(2.0))
    bb.fill.solid(); bb.fill.fore_color.rgb = COLOR_WHITE
    bb.line.color.rgb = COLOR_BORDER; bb.line.width = Pt(1.2)

    tbb = s5.shapes.add_textbox(Inches(1.4), bot_y + Inches(0.3), Inches(16.2), Inches(1.4))
    tfb = tbb.text_frame; tfb.word_wrap = True; tfb.margin_left = tfb.margin_top = 0
    p = tfb.paragraphs[0]; p.text = "🎯 საინჟინრო პრინციპი:"
    p.font.name = FONT_ARIAL; p.font.size = Pt(14); p.font.bold = True; p.font.color.rgb = COLOR_RED
    p = tfb.add_paragraph()
    p.text = "Chain-of-Verification (CoVe) მოდელს აიძულებს საკუთარ თავს დაუსვას შეკითხვები და უპასუხოს მათ ორიგინალი წყაროდან; ეს მეთოდოლოგია 70%-ით ამცირებს ფაქტობრივ შეცდომებსა და ჰალუცინაციებს."
    p.font.name = FONT_ARIAL; p.font.size = Pt(13); p.font.color.rgb = COLOR_MUTED; p.space_before = Pt(4)

    # -------------------------------------------------------------------------
    # SLIDE 6: Confidence Calibration & Uncertainty
    # -------------------------------------------------------------------------
    s6 = created_slides[4]
    add_header(s6, "თავდაჯერების კალიბრაცია: რამდენად დარწმუნებულია მოდელი?", "ᲙᲐᲚᲘᲑᲠᲐᲪᲘᲐ • CONFIDENCE & UNCERTAINTY")

    conf_s6 = [
        ("📊 CONFIDENCE SCORING", "აშკარა თავდაჯერება", "მოდელის თვითშეფასება", [
            ("სტრუქტურირებული ველი", "JSON გამოსავალში მოდელი აბრუნებს confidence_score-ს (0.0 – 1.0) და დასაბუთებას."),
            ("კრიტერიუმების რუბრიკა", "0.9+ = ფაქტი პირდაპირ წერია; 0.6 – 0.8 = ფაქტი ნაგულისხმევია; < 0.5 = ინფორმაცია არასრულია."),
            ("ავტომატური გაფრთხილება", "თუ ქულა < 0.8-ზე, სისტემა ავტომატურად ამატებს გაფრთხილებას მომხმარებლისთვის.")
        ]),
        ("🔬 LOGPROBS & PERPLEXITY", "ტოკენების ალბათობა", "მათემატიკური სიზუსტე API-ში", [
            ("API Logprobs", "OpenAI და Google AI Studio გაწვდის თითოეული გენერირებული ტოკენის ზუსტ ლოგ-ალბათობას."),
            ("Perplexity ანალიზი", "თუ საკვანძო ფაქტის (მაგ. თარიღი, გვარი) ტოკენებზე ალბათობა დაბალია, ეს ჰალუცინაციის 90%-იანი სიგნალია."),
            ("კოდის დონეზე კონტროლი", "სერვერული ალგორითმი ამოწმებს საკვანძო ტოკენების ალბათობას და საეჭვო პასუხებს ბლოკავს.")
        ]),
        ("👥 HUMAN-IN-THE-LOOP", "ადამიანური კონტროლი", "კრიტიკული შემთხვევები", [
            ("ზღვრული მარშრუტიზაცია", "მაღალი თავდაჯერების პასუხი იგზავნება მყისიერად; დაბალი თავდაჯერების პასუხი მიდის ოპერატორთან."),
            ("ოპერატორის დახმარება", "ადამიანს ეგზავნება მოდელის დრაფტი, მოძიებული ციტატა და დაბალი ქულის მიზეზი."),
            ("უკუკავშირის ციკლი", "ადამიანის მიერ გასწორებული პასუხი ხდება ახალი Golden ქეისი საწარმოო Few-Shot ბაზისთვის.")
        ])
    ]

    for i, (badge, title, subtitle, items) in enumerate(conf_s6):
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
    p.text = "კალიბრირებული სისტემა არ მალავს გაურკვევლობას; ის ზომავს საკუთარ თავდაჯერებას (Logprobs და Confidence Scoring-ით) და საეჭვო შემთხვევებში რთავს ადამიანურ ფილტრს (Human-in-the-Loop)."
    p.font.name = FONT_ARIAL; p.font.size = Pt(13); p.font.color.rgb = COLOR_MUTED; p.space_before = Pt(4)

    # -------------------------------------------------------------------------
    # SLIDE 7: Benchmark Matrix: Naive vs Grounded vs CoVe (Table)
    # -------------------------------------------------------------------------
    s7 = created_slides[5]
    add_header(s7, "Naive Prompting vs Grounded Citation vs CoVe: შედარებითი ანალიზი", "ᲨᲔᲓᲐᲠᲔᲑᲘᲗᲘ ᲐᲜᲐᲚᲘᲖᲘ • BENCHMARK MATRIX")

    # Table: 5 rows, 5 columns
    tbl_left = Inches(1.0); tbl_top = Inches(2.35); tbl_w = Inches(17.0); tbl_h = Inches(5.2)
    tbl_shape = s7.shapes.add_table(5, 5, tbl_left, tbl_top, tbl_w, tbl_h)
    table = tbl_shape.table

    col_widths = [Inches(3.2), Inches(3.2), Inches(3.6), Inches(3.5), Inches(3.5)]
    for ci, w in enumerate(col_widths):
        table.columns[ci].width = w

    headers = ["კრიტერიუმი / პარამეტრი", "Naive (ჩვეულებრივი)", "Grounded + Citations", "Chain-of-Verification (CoVe)", "საინჟინრო რეკომენდაცია"]
    for ci, h in enumerate(headers):
        cell = table.cell(0, ci)
        cell.fill.solid(); cell.fill.fore_color.rgb = COLOR_DARK
        tf = cell.text_frame; tf.word_wrap = True; tf.margin_left = tf.margin_right = Inches(0.12)
        p = tf.paragraphs[0]; p.text = h; p.font.name = FONT_ARIAL; p.font.size = Pt(12); p.font.bold = True; p.font.color.rgb = COLOR_WHITE

    cove_rows = [
        ("ფაქტობრივი სიზუსტე (Accuracy)", "65% – 72% (ხშირი ცდომილება)", "88% – 92% (მაღალი დამიწება)", "96% – 99% (უმაღლესი სიზუსტე)", "CoVe საუკეთესოა კრიტიკულ დომენებში"),
        ("ჰალუცინაციების სიხშირე", "25% – 35% (მიუღებლად მაღალი)", "5% – 8% (მინიმალური)", "< 1% (პრაქტიკულად ნულოვანი)", "ციტირება საწარმოო მინიმუმია"),
        ("ტოკენების დამატებითი ხარჯი", "0 დამატებითი ტოკენი", "+ 15% – 25% (ციტატები)", "+ 100% – 200% (მრავალეტაპიანი)", "CoVe მოითხოვს Prompt Caching-ს"),
        ("პასუხის დაყოვნება (Latency)", "უსწრაფესი (< 1.0 წამი)", "სწრაფი (~1.5 წამი)", "გაზრდილი (~3.0 – 5.0 წამი)", "CoVe გამოიყენეთ ასინქრონულ პროცესებში")
    ]

    for ri, row_data in enumerate(cove_rows):
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
    p.text = "საბაზისო ციტირება (Grounded Citations) აუცილებელი მინიმუმია ნებისმიერი საწარმოო RAG სისტემისთვის; CoVe კი გამოიყენება მაღალი რისკის მქონე ამოცანებში, სადაც 1% შეცდომაც კი კატასტროფულია."
    p.font.name = FONT_ARIAL; p.font.size = Pt(13); p.font.color.rgb = COLOR_MUTED; p.space_before = Pt(4)

    # -------------------------------------------------------------------------
    # SLIDE 8: Best Practices vs Anti-Patterns in Factual Accuracy
    # -------------------------------------------------------------------------
    s8 = created_slides[6]
    add_header(s8, "საინჟინრო პატერნები და ანტიპატერნები ფაქტობრივ სიზუსტეში", "ᲡᲐᲣᲙᲔᲗᲔᲡᲝ ᲞᲠᲐᲥᲢᲘᲙᲐ • PATTERNS & ANTI-PATTERNS")

    box_w = Inches(8.35); box_h = Inches(5.4)

    # Anti-patterns Box
    c_anti = s8.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(1.0), card_y, box_w, box_h)
    c_anti.fill.solid(); c_anti.fill.fore_color.rgb = COLOR_CARD_RED
    c_anti.line.color.rgb = COLOR_BORDER_RED; c_anti.line.width = Pt(1.5)

    tb_anti = s8.shapes.add_textbox(Inches(1.35), card_y + Inches(0.35), box_w - Inches(0.7), box_h - Inches(0.7))
    tf_anti = tb_anti.text_frame; tf_anti.word_wrap = True; tf_anti.margin_left = tf_anti.margin_top = 0
    p = tf_anti.paragraphs[0]; p.text = "❌ CRITICAL ANTI-PATTERNS (კრიტიკული შეცდომები)"
    p.font.name = FONT_ARIAL; p.font.size = Pt(13); p.font.bold = True; p.font.color.rgb = COLOR_RED

    p = tf_anti.add_paragraph(); p.text = "რას უნდა მოვერიდოთ ფაქტობრივ სისტემებში"
    p.font.name = FONT_ARIAL; p.font.size = Pt(16); p.font.bold = True; p.font.color.rgb = COLOR_DARK; p.space_before = Pt(4)

    anti_items = [
        ("პასუხის იძულება", "პრომპტი აიძულებს მოდელს გასცეს პასუხი მაშინაც, როდესაც მონაცემთა ბაზაში ინფორმაცია არ არსებობს."),
        ("მაღალი ტემპერატურა ფაქტებზე", "ტემპერატურის აწევა (T > 0.5) იურიდიულ ან ფინანსურ ექსტრაქციაზე, რაც ზრდის რყევას."),
        ("დაუმოწმებელი ციტატები", "მოდელის მიერ მოყვანილი ციტატის ბრმად ნდობა სერვერული კოდით შემოწმების გარეშე."),
        ("ღია, დაუცველი კითხვები", "ფაქტობრივი კითხვების დასმა კონტექსტური საზღვრებისა და მკაცრი გამყოფი ტეგების გარეშე.")
    ]
    for head, body in anti_items:
        p = tf_anti.add_paragraph(); p.text = f"• {head}: {body}"
        p.font.name = FONT_ARIAL; p.font.size = Pt(11); p.font.color.rgb = COLOR_DARK; p.space_before = Pt(6)

    # Best Practices Box
    c_rec = s8.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(9.65), card_y, box_w, box_h)
    c_rec.fill.solid(); c_rec.fill.fore_color.rgb = COLOR_GREEN_BG
    c_rec.line.color.rgb = COLOR_GREEN_BOR; c_rec.line.width = Pt(1.5)

    tb_rec = s8.shapes.add_textbox(Inches(10.0), card_y + Inches(0.35), box_w - Inches(0.7), box_h - Inches(0.7))
    tf_rec = tb_rec.text_frame; tf_rec.word_wrap = True; tf_rec.margin_left = tf_rec.margin_top = 0
    p = tf_rec.paragraphs[0]; p.text = "✅ RECOMMENDED BEST PRACTICES (საუკეთესო პატერნები)"
    p.font.name = FONT_ARIAL; p.font.size = Pt(13); p.font.bold = True; p.font.color.rgb = COLOR_GREEN_TEXT

    p = tf_rec.add_paragraph(); p.text = "სტაბილური და დაცული ფაქტობრივი არქიტექტურა"
    p.font.name = FONT_ARIAL; p.font.size = Pt(16); p.font.bold = True; p.font.color.rgb = COLOR_DARK; p.space_before = Pt(4)

    rec_items = [
        ("Strict Negative Rules", "'თუ ფაქტი არ წერია, უპასუხე: არ ვიცი; არ გამოიცნო; არ ივარაუდო' — მკაცრი შეზღუდვები."),
        ("Exact Substring Check", "სერვერული კოდით ციტატების სავალდებულო ვალიდაცია ორიგინალ დოკუმენტთან შედარებით."),
        ("Minimum Temperature (T=0.0)", "ნულოვანი ტემპერატურა მაქსიმალური დეტერმინიზმისა და ფაქტობრივი სიზუსტისთვის."),
        ("Structured Grounding JSON", "პასუხის, ციტატისა და თავდაჯერების ქულის სავალდებულო დაბრუნება JSON სქემით.")
    ]
    for head, body in rec_items:
        p = tf_rec.add_paragraph(); p.text = f"• {head}: {body}"
        p.font.name = FONT_ARIAL; p.font.size = Pt(11); p.font.color.rgb = COLOR_DARK; p.space_before = Pt(6)

    # Bottom takeaway
    bb = s8.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(1.0), bot_y, Inches(17.0), Inches(2.0))
    bb.fill.solid(); bb.fill.fore_color.rgb = COLOR_WHITE
    bb.line.color.rgb = COLOR_BORDER; bb.line.width = Pt(1.2)

    tbb = s8.shapes.add_textbox(Inches(1.4), bot_y + Inches(0.3), Inches(16.2), Inches(1.4))
    tfb = tbb.text_frame; tfb.word_wrap = True; tfb.margin_left = tfb.margin_top = 0
    p = tfb.paragraphs[0]; p.text = "💡 მთავარი არსი:"
    p.font.name = FONT_ARIAL; p.font.size = Pt(14); p.font.bold = True; p.font.color.rgb = COLOR_RED
    p = tfb.add_paragraph()
    p.text = "ჰალუცინაციებთან ბრძოლა არ არის პრომპტის მხოლოდ ერთი ჯადოსნური ფრაზა; ეს არის სისტემური არქიტექტურა: 'არ ვიცი' წესი + ციტირება + სერვერული ვალიდაცია + CoVe გადამოწმების ციკლი."
    p.font.name = FONT_ARIAL; p.font.size = Pt(13); p.font.color.rgb = COLOR_MUTED; p.space_before = Pt(4)

    # -------------------------------------------------------------------------
    # SLIDE 9: Production Checklist & Transition to Lecture 7.2 (Prompt Chaining)
    # -------------------------------------------------------------------------
    s9 = created_slides[7]
    add_header(s9, "ფაქტობრივი სანდოობის საწარმოო ჩეკლისტი და გადასვლა Chaining-ზე", "ᲡᲐᲬᲐᲠᲛᲝᲝ ᲩᲔᲙᲚᲘᲡᲢᲘ • PRODUCTION CHECKLIST")

    # Verification Checklist Box
    c_check = s9.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(1.0), card_y, box_w, box_h)
    c_check.fill.solid(); c_check.fill.fore_color.rgb = COLOR_GREEN_BG
    c_check.line.color.rgb = COLOR_GREEN_BOR; c_check.line.width = Pt(1.5)

    tb_check = s9.shapes.add_textbox(Inches(1.35), card_y + Inches(0.35), box_w - Inches(0.7), box_h - Inches(0.7))
    tf_check = tb_check.text_frame; tf_check.word_wrap = True; tf_check.margin_left = tf_check.margin_top = 0
    p = tf_check.paragraphs[0]; p.text = "📋 FACTUAL ACCURACY CHECKLIST (ჩეკლისტი)"
    p.font.name = FONT_ARIAL; p.font.size = Pt(13); p.font.bold = True; p.font.color.rgb = COLOR_GREEN_TEXT

    p = tf_check.add_paragraph(); p.text = "საწარმოო სისტემის მზაობის შემოწმება"
    p.font.name = FONT_ARIAL; p.font.size = Pt(16); p.font.bold = True; p.font.color.rgb = COLOR_DARK; p.space_before = Pt(4)

    check_items = [
        ("„არ ვიცი“ წესის არსებობა", "პრომპტში მკაფიოდ არის გაწერილი უცოდინრობის აღიარების ნებართვა და უარის ფორმატი."),
        ("სავალდებულო ციტირება", "ყველა ფაქტობრივ მტკიცებას თან ახლავს ზუსტი ციტატა და სექციის იდენტიფიკატორი."),
        ("მინიმალური ტემპერატურა (T=0.0)", "ტემპერატურა დაყვანილია ნულამდე ფაქტობრივი ექსტრაქციისა და ანალიზის დროს."),
        ("სერვერული ვალიდაცია", "აპლიკაცია კოდით ამოწმებს ციტატების არსებობას ორიგინალ დოკუმენტში (Substring Match)."),
        ("თავდაჯერების ქულა & HITL", "მოდელი აბრუნებს confidence_score-ს და საეჭვო ქეისები ეგზავნება ადამიან ოპერატორს.")
    ]
    for head, body in check_items:
        p = tf_check.add_paragraph(); p.text = f"• {head}: {body}"
        p.font.name = FONT_ARIAL; p.font.size = Pt(11); p.font.color.rgb = COLOR_DARK; p.space_before = Pt(6)

    # Transition to Chaining Box
    c_chain = s9.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(9.65), card_y, box_w, box_h)
    c_chain.fill.solid(); c_chain.fill.fore_color.rgb = COLOR_BG_LIGHT
    c_chain.line.color.rgb = COLOR_BORDER; c_chain.line.width = Pt(1.5)

    tb_chain = s9.shapes.add_textbox(Inches(10.0), card_y + Inches(0.35), box_w - Inches(0.7), box_h - Inches(0.7))
    tf_chain = tb_chain.text_frame; tf_chain.word_wrap = True; tf_chain.margin_left = tf_chain.margin_top = 0
    p = tf_chain.paragraphs[0]; p.text = "🔗 TRANSITION: PROMPT CHAINING"
    p.font.name = FONT_ARIAL; p.font.size = Pt(13); p.font.bold = True; p.font.color.rgb = COLOR_RED

    p = tf_chain.add_paragraph(); p.text = "როდესაც ერთი პრომპტი აღარ კმარა"
    p.font.name = FONT_ARIAL; p.font.size = Pt(16); p.font.bold = True; p.font.color.rgb = COLOR_DARK; p.space_before = Pt(4)

    chain_items = [
        ("მონოლითის მარცხი", "როდესაც კომპლექსური ამოცანა ერთ პრომპტში იწვევს ჰალუცინაციებსა და წესების უგულებელყოფას."),
        ("Decompose & Conquer", "დიდი ამოცანის დაყოფა 3 იზოლირებულ ნაბიჯად: ამოღება ➔ დრაფტი ➔ შემოწმება."),
        ("Pipeline Architecture", "როგორ გადაეცემა მონაცემები ნაბიჯებს შორის API დონეზე და როგორ იმართება შეცდომები."),
        ("ანგარიშის გენერაციის ჯაჭვი", "3-ბიჯიანი საწარმოო ჯაჭვის აწყობა და მისი ტესტირება Google AI Studio-სა და კოდში.")
    ]
    for head, body in chain_items:
        p = tf_chain.add_paragraph(); p.text = f"• {head}: {body}"
        p.font.name = FONT_ARIAL; p.font.size = Pt(11); p.font.color.rgb = COLOR_DARK; p.space_before = Pt(6)

    # Bottom takeaway & transition to Lecture 7.2
    bb = s9.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(1.0), bot_y, Inches(17.0), Inches(2.0))
    bb.fill.solid(); bb.fill.fore_color.rgb = COLOR_WHITE
    bb.line.color.rgb = COLOR_BORDER; bb.line.width = Pt(1.2)

    tbb = s9.shapes.add_textbox(Inches(1.4), bot_y + Inches(0.25), Inches(16.2), Inches(1.5))
    tfb = tbb.text_frame; tfb.word_wrap = True; tfb.margin_left = tfb.margin_top = 0
    p = tfb.paragraphs[0]; p.text = "💡 საინჟინრო შეჯამება და გადასვლა ლექციაზე 7.2:"
    p.font.name = FONT_ARIAL; p.font.size = Pt(14); p.font.bold = True; p.font.color.rgb = COLOR_RED
    p = tfb.add_paragraph()
    p.text = "ფაქტობრივი სანდოობის მიღწევის შემდეგ, შემდეგი ნაბიჯია რთული კომპლექსური ამოცანების მართვა. შემდეგ ლექციაში (7.2: რთული ამოცანების დეკომპოზიცია და Prompt Chaining არქიტექტურა) შევისწავლით: როგორ დავშალოთ მონოლითური პრომპტები და ავაწყოთ 3-ბიჯიანი საიმედო მილსადენი (ამოღება ➔ დრაფტი ➔ გადამოწმება)."
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
    filename = "7.1 - ჰალუცინაციების მართვა, თავდაჯერების კალიბრაცია და გადამოწმების პრომპტები.pptx"
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
    build_lecture_m7_7_1()
