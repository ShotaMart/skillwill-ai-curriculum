import os
import sys
import shutil
import pptx
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN
from pptx.enum.shapes import MSO_SHAPE

def build_lecture_m7_4_1():
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
            p.text = "4.1 - Prompt Templates, პარამეტრიზაცია და გუნდური სტანდარტიზაცია"
            p.alignment = PP_ALIGN.CENTER
            p.font.name = FONT_ARIAL
            p.font.size = Pt(32)
            p.font.bold = True
            p.font.color.rgb = COLOR_RED

            p2 = tf.add_paragraph()
            p2.text = "დინამიკური შაბლონების არქიტექტურა, ცვლადების ვალიდაცია და ცენტრალიზებული Prompt Registry 2026 წელს"
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
    def add_header(slide, title, category="ᲛᲝᲓᲣᲚᲘ 7 • PROMPT ENGINEERING • 4.1"):
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
    # SLIDE 2: Paradigm Shift (Hardcoded Prompts vs Prompt Templates)
    # -------------------------------------------------------------------------
    s2 = created_slides[0]
    add_header(s2, "სტატიკური ტექსტიდან პროგრამულ სისტემამდე: რატომ გვჭირდება შაბლონები?", "ᲞᲐᲠᲐᲓᲘᲒᲛᲐ • HARDCODED PROMPTS VS PROMPT TEMPLATES")

    cards_s2 = [
        ("📄 HARDCODED STRINGS", "სტატიკური ტექსტის ჩიხი", "ტექნიკური ვალი და ქაოსი", [
            ("Code Smell & ფრაგმენტაცია", "პრომპტები გაფანტულია კოდის სხვადასხვა ფაილში სტატიკურ ტექსტურ სტრიქონებად."),
            ("დუბლირება და უცვლელობა", "ერთი და იგივე პრომპტის კოპირება რამდენიმე სერვისში, რაც რთულს ხდის ცვლილებების სინქრონიზაციას."),
            ("ლოგიკისა და შინაარსის აღრევა", "დეველოპერს უწევს მთელი აპლიკაციის ხელახალი Deploy მარტივი ტექსტური კორექტირებისთვის.")
        ]),
        ("🧩 TEMPLATE ARCHITECTURE", "დინამიკური შაბლონი", "მოდულური ჩონჩხი და სლოტები", [
            ("ინვარიანტული ჩონჩხი", "სისტემური ინსტრუქცია, როლი, ქცევის წესები და შეზღუდვები გამოყოფილია ცვლადი მონაცემებისგან."),
            ("Dynamic Slots", "ცვლადი პარამეტრების ({{variable}}) იზოლირებული ინექცია შესრულების დროს (Runtime)."),
            ("გარემოზე მორგება", "შაბლონების მოქნილი მართვა გარემოს ცვლადებითა და ცენტრალიზებული კონფიგურაციით.")
        ]),
        ("🚀 PRODUCTION SCALABILITY", "მასშტაბირებადობა", "ათასობით მოთხოვნის მომსახურება", [
            ("ავტომატური პარამეტრიზაცია", "მომხმარებლის პროფილის, ენის და ისტორიის დინამიკური ინტეგრაცია მილიწამებში."),
            ("მრავალენოვანი ლოკალიზაცია", "ერთიანი ბიზნეს-სტრუქტურის გამოყენება სხვადასხვა ენობრივ ბაზაზე."),
            ("უსაფრთხოების ფილტრაცია", "შემავალი ცვლადების წინასწარი სანიტარიზაცია და ვალიდაცია LLM-თან გაგზავნამდე.")
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
    p = tfb.paragraphs[0]; p.text = "💡 მთავარი არსი (საინჟინრო ევოლუცია):"
    p.font.name = FONT_ARIAL; p.font.size = Pt(14); p.font.bold = True; p.font.color.rgb = COLOR_RED
    p = tfb.add_paragraph()
    p.text = "Prompt Template არ არის უბრალო ტექსტი – ეს არის პროგრამული ფუნქცია, რომელიც იღებს პარამეტრებს (Input Variables), ახდენს მათ ვალიდაციას და აგენერირებს მოდელისთვის უსაფრთხო, სტრუქტურირებულ კონტექსტს."
    p.font.name = FONT_ARIAL; p.font.size = Pt(13); p.font.color.rgb = COLOR_MUTED; p.space_before = Pt(4)

    # -------------------------------------------------------------------------
    # SLIDE 3: Variable Syntax & Structural Slotting
    # -------------------------------------------------------------------------
    s3 = created_slides[1]
    add_header(s3, "ცვლადების არქიტექტურა, სინტაქსური სტანდარტები და სტრუქტურული სლოტები", "ᲐᲠᲥᲘᲢᲔᲥᲢᲣᲠᲐ • VARIABLE SYNTAX & SLOTTING")

    slots_s3 = [
        ("🔤 SYNTAX STANDARDS", "სინტაქსური ფორმატები", "ინდუსტრიული სტანდარტები", [
            ("Mustache / Jinja2: {{var}}", "2026 წლის ოქროს სტანდარტი template engines-სა და LLM API-ებში; მკაფიოდ განასხვავებს ენობრივ ცვლადს."),
            ("Python f-string: {var}", "გამოიყენება მარტივ სკრიპტებში, თუმცა შეიცავს კონფლიქტის მაღალ რისკს JSON ფორმატის ფიგურულ ფრჩხილებთან."),
            ("Strict Escaping წესები", "JSON სქემებთან მუშაობისას აუცილებელია ფრჩხილების სწორი ეკრანირება ({ -> {{), რათა პარსერმა შეცდომა არ დაუშვას.")
        ]),
        ("📊 VARIABLE TYPES", "ცვლადების ტიპოლოგია", "მონაცემთა მრავალფეროვნება", [
            ("Primitive Data Types", "სტრიქონები (მომხმარებლის სახელი, ID, თარიღი) და რიცხვითი მონაცემები (ფასი, ასაკი, რეიტინგი)."),
            ("Structured Objects", "JSON ობიექტები და მასივები (მაგ. კალათაში არსებული პროდუქტების სია, შეკვეთის ისტორია)."),
            ("System Metadata", "მომხმარებლის როლი, გარემოს მეტამონაცემები, სესიის ID და ავტორიზაციის დონე.")
        ]),
        ("🏷️ ISOLATION TAGS", "სტრუქტურული იზოლაცია", "2026 წლის საუკეთესო პრაქტიკა", [
            ("XML/HTML სტილის ტეგები", "ცვლადის მკაცრი შეფუთვა კონტეინერში (<user_query>{{query}}</user_query>)."),
            ("Prompt Boundaries", "ტეგები მოდელს უჩვენებს ურყევ ზღვარს სისტემურ დირექტივასა და შემავალ მონაცემებს შორის."),
            ("Frontier მოდელების ოპტიმიზაცია", "GPT-6 და Claude 5.x მკაფიოდ აღიქვამენ ტეგებში მოქცეულ ტექსტს როგორც მონაცემს და არა ბრძანებას.")
        ])
    ]

    for i, (badge, title, subtitle, items) in enumerate(slots_s3):
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
    p.text = "არასოდეს ჩასვათ მომხმარებლის ცვლადი დაუცველად შაბლონის შუაგულში; ყოველთვის გამოიყენეთ სტრუქტურული გამყოფი ტეგები (<context>, <input>), რათა მოდელმა შემავალი მონაცემი ინსტრუქციაში არ აურიოს."
    p.font.name = FONT_ARIAL; p.font.size = Pt(13); p.font.color.rgb = COLOR_MUTED; p.space_before = Pt(4)

    # -------------------------------------------------------------------------
    # SLIDE 4: Variable Validation & Input Sanitization
    # -------------------------------------------------------------------------
    s4 = created_slides[2]
    add_header(s4, "ცვლადების ვალიდაცია და უსაფრთხოება: ნუ ენდობი ცვლადის მნიშვნელობას!", "ᲣᲡᲐᲤᲠᲗᲮᲝᲔᲑᲐ • INPUT SANITIZATION & INJECTION DEFENSE")

    val_s4 = [
        ("🛡️ INPUT SANITIZATION", "მონაცემთა გაწმენდა", "Delimiter Escape დაცვა", [
            ("ტეგების ეკრანირება", "თუ მომხმარებლის ტექსტი შეიცავს </user_query> ან <system> ტეგებს, მათი ავტომატური განეიტრალება."),
            ("საკვანძო ფრაზების შემოწმება", "ცვლადის სკანირება შეტევის პატერნებზე: 'Ignore previous instructions', 'System override'."),
            ("Encoding & Normalization", "უნიკოდის ნორმალიზაცია (NFKC) და მალული სიმბოლოების (Zero-width characters) მოცილება.")
        ]),
        ("📏 BOUNDS & SCHEMA CHECK", "ზომისა და ტიპის კონტროლი", "რესურსების დაცვა", [
            ("სიგრძის ლიმიტები (Max Length)", "ცვლადის მაქსიმალური ზომის შეზღუდვა ტოკენების ამოწურვის (Context Exhaustion) თავიდან ასაცილებლად."),
            ("Pydantic / Zod ვალიდაცია", "ცვლადის ტიპობრივი შემოწმება პროგრამულ დონეზე, სანამ ის LLM-თან გაიგზავნება."),
            ("ცარიელი ველების ბლოკირება", "შეცდომის დაფიქსირება და ვალიდაცია, თუ სავალდებულო პარამეტრი ცარიელია ან არასწორი ტიპისაა.")
        ]),
        ("🔄 DEFAULTS & FALLBACKS", "ნაგულისხმევი ლოგიკა", "გრაციოზული დეგრადაცია", [
            ("Default Arguments", "სათადარიგო მნიშვნელობები არასავალდებულო ცვლადებისთვის (მაგ. language = 'ka', tone = 'formal')."),
            ("Graceful Degradation", "თუ კონტექსტური ცვლადი მიუწვდომელია (მაგ. CRM დროებით გათიშულია), შაბლონი გადადის ზოგად რეჟიმზე."),
            ("Fallback Prompt", "ავარიული პრომპტის გააქტიურება სისტემური კრიტიკული ხარვეზების დროს.")
        ])
    ]

    for i, (badge, title, subtitle, items) in enumerate(val_s4):
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
    p = tfb.paragraphs[0]; p.text = "💡 მთავარი არსი (Zero-Trust Prompting):"
    p.font.name = FONT_ARIAL; p.font.size = Pt(14); p.font.bold = True; p.font.color.rgb = COLOR_RED
    p = tfb.add_paragraph()
    p.text = "Zero-Trust მიდგომა Prompt Engineering-ში: ნებისმიერი დინამიკური ცვლადი პოტენციური შეტევის ვექტორია; ვალიდაცია, ტიპიზაცია და სანიტარიზაცია უნდა მოხდეს აპლიკაციის დონეზე, მოდელთან მიმართვამდე!"
    p.font.name = FONT_ARIAL; p.font.size = Pt(13); p.font.color.rgb = COLOR_MUTED; p.space_before = Pt(4)

    # -------------------------------------------------------------------------
    # SLIDE 5: Enterprise Prompt Registry (Prompt-as-Code)
    # -------------------------------------------------------------------------
    s5 = created_slides[3]
    add_header(s5, "Enterprise Prompt Registry: პრომპტები როგორც კოდი (Prompt-as-Code)", "ᲒᲣᲜᲓᲣᲠᲘ ᲡᲢᲐᲜᲓᲐᲠᲢᲘ • PROMPT REGISTRY & COLLABORATION")

    steps_s5 = [
        ("🗄️ REPOSITORY", "Prompt-as-Code", "ცენტრალიზებული ბაზა", [
            ("ერთიანი საცავი", "ყველა საწარმოო პრომპტი ინახება ცენტრალიზებულ საცავში (Git ან გამოყოფილი Registry)."),
            ("ქაოსის აღკვეთა", "აღარანაირი გაფანტული ჩანაწერები დეველოპერების პირად ჩანაწერებსა თუ პირად ფაილებში."),
            ("კოდთან ინტეგრაცია", "პრომპტების ჩატვირთვა ხდება SDK-ის ან API-ის მეშვეობით დინამიკურად.")
        ]),
        ("🏷️ METADATA & GOV", "მეტამონაცემები", "მოდელების თავსებადობა", [
            ("სრული დოკუმენტაცია", "თითოეულ შაბლონს თან ახლავს: ავტორი, შექმნის თარიღი, ბიზნეს-დანიშნულება."),
            ("თავსებადობის მატრიცა", "მითითებულია სამიზნე frontier მოდელები (GPT-6 Astra, Claude 5.x, Gemini 4)."),
            ("პარამეტრების სქემა", "განსაზღვრულია სავალდებულო ცვლადების ზუსტი სია და ტიპები.")
        ]),
        ("🔍 DISCOVERY & REUSE", "ძიება და გაზიარება", "დუბლირების აღმოფხვრა", [
            ("გუნდური ხელმისაწვდომობა", "ინჟინრები და ანალიტიკოსები მარტივად პოულობენ უკვე აპრობირებულ შაბლონებს."),
            ("კომპონენტების რეუზაბილურობა", "Reusability საგრძნობლად აჩქარებს ახალი AI ფუნქციონალის დანერგვას."),
            ("ხარისხის სტანდარტი", "გუნდი იყენებს მხოლოდ შემოწმებულ და ვალიდირებულ ფორმულირებებს.")
        ]),
        ("🔒 ACCESS CONTROL", "უსაფრთხოება და უფლებები", "გარემოების იზოლაცია", [
            ("RBAC მოდელი", "Role-Based Access Control – ვინ ფლობს პროდაქშენ პრომპტის შეცვლის უფლებას."),
            ("გარემოების გამიჯვნა", "Development, Staging და Production პრომპტების მკაცრი იზოლაცია."),
            ("Audit Log", "ყველა ცვლილებისა და წვდომის სრული ისტორიული ჟურნალიზაცია.")
        ])
    ]

    for i, (badge, title, subtitle, items) in enumerate(steps_s5):
        bg_col = COLOR_CARD_RED if i == 0 else COLOR_BG_LIGHT
        bor_col = COLOR_BORDER_RED if i == 0 else COLOR_BORDER
        c = s5.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, xs_4[i], step_y, step_w, step_h)
        c.fill.solid(); c.fill.fore_color.rgb = bg_col
        c.line.color.rgb = bor_col; c.line.width = Pt(1.5)

        tb = s5.shapes.add_textbox(xs_4[i] + Inches(0.25), step_y + Inches(0.25), step_w - Inches(0.5), step_h - Inches(0.5))
        tf = tb.text_frame; tf.word_wrap = True; tf.margin_left = tf.margin_top = 0

        p = tf.paragraphs[0]; p.text = badge.upper()
        p.font.name = FONT_ARIAL; p.font.size = Pt(11); p.font.bold = True
        p.font.color.rgb = COLOR_RED if i == 0 else COLOR_MUTED

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
    p.text = "Enterprise Prompt Registry პრომპტების მართვას გარდაქმნის პროგრამული უზრუნველყოფის სტანდარტულ სასიცოცხლო ციკლად: კოდის რევიზია, წვდომის კონტროლი და ცენტრალიზებული საცავი გამორიცხავს გაუთვალისწინებელ რისკებს."
    p.font.name = FONT_ARIAL; p.font.size = Pt(13); p.font.color.rgb = COLOR_MUTED; p.space_before = Pt(4)

    # -------------------------------------------------------------------------
    # SLIDE 6: Versioning & CI/CD for Prompts (SemVer & Golden Datasets)
    # -------------------------------------------------------------------------
    s6 = created_slides[4]
    add_header(s6, "შაბლონების ვერსიონირება (SemVer) და ავტომატური რეგრესიული ტესტირება", "DevOps & AI • SEMANTIC VERSIONING & CI/CD")

    cards_s6 = [
        ("🏷️ SEMANTIC VERSIONING", "SemVer პრომპტებში", "ცვლილებების აღრიცხვა", [
            ("Major (v2.0.0)", "სისტემური როლის რადიკალური შეცვლა, გამოსავლის JSON სქემის შეცვლა (Breaking Change)."),
            ("Minor (v1.1.0)", "ახალი ცვლადის ან დამატებითი შეზღუდვის დამატება არსებული სქემის შენარჩუნებით."),
            ("Patch (v1.0.1)", "ტექსტური ფორმულირების დახვეწა, მცირე ორთოგრაფიული ან სტილისტური კორექტირება.")
        ]),
        ("🔀 GIT WORKFLOWS", "Pull Requests & Review", "თანამშრომლობის კულტურა", [
            ("Peer Review", "პრომპტის ცვლილება გადის სავალდებულო კოდის რევიზიას (Code Review) უფროს ინჟინერთან."),
            ("Commit History & Context", "დეტალური ჩანაწერი, თუ რატომ შეიცვალა კონკრეტული ინსტრუქცია და რა ბიზნეს-ქეისი გადაიჭრა."),
            ("Instant Rollback", "ნებისმიერ დროს შესაძლებელია წინა სტაბილურ ვერსიაზე მყისიერი დაბრუნება (Rollback).")
        ]),
        ("🧪 CI/CD EVALUATION", "ავტომატური ტესტები", "რეგრესიისგან დაცვა", [
            ("Golden Evaluation Set", "50-100 ეტალონური ქეისის ავტომატური გაშვება ახალი ვერსიის შემოწმებისას CI მილსადენში."),
            ("Threshold Checks", "ახალი შაბლონის დაშვება პროდაქშენში მხოლოდ მაშინ, თუ ხარისხის მეტრიკა არ გაუარესდა."),
            ("Model Drift Protection", "რეგულარული ტესტირება ადასტურებს, რომ მოდელის განახლებამ არ დააზიანა შაბლონის მუშაობა.")
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
    p = tfb.paragraphs[0]; p.text = "💡 მთავარი არსი:"
    p.font.name = FONT_ARIAL; p.font.size = Pt(14); p.font.bold = True; p.font.color.rgb = COLOR_RED
    p = tfb.add_paragraph()
    p.text = "პრომპტის შეცვლა პროდაქშენში ისეთივე საპასუხისმგებლო ოპერაციაა, როგორც მონაცემთა ბაზის მიგრაცია; SemVer და ავტომატური Golden Set ტესტები გამორიცხავს გაუთვალისწინებელ რეგრესიებს."
    p.font.name = FONT_ARIAL; p.font.size = Pt(13); p.font.color.rgb = COLOR_MUTED; p.space_before = Pt(4)

    # -------------------------------------------------------------------------
    # SLIDE 7: Modern Tech Stack (Engines, CMS & Observability 2026)
    # -------------------------------------------------------------------------
    s7 = created_slides[5]
    add_header(s7, "თანამედროვე ტექნოლოგიური დასტა: Template Engines და AI პლატფორმები", "ᲢᲔᲥᲜᲝᲚᲝᲒᲘᲣᲠᲘ ᲡᲢᲔᲙᲘ • ENGINES & PLATFORMS 2026")

    stack_s7 = [
        ("⚙️ TEMPLATE ENGINES", "შაბლონიზატორები", "პროგრამული ბიბლიოთეკები", [
            ("Jinja2 (Python) & Liquid (JS)", "ყველაზე მძლავრი და გავრცელებული ღია ძრავები პირობითი ლოგიკისა და ციკლების მხარდაჭერით."),
            ("LangChain PromptTemplate", "მზა აბსტრაქცია პარამეტრების ავტომატური შემოწმებით, სერიალიზაციითა და კომპოზიციით."),
            ("LlamaIndex Prompts", "მონაცემთა ჩარჩოებთან (RAG) და დოკუმენტებთან მუშაობაზე ოპტიმიზებული შაბლონური სისტემა.")
        ]),
        ("🏢 ENTERPRISE PLATFORMS", "Prompt CMS & Observability", "2026 წლის წამყვანი პლატფორმები", [
            ("Langfuse", "Open-source Prompt Management, ვერსიონირება, API წვდომა და Observability ტრასირება ერთ სივრცეში."),
            ("Promptfoo", "ავტომატური CI/CD ტესტირება, უსაფრთხოების წითელი გუნდის (Red-teaming) შემოწმება და შეფასების მატრიცა."),
            ("Agenta & Humanloop", "პრომპტების ერთობლივი რედაქტორი ინჟინრებისა და დომენის ექსპერტებისთვის UI ინტერფეისით.")
        ]),
        ("🔀 CONDITIONAL LOGIC", "დინამიკური ლოგიკა", "პირობითი ინსტრუქციები", [
            ("Conditional Blocks", "{% if user.is_enterprise %} – დამატებითი მკაცრი SLA წესებისა და ტონის ჩასმა."),
            ("Loops & Collections", "{% for item in order.items %} – კატალოგის ან კალათის პროდუქტების დინამიკური ფორმატირება."),
            ("Macro Components", "ხშირად გამოყენებადი სისტემური ფრაგმენტების (მაგ. უარის თქმის წესების) ცალკე მოდულებად გატანა.")
        ])
    ]

    for i, (badge, title, subtitle, items) in enumerate(stack_s7):
        bg_col = COLOR_CARD_RED if i == 1 else COLOR_BG_LIGHT
        bor_col = COLOR_BORDER_RED if i == 1 else COLOR_BORDER
        c = s7.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, xs_3[i], card_y, card_w, card_h)
        c.fill.solid(); c.fill.fore_color.rgb = bg_col
        c.line.color.rgb = bor_col; c.line.width = Pt(1.5)

        tb = s7.shapes.add_textbox(xs_3[i] + Inches(0.35), card_y + Inches(0.35), card_w - Inches(0.7), card_h - Inches(0.7))
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
    bb = s7.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(1.0), bot_y, Inches(17.0), Inches(2.0))
    bb.fill.solid(); bb.fill.fore_color.rgb = COLOR_WHITE
    bb.line.color.rgb = COLOR_BORDER; bb.line.width = Pt(1.2)

    tbb = s7.shapes.add_textbox(Inches(1.4), bot_y + Inches(0.3), Inches(16.2), Inches(1.4))
    tfb = tbb.text_frame; tfb.word_wrap = True; tfb.margin_left = tfb.margin_top = 0
    p = tfb.paragraphs[0]; p.text = "🎯 საინჟინრო პრინციპი:"
    p.font.name = FONT_ARIAL; p.font.size = Pt(14); p.font.bold = True; p.font.color.rgb = COLOR_RED
    p = tfb.add_paragraph()
    p.text = "2026 წელს პრომპტების მართვა გასცდა მარტივ ტექსტურ ფაილებს; დღეს საწარმოო სტანდარტია სპეციალიზებული Prompt CMS (Langfuse, Promptfoo, Agenta), რომელიც API-ით უკავშირდება CI/CD მილსადენს."
    p.font.name = FONT_ARIAL; p.font.size = Pt(13); p.font.color.rgb = COLOR_MUTED; p.space_before = Pt(4)

    # -------------------------------------------------------------------------
    # SLIDE 8: Comparative Matrix (Hardcoded Prompts vs Enterprise Templates)
    # -------------------------------------------------------------------------
    s8 = created_slides[6]
    add_header(s8, "Hardcoded Prompts vs Enterprise Prompt Templates: არქიტექტურული შედარება", "ᲨᲔᲓᲐᲠᲔᲑᲘᲗᲘ ᲐᲜᲐᲚᲘᲖᲘ • ARCHITECTURAL COMPARISON")

    # Table: 6 rows, 5 columns
    tbl_left = Inches(1.0); tbl_top = Inches(2.35); tbl_w = Inches(17.0); tbl_h = Inches(5.2)
    tbl_shape = s8.shapes.add_table(6, 5, tbl_left, tbl_top, tbl_w, tbl_h)
    table = tbl_shape.table

    col_widths = [Inches(2.8), Inches(3.5), Inches(3.7), Inches(3.4), Inches(3.6)]
    for ci, w in enumerate(col_widths):
        table.columns[ci].width = w

    headers = ["კრიტერიუმი / პარამეტრი", "Hardcoded Prompt (საწყისი)", "Prompt Template (საწარმოო)", "რისკი / ნეგატიური გავლენა", "2026 წლის რეკომენდაცია"]
    for ci, h in enumerate(headers):
        cell = table.cell(0, ci)
        cell.fill.solid(); cell.fill.fore_color.rgb = COLOR_DARK
        tf = cell.text_frame; tf.word_wrap = True; tf.margin_left = tf.margin_right = Inches(0.12)
        p = tf.paragraphs[0]; p.text = h; p.font.name = FONT_ARIAL; p.font.size = Pt(12); p.font.bold = True; p.font.color.rgb = COLOR_WHITE

    matrix_rows = [
        ("ცვლილების სიჩქარე (Agility)", "მოითხოვს კოდის ხელახალ Build & Deploy ციკლს", "მყისიერი განახლება Registry API-ით კოდის შეცვლის გარეშე", "სისტემის დაყოვნება და დეველოპერული რესურსის კარგვა", "დინამიკური Prompt Registry API-ის გამოყენება"),
        ("ცვლადების უსაფრთხოება", "სტრიქონების პირდაპირი მიერთება, მაღალი Injection რისკი", "წინასწარი სანიტარიზაცია, XML ტეგირება და ტიპების შემოწმება", "კრიტიკული უსაფრთხოების ხვრელი და მონაცემთა გაჟონვა", "მკაცრი Delimiter Isolation და ტიპობრივი შემოწმება"),
        ("გუნდური კოლაბორაცია", "გაფანტულია დეველოპერების ლოკალურ ფაილებში", "ცენტრალიზებული რეესტრი, ხელმისაწვდომი მთელი გუნდისთვის", "შრომის დუბლირება და არაერთგვაროვანი პასუხები", "Prompt-as-Code ერთიან რეესტრში Git კონტროლით"),
        ("ტესტირებადობა (Testing)", "სპონტანური, ხელით ტესტირება ჩატბოტის ინტერფეისში", "ავტომატური CI/CD ტესტები 50-100 Golden Case-ზე", "შეუმჩნეველი რეგრესია და პასუხების ხარისხის ვარდნა", "ავტომატური Regression Pipeline ტესტები ყოველ Build-ზე"),
        ("ვერსიონირება & Audit", "Git Commit-ებში არეული, რთულად აღსაქმელი ისტორია", "SemVer ვერსიები (v1.2.0), სრული ისტორია და მყისიერი Rollback", "უმართავი ცვლილებები და პრობლემის მიზეზის ძიების სირთულე", "სემანტიკური ვერსიონირება და Tracing ინტეგრაცია")
    ]

    for ri, row_data in enumerate(matrix_rows):
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
    bb = s8.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(1.0), bot_y, Inches(17.0), Inches(2.0))
    bb.fill.solid(); bb.fill.fore_color.rgb = COLOR_WHITE
    bb.line.color.rgb = COLOR_BORDER; bb.line.width = Pt(1.2)

    tbb = s8.shapes.add_textbox(Inches(1.4), bot_y + Inches(0.3), Inches(16.2), Inches(1.4))
    tfb = tbb.text_frame; tfb.word_wrap = True; tfb.margin_left = tfb.margin_top = 0
    p = tfb.paragraphs[0]; p.text = "💡 შედარების მთავარი დასკვნა:"
    p.font.name = FONT_ARIAL; p.font.size = Pt(14); p.font.bold = True; p.font.color.rgb = COLOR_RED
    p = tfb.add_paragraph()
    p.text = "Hardcoded მიდგომა მისაღებია მხოლოდ ერთჯერადი სწრაფი ექსპერიმენტებისთვის; ნებისმიერი საწარმოო აპლიკაცია მოითხოვს Prompt Templates-ს, რათა უზრუნველყოფილი იყოს უსაფრთხოება, ტესტირებადობა და მყისიერი განახლება."
    p.font.name = FONT_ARIAL; p.font.size = Pt(13); p.font.color.rgb = COLOR_MUTED; p.space_before = Pt(4)

    # -------------------------------------------------------------------------
    # SLIDE 9: Engineering Patterns vs Anti-Patterns
    # -------------------------------------------------------------------------
    s9 = created_slides[7]
    add_header(s9, "არქიტექტურული პატერნები და თავიდან ასაცილებელი ანტიპატერნები", "ᲡᲐᲣᲙᲔᲗᲔᲡᲝ ᲞᲠᲐᲥᲢᲘᲙᲐ • PATTERNS & ANTI-PATTERNS")

    box_w = Inches(8.35); box_h = Inches(5.4)

    # Recommended Patterns Card
    c_rec = s9.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(1.0), card_y, box_w, box_h)
    c_rec.fill.solid(); c_rec.fill.fore_color.rgb = COLOR_GREEN_BG
    c_rec.line.color.rgb = COLOR_GREEN_BOR; c_rec.line.width = Pt(1.5)

    tb_rec = s9.shapes.add_textbox(Inches(1.35), card_y + Inches(0.35), box_w - Inches(0.7), box_h - Inches(0.7))
    tf_rec = tb_rec.text_frame; tf_rec.word_wrap = True; tf_rec.margin_left = tf_rec.margin_top = 0
    p = tf_rec.paragraphs[0]; p.text = "✅ RECOMMENDED PATTERNS (საუკეთესო პატერნები)"
    p.font.name = FONT_ARIAL; p.font.size = Pt(13); p.font.bold = True; p.font.color.rgb = COLOR_GREEN_TEXT

    p = tf_rec.add_paragraph(); p.text = "სტაბილური და დაცული არქიტექტურა"
    p.font.name = FONT_ARIAL; p.font.size = Pt(16); p.font.bold = True; p.font.color.rgb = COLOR_DARK; p.space_before = Pt(4)

    rec_items = [
        ("Modular Blocks (მოდულური დაყოფა)", "დიდი პრომპტების დაყოფა დამოუკიდებელ მოდულებად (როლი, წესები, კონტექსტი)."),
        ("Strict Tag Wrapping (ტეგირება)", "ყველა ცვლადის მკაცრი შეფუთვა XML ტეგებში (<context>{{data}}</context>)."),
        ("Explicit Variable Contracts", "თითოეული ცვლადის ტიპისა და სავალდებულოობის დეკლარაცია Pydantic/Zod სქემით."),
        ("Graceful Fallback Defaults", "ნაგულისხმევი უსაფრთხო მნიშვნელობების განსაზღვრა არასავალდებულო ველებისთვის."),
        ("Metadata & Version Logging", "ყოველ API გამოძახებაზე გამოყენებული შაბლონის ვერსიისა და ჰეშის ლოგირება.")
    ]
    for head, body in rec_items:
        p = tf_rec.add_paragraph(); p.text = f"• {head}: {body}"
        p.font.name = FONT_ARIAL; p.font.size = Pt(11); p.font.color.rgb = COLOR_DARK; p.space_before = Pt(6)

    # Anti-Patterns Card
    c_anti = s9.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(9.65), card_y, box_w, box_h)
    c_anti.fill.solid(); c_anti.fill.fore_color.rgb = COLOR_CARD_RED
    c_anti.line.color.rgb = COLOR_BORDER_RED; c_anti.line.width = Pt(1.5)

    tb_anti = s9.shapes.add_textbox(Inches(10.0), card_y + Inches(0.35), box_w - Inches(0.7), box_h - Inches(0.7))
    tf_anti = tb_anti.text_frame; tf_anti.word_wrap = True; tf_anti.margin_left = tf_anti.margin_top = 0
    p = tf_anti.paragraphs[0]; p.text = "❌ DANGEROUS ANTI-PATTERNS (კრიტიკული შეცდომები)"
    p.font.name = FONT_ARIAL; p.font.size = Pt(13); p.font.bold = True; p.font.color.rgb = COLOR_RED

    p = tf_anti.add_paragraph(); p.text = "მოერიდეთ ამ პრაქტიკებს პროდაქშენში"
    p.font.name = FONT_ARIAL; p.font.size = Pt(16); p.font.bold = True; p.font.color.rgb = COLOR_DARK; p.space_before = Pt(4)

    anti_items = [
        ("Monolithic Spaghetti Prompts", "ერთი გიგანტური 500-სტრიქონიანი შაბლონი, სადაც ერთად წერია ათობით სხვადასხვა ამოცანა."),
        ("Raw String Interpolation", "მომხმარებლის შემავალი ტექსტის პირდაპირი კონკატენაცია ვალიდაციისა და ტეგების გარეშე."),
        ("Untyped Missing Variables", "აპლიკაციის მოულოდნელი გაჩერება (Crash), როდესაც ცვლადი None ან Undefined-ია."),
        ("Shadow Updates (ჩუმი ცვლილებები)", "პრომპტის შეცვლა პროდაქშენ გარემოში ვერსიის გაზრდისა და ტესტების გარეშე."),
        ("JSON Schema Conflicts", "f-string-ების უდიერი გამოყენება JSON სქემებთან ორმაგი ფრჩხილების ეკრანირების გარეშე.")
    ]
    for head, body in anti_items:
        p = tf_anti.add_paragraph(); p.text = f"• {head}: {body}"
        p.font.name = FONT_ARIAL; p.font.size = Pt(11); p.font.color.rgb = COLOR_DARK; p.space_before = Pt(6)

    # Bottom takeaway & transition to 4.2
    bb = s9.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(1.0), bot_y, Inches(17.0), Inches(2.0))
    bb.fill.solid(); bb.fill.fore_color.rgb = COLOR_WHITE
    bb.line.color.rgb = COLOR_BORDER; bb.line.width = Pt(1.2)

    tbb = s9.shapes.add_textbox(Inches(1.4), bot_y + Inches(0.25), Inches(16.2), Inches(1.5))
    tfb = tbb.text_frame; tfb.word_wrap = True; tfb.margin_left = tfb.margin_top = 0
    p = tfb.paragraphs[0]; p.text = "💡 საინჟინრო შეჯამება და გადასვლა ლექციაზე 4.2:"
    p.font.name = FONT_ARIAL; p.font.size = Pt(14); p.font.bold = True; p.font.color.rgb = COLOR_RED
    p = tfb.add_paragraph()
    p.text = "Prompt Templates აქცევს AI-ს სპონტანური ექსპერიმენტებიდან საიმედო პროგრამულ კომპონენტად. შემდეგ ლექციაში (4.2) შევისწავლით: როდის კმარა Zero-Shot მიდგომა და რა ეტაპზე იწყება მისი დეგრადაცია, როგორ ავაწყოთ V1 სისტემური პრომპტის არქიტექტურა მომხმარებლის მხარდაჭერის ასისტენტისთვის და მოდელის თვითშემოწმების (Self-Evaluation) 6-პუნქტიანი ჩეკლისტი."
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
    filename = "4.1 - Prompt Templates, პარამეტრიზაცია და გუნდური სტანდარტიზაცია.pptx"
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
    build_lecture_m7_4_1()
