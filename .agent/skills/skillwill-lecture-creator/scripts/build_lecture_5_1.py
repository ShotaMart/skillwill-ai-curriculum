import os
import sys
import shutil
import pptx
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN
from pptx.enum.shapes import MSO_SHAPE

def build_lecture_5_1():
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
    COLOR_BLUE = RGBColor(0x1A, 0x73, 0xE8)
    COLOR_BLUE_BG = RGBColor(0xE8, 0xF0, 0xFE)
    COLOR_GREEN = RGBColor(0x13, 0x73, 0x33)
    COLOR_GREEN_BG = RGBColor(0xE6, 0xF4, 0xEA)

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
            p.text = "5.1 - Claude სრულად II: Projects, Memory, Custom Instructions და Connectors"
            p.alignment = PP_ALIGN.CENTER
            p.font.name = FONT_ARIAL
            p.font.size = Pt(36)
            p.font.bold = True
            p.font.color.rgb = COLOR_RED

            p2 = tf.add_paragraph()
            p2.text = "პერსონალიზებული სამუშაო სივრცეები, ცოდნის ბაზა და გარე ინტეგრაციები"
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
    def add_header(slide, title, category="ᲛᲝᲓᲣᲚᲘ 5 • CLAUDE ეკოსისტემა • 5.1"):
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
    # SLIDE 2: Standalone Chat vs Claude Projects
    # -------------------------------------------------------------------------
    s2 = created_slides[0]
    add_header(s2, "ჩვეულებრივი ჩატი vs Claude Projects (იზოლირებული სივრცეები)", "კონტექსტის მართვა • რატომ აღარ არის საკმარისი ერთიანი ჩატის ისტორია?")

    box_w = Inches(8.35); box_h = Inches(5.4); box_y = Inches(2.35)

    # Left: Standalone Chat
    b1 = s2.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(1.0), box_y, box_w, box_h)
    b1.fill.solid(); b1.fill.fore_color.rgb = COLOR_BG_LIGHT
    b1.line.color.rgb = COLOR_BORDER; b1.line.width = Pt(1.5)

    tb1 = s2.shapes.add_textbox(Inches(1.4), box_y + Inches(0.4), box_w - Inches(0.8), box_h - Inches(0.8))
    tf1 = tb1.text_frame; tf1.word_wrap = True; tf1.margin_left = tf1.margin_top = 0
    p = tf1.paragraphs[0]; p.text = "ᲔᲤᲔᲛᲔᲠᲣᲚᲘ ᲡᲐᲣᲑᲐᲠᲘ"
    p.font.name = FONT_ARIAL; p.font.size = Pt(12); p.font.bold = True; p.font.color.rgb = COLOR_MUTED

    p = tf1.add_paragraph(); p.text = "ჩვეულებრივი ჩატი (Standalone Chat)"
    p.font.name = FONT_ARIAL; p.font.size = Pt(22); p.font.bold = True; p.font.color.rgb = COLOR_DARK; p.space_before = Pt(4)

    chat_points = [
        ("Ephemeral Context (დროებითი კონტექსტი)", "ყოველი ახალი ჩატის გახსნისას Claude იწყებს „სუფთა ფურცლიდან“ და არ ახსოვს წინა საუბრების დეტალები."),
        ("განმეორებითი ატვირთვები", "საჭირო ფაილების, კოდისა და ინსტრუქციების ხელით ატვირთვა გიწევთ ყოველ ჯერზე თავიდან."),
        ("Context Sprawl (კონტექსტის გაბნევა)", "სხვადასხვა თემისა და დავალების ერთ ჩატში შერევა იწვევს მოდელის ყურადღების გაფანტვას (Attention Degradation)."),
        ("ინდივიდუალური გამოყენება", "საუბარი ეკუთვნის მხოლოდ ერთ მომხმარებელს, რთულია მისი სისტემატიზაცია და გუნდური გამოყენება.")
    ]
    for t, d in chat_points:
        p = tf1.add_paragraph(); p.text = f"• {t}: {d}"
        p.font.name = FONT_ARIAL; p.font.size = Pt(13); p.font.color.rgb = COLOR_TEXT_MUTED; p.space_before = Pt(8)

    # Right: Dedicated Projects
    b2 = s2.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(9.65), box_y, box_w, box_h)
    b2.fill.solid(); b2.fill.fore_color.rgb = COLOR_CARD_RED
    b2.line.color.rgb = COLOR_BORDER_RED; b2.line.width = Pt(1.5)

    tb2 = s2.shapes.add_textbox(Inches(10.05), box_y + Inches(0.4), box_w - Inches(0.8), box_h - Inches(0.8))
    tf2 = tb2.text_frame; tf2.word_wrap = True; tf2.margin_left = tf2.margin_top = 0
    p = tf2.paragraphs[0]; p.text = "ᲡᲞᲔᲪᲘᲐᲚᲘᲖᲔᲑᲣᲚᲘ ᲡᲐᲛᲣᲨᲐᲝ ᲒᲐᲠᲔᲛᲝ"
    p.font.name = FONT_ARIAL; p.font.size = Pt(12); p.font.bold = True; p.font.color.rgb = COLOR_RED

    p = tf2.add_paragraph(); p.text = "Claude Projects (Dedicated Workspace)"
    p.font.name = FONT_ARIAL; p.font.size = Pt(22); p.font.bold = True; p.font.color.rgb = COLOR_DARK; p.space_before = Pt(4)

    proj_points = [
        ("Persistent Knowledge (მუდმივი ცოდნის ბაზა)", "ერთხელ ატვირთული დოკუმენტები, გაიდლაინები და ცხრილები ხელმისაწვდომია პროექტის ყველა ჩატისთვის."),
        ("Custom Project Instructions (სპეციფიკური როლი)", "პროექტს აქვს საკუთარი ქცევის წესები, ტონი და ფორმატი, რომელსაც მოდელი ყოველთვის მიჰყვება."),
        ("საუბრების სტრუქტურირება", "პროექტში შეგიძლიათ შექმნათ ათობით პარალელური ჩატი სხვადასხვა ქვეამოცანაზე, საერთო ცოდნის დაკარგვის გარეშე."),
        ("გუნდური გაზიარება (Team Sharing)", "შესაძლებელია პროექტის ცოდნისა და ინსტრუქციების მთელ გუნდზე გაზიარება ერთი დაწკაპუნებით.")
    ]
    for t, d in proj_points:
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
    p.text = "თუ ჩვეულებრივი ჩატი ჰგავს სწრაფ სატელეფონო ზარს, სადაც ყოველ ჯერზე თავიდან უნდა აუხსნათ ვინ ხართ და რა გინდათ; Claude Project არის გამოყოფილი სამუშაო კაბინეტი საკუთარი კარადით (დოკუმენტები), სამუშაო წესებითა და საერთო სამუშაო მაგიდით."
    p.font.name = FONT_ARIAL; p.font.size = Pt(13); p.font.color.rgb = COLOR_MUTED; p.space_before = Pt(4)

    # -------------------------------------------------------------------------
    # SLIDE 3: Project Knowledge - პროექტის ცოდნის ბაზა
    # -------------------------------------------------------------------------
    s3 = created_slides[1]
    add_header(s3, "Project Knowledge (პროექტის ცოდნის ბაზა და ფაილების მართვა)", "არქიტექტურა • როგორ მუშაობს დოკუმენტების ჩატვირთვა პროექტში")

    card_w = Inches(5.4); card_h = Inches(5.4); card_y = Inches(2.35)
    xs = [Inches(1.0), Inches(6.8), Inches(12.6)]

    cards_s3 = [
        ("📁 მასშტაბი", "ფაილების ტიპები და მოცულობა", "Document Ingestion", [
            ("მრავალფეროვანი ფორმატები", "PDF, DOCX, TXT, CSV, JSON, Markdown და კოდის სრული რეპოზიტორიები."),
            ("მასიური ტევადობა", "პროექტის ცოდნის ბაზაში შესაძლებელია ასობით დოკუმენტის (მილიონობით ტოკენის) ერთდროულად განთავსება."),
            ("ავტომატური სტრუქტურირება", "Claude ავტომატურად ინდექსირებს ატვირთულ ფაილებს და მათზე მყისიერ წვდომას ინარჩუნებს.")
        ]),
        ("🔍 ამოღება", "RAG და კონტექსტის ფანჯარა", "Context Retrieval", [
            ("Knowledge Retrieval (ცოდნის ამოღება)", "კითხვის დასმისას მოდელი პოულობს შესაბამის მონაკვეთებს პროექტის ფაილებიდან და აწვდის კონტექსტს."),
            ("Context Window Allocation", "პროექტის ფაილები ოპტიმიზებულად იკავებს ადგილს კონტექსტის ფანჯარაში, რათა დარჩეს სივრცე საუბრისთვის."),
            ("მაღალი სიზუსტე", "მოდელი პასუხობს კონკრეტულ ციტატებზე დაყრდნობით, რაც ჰალუცინაციებს მინიმუმამდე ამცირებს.")
        ]),
        ("⚡ სისწრაფე", "Prompt Caching ტექნოლოგია", "Performance & Latency", [
            ("Prompt Caching (პრომპტის ქეშირება)", "ატვირთული ცოდნის ბაზა ინახება ქეშში, რაც გამორიცხავს ფაილების ხელახალ გადამუშავებას ყოველ რეპლიკაზე."),
            ("მყისიერი პასუხები", "პასუხის დაწყების დრო (Time to First Token) მცირდება 2-3-ჯერ, დიდი მოცულობის დოკუმენტების მიუხედავად."),
            ("რესურსების ეკონომია", "ქეშირებული ტოკენების დამუშავება მოითხოვს გაცილებით ნაკლებ გამოთვლით სიმძლავრეს.")
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
    p = tfb.paragraphs[0]; p.text = "💡 საუკეთესო პრაქტიკა (Best Practice):"
    p.font.name = FONT_ARIAL; p.font.size = Pt(14); p.font.bold = True; p.font.color.rgb = COLOR_RED
    p = tfb.add_paragraph()
    p.text = "პროექტში ატვირთეთ მხოლოდ კონკრეტული ამოცანის მარეგულირებელი, მეთოდოლოგიური და საკვანძო დოკუმენტები. ზედმეტი, უსარგებლო ინფორმაციის ატვირთვა ფანტავს მოდელის ყურადღებას და აუარესებს პასუხის სიზუსტეს."
    p.font.name = FONT_ARIAL; p.font.size = Pt(13); p.font.color.rgb = COLOR_MUTED; p.space_before = Pt(4)

    # -------------------------------------------------------------------------
    # SLIDE 4: Custom Instructions & Project Instructions
    # -------------------------------------------------------------------------
    s4 = created_slides[2]
    add_header(s4, "Custom Instructions და Project Instructions (ქცევის წესები)", "პერსონალიზაცია • როგორ ვასწავლით Claude-ს ჩვენს როლსა და მოთხოვნებს")

    # Left: Account-Level Custom Instructions
    b1 = s4.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(1.0), box_y, box_w, box_h)
    b1.fill.solid(); b1.fill.fore_color.rgb = COLOR_BG_LIGHT
    b1.line.color.rgb = COLOR_BORDER; b1.line.width = Pt(1.5)

    tb1 = s4.shapes.add_textbox(Inches(1.4), box_y + Inches(0.4), box_w - Inches(0.8), box_h - Inches(0.8))
    tf1 = tb1.text_frame; tf1.word_wrap = True; tf1.margin_left = tf1.margin_top = 0
    p = tf1.paragraphs[0]; p.text = "ᲐᲜᲒᲐᲠᲘᲨᲘᲡ ᲓᲝᲜᲔ"
    p.font.name = FONT_ARIAL; p.font.size = Pt(12); p.font.bold = True; p.font.color.rgb = COLOR_MUTED

    p = tf1.add_paragraph(); p.text = "Account-Level Instructions (გლობალური წესები)"
    p.font.name = FONT_ARIAL; p.font.size = Pt(21); p.font.bold = True; p.font.color.rgb = COLOR_DARK; p.space_before = Pt(4)

    acc_points = [
        ("უნივერსალური მოქმედება", "მოქმედებს თქვენი ანგარიშის ყველა ახალ ჩატზე, ნებისმიერ პროექტში თუ პროექტს გარეთ."),
        ("მომხმარებლის პროფილი", "ვინ ხართ თქვენ (პროფესია, კომპეტენცია, როლი), რა ენაზე გირჩევნიათ პასუხების მიღება."),
        ("სტილისტური პრეფერენციები", "პასუხების სიგრძე, ტონი (ლაკონიური, აკადემიური, პირდაპირი), ფორმატირების წესები."),
        ("ტიპური მაგალითები", "„ყოველთვის მიპასუხე ქართულად“, „კოდის წერისას გამოიყენე თანამედროვე TypeScript და Tailwind CSS“.")
    ]
    for t, d in acc_points:
        p = tf1.add_paragraph(); p.text = f"• {t}: {d}"
        p.font.name = FONT_ARIAL; p.font.size = Pt(13); p.font.color.rgb = COLOR_TEXT_MUTED; p.space_before = Pt(8)

    # Right: Project-Level Custom Instructions
    b2 = s4.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(9.65), box_y, box_w, box_h)
    b2.fill.solid(); b2.fill.fore_color.rgb = COLOR_CARD_RED
    b2.line.color.rgb = COLOR_BORDER_RED; b2.line.width = Pt(1.5)

    tb2 = s4.shapes.add_textbox(Inches(10.05), box_y + Inches(0.4), box_w - Inches(0.8), box_h - Inches(0.8))
    tf2 = tb2.text_frame; tf2.word_wrap = True; tf2.margin_left = tf2.margin_top = 0
    p = tf2.paragraphs[0]; p.text = "ᲞᲠᲝᲔᲥᲢᲘᲡ ᲓᲝᲜᲔ"
    p.font.name = FONT_ARIAL; p.font.size = Pt(12); p.font.bold = True; p.font.color.rgb = COLOR_RED

    p = tf2.add_paragraph(); p.text = "Project-Level Instructions (სპეციფიკური როლი)"
    p.font.name = FONT_ARIAL; p.font.size = Pt(21); p.font.bold = True; p.font.color.rgb = COLOR_DARK; p.space_before = Pt(4)

    pinst_points = [
        ("იზოლირებული მოქმედება", "მოქმედებს მხოლოდ მოცემული პროექტის ფარგლებში შექმნილ ჩატებზე."),
        ("ვიწრო ექსპერტული როლი", "განსაზღვრავს მოდელის სპეციფიკურ როლს (მაგ. „შენ ხარ Skillwill-ის სასწავლო პროგრამების აუდიტორი“)."),
        ("მკაცრი შეზღუდვები (Negative Constraints)", "რა არ უნდა გააკეთოს მოდელმა (მაგ. „არ გამოიყენო ტექნიკური ჟარგონი განმარტების გარეშე“)."),
        ("Output Template (შედეგის ფორმატი)", "პასუხების მკაცრი სტრუქტურა (ცხრილები, ეტაპობრივი გეგმები, ციტატები წყაროდან).")
    ]
    for t, d in pinst_points:
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
    p = tfb.paragraphs[0]; p.text = "💡 ეფექტური პროექტის ინსტრუქციის ფორმულა:"
    p.font.name = FONT_ARIAL; p.font.size = Pt(14); p.font.bold = True; p.font.color.rgb = COLOR_RED
    p = tfb.add_paragraph()
    p.text = "იდეალური ინსტრუქცია მოიცავს 3 ნაწილს: (1) როლი და ამოცანა („შენ ხარ ფინანსური ანალიტიკოსი...“), (2) სამუშაო წყაროები („დაეყრდენი მხოლოდ Project Knowledge-ის ფაილებს...“), (3) გამოსატანი შედეგის სტანდარტი („პასუხი დაყავი SWOT ბლოკებად და წარმოადგინე Markdown ცხრილში“)."
    p.font.name = FONT_ARIAL; p.font.size = Pt(13); p.font.color.rgb = COLOR_MUTED; p.space_before = Pt(4)

    # -------------------------------------------------------------------------
    # SLIDE 5: Claude Memory - მეხსიერება ჩატებს შორის
    # -------------------------------------------------------------------------
    s5 = created_slides[3]
    add_header(s5, "Claude Memory (მეხსიერება და უწყვეტი პერსონალიზაცია)", "ფუნქციონალი • როგორ იმახსოვრებს მოდელი მომხმარებლის პრეფერენციებს")

    cards_s5 = [
        ("🧠 მეხსიერება", "რა არის Claude Memory?", "Cross-Chat Personalization", [
            ("ფაქტებისა და ჩვევების დამახსოვრება", "მოდელს შეუძლია ავტომატურად დაიმახსოვროს მნიშვნელოვანი დეტალები საუბრის პროცესში."),
            ("პირდაპირი ბრძანებები", "მომხმარებელს შეუძლია პირდაპირ უთხრას: „დაიმახსოვრე, რომ ჩემი კომპანიის სახელია X და ჩვენი ფერებია Y“."),
            ("უწყვეტი კონტექსტი", "ახალ ჩატში გადასვლისას მოდელს ახსოვს თქვენი როლი და საჭიროებები თავიდან ახსნის გარეშე.")
        ]),
        ("⚙️ მართვა", "მომხმარებლის სრული კონტროლი", "Transparency & Control", [
            ("გამჭვირვალე სია (Memory List)", "პარამეტრებში ჩანს ყველა ფაქტი, რაც Claude-ს თქვენ შესახებ აქვს დამახსოვრებული."),
            ("რედაქტირება და წაშლა", "შეგიძლიათ ნებისმიერი ფაქტის წაშლა ან შესწორება ერთი დაწკაპუნებით."),
            ("ფუნქციის გამორთვა", "მომხმარებელს შეუძლია ნებისმიერ დროს სრულად გათიშოს მეხსიერება ან წაშალოს მთელი ისტორია.")
        ]),
        ("⚖️ შედარება", "Memory vs Project Knowledge", "სწორი გამიჯვნა", [
            ("Claude Memory", "მოკლე, პერსონალური ფაქტები და პრეფერენციები მომხმარებლის შესახებ (როლი, ტონი, ხშირად გამოყენებული ტექნოლოგიები)."),
            ("Project Knowledge", "მასიური, დოკუმენტური ცოდნის ბაზა კონკრეტული ბიზნეს ამოცანის ან პროექტისთვის."),
            ("კომბინირებული ძალა", "ორივეს ერთობლივი გამოყენება ქმნის მაქსიმალურად მორგებულ და ჭკვიან სამუშაო პარტნიორს.")
        ])
    ]

    for i, (badge, title, subtitle, items) in enumerate(cards_s5):
        bg_col = COLOR_CARD_RED if i == 0 else COLOR_BG_LIGHT
        bor_col = COLOR_BORDER_RED if i == 0 else COLOR_BORDER
        c = s5.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, xs[i], card_y, card_w, card_h)
        c.fill.solid(); c.fill.fore_color.rgb = bg_col
        c.line.color.rgb = bor_col; c.line.width = Pt(1.5)

        tb = s5.shapes.add_textbox(xs[i] + Inches(0.35), card_y + Inches(0.35), card_w - Inches(0.7), card_h - Inches(0.7))
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
    bb = s5.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(1.0), bot_y, Inches(17.0), Inches(2.0))
    bb.fill.solid(); bb.fill.fore_color.rgb = COLOR_WHITE
    bb.line.color.rgb = COLOR_BORDER; bb.line.width = Pt(1.2)

    tbb = s5.shapes.add_textbox(Inches(1.4), bot_y + Inches(0.3), Inches(16.2), Inches(1.4))
    tfb = tbb.text_frame; tfb.word_wrap = True; tfb.margin_left = tfb.margin_top = 0
    p = tfb.paragraphs[0]; p.text = "💡 მთავარი არსი (მარტივი ანალოგია):"
    p.font.name = FONT_ARIAL; p.font.size = Pt(14); p.font.bold = True; p.font.color.rgb = COLOR_RED
    p = tfb.add_paragraph()
    p.text = "Claude Memory არის თანაშემწის პირადი ბლოკნოტი, სადაც წერია თქვენი სამუშაო გრაფიკი, მოთხოვნები და ჩვევები, ხოლო Project Knowledge არის საარქივო კარადა კონკრეტული საქმის დოკუმენტაციით."
    p.font.name = FONT_ARIAL; p.font.size = Pt(13); p.font.color.rgb = COLOR_MUTED; p.space_before = Pt(4)

    # -------------------------------------------------------------------------
    # SLIDE 6: Connectors არქიტექტურა
    # -------------------------------------------------------------------------
    s6 = created_slides[4]
    add_header(s6, "Connectors არქიტექტურა (პირდაპირი კავშირი გარე სერვისებთან)", "ინტეგრაცია • ფაილების ხელით ატვირთვის ეპოქიდან პირდაპირ სინქრონიზაციამდე")

    # Left: Manual Uploads
    b1 = s6.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(1.0), box_y, box_w, box_h)
    b1.fill.solid(); b1.fill.fore_color.rgb = COLOR_BG_LIGHT
    b1.line.color.rgb = COLOR_BORDER; b1.line.width = Pt(1.5)

    tb1 = s6.shapes.add_textbox(Inches(1.4), box_y + Inches(0.4), box_w - Inches(0.8), box_h - Inches(0.8))
    tf1 = tb1.text_frame; tf1.word_wrap = True; tf1.margin_left = tf1.margin_top = 0
    p = tf1.paragraphs[0]; p.text = "ᲫᲕᲔᲚᲘ ᲛᲝᲓᲔᲚᲘ"
    p.font.name = FONT_ARIAL; p.font.size = Pt(12); p.font.bold = True; p.font.color.rgb = COLOR_MUTED

    p = tf1.add_paragraph(); p.text = "Manual File Upload (ფაილების ხელით ატვირთვა)"
    p.font.name = FONT_ARIAL; p.font.size = Pt(21); p.font.bold = True; p.font.color.rgb = COLOR_DARK; p.space_before = Pt(4)

    old_points = [
        ("ფრაგმენტული პროცესი", "მომხმარებელს უწევს ფაილის გადმოწერა კომპიუტერში და შემდეგ ხელით ატვირთვა Claude-ში."),
        ("მოძველებული მონაცემები", "თუ Google Docs-ში დოკუმენტი განახლდა, Claude-ში რჩება ძველი ვერსია, სანამ თავიდან არ ატვირთავთ."),
        ("დროის დაკარგვა", "მრავალი დოკუმენტის მართვა და სინქრონიზაცია მოითხოვს მუდმივ მექანიკურ შრომას."),
        ("ლოკალური შენახვის რისკები", "სენსიტიური ფაილების ლოკალურ დისკზე შენახვა ქმნის უსაფრთხოების დამატებით რისკებს.")
    ]
    for t, d in old_points:
        p = tf1.add_paragraph(); p.text = f"• {t}: {d}"
        p.font.name = FONT_ARIAL; p.font.size = Pt(13); p.font.color.rgb = COLOR_TEXT_MUTED; p.space_before = Pt(8)

    # Right: Native Connectors
    b2 = s6.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(9.65), box_y, box_w, box_h)
    b2.fill.solid(); b2.fill.fore_color.rgb = COLOR_CARD_RED
    b2.line.color.rgb = COLOR_BORDER_RED; b2.line.width = Pt(1.5)

    tb2 = s6.shapes.add_textbox(Inches(10.05), box_y + Inches(0.4), box_w - Inches(0.8), box_h - Inches(0.8))
    tf2 = tb2.text_frame; tf2.word_wrap = True; tf2.margin_left = tf2.margin_top = 0
    p = tf2.paragraphs[0]; p.text = "ᲗᲐᲜᲐᲛᲔᲓᲠᲝᲕᲔ ᲛᲝᲓᲔᲚᲘ (2026)"
    p.font.name = FONT_ARIAL; p.font.size = Pt(12); p.font.bold = True; p.font.color.rgb = COLOR_RED

    p = tf2.add_paragraph(); p.text = "Native Connectors (პირდაპირი API სინქრონიზაცია)"
    p.font.name = FONT_ARIAL; p.font.size = Pt(21); p.font.bold = True; p.font.color.rgb = COLOR_DARK; p.space_before = Pt(4)

    conn_points = [
        ("Direct Integration (პირდაპირი ინტეგრაცია)", "Claude პირდაპირ უკავშირდება თქვენს სამუშაო სერვისებს (Google Drive, Gmail, Notion)."),
        ("Live Data Access (ცოცხალი მონაცემები)", "მოდელი ყოველთვის კითხულობს დოკუმენტების უახლეს, ცოცხალ ვერსიას პირდაპირ ღრუბლიდან."),
        ("Secure OAuth 2.0 (დაცული არხი)", "წვდომა რეგულირდება საერთაშორისო უსაფრთხოების სტანდარტებით, პაროლების შენახვის გარეშე."),
        ("Granular Permissions (დეტალური უფლებები)", "თქვენ თავად ირჩევთ, რომელ საქაღალდეზე ან ეტიკეტზე (Label) მიანიჭოთ წვდომა.")
    ]
    for t, d in conn_points:
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
    p = tfb.paragraphs[0]; p.text = "💡 უსაფრთხოების სტანდარტი:"
    p.font.name = FONT_ARIAL; p.font.size = Pt(14); p.font.bold = True; p.font.color.rgb = COLOR_RED
    p = tfb.add_paragraph()
    p.text = "Connectors იყენებს მკაცრად შეზღუდულ წაკითხვის (Read-Only) უფლებებს. Claude არ ცვლის თქვენს ორიგინალ ფაილებს Google Drive-ში და არ აგზავნის წერილებს თქვენი ნებართვის გარეშე."
    p.font.name = FONT_ARIAL; p.font.size = Pt(13); p.font.color.rgb = COLOR_MUTED; p.space_before = Pt(4)

    # -------------------------------------------------------------------------
    # SLIDE 7: Google Drive და Gmail Connectors
    # -------------------------------------------------------------------------
    s7 = created_slides[5]
    add_header(s7, "Google Drive და Gmail Connectors (საოფისე ეკოსისტემა)", "პროდუქტიულობა • როგორ ვმუშაობთ დოკუმენტებსა და იმეილებზე Claude-ის დახმარებით")

    # Left: Google Drive Connector
    b1 = s7.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(1.0), box_y, box_w, box_h)
    b1.fill.solid(); b1.fill.fore_color.rgb = COLOR_BG_LIGHT
    b1.line.color.rgb = COLOR_BORDER; b1.line.width = Pt(1.5)

    tb1 = s7.shapes.add_textbox(Inches(1.4), box_y + Inches(0.4), box_w - Inches(0.8), box_h - Inches(0.8))
    tf1 = tb1.text_frame; tf1.word_wrap = True; tf1.margin_left = tf1.margin_top = 0
    p = tf1.paragraphs[0]; p.text = "ᲓᲝᲙᲣᲛᲔᲜᲢᲔᲑᲘᲡ ᲛᲐᲠᲗᲕᲐ"
    p.font.name = FONT_ARIAL; p.font.size = Pt(12); p.font.bold = True; p.font.color.rgb = COLOR_MUTED

    p = tf1.add_paragraph(); p.text = "Google Drive Connector (ღრუბლოვანი ფაილები)"
    p.font.name = FONT_ARIAL; p.font.size = Pt(21); p.font.bold = True; p.font.color.rgb = COLOR_DARK; p.space_before = Pt(4)

    gdrive_points = [
        ("უნივერსალური ფორმატების კითხვა", "Google Docs, Google Sheets, Google Slides და Drive-ში შენახული PDF ფაილები."),
        ("Cross-Document Search (ჭკვიანი ძებნა)", "Claude პოულობს საჭირო ინფორმაციას ათობით დოკუმენტში შინაარსობრივი კონტექსტის მიხედვით."),
        ("შედარებითი ანალიზი", "შეუძლია შეადაროს სხვადასხვა ვერსიის ანგარიშები ან ფინანსური ცხრილები და გამოავლინოს ცვლილებები."),
        ("პირდაპირი ციტირება", "პასუხში უთითებს კონკრეტულ დოკუმენტსა და აბზაცს, საიდანაც ამოიღო მონაცემი.")
    ]
    for t, d in gdrive_points:
        p = tf1.add_paragraph(); p.text = f"• {t}: {d}"
        p.font.name = FONT_ARIAL; p.font.size = Pt(13); p.font.color.rgb = COLOR_TEXT_MUTED; p.space_before = Pt(8)

    # Right: Gmail Connector
    b2 = s7.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(9.65), box_y, box_w, box_h)
    b2.fill.solid(); b2.fill.fore_color.rgb = COLOR_CARD_RED
    b2.line.color.rgb = COLOR_BORDER_RED; b2.line.width = Pt(1.5)

    tb2 = s7.shapes.add_textbox(Inches(10.05), box_y + Inches(0.4), box_w - Inches(0.8), box_h - Inches(0.8))
    tf2 = tb2.text_frame; tf2.word_wrap = True; tf2.margin_left = tf2.margin_top = 0
    p = tf2.paragraphs[0]; p.text = "ᲙᲝᲛᲣᲜᲘᲙᲐᲪᲘᲘᲡ ᲛᲐᲠᲗᲕᲐ"
    p.font.name = FONT_ARIAL; p.font.size = Pt(12); p.font.bold = True; p.font.color.rgb = COLOR_RED

    p = tf2.add_paragraph(); p.text = "Gmail Connector (სამუშაო მიმოწერა)"
    p.font.name = FONT_ARIAL; p.font.size = Pt(21); p.font.bold = True; p.font.color.rgb = COLOR_DARK; p.space_before = Pt(4)

    gmail_points = [
        ("Email Thread Summarization (თრედების შეჯამება)", "მრავალდღიანი და მრავალმომხმარებლიანი რთული მიმოწერის 2-აბზაციანი მოკლე რეზიუმე."),
        ("Action Items Extraction (ამოცანების ამოღება)", "წერილებიდან ავტომატურად აგენერირებს შესასრულებელი დავალებების ჩამონათვალს (To-Do List)."),
        ("Draft Responses (პასუხის დრაფტის მომზადება)", "ამზადებს გამართულ, პროფესიონალურ პასუხს Drive-ის დოკუმენტებზე დაყრდნობით."),
        ("ჭკვიანი ფილტრაცია", "გამოყოფს გადაუდებელ შეტყობინებებს მეორეხარისხოვანი საინფორმაციო წერილებისგან.")
    ]
    for t, d in gmail_points:
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
    p = tfb.paragraphs[0]; p.text = "💡 რეალური სამუშაო პროცესი (Unified Workflow):"
    p.font.name = FONT_ARIAL; p.font.size = Pt(14); p.font.bold = True; p.font.color.rgb = COLOR_RED
    p = tfb.add_paragraph()
    p.text = "დილის სცენარი: Claude აანალიზებს Gmail-ის ბოლო მიმოწერას ➔ Drive-ში პოულობს შესაბამის ბრიფს ➔ აგენერირებს პროექტის სტატუსის განახლებულ ანგარიშს Artifact-ის სახით."
    p.font.name = FONT_ARIAL; p.font.size = Pt(13); p.font.color.rgb = COLOR_MUTED; p.space_before = Pt(4)

    # -------------------------------------------------------------------------
    # SLIDE 8: Team Projects და კოლაბორაცია
    # -------------------------------------------------------------------------
    s8 = created_slides[6]
    add_header(s8, "Team Projects და კოლაბორაცია (გუნდური მუშაობის სივრცე)", "კოლაბორაცია • ცოდნის გაზიარება და გუნდური პროდუქტიულობა")

    cards_s8 = [
        ("📚 საერთო ცოდნა", "Shared Project Knowledge", "ერთიანი სტანდარტი", [
            ("სინქრონიზებული ცოდნის ბაზა", "გუნდის ყველა წევრი მუშაობს კომპანიის ერთიან სტანდარტებზე, გაიდლაინებსა და დოკუმენტებზე დაყრდნობით."),
            ("ინფორმაციის დუბლირების აღმოფხვრა", "აღარ არის საჭირო თითოეული თანამშრომლის მიერ ფაილების ცალ-ცალკე ატვირთვა."),
            ("მყისიერი განახლება", "ცოდნის ბაზის განახლებისთანავე ახალი ინფორმაცია ხელმისაწვდომია მთელი გუნდისთვის.")
        ]),
        ("🔒 პირადი ჩატი", "Private Chats in Shared Space", "კონფიდენციალურობა", [
            ("იზოლირებული საუბრები", "პროექტის შიგნით თითოეული მომხმარებლის ჩატი რჩება მკაცრად პირადი და მიუწვდომელი სხვებისთვის."),
            ("Selective Sharing (შერჩევითი გაზიარება)", "მომხმარებელს შეუძლია გააზიაროს კონკრეტული წარმატებული საუბარი (Share Chat) კოლეგებთან."),
            ("უსაფრთხო ექსპერიმენტები", "გუნდის წევრებს შეუძლიათ თავისუფლად მოსინჯონ განსხვავებული მიდგომები საერთო ცოდნის დაზიანების გარეშე.")
        ]),
        ("👥 წვდომები", "Role-Based Access Control", "ადმინისტრირება", [
            ("როლების მართვა (RBAC)", "პროექტის მფლობელები (Owners), რედაქტორები (Editors) და დამკვირვებლები (Viewers)."),
            ("ცოდნის ბაზის დაცვა", "მხოლოდ ავტორიზებულ პირებს შეუძლიათ პროექტის ინსტრუქციების შეცვლა ან დოკუმენტების დამატება/წაშლა."),
            ("Fast Onboarding", "ახალი თანამშრომელი პირველივე დღიდან ერთვება მუშაობაში და მყისიერად ფლობს კომპანიის სრულ ცოდნას.")
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
    p = tfb.paragraphs[0]; p.text = "💡 გუნდური პროდუქტიულობის ეფექტი:"
    p.font.name = FONT_ARIAL; p.font.size = Pt(14); p.font.bold = True; p.font.color.rgb = COLOR_RED
    p = tfb.add_paragraph()
    p.text = "Team Projects აქცევს Claude-ს ორგანიზაციის ცენტრალურ ციფრულ ტვინად: ცოდნა დაგროვილია ერთ დაცულ სივრცეში და აღარ იკარგება თანამშრომლების როტაციის დროს."
    p.font.name = FONT_ARIAL; p.font.size = Pt(13); p.font.color.rgb = COLOR_MUTED; p.space_before = Pt(4)

    # -------------------------------------------------------------------------
    # SLIDE 9: Projects ეკოსისტემის შემაჯამებელი რუკა
    # -------------------------------------------------------------------------
    s9 = created_slides[7]
    add_header(s9, "Projects ეკოსისტემის სრული რუკა (3 საყრდენი სვეტი)", "არქიტექტურა • ცოდნა, ინსტრუქციები და ინტეგრაციები ერთ სივრცეში")

    pillars_s9 = [
        ("ᲡᲕᲔᲢᲘ 1 • ᲠᲐ ᲘᲪᲘᲡ CLAUDE-Მ?", "Project Knowledge (ცოდნის ბაზა)", "დოკუმენტური საფუძველი", [
            ("დოკუმენტები და გაიდლაინები", "PDF, Word, Markdown, კვლევები და სტანდარტები."),
            ("კოდის რეპოზიტორიები", "პროგრამული ბაზები, სქემები და არქიტექტურა."),
            ("Prompt Caching", "მყისიერი წვდომა და დაბალი შეყოვნება (Latency)."),
            ("ფაქტობრივი საფუძველი", "პასუხები ჰალუცინაციების გარეშე, მკაცრად ფაქტებზე.")
        ]),
        ("ᲡᲕᲔᲢᲘ 2 • ᲠᲝᲒᲝᲠ ᲘᲥᲪᲔᲕᲐ CLAUDE?", "Custom Instructions (წესები და როლი)", "ქცევის ჩარჩო", [
            ("პროფესიული როლი", "ვიწრო ექსპერტიზა და კომპეტენციის განსაზღვრა."),
            ("ტონი და ენა", "აკადემიური, ლაკონიური, პროფესიონალური სტილი."),
            ("მკაცრი შეზღუდვები", "რა არის დაუშვებელი პასუხებში (Negative Rules)."),
            ("Output Standard", "ცხრილები, სტრუქტურა და ინტერაქტიული Artifacts.")
        ]),
        ("ᲡᲕᲔᲢᲘ 3 • ᲠᲐᲡ ᲣᲙᲐᲕᲨᲘᲠᲓᲔᲑᲐ CLAUDE?", "Connectors & Memory (კავშირები)", "გარე ეკოსისტემა", [
            ("Google Drive Connector", "ცოცხალი საოფისე ფაილების პირდაპირი კითხვა."),
            ("Gmail Connector", "სამუშაო მიმოწერის ანალიზი და დრაფტების შედგენა."),
            ("Claude Memory", "მომხმარებლის ჩვევებისა და წესების შენახვა."),
            ("Team Collaboration", "გაზიარებული სამუშაო სივრცე და როლების მართვა.")
        ])
    ]

    for i, (badge, title, subtitle, items) in enumerate(pillars_s9):
        bg_col = COLOR_CARD_RED if i == 1 else COLOR_BG_LIGHT
        bor_col = COLOR_BORDER_RED if i == 1 else COLOR_BORDER
        c = s9.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, xs[i], card_y, card_w, card_h)
        c.fill.solid(); c.fill.fore_color.rgb = bg_col
        c.line.color.rgb = bor_col; c.line.width = Pt(1.5)

        tb = s9.shapes.add_textbox(xs[i] + Inches(0.35), card_y + Inches(0.35), card_w - Inches(0.7), card_h - Inches(0.7))
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
    p = tfb.paragraphs[0]; p.text = "🎯 შემაჯამებელი დასკვნა:"
    p.font.name = FONT_ARIAL; p.font.size = Pt(14); p.font.bold = True; p.font.color.rgb = COLOR_RED
    p = tfb.add_paragraph()
    p.text = "ეს 3 კომპონენტი ერთობლიობაში ჩვეულებრივ ჩატბოტს გარდაქმნის თქვენს პირად, მაღალპროდუქტიულ ციფრულ პარტნიორად, რომელმაც იცის თქვენი ბიზნესი, იცავს თქვენს წესებს და დაკავშირებულია თქვენს სამუშაო ხელსაწყოებთან."
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
    filename = "5.1 - Claude სრულად II - Projects, Memory, Custom Instructions და Connectors.pptx"
    out_master = os.path.join(r'd:\ხელოვნური ინტელექტის ინსტრუმენტები', filename)
    prs.save(out_master)
    print(f"Master PPTX for Lecture 5.1 generated successfully: {out_master}")

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
    build_lecture_5_1()
