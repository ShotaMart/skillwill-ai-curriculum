import os
import sys
import shutil
import pptx
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN
from pptx.enum.shapes import MSO_SHAPE

def build_lecture_3():
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
    COLOR_GREEN = RGBColor(0x13, 0x73, 0x33)
    COLOR_GREEN_BG = RGBColor(0xE6, 0xF4, 0xEA)
    COLOR_BLUE = RGBColor(0x1A, 0x73, 0xE8)
    COLOR_BLUE_BG = RGBColor(0xE8, 0xF0, 0xFE)

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
            p.text = "3 - Machine Learning და Deep Learning მომხმარებლის დონეზე"
            p.alignment = PP_ALIGN.CENTER
            p.font.name = FONT_ARIAL
            p.font.size = Pt(38)
            p.font.bold = True
            p.font.color.rgb = COLOR_RED

            p2 = tf.add_paragraph()
            p2.text = "როგორ სწავლობს მანქანა გამოცდილებიდან — კოდის წერის გარეშე"
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
    def add_header(slide, title, category="ᲛᲝᲓᲣᲚᲘ 5 • ᲨᲔᲡᲐᲕᲐᲚᲘ AI • 3"):
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
    # SLIDE 2: Traditional Programming vs Machine Learning (Bilingual Paradigm)
    # -------------------------------------------------------------------------
    s2 = created_slides[0]
    add_header(s2, "Traditional Programming vs Machine Learning (მანქანური სწავლება)", "პარადიგმა • აზროვნების ფუნდამენტური ცვლილება")

    box_w = Inches(8.35); box_h = Inches(5.4); box_y = Inches(2.35)

    # Left: Traditional
    b1 = s2.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(1.0), box_y, box_w, box_h)
    b1.fill.solid(); b1.fill.fore_color.rgb = COLOR_BG_LIGHT
    b1.line.color.rgb = COLOR_BORDER; b1.line.width = Pt(1.5)

    tb1 = s2.shapes.add_textbox(Inches(1.4), box_y + Inches(0.4), box_w - Inches(0.8), box_h - Inches(0.8))
    tf1 = tb1.text_frame; tf1.word_wrap = True; tf1.margin_left = tf1.margin_top = 0
    p = tf1.paragraphs[0]; p.text = "ᲢᲠᲐᲓᲘᲪᲘᲣᲚᲘ ᲛᲘᲓᲒᲝᲛᲐ"
    p.font.name = FONT_ARIAL; p.font.size = Pt(12); p.font.bold = True; p.font.color.rgb = COLOR_MUTED

    p = tf1.add_paragraph(); p.text = "Traditional Programming (წესების კოდი)"
    p.font.name = FONT_ARIAL; p.font.size = Pt(22); p.font.bold = True; p.font.color.rgb = COLOR_DARK; p.space_before = Pt(4)

    p = tf1.add_paragraph(); p.text = "მეტაფორა: კულინარიული რეცეპტი — მკაცრი ინსტრუქცია"
    p.font.name = FONT_ARIAL; p.font.size = Pt(13); p.font.bold = True; p.font.color.rgb = COLOR_RED; p.space_before = Pt(8)

    t_points = [
        ("ლოგიკა", "პროგრამისტი წინასწარ ხელით წერს ყველა შესაძლო წესს (If / Else)."),
        ("ფორმულა", "Rules (წესები) + Data (მონაცემები) ➔ Answers (პასუხები)."),
        ("შეზღუდვა", "თუ გაუთვალისწინებელი სიტუაცია შეიქმნა ან წესი არ ჩაიწერა, პროგრამა შეცდომას უშვებს და ჩერდება.")
    ]
    for t, d in t_points:
        p = tf1.add_paragraph(); p.text = f"• {t}: {d}"
        p.font.name = FONT_ARIAL; p.font.size = Pt(14); p.font.color.rgb = COLOR_TEXT_MUTED; p.space_before = Pt(10)

    # Right: Machine Learning
    b2 = s2.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(9.65), box_y, box_w, box_h)
    b2.fill.solid(); b2.fill.fore_color.rgb = COLOR_CARD_RED
    b2.line.color.rgb = COLOR_BORDER_RED; b2.line.width = Pt(1.5)

    tb2 = s2.shapes.add_textbox(Inches(10.05), box_y + Inches(0.4), box_w - Inches(0.8), box_h - Inches(0.8))
    tf2 = tb2.text_frame; tf2.word_wrap = True; tf2.margin_left = tf2.margin_top = 0
    p = tf2.paragraphs[0]; p.text = "AI ᲛᲘᲓᲒᲝᲛᲐ"
    p.font.name = FONT_ARIAL; p.font.size = Pt(12); p.font.bold = True; p.font.color.rgb = COLOR_RED

    p = tf2.add_paragraph(); p.text = "Machine Learning (მანქანური სწავლება)"
    p.font.name = FONT_ARIAL; p.font.size = Pt(22); p.font.bold = True; p.font.color.rgb = COLOR_DARK; p.space_before = Pt(4)

    p = tf2.add_paragraph(); p.text = "მეტაფორა: ბავშვის სწავლება — მაგალითებზე დაკვირვება"
    p.font.name = FONT_ARIAL; p.font.size = Pt(13); p.font.bold = True; p.font.color.rgb = COLOR_GREEN; p.space_before = Pt(8)

    ml_points = [
        ("ლოგიკა", "კომპიუტერს წესებს არ ვუწერთ. ვაჩვენებთ ათასობით მაგალითს და ის თავად პოულობს კანონზომიერებას."),
        ("ფორმულა", "Data (მონაცემები) + Answers (პასუხები) ➔ Model / Rules (ნასწავლი წესები)."),
        ("უპირატესობა", "მოდელს შეუძლია გადაჭრას ამოცანები, რომელთა წესებად ჩამოყალიბება ადამიანს უჭირს (მაგ. ხმის ან სახის ცნობა).")
    ]
    for t, d in ml_points:
        p = tf2.add_paragraph(); p.text = f"• {t}: {d}"
        p.font.name = FONT_ARIAL; p.font.size = Pt(14); p.font.color.rgb = COLOR_TEXT_MUTED; p.space_before = Pt(10)

    # Bottom takeaway
    bot_y = Inches(8.0)
    bb = s2.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(1.0), bot_y, Inches(17.0), Inches(2.0))
    bb.fill.solid(); bb.fill.fore_color.rgb = COLOR_WHITE
    bb.line.color.rgb = COLOR_BORDER; bb.line.width = Pt(1.2)

    tbb = s2.shapes.add_textbox(Inches(1.4), bot_y + Inches(0.3), Inches(16.2), Inches(1.4))
    tfb = tbb.text_frame; tfb.word_wrap = True; tfb.margin_left = tfb.margin_top = 0
    p = tfb.paragraphs[0]; p.text = "💡 მთავარი არსი (მარტივი ანალოგია):"
    p.font.name = FONT_ARIAL; p.font.size = Pt(15); p.font.bold = True; p.font.color.rgb = COLOR_RED

    p = tfb.add_paragraph()
    p.text = "ტრადიციულ სამყაროში ადამიანი იყო „მზარეული“, რომელიც რეცეპტს წერდა. Machine Learning-ში ადამიანი არის „დამკვირვებელი“, რომელიც კომპიუტერს აძლევს ბევრ ინგრედიენტსა და მზა კერძს, კომპიუტერი კი თავად აღმოაჩენს საუკეთესო რეცეპტს."
    p.font.name = FONT_ARIAL; p.font.size = Pt(13); p.font.color.rgb = COLOR_DARK; p.space_before = Pt(6)

    # -------------------------------------------------------------------------
    # SLIDE 3: 3 Types of Machine Learning (Supervised, Unsupervised, RL)
    # -------------------------------------------------------------------------
    s3 = created_slides[1]
    add_header(s3, "Machine Learning-ის 3 ძირითადი მიმართულება (სწავლების ტიპები)", "არქიტექტურა • როგორ სწავლობს ალგორითმი")

    types_ml = [
        {
            "tag": "01 • ᲛᲐᲡᲬᲐᲕᲚᲔᲑᲚᲘᲗ",
            "title": "Supervised Learning",
            "geo": "(ზედამხედველობითი სწავლება)",
            "meta": "სწავლება „ნიშანდებული პასუხებით“ (Labeled Data)",
            "desc": "მოდელს ვაწვდით მონაცემს და მზა სწორ პასუხს (მაგ. სურათი + წარწერა „ვაშლი“). ალგორითმი სწავლობს მათ დაკავშირებას.",
            "examples": "სპამის ფილტრაცია, უძრავი ქონების ფასის პროგნოზი, სახის ამოცნობა.",
            "color": COLOR_CARD_RED, "bcolor": COLOR_BORDER_RED, "tcolor": COLOR_RED
        },
        {
            "tag": "02 • ᲓᲐᲛᲝᲣᲙᲘᲓᲔᲑᲚᲐᲓ",
            "title": "Unsupervised Learning",
            "geo": "(ზედამხედველობის გარეშე სწავლება)",
            "meta": "სწავლება „პასუხების გარეშე“ (Clustering)",
            "desc": "მოდელს ვაძლევთ ნედლ მონაცემებს მინიშნებების გარეშე. ის თავად ეძებს მსგავსებებს, ჯგუფებსა და ფარულ კანონზომიერებებს.",
            "examples": "მომხმარებელთა სეგმენტაცია მარკეტინგში, რეკომენდაციები (Netflix/Spotify).",
            "color": COLOR_BG_LIGHT, "bcolor": COLOR_BORDER, "tcolor": COLOR_MUTED
        },
        {
            "tag": "03 • ᲗᲐᲛᲐᲨᲘᲗ",
            "title": "Reinforcement Learning",
            "geo": "(განმტკიცებითი სწავლება)",
            "meta": "სწავლება „ცდისა და შეცდომის გზით“ (Trial & Error)",
            "desc": "აგენტი მოქმედებს ციფრულ გარემოში. სწორი ნაბიჯისთვის იღებს „ჯილდოს“ (Reward), შეცდომისთვის — „ჯარიმას“. მიზანია ქულების მაქსიმიზაცია.",
            "examples": "ჭადრაკი (AlphaZero), Tesla ავტოპილოტი, რობოტიკა.",
            "color": COLOR_BG_LIGHT, "bcolor": COLOR_BORDER, "tcolor": COLOR_MUTED
        }
    ]

    card_w = Inches(5.4); card_gap = Inches(0.4); start_x = Inches(1.0); card_y = Inches(2.35)
    for i, item in enumerate(types_ml):
        cx = start_x + i * (card_w + card_gap)

        c = s3.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, cx, card_y, card_w, Inches(7.7))
        c.fill.solid(); c.fill.fore_color.rgb = item["color"]
        c.line.color.rgb = item["bcolor"]; c.line.width = Pt(1.5)

        tb = s3.shapes.add_textbox(cx + Inches(0.4), card_y + Inches(0.4), card_w - Inches(0.8), Inches(6.9))
        tf = tb.text_frame; tf.word_wrap = True; tf.margin_left = tf.margin_top = 0

        p = tf.paragraphs[0]; p.text = item["tag"]
        p.font.name = FONT_ARIAL; p.font.size = Pt(12); p.font.bold = True; p.font.color.rgb = item["tcolor"]

        p = tf.add_paragraph(); p.text = item["title"]
        p.font.name = FONT_ARIAL; p.font.size = Pt(22); p.font.bold = True; p.font.color.rgb = COLOR_DARK; p.space_before = Pt(4)

        p = tf.add_paragraph(); p.text = item["geo"]
        p.font.name = FONT_ARIAL; p.font.size = Pt(14); p.font.bold = True; p.font.color.rgb = COLOR_MUTED; p.space_before = Pt(2)

        p = tf.add_paragraph(); p.text = f"\n💡 {item['meta']}"
        p.font.name = FONT_ARIAL; p.font.size = Pt(13); p.font.bold = True; p.font.color.rgb = COLOR_DARK; p.space_before = Pt(6)

        p = tf.add_paragraph(); p.text = item["desc"]
        p.font.name = FONT_ARIAL; p.font.size = Pt(13); p.font.color.rgb = COLOR_TEXT_MUTED; p.space_before = Pt(4)

        p = tf.add_paragraph(); p.text = f"\n🎯 რეალური მაგალითები:\n{item['examples']}"
        p.font.name = FONT_ARIAL; p.font.size = Pt(13); p.font.color.rgb = COLOR_DARK; p.space_before = Pt(8)

    # -------------------------------------------------------------------------
    # SLIDE 4: Train / Test Split & Overfitting (Bilingual Terminology)
    # -------------------------------------------------------------------------
    s4 = created_slides[2]
    add_header(s4, "Train / Test Split და Overfitting (გადაზეპირების ხაფანგი)", "მეთოდოლოგია • როგორ ვამოწმებთ მოდელის რეალურ ცოდნას")

    top_y = Inches(2.35)
    tb_intro = s4.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(1.0), top_y, Inches(17.0), Inches(1.3))
    tb_intro.fill.solid(); tb_intro.fill.fore_color.rgb = COLOR_BG_LIGHT
    tb_intro.line.color.rgb = COLOR_BORDER; tb_intro.line.width = Pt(1.2)

    tbox_i = s4.shapes.add_textbox(Inches(1.4), top_y + Inches(0.2), Inches(16.2), Inches(0.9))
    tfi = tbox_i.text_frame; tfi.word_wrap = True; tfi.margin_left = tfi.margin_top = 0
    p = tfi.paragraphs[0]; p.text = "🎓 საგამოცდო ანალოგია: როგორ ვარკვევთ, სტუდენტმა მართლა გაიგო მასალა თუ უბრალოდ ტესტები დაიზეპირა?"
    p.font.name = FONT_ARIAL; p.font.size = Pt(14); p.font.bold = True; p.font.color.rgb = COLOR_RED
    p = tfi.add_paragraph(); p.text = "მონაცემებს ყოველთვის ვყოფთ ორ ნაწილად: ერთით ვასწავლით, მეორით კი ვატარებთ ობიექტურ გამოცდას."
    p.font.name = FONT_ARIAL; p.font.size = Pt(13); p.font.color.rgb = COLOR_TEXT_MUTED; p.space_before = Pt(2)

    card_y = Inches(3.9); card_h = Inches(4.2); card_w = Inches(8.35)

    # Left: Split
    c1 = s4.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(1.0), card_y, card_w, card_h)
    c1.fill.solid(); c1.fill.fore_color.rgb = COLOR_BG_LIGHT
    c1.line.color.rgb = COLOR_BORDER; c1.line.width = Pt(1.5)

    tbc1 = s4.shapes.add_textbox(Inches(1.4), card_y + Inches(0.35), card_w - Inches(0.8), card_h - Inches(0.7))
    tfc1 = tbc1.text_frame; tfc1.word_wrap = True; tfc1.margin_left = tfc1.margin_top = 0
    p = tfc1.paragraphs[0]; p.text = "01 • DATA SPLIT (მონაცემების გაყოფა)"
    p.font.name = FONT_ARIAL; p.font.size = Pt(12); p.font.bold = True; p.font.color.rgb = COLOR_MUTED

    p = tfc1.add_paragraph(); p.text = "Training Set vs Test Set"
    p.font.name = FONT_ARIAL; p.font.size = Pt(20); p.font.bold = True; p.font.color.rgb = COLOR_DARK; p.space_before = Pt(4)

    split_points = [
        ("Training Set (საწვრთნელი ნაკრები ~80%)", "სახელმძღვანელო და სავარჯიშოები. მოდელი ამ მონაცემებს ათვალიერებს და სწავლობს კანონზომიერებებს."),
        ("Test Set (სატესტო ნაკრები ~20%)", "საგამოცდო ბილეთი. სრულიად ახალი მონაცემები, რომლებიც მოდელს აქამდე თვალით არ უნახავს."),
        ("ოქროს წესი", "მოდელი ფასდება მხოლოდ Test ნაკრებზე ნაჩვენები სიზუსტით!")
    ]
    for h, d in split_points:
        p = tfc1.add_paragraph(); p.text = f"• {h}: {d}"
        p.font.name = FONT_ARIAL; p.font.size = Pt(13); p.font.color.rgb = COLOR_TEXT_MUTED; p.space_before = Pt(6)

    # Right: Overfitting
    c2 = s4.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(9.65), card_y, card_w, card_h)
    c2.fill.solid(); c2.fill.fore_color.rgb = COLOR_CARD_RED
    c2.line.color.rgb = COLOR_BORDER_RED; c2.line.width = Pt(1.5)

    tbc2 = s4.shapes.add_textbox(Inches(10.05), card_y + Inches(0.35), card_w - Inches(0.8), card_h - Inches(0.7))
    tfc2 = tbc2.text_frame; tfc2.word_wrap = True; tfc2.margin_left = tfc2.margin_top = 0
    p = tfc2.paragraphs[0]; p.text = "02 • ᲫᲘᲠᲘᲗᲐᲓᲘ ᲡᲐᲤᲠᲗᲮᲔ"
    p.font.name = FONT_ARIAL; p.font.size = Pt(12); p.font.bold = True; p.font.color.rgb = COLOR_RED

    p = tfc2.add_paragraph(); p.text = "Overfitting (გადაზეპირება / ზედმეტი მორგება)"
    p.font.name = FONT_ARIAL; p.font.size = Pt(20); p.font.bold = True; p.font.color.rgb = COLOR_DARK; p.space_before = Pt(4)

    fit_points = [
        ("რა არის Overfitting?", "როცა მოდელმა საწვრთნელი მონაცემები იმდენად დეტალურად და შემთხვევითი ხმაურიანად დაიზეპირა, რომ ახალ კითხვაზე უძლურია."),
        ("Generalization (განზოგადების უნარი)", "ML-ის უმთავრესი მიზანი — მოდელმა უნდა გაიგოს ზოგადი პრინციპი და არა კონკრეტული სურათის პიქსელები."),
        ("შედეგი", "კარგი მოდელი გამოცდაზე (Test Set) თავს ისევე თავდაჯერებულად გრძნობს, როგორც სავარჯიშოებში.")
    ]
    for h, d in fit_points:
        p = tfc2.add_paragraph(); p.text = f"• {h}: {d}"
        p.font.name = FONT_ARIAL; p.font.size = Pt(13); p.font.color.rgb = COLOR_TEXT_MUTED; p.space_before = Pt(6)

    # Bottom takeaway
    bot_y = Inches(8.35)
    bs = s4.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(1.0), bot_y, Inches(17.0), Inches(1.65))
    bs.fill.solid(); bs.fill.fore_color.rgb = COLOR_WHITE
    bs.line.color.rgb = COLOR_BORDER; bs.line.width = Pt(1.2)

    tb_s = s4.shapes.add_textbox(Inches(1.4), bot_y + Inches(0.25), Inches(16.2), Inches(1.2))
    tfs = tb_s.text_frame; tfs.word_wrap = True; tfs.margin_left = tfs.margin_top = 0
    p = tfs.paragraphs[0]; p.text = "⚡ შეჯამება:"
    p.font.name = FONT_ARIAL; p.font.size = Pt(14); p.font.bold = True; p.font.color.rgb = COLOR_DARK
    p = tfs.add_paragraph()
    p.text = "Machine Learning-ის მიზანი არასდროს არის წარსული მონაცემების 100%-ით დაზეპირება. მისი ერთადერთი მიზანია Generalization (განზოგადება) — მომავალი, უცნობი სიტუაციების ზუსტი ამოცნობა."
    p.font.name = FONT_ARIAL; p.font.size = Pt(13); p.font.color.rgb = COLOR_TEXT_MUTED; p.space_before = Pt(3)

    # -------------------------------------------------------------------------
    # SLIDE 5: Deep Learning & The Artificial Neuron (Bilingual Terminology)
    # -------------------------------------------------------------------------
    s5 = created_slides[3]
    add_header(s5, "Deep Learning და ხელოვნური Neuron (ბიოლოგიური შთაგონება)", "ნეირონული ქსელები • ტვინის არქიტექტურიდან მათემატიკამდე")

    # Top intro
    top_y = Inches(2.35)
    tb_intro5 = s5.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(1.0), top_y, Inches(17.0), Inches(1.3))
    tb_intro5.fill.solid(); tb_intro5.fill.fore_color.rgb = COLOR_CARD_RED
    tb_intro5.line.color.rgb = COLOR_BORDER_RED; tb_intro5.line.width = Pt(1.5)

    tb5_t = s5.shapes.add_textbox(Inches(1.4), top_y + Inches(0.2), Inches(16.2), Inches(0.9))
    tf5 = tb5_t.text_frame; tf5.word_wrap = True; tf5.margin_left = tf5.margin_top = 0
    p = tf5.paragraphs[0]; p.text = "🧠 საიდან გაჩნდა იდეა? ადამიანის თავის ტვინში 86 მილიარდი ბიოლოგიური ნეირონია."
    p.font.name = FONT_ARIAL; p.font.size = Pt(14); p.font.bold = True; p.font.color.rgb = COLOR_RED
    p = tf5.add_paragraph(); p.text = "ხელოვნური ნეირონი (Perceptron) არის გამარტივებული მათემატიკური უჯრედი, რომელიც იღებს რამდენიმე სიგნალს, აფასებს მათ და იღებს გადაწყვეტილებას."
    p.font.name = FONT_ARIAL; p.font.size = Pt(13); p.font.color.rgb = COLOR_DARK; p.space_before = Pt(2)

    # 3 parts of Neuron
    neuron_parts = [
        {
            "num": "01 • ᲨᲔᲛᲐᲕᲐᲚᲘ ᲡᲘᲒᲜᲐᲚᲔᲑᲘ",
            "title": "Inputs (შემავალი მონაცემები)",
            "sub": "ობიექტის კონკრეტული მახასიათებლები",
            "desc": "ციფრული მნიშვნელობები, რომლებსაც ნეირონი აკვირდება. მაგალითად: სურათის პიქსელების ფერი, ყურების ზომა, ბეწვის სიგრძე."
        },
        {
            "num": "02 • ᲛᲜᲘᲨᲕᲜᲔᲚᲝᲑᲘᲡ ᲙᲝᲔᲤᲘᲪᲘᲔᲜᲢᲘ",
            "title": "Weights & Bias (წონები და გადახრა)",
            "sub": "რამდენად ენდობა ნეირონი თითოეულ ნიშანს?",
            "desc": "თითოეულ სიგნალს აქვს საკუთარი „წონა“ (Weight). თუ ყურების ფორმა გადამწყვეტია კატის საცნობად, ამ ნიშნის წონა ავტომატურად იზრდება."
        },
        {
            "num": "03 • ᲒᲐᲓᲐᲬᲧᲕᲔᲢᲘᲚᲔᲑᲐ",
            "title": "Activation Function (აქტივაცია)",
            "sub": "გადაწყვეტილების მიღების ზღვარი",
            "desc": "ნეირონი აჯამებს შეწონილ სიგნალებს. თუ ჯამი გარკვეულ ზღვარს გადასცდა, ნეირონი „აქტიურდება“ და აგზავნის სიგნალს: „დიახ, ეს კატაა!“"
        }
    ]

    card_w = Inches(5.4); card_gap = Inches(0.4); start_x = Inches(1.0); card_y = Inches(3.9)
    for i, part in enumerate(neuron_parts):
        cx = start_x + i * (card_w + card_gap)

        c = s5.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, cx, card_y, card_w, Inches(4.2))
        c.fill.solid(); c.fill.fore_color.rgb = COLOR_BG_LIGHT
        c.line.color.rgb = COLOR_BORDER; c.line.width = Pt(1.5)

        tb = s5.shapes.add_textbox(cx + Inches(0.4), card_y + Inches(0.35), card_w - Inches(0.8), Inches(3.5))
        tf = tb.text_frame; tf.word_wrap = True; tf.margin_left = tf.margin_top = 0

        p = tf.paragraphs[0]; p.text = part["num"]
        p.font.name = FONT_ARIAL; p.font.size = Pt(11); p.font.bold = True; p.font.color.rgb = COLOR_MUTED

        p = tf.add_paragraph(); p.text = part["title"]
        p.font.name = FONT_ARIAL; p.font.size = Pt(18); p.font.bold = True; p.font.color.rgb = COLOR_DARK; p.space_before = Pt(3)

        p = tf.add_paragraph(); p.text = part["sub"]
        p.font.name = FONT_ARIAL; p.font.size = Pt(12); p.font.bold = True; p.font.color.rgb = COLOR_RED; p.space_before = Pt(2)

        p = tf.add_paragraph(); p.text = f"\n{part['desc']}"
        p.font.name = FONT_ARIAL; p.font.size = Pt(13); p.font.color.rgb = COLOR_TEXT_MUTED; p.space_before = Pt(4)

    # Bottom takeaway
    bot_y = Inches(8.35)
    bb5 = s5.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(1.0), bot_y, Inches(17.0), Inches(1.65))
    bb5.fill.solid(); bb5.fill.fore_color.rgb = COLOR_WHITE
    bb5.line.color.rgb = COLOR_BORDER; bb5.line.width = Pt(1.2)

    tbb5 = s5.shapes.add_textbox(Inches(1.4), bot_y + Inches(0.25), Inches(16.2), Inches(1.2))
    tfb5 = tbb5.text_frame; tfb5.word_wrap = True; tfb5.margin_left = tfb5.margin_top = 0
    p = tfb5.paragraphs[0]; p.text = "💡 მარტივი დასკვნა:"
    p.font.name = FONT_ARIAL; p.font.size = Pt(14); p.font.bold = True; p.font.color.rgb = COLOR_DARK
    p = tfb5.add_paragraph()
    p.text = "ერთი ნეირონი მარტივ არჩევანს აკეთებს. მაგრამ როცა მილიარდობით ასეთ ნეირონს ერთმანეთს გადავაბამთ, მივიღებთ Deep Learning სისტემას, რომელსაც შეუძლია საუბარი, ხატვა და სამედიცინო დიაგნოზის დასმა."
    p.font.name = FONT_ARIAL; p.font.size = Pt(13); p.font.color.rgb = COLOR_TEXT_MUTED; p.space_before = Pt(3)

    # -------------------------------------------------------------------------
    # SLIDE 6: Neural Network Layers - Why "Deep"? (Bilingual Terminology)
    # -------------------------------------------------------------------------
    s6 = created_slides[4]
    add_header(s6, "Neural Network Layers: რატომ ჰქვია „Deep Learning“ (ღრმა სწავლება)?", "იერარქია • აბსტრაქციის 3 თანმიმდევრული დონე")

    layers = [
        {
            "step": "ᲤᲔᲜᲐ 01 • ᲓᲐᲑᲐᲚᲘ ᲓᲝᲜᲔ",
            "title": "Input & Low-Level Layers",
            "geo": "(საწყისი ფენები: ხაზები და კონტურები)",
            "desc": "ქსელი ხედავს მხოლოდ ნედლ პიქსელებს. პირველი ფენის ნეირონები სწავლობენ უმარტივეს ელემენტებს: სად არის ჰორიზონტალური ხაზი, სად ვერტიკალური, სად სინათლე და ჩრდილი.",
            "meta": "ვიზუალი: წერტილები, ზოლები, კიდეები"
        },
        {
            "step": "ᲤᲔᲜᲐ 02 • ᲡᲐᲨᲣᲐᲚᲝ ᲓᲝᲜᲔ",
            "title": "Hidden Layers (შუალედური ფენები)",
            "geo": "(გეომეტრია და ნაწილები)",
            "desc": "წინა ფენის ხაზები ერთიანდება უფრო რთულ ფორმებად. ნეირონები უკვე ცნობენ გეომეტრიულ ფიგურებს: კუთხეებს, წრეებს, თვალის ფორმას, ცხვირის კონტურს, ავტომობილის ბორბალს.",
            "meta": "ვიზუალი: ფორმები, ტექსტურა, ობიექტის ნაწილები"
        },
        {
            "step": "ᲤᲔᲜᲐ 03 • ᲛᲐᲦᲐᲚᲘ ᲓᲝᲜᲔ",
            "title": "Output & High-Level Layers",
            "geo": "(გამომავალი ფენა: სრული ობიექტები)",
            "desc": "ნაწილები ერთიანდება სრულ ობიექტებად. ქსელი სრულყოფილად ცნობს: ადამიანის კონკრეტულ სახეს, ძაღლის ჯიშს ან ავტომობილის მოდელს და გასცემს საბოლოო პასუხს.",
            "meta": "ვიზუალი: მთლიანი სახე, ცხოველი, ტექსტი"
        }
    ]

    card_w = Inches(5.4); card_gap = Inches(0.4); start_x = Inches(1.0); card_y = Inches(2.35)
    for i, lyr in enumerate(layers):
        cx = start_x + i * (card_w + card_gap)

        c = s6.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, cx, card_y, card_w, Inches(7.7))
        c.fill.solid()
        c.fill.fore_color.rgb = COLOR_CARD_RED if i == 0 else COLOR_BG_LIGHT
        c.line.color.rgb = COLOR_BORDER_RED if i == 0 else COLOR_BORDER
        c.line.width = Pt(1.5)

        tb = s6.shapes.add_textbox(cx + Inches(0.4), card_y + Inches(0.4), card_w - Inches(0.8), Inches(6.9))
        tf = tb.text_frame; tf.word_wrap = True; tf.margin_left = tf.margin_top = 0

        p = tf.paragraphs[0]; p.text = lyr["step"]
        p.font.name = FONT_ARIAL; p.font.size = Pt(12); p.font.bold = True
        p.font.color.rgb = COLOR_RED if i == 0 else COLOR_MUTED

        p = tf.add_paragraph(); p.text = lyr["title"]
        p.font.name = FONT_ARIAL; p.font.size = Pt(20); p.font.bold = True; p.font.color.rgb = COLOR_DARK; p.space_before = Pt(4)

        p = tf.add_paragraph(); p.text = lyr["geo"]
        p.font.name = FONT_ARIAL; p.font.size = Pt(13); p.font.bold = True; p.font.color.rgb = COLOR_MUTED; p.space_before = Pt(2)

        p = tf.add_paragraph(); p.text = f"\n{lyr['desc']}"
        p.font.name = FONT_ARIAL; p.font.size = Pt(13); p.font.color.rgb = COLOR_TEXT_MUTED; p.space_before = Pt(6)

        p = tf.add_paragraph(); p.text = f"\n🔍 {lyr['meta']}"
        p.font.name = FONT_ARIAL; p.font.size = Pt(12); p.font.bold = True; p.font.color.rgb = COLOR_DARK; p.space_before = Pt(8)

    # -------------------------------------------------------------------------
    # SLIDE 7: The Training Loop (Forward, Loss, Backprop, Weights)
    # -------------------------------------------------------------------------
    s7 = created_slides[5]
    add_header(s7, "Training Loop (წვრთნის ციკლი): როგორ ჭკვიანდება მოდელი?", "პროცესი • 4-საფეხურიანი უწყვეტი გაუმჯობესება")

    steps = [
        ("01 • FORWARD PASS", "Forward Pass (პირდაპირი პროგნოზი)", "მონაცემი (სურათი) გადის ქსელის ყველა ფენას. მოდელი აკეთებს პირველ ვარაუდს (დასაწყისში სრულიად შემთხვევითს)."),
        ("02 • LOSS FUNCTION", "Loss Function (შეცდომის გაზომვა)", "სპეციალური მათემატიკური ფუნქცია ზომავს განსხვავებას: რა ივარაუდა მოდელმა vs რა იყო სინამდვილეში."),
        ("03 • BACKPROPAGATION", "Backpropagation (შეცდომის უკუგავრცელება)", "შეცდომის სიგნალი უკან ბრუნდება ქსელში ბოლოდან დასაწყისისკენ, რათა გაირკვეს, რომელმა ნეირონმა დააშავა."),
        ("04 • WEIGHT UPDATE", "Weight Update (წონების კორექტირება)", "ნეირონების წონები (Weights) ოდნავ იცვლება (Gradient Descent), რათა შემდეგ ცდაზე შეცდომა შემცირდეს.")
    ]

    sw = Inches(3.95); sg = Inches(0.4); sx = Inches(1.0); sy = Inches(2.35)
    for i, (tag, stitle, sdesc) in enumerate(steps):
        cx = sx + i * (sw + sg)

        c = s7.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, cx, sy, sw, Inches(5.4))
        c.fill.solid()
        c.fill.fore_color.rgb = COLOR_CARD_RED if i == 0 else COLOR_BG_LIGHT
        c.line.color.rgb = COLOR_BORDER_RED if i == 0 else COLOR_BORDER
        c.line.width = Pt(1.5)

        tb = s7.shapes.add_textbox(cx + Inches(0.3), sy + Inches(0.35), sw - Inches(0.6), Inches(4.7))
        tf = tb.text_frame; tf.word_wrap = True; tf.margin_left = tf.margin_top = 0

        p = tf.paragraphs[0]; p.text = tag
        p.font.name = FONT_ARIAL; p.font.size = Pt(11); p.font.bold = True
        p.font.color.rgb = COLOR_RED if i == 0 else COLOR_MUTED

        p = tf.add_paragraph(); p.text = stitle
        p.font.name = FONT_ARIAL; p.font.size = Pt(16); p.font.bold = True; p.font.color.rgb = COLOR_DARK; p.space_before = Pt(4)

        p = tf.add_paragraph(); p.text = f"\n{sdesc}"
        p.font.name = FONT_ARIAL; p.font.size = Pt(13); p.font.color.rgb = COLOR_TEXT_MUTED; p.space_before = Pt(6)

    # Bottom takeaway
    bot_y = Inches(8.0)
    bb7 = s7.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(1.0), bot_y, Inches(17.0), Inches(2.0))
    bb7.fill.solid(); bb7.fill.fore_color.rgb = COLOR_WHITE
    bb7.line.color.rgb = COLOR_BORDER; bb7.line.width = Pt(1.2)

    tbb7 = s7.shapes.add_textbox(Inches(1.4), bot_y + Inches(0.3), Inches(16.2), Inches(1.4))
    tfb7 = tbb7.text_frame; tfb7.word_wrap = True; tfb7.margin_left = tfb7.margin_top = 0
    p = tfb7.paragraphs[0]; p.text = "⚡ რა ხდება სინამდვილეში?"
    p.font.name = FONT_ARIAL; p.font.size = Pt(14); p.font.bold = True; p.font.color.rgb = COLOR_RED
    p = tfb7.add_paragraph()
    p.text = "ეს 4-საფეხურიანი ციკლი მეორდება მილიონობითჯერ. თავიდან მოდელი ბრმად მარჩიელობს, მაგრამ ყოველ ნაბიჯზე შეცდომის გამოსწორებით, წონები სრულყოფილებამდე მიდის — ასე სწავლობს AI ნებისმიერ ამოცანას."
    p.font.name = FONT_ARIAL; p.font.size = Pt(13); p.font.color.rgb = COLOR_DARK; p.space_before = Pt(4)

    # -------------------------------------------------------------------------
    # SLIDE 8: Why GPU and not CPU? (Bilingual Terminology)
    # -------------------------------------------------------------------------
    s8 = created_slides[6]
    add_header(s8, "CPU vs GPU (ტექნოლოგიური საწვავი: რატომ გახდა Deep Learning შესაძლებელი?)", "ინფრასტრუქტურა • კომპიუტერული სიმძლავრის ევოლუცია")

    col_w = Inches(8.35); col_h = Inches(5.4); col_y = Inches(2.35)

    # Left: CPU
    b1 = s8.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(1.0), col_y, col_w, col_h)
    b1.fill.solid(); b1.fill.fore_color.rgb = COLOR_BG_LIGHT
    b1.line.color.rgb = COLOR_BORDER; b1.line.width = Pt(1.5)

    tb1 = s8.shapes.add_textbox(Inches(1.4), col_y + Inches(0.4), col_w - Inches(0.8), col_h - Inches(0.8))
    tf1 = tb1.text_frame; tf1.word_wrap = True; tf1.margin_left = tf1.margin_top = 0

    p = tf1.paragraphs[0]; p.text = "CPU (CENTRAL PROCESSING UNIT)"
    p.font.name = FONT_ARIAL; p.font.size = Pt(12); p.font.bold = True; p.font.color.rgb = COLOR_MUTED

    p = tf1.add_paragraph(); p.text = "ცენტრალური პროცესორი"
    p.font.name = FONT_ARIAL; p.font.size = Pt(22); p.font.bold = True; p.font.color.rgb = COLOR_DARK; p.space_before = Pt(4)

    p = tf1.add_paragraph(); p.text = "მეტაფორა: „აკადემიკოსი პროფესორი“"
    p.font.name = FONT_ARIAL; p.font.size = Pt(14); p.font.bold = True; p.font.color.rgb = COLOR_MUTED; p.space_before = Pt(6)

    cpu_desc = [
        ("ბირთვების რაოდენობა", "4-დან 16-მდე ულტრა-ძლიერი ბირთვი."),
        ("მუშაობის სტილი", "ასრულებს ძალიან რთულ ამოცანებს, მაგრამ თანმიმდევრულად — სათითაოდ (Sequential Processing)."),
        ("AI-ს პრობლემა", "ნეირონული ქსელების მილიარდობით მარტივი გამრავლების რიგ-რიგობით შესრულებას CPU-ზე თვეები და წლები სჭირდება.")
    ]
    for h, d in cpu_desc:
        p = tf1.add_paragraph(); p.text = f"• {h}: {d}"
        p.font.name = FONT_ARIAL; p.font.size = Pt(13); p.font.color.rgb = COLOR_TEXT_MUTED; p.space_before = Pt(8)

    # Right: GPU
    b2 = s8.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(9.65), col_y, col_w, col_h)
    b2.fill.solid(); b2.fill.fore_color.rgb = COLOR_CARD_RED
    b2.line.color.rgb = COLOR_BORDER_RED; b2.line.width = Pt(1.5)

    tb2 = s8.shapes.add_textbox(Inches(10.05), col_y + Inches(0.4), col_w - Inches(0.8), col_h - Inches(0.8))
    tf2 = tb2.text_frame; tf2.word_wrap = True; tf2.margin_left = tf2.margin_top = 0

    p = tf2.paragraphs[0]; p.text = "GPU (GRAPHICS PROCESSING UNIT)"
    p.font.name = FONT_ARIAL; p.font.size = Pt(12); p.font.bold = True; p.font.color.rgb = COLOR_RED

    p = tf2.add_paragraph(); p.text = "გრაფიკული პროცესორი"
    p.font.name = FONT_ARIAL; p.font.size = Pt(22); p.font.bold = True; p.font.color.rgb = COLOR_DARK; p.space_before = Pt(4)

    p = tf2.add_paragraph(); p.text = "მეტაფორა: „10,000 სკოლის მოსწავლე ერთ ოთახში“"
    p.font.name = FONT_ARIAL; p.font.size = Pt(14); p.font.bold = True; p.font.color.rgb = COLOR_GREEN; p.space_before = Pt(6)

    gpu_desc = [
        ("ბირთვების რაოდენობა", "5,000-დან 15,000-მდე პარალელური ბირთვი ერთ ჩიპზე."),
        ("მუშაობის სტილი", "თითოეული ბირთვი აკეთებს უმარტივეს გამრავლებას, მაგრამ ყველა მათგანი მუშაობს ერთდროულად (Parallel Processing)."),
        ("რატომ გახდა AI-ს საწვავი?", "ნეირონული ქსელი სხვა არაფერია, თუ არა მატრიცების პარალელური გამრავლება — რაც GPU-ზე წამებში სრულდება.")
    ]
    for h, d in gpu_desc:
        p = tf2.add_paragraph(); p.text = f"• {h}: {d}"
        p.font.name = FONT_ARIAL; p.font.size = Pt(13); p.font.color.rgb = COLOR_TEXT_MUTED; p.space_before = Pt(8)

    # Bottom comparison summary
    bot_y = Inches(8.0)
    bb8 = s8.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(1.0), bot_y, Inches(17.0), Inches(2.0))
    bb8.fill.solid(); bb8.fill.fore_color.rgb = COLOR_WHITE
    bb8.line.color.rgb = COLOR_BORDER; bb8.line.width = Pt(1.2)

    tbb8 = s8.shapes.add_textbox(Inches(1.4), bot_y + Inches(0.3), Inches(16.2), Inches(1.4))
    tfb8 = tbb8.text_frame; tfb8.word_wrap = True; tfb8.margin_left = tfb8.margin_top = 0
    p = tfb8.paragraphs[0]; p.text = "🚀 Nvidia-ს ფენომენი:"
    p.font.name = FONT_ARIAL; p.font.size = Pt(14); p.font.bold = True; p.font.color.rgb = COLOR_RED
    p = tfb8.add_paragraph()
    p.text = "ვიდეობარათები თავდაპირველად კომპიუტერული თამაშების გრაფიკის დასახატად შეიქმნა. მაგრამ 2012 წელს აღმოაჩინეს, რომ ნეირონული ქსელების წვრთნისთვის ეს არქიტექტურა იდეალური იყო — სწორედ აქედან დაიწყო თანამედროვე AI ბუმი."
    p.font.name = FONT_ARIAL; p.font.size = Pt(13); p.font.color.rgb = COLOR_DARK; p.space_before = Pt(4)

    # -------------------------------------------------------------------------
    # SLIDE 9: Ecosystem & Tools (512x512 Logos)
    # -------------------------------------------------------------------------
    s9 = created_slides[7]
    add_header(s9, "თანამედროვე ხელსაწყოები და ეკოსისტემა (პროდუქტების რუკა)", "ინსტრუმენტები • ვის რა როლი აქვს Machine Learning-ში")

    eco_columns = [
        {
            "cat": "01 • VISUAL & NO-CODE AI",
            "title": "No-Code Learning",
            "sub": "ვიზუალური სწავლება კოდის გარეშე",
            "tools": [
                ("google_cloud.png", "Google Teachable Machine", "მოდელის გაწვრთნა ბრაუზერში ვებკამერით"),
                ("tensorflow.png", "TensorFlow Playground", "ნეირონული ქსელის ინტერაქტიული სიმულატორი")
            ]
        },
        {
            "cat": "02 • CORE FRAMEWORKS",
            "title": "ML Frameworks",
            "sub": "საწარმოო ბიბლიოთეკები დეველოპერებისთვის",
            "tools": [
                ("pytorch.png", "PyTorch (Meta)", "მსოფლიო კვლევითი სტანდარტი და LLM ძრავა"),
                ("tensorflow.png", "TensorFlow (Google)", "საწარმოო და ინდუსტრიული დანერგვის პლატფორმა")
            ]
        },
        {
            "cat": "03 • COMPUTE & HARDWARE",
            "title": "Hardware & Cloud",
            "sub": "სუპერკომპიუტერები და ინფრასტრუქტურა",
            "tools": [
                ("nvidia.png", "Nvidia CUDA & H100", "AI გამოთვლების გლობალური ლიდერი"),
                ("aws.png", "AWS Bedrock / Cloud", "ღრუბლოვანი სერვერები მასშტაბური წვრთნისთვის")
            ]
        }
    ]

    col_w = Inches(5.4); col_gap = Inches(0.4); start_x = Inches(1.0)
    for i, col in enumerate(eco_columns):
        cx = start_x + i * (col_w + col_gap)

        cb = s9.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, cx, Inches(2.35), col_w, Inches(7.7))
        cb.fill.solid()
        cb.fill.fore_color.rgb = COLOR_CARD_RED if i == 0 else COLOR_BG_LIGHT
        cb.line.color.rgb = COLOR_BORDER_RED if i == 0 else COLOR_BORDER
        cb.line.width = Pt(1.5)

        tb_head = s9.shapes.add_textbox(cx + Inches(0.35), Inches(2.55), col_w - Inches(0.7), Inches(1.6))
        tf = tb_head.text_frame; tf.word_wrap = True; tf.margin_left = tf.margin_top = 0

        p = tf.paragraphs[0]; p.text = col["cat"]
        p.font.name = FONT_ARIAL; p.font.size = Pt(11); p.font.bold = True
        p.font.color.rgb = COLOR_RED if i == 0 else COLOR_MUTED

        p = tf.add_paragraph(); p.text = col["title"]
        p.font.name = FONT_ARIAL; p.font.size = Pt(20); p.font.bold = True; p.font.color.rgb = COLOR_DARK; p.space_before = Pt(2)

        p = tf.add_paragraph(); p.text = col["sub"]
        p.font.name = FONT_ARIAL; p.font.size = Pt(12); p.font.color.rgb = COLOR_TEXT_MUTED; p.space_before = Pt(3)

        num_tools = len(col["tools"])
        card_h = Inches(1.9)
        spacing = Inches(2.2)
        base_y = Inches(4.5)

        for j, (l_file, t_name, t_desc) in enumerate(col["tools"]):
            item_y = base_y + j * spacing

            t_card = s9.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, cx + Inches(0.3), item_y, col_w - Inches(0.6), card_h)
            t_card.fill.solid(); t_card.fill.fore_color.rgb = COLOR_WHITE
            t_card.line.color.rgb = COLOR_BORDER_RED if i == 0 else COLOR_BORDER
            t_card.line.width = Pt(1.2)

            # Square Logo 1.1" x 1.1"
            lp = os.path.join(logo_dir, l_file)
            logo_sz = Inches(1.1)
            logo_y = item_y + (card_h - logo_sz) / 2
            if os.path.exists(lp):
                s9.shapes.add_picture(lp, cx + Inches(0.5), logo_y, width=logo_sz, height=logo_sz)

            tb_item = s9.shapes.add_textbox(cx + Inches(1.8), item_y + Inches(0.3), col_w - Inches(2.2), card_h - Inches(0.5))
            tf = tb_item.text_frame; tf.word_wrap = True; tf.margin_left = tf.margin_top = 0

            p = tf.paragraphs[0]; p.text = t_name
            p.font.name = FONT_ARIAL; p.font.size = Pt(14); p.font.bold = True; p.font.color.rgb = COLOR_DARK

            p = tf.add_paragraph(); p.text = t_desc
            p.font.name = FONT_ARIAL; p.font.size = Pt(11); p.font.color.rgb = COLOR_TEXT_MUTED; p.space_before = Pt(3)

    # -------------------------------------------------------------------------
    # SAVE PPTX AND DISTRIBUTE TO TARGET DIRECTORIES
    # -------------------------------------------------------------------------
    filename = "3 - Machine Learning და Deep Learning - როგორ სწავლობს AI.pptx"
    out_master = os.path.join(r'd:\ხელოვნური ინტელექტის ინსტრუმენტები', filename)
    prs.save(out_master)
    print("Master PPTX for Lecture 3 generated successfully!")

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
            print(f"Copied to Drive target: {dst_file}")
        except Exception as e:
            print(f"Error copying to {td}: {e}")

if __name__ == '__main__':
    build_lecture_3()
