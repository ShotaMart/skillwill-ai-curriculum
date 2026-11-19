import os
import sys
import shutil
import pptx
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN
from pptx.enum.shapes import MSO_SHAPE

def build_lecture_2_2():
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
            p.text = "2.2 - LLM არქიტექტურა: Transformer, Attention და მოდელის წვრთნა"
            p.alignment = PP_ALIGN.CENTER
            p.font.name = FONT_ARIAL
            p.font.size = Pt(38)
            p.font.bold = True
            p.font.color.rgb = COLOR_RED

            p2 = tf.add_paragraph()
            p2.text = "როგორ აზროვნებს და სწავლობს თანამედროვე ხელოვნური ინტელექტი"
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
    def add_header(slide, title, category="ᲛᲝᲓᲣᲚᲘ 5 • ᲨᲔᲡᲐᲕᲐᲚᲘ AI • 2.2"):
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
    # SLIDE 2: Bridging from Lecture 2.1 to 2.2 (Words to "Thinking Engine")
    # -------------------------------------------------------------------------
    s2 = created_slides[0]
    add_header(s2, "ტექსტის ნედლეულიდან „აზროვნების ძრავამდე“", "კონტექსტი • წინა ლექციის გაგრძელება")

    box_w = Inches(8.35); box_h = Inches(5.4); box_y = Inches(2.35)

    # Left: Lecture 2.1 Recap
    b1 = s2.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(1.0), box_y, box_w, box_h)
    b1.fill.solid(); b1.fill.fore_color.rgb = COLOR_BG_LIGHT
    b1.line.color.rgb = COLOR_BORDER; b1.line.width = Pt(1.5)

    tb1 = s2.shapes.add_textbox(Inches(1.4), box_y + Inches(0.4), box_w - Inches(0.8), box_h - Inches(0.8))
    tf1 = tb1.text_frame; tf1.word_wrap = True; tf1.margin_left = tf1.margin_top = 0
    p = tf1.paragraphs[0]; p.text = "ᲚᲔᲥᲪᲘᲐ 2.1 • ᲠᲐ ᲕᲘᲡᲬᲐᲕᲚᲔᲗ?"
    p.font.name = FONT_ARIAL; p.font.size = Pt(13); p.font.bold = True; p.font.color.rgb = COLOR_MUTED

    p = tf1.add_paragraph(); p.text = "ენის გაციფრულება და რუკა"
    p.font.name = FONT_ARIAL; p.font.size = Pt(24); p.font.bold = True; p.font.color.rgb = COLOR_DARK
    p.space_before = Pt(4)

    items_2_1 = [
        ("ტოკენიზაცია", "ადამიანური ტექსტი იჭრება ციფრულ „ლეგოს კუბიკებად“."),
        ("Embeddings (რუკა)", "თითოეულ სიტყვას ენიჭება GPS კოორდინატები მნიშვნელობის მიხედვით."),
        ("სემანტიკური ძებნა", "AI-ს შეუძლია იპოვოს მსგავსი შინაარსის მონაკვეთები რუკაზე.")
    ]
    for title, desc in items_2_1:
        p = tf1.add_paragraph(); p.text = f"• {title}: {desc}"
        p.font.name = FONT_ARIAL; p.font.size = Pt(14); p.font.color.rgb = COLOR_TEXT_MUTED
        p.space_before = Pt(12)

    # Right: Lecture 2.2 Focus (LLM Core)
    b2 = s2.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(9.65), box_y, box_w, box_h)
    b2.fill.solid(); b2.fill.fore_color.rgb = COLOR_CARD_RED
    b2.line.color.rgb = COLOR_BORDER_RED; b2.line.width = Pt(1.5)

    tb2 = s2.shapes.add_textbox(Inches(10.05), box_y + Inches(0.4), box_w - Inches(0.8), box_h - Inches(0.8))
    tf2 = tb2.text_frame; tf2.word_wrap = True; tf2.margin_left = tf2.margin_top = 0
    p = tf2.paragraphs[0]; p.text = "ᲚᲔᲥᲪᲘᲐ 2.2 • ᲓᲦᲔᲕᲐᲜᲓᲔᲚᲘ ᲛᲘᲖᲐᲜᲘ"
    p.font.name = FONT_ARIAL; p.font.size = Pt(13); p.font.bold = True; p.font.color.rgb = COLOR_RED

    p = tf2.add_paragraph(); p.text = "აზროვნების ძრავა — LLM"
    p.font.name = FONT_ARIAL; p.font.size = Pt(24); p.font.bold = True; p.font.color.rgb = COLOR_DARK
    p.space_before = Pt(4)

    items_2_2 = [
        ("Transformer არქიტექტურა", "ქსელი, რომელიც ამ ციფრულ კოორდინატებს ერთიან კონტექსტად კრავს."),
        ("Self-Attention მექანიზმი", "„ყურადღების პროჟექტორი“, რომელიც ხვდება, რომელი სიტყვა რომელს უკავშირდება."),
        ("მოდელის წვრთნა (Training)", "როგორ იქცევა მილიარდობით პარამეტრი გონიერ ასისტენტად.")
    ]
    for title, desc in items_2_2:
        p = tf2.add_paragraph(); p.text = f"• {title}: {desc}"
        p.font.name = FONT_ARIAL; p.font.size = Pt(14); p.font.color.rgb = COLOR_TEXT_MUTED
        p.space_before = Pt(12)

    # Bottom highlight banner
    bot_y = Inches(8.0)
    bb = s2.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(1.0), bot_y, Inches(17.0), Inches(2.0))
    bb.fill.solid(); bb.fill.fore_color.rgb = COLOR_WHITE
    bb.line.color.rgb = COLOR_BORDER; bb.line.width = Pt(1.2)

    tbb = s2.shapes.add_textbox(Inches(1.4), bot_y + Inches(0.3), Inches(16.2), Inches(1.4))
    tfb = tbb.text_frame; tfb.word_wrap = True; tfb.margin_left = tfb.margin_top = 0
    p = tfb.paragraphs[0]; p.text = "💡 მთავარი პარადიგმა (0-ლეველის მენტალური მოდელი):"
    p.font.name = FONT_ARIAL; p.font.size = Pt(15); p.font.bold = True; p.font.color.rgb = COLOR_RED

    p = tfb.add_paragraph()
    p.text = "LLM არ არის ფაქტების ან ენციკლოპედიური მონაცემების ტრადიციული ბაზა. ის არის უზარმაზარი მათემატიკური ალგორითმი, რომელმაც ისწავლა ადამიანური ენის ლოგიკა, სტრუქტურა და წინადადების აზრობრივი გაგრძელება."
    p.font.name = FONT_ARIAL; p.font.size = Pt(14); p.font.color.rgb = COLOR_DARK; p.space_before = Pt(6)

    # -------------------------------------------------------------------------
    # SLIDE 3: The Transformer Revolution (2017)
    # -------------------------------------------------------------------------
    s3 = created_slides[1]
    add_header(s3, "Transformer არქიტექტურა: რევოლუცია AI-ში (2017)", "ისტორიული ნახტომი • როგორ შეიცვალა სამყარო")

    col_w = Inches(8.35); col_h = Inches(7.7); col_y = Inches(2.35)

    # Left: Before 2017 (RNN / Sequential)
    c1 = s3.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(1.0), col_y, col_w, col_h)
    c1.fill.solid(); c1.fill.fore_color.rgb = COLOR_BG_LIGHT
    c1.line.color.rgb = COLOR_BORDER; c1.line.width = Pt(1.5)

    tbc1 = s3.shapes.add_textbox(Inches(1.4), col_y + Inches(0.4), col_w - Inches(0.8), col_h - Inches(0.8))
    tfc1 = tbc1.text_frame; tfc1.word_wrap = True; tfc1.margin_left = tfc1.margin_top = 0

    p = tfc1.paragraphs[0]; p.text = "2017 წლამდე • ᲫᲕᲔᲚᲘ ᲗᲐᲝᲑᲐ"
    p.font.name = FONT_ARIAL; p.font.size = Pt(13); p.font.bold = True; p.font.color.rgb = COLOR_MUTED

    p = tfc1.add_paragraph(); p.text = "RNN და თანმიმდევრული კითხვა"
    p.font.name = FONT_ARIAL; p.font.size = Pt(22); p.font.bold = True; p.font.color.rgb = COLOR_DARK; p.space_before = Pt(4)

    p = tfc1.add_paragraph(); p.text = "მეტაფორა: წიგნის კითხვა ლუპით — სიტყვასიტყვით"
    p.font.name = FONT_ARIAL; p.font.size = Pt(14); p.font.bold = True; p.font.color.rgb = COLOR_RED; p.space_before = Pt(8)

    rnn_points = [
        ("თანმიმდევრული დამუშავება", "მოდელი კითხულობდა სიტყვებს რიგრიგობით (სიტყვა 1 -> სიტყვა 2 -> სიტყვა 3). სანამ პირველს არ მორჩებოდა, მეორეს ვერ იწყებდა."),
        ("მოკლე მეხსიერება", "წინადადების ბოლოს მისულ მოდელს დასაწყისი ავიწყდებოდა (Vanishing Gradient). აბზაცზე დიდ ტექსტს ვერ იაზრებდა."),
        ("შეუძლებელი პარალელიზაცია", "სუპერკომპიუტერების GPU-ებს არ შეეძლოთ ერთდროულად მუშაობა, რადგან პროცესი მკაცრად თანმიმდევრული იყო.")
    ]
    for t, d in rnn_points:
        p = tfc1.add_paragraph(); p.text = f"\n❌ {t}"
        p.font.name = FONT_ARIAL; p.font.size = Pt(14); p.font.bold = True; p.font.color.rgb = COLOR_DARK; p.space_before = Pt(8)
        p = tfc1.add_paragraph(); p.text = d
        p.font.name = FONT_ARIAL; p.font.size = Pt(13); p.font.color.rgb = COLOR_TEXT_MUTED; p.space_before = Pt(2)

    # Right: 2017+ (Transformer)
    c2 = s3.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(9.65), col_y, col_w, col_h)
    c2.fill.solid(); c2.fill.fore_color.rgb = COLOR_CARD_RED
    c2.line.color.rgb = COLOR_BORDER_RED; c2.line.width = Pt(1.5)

    tbc2 = s3.shapes.add_textbox(Inches(10.05), col_y + Inches(0.4), col_w - Inches(0.8), col_h - Inches(0.8))
    tfc2 = tbc2.text_frame; tfc2.word_wrap = True; tfc2.margin_left = tfc2.margin_top = 0

    p = tfc2.paragraphs[0]; p.text = "2017 ᲬᲚᲘᲓᲐᲜ • ᲡᲠᲣᲚᲘ ᲠᲔᲕᲝᲚᲣᲪᲘᲐ"
    p.font.name = FONT_ARIAL; p.font.size = Pt(13); p.font.bold = True; p.font.color.rgb = COLOR_RED

    p = tfc2.add_paragraph(); p.text = "Transformer (Attention Is All You Need)"
    p.font.name = FONT_ARIAL; p.font.size = Pt(22); p.font.bold = True; p.font.color.rgb = COLOR_DARK; p.space_before = Pt(4)

    p = tfc2.add_paragraph(); p.text = "მეტაფორა: მთლიანი გვერდის ერთი შეხედვით აღქმა"
    p.font.name = FONT_ARIAL; p.font.size = Pt(14); p.font.bold = True; p.font.color.rgb = COLOR_GREEN; p.space_before = Pt(8)

    tr_points = [
        ("პარალელური დამუშავება", "მთელ წინადადებას (და მთელ წიგნს) მოდელი ხედავს ერთდროულად, წამიერად, ათასობით GPU პროცესორის პარალელური სიმძლავრით."),
        ("უსაზღვრო კონტექსტი", "სიტყვები ერთმანეთს უკავშირდება პირდაპირ, მანძილის მიუხედავად (Self-Attention). შესაძლებელი გახდა 100+ გვერდის ერთდროული მეხსიერება."),
        ("მასშტაბურობა (Scaling Laws)", "რაც უფრო მეტ მონაცემს და კომპიუტერულ სიმძლავრეს მივცემთ Transformer-ს, მით უფრო ჭკვიანი ხდება ის.")
    ]
    for t, d in tr_points:
        p = tfc2.add_paragraph(); p.text = f"\n✅ {t}"
        p.font.name = FONT_ARIAL; p.font.size = Pt(14); p.font.bold = True; p.font.color.rgb = COLOR_DARK; p.space_before = Pt(8)
        p = tfc2.add_paragraph(); p.text = d
        p.font.name = FONT_ARIAL; p.font.size = Pt(13); p.font.color.rgb = COLOR_TEXT_MUTED; p.space_before = Pt(2)

    # -------------------------------------------------------------------------
    # SLIDE 4: Self-Attention Mechanism (Spotlight of meaning)
    # -------------------------------------------------------------------------
    s4 = created_slides[2]
    add_header(s4, "Self-Attention: სიტყვების „ყურადღების პროჟექტორი“", "მექანიზმი • როგორ იგებს AI რეალურ კონტექსტს")

    # Top intro card
    top_y = Inches(2.35)
    tb_intro = s4.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(1.0), top_y, Inches(17.0), Inches(1.4))
    tb_intro.fill.solid(); tb_intro.fill.fore_color.rgb = COLOR_BG_LIGHT
    tb_intro.line.color.rgb = COLOR_BORDER; tb_intro.line.width = Pt(1.2)

    tbox_i = s4.shapes.add_textbox(Inches(1.4), top_y + Inches(0.2), Inches(16.2), Inches(1.0))
    tfi = tbox_i.text_frame; tfi.word_wrap = True; tfi.margin_left = tfi.margin_top = 0
    p = tfi.paragraphs[0]; p.text = "🎯 პრობლემა: ადამიანურ ენაში ერთი და იგივე სიტყვა კონტექსტის მიხედვით სრულიად განსხვავებულ რამეს ნიშნავს."
    p.font.name = FONT_ARIAL; p.font.size = Pt(14); p.font.bold = True; p.font.color.rgb = COLOR_RED
    p = tfi.add_paragraph(); p.text = "Self-Attention მექანიზმი წინადადების ყოველ სიტყვას აძლევს „ყურადღების პროჟექტორს“, რათა გაარკვიოს, რომელი მეზობელი სიტყვები ხსნის მის აზრს."
    p.font.name = FONT_ARIAL; p.font.size = Pt(13); p.font.color.rgb = COLOR_TEXT_MUTED; p.space_before = Pt(3)

    # Two comparison cards
    card_y = Inches(4.0); card_h = Inches(4.1); card_w = Inches(8.35)

    # Case 1: Finance Bank
    c1 = s4.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(1.0), card_y, card_w, card_h)
    c1.fill.solid(); c1.fill.fore_color.rgb = COLOR_CARD_RED
    c1.line.color.rgb = COLOR_BORDER_RED; c1.line.width = Pt(1.5)

    tbc1 = s4.shapes.add_textbox(Inches(1.4), card_y + Inches(0.35), card_w - Inches(0.8), card_h - Inches(0.7))
    tfc1 = tbc1.text_frame; tfc1.word_wrap = True; tfc1.margin_left = tfc1.margin_top = 0
    p = tfc1.paragraphs[0]; p.text = "ᲛᲐᲒᲐᲚᲘᲗᲘ 1 • ᲡᲐᲤᲘᲜᲐᲜᲡᲝ ᲙᲝᲜᲢᲔᲥᲡᲢᲘ"
    p.font.name = FONT_ARIAL; p.font.size = Pt(12); p.font.bold = True; p.font.color.rgb = COLOR_RED

    p = tfc1.add_paragraph(); p.text = "„მდინარის პირას ბანკი იდგა, სადაც ფული შევიტანე“"
    p.font.name = FONT_ARIAL; p.font.size = Pt(17); p.font.bold = True; p.font.color.rgb = COLOR_DARK; p.space_before = Pt(6)

    p = tfc1.add_paragraph(); p.text = "\n🔍 ყურადღების პროჟექტორის ფოკუსი:"
    p.font.name = FONT_ARIAL; p.font.size = Pt(14); p.font.bold = True; p.font.color.rgb = COLOR_DARK
    p = tfc1.add_paragraph(); p.text = "სიტყვა „ბანკი“ ძლიერ კავშირს ამყარებს სიტყვასთან „ფული“ და „შევიტანე“.\nკავშირი „მდინარესთან“ სუსტია.\n\n➡ AI ასკვნის: საუბარია საფინანსო ინსტიტუტზე."
    p.font.name = FONT_ARIAL; p.font.size = Pt(13); p.font.color.rgb = COLOR_TEXT_MUTED; p.space_before = Pt(4)

    # Case 2: River Bank
    c2 = s4.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(9.65), card_y, card_w, card_h)
    c2.fill.solid(); c2.fill.fore_color.rgb = COLOR_BLUE_BG
    c2.line.color.rgb = COLOR_LINE; c2.line.width = Pt(1.5)

    tbc2 = s4.shapes.add_textbox(Inches(10.05), card_y + Inches(0.35), card_w - Inches(0.8), card_h - Inches(0.7))
    tfc2 = tbc2.text_frame; tfc2.word_wrap = True; tfc2.margin_left = tfc2.margin_top = 0
    p = tfc2.paragraphs[0]; p.text = "ᲛᲐᲒᲐᲚᲘᲗᲘ 2 • ᲑᲣᲜᲔᲑᲠᲘᲕᲘ ᲙᲝᲜᲢᲔᲥᲡᲢᲘ"
    p.font.name = FONT_ARIAL; p.font.size = Pt(12); p.font.bold = True; p.font.color.rgb = COLOR_BLUE

    p = tfc2.add_paragraph(); p.text = "„მდინარის პირას ბანკი იდგა, სადაც თევზი დავიჭირე“"
    p.font.name = FONT_ARIAL; p.font.size = Pt(17); p.font.bold = True; p.font.color.rgb = COLOR_DARK; p.space_before = Pt(6)

    p = tfc2.add_paragraph(); p.text = "\n🔍 ყურადღების პროჟექტორის ფოკუსი:"
    p.font.name = FONT_ARIAL; p.font.size = Pt(14); p.font.bold = True; p.font.color.rgb = COLOR_DARK
    p = tfc2.add_paragraph(); p.text = "სიტყვა „ბანკი“ (Bank - ნაპირი) ძლიერ კავშირს ამყარებს სიტყვასთან „მდინარე“ და „თევზი“.\n\n➡ AI ასკვნის: საუბარია წყლის სანაპირო ზოლზე."
    p.font.name = FONT_ARIAL; p.font.size = Pt(13); p.font.color.rgb = COLOR_TEXT_MUTED; p.space_before = Pt(4)

    # Bottom summary box
    bot_y = Inches(8.35)
    bs = s4.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(1.0), bot_y, Inches(17.0), Inches(1.65))
    bs.fill.solid(); bs.fill.fore_color.rgb = COLOR_WHITE
    bs.line.color.rgb = COLOR_BORDER; bs.line.width = Pt(1.2)

    tb_s = s4.shapes.add_textbox(Inches(1.4), bot_y + Inches(0.25), Inches(16.2), Inches(1.2))
    tfs = tb_s.text_frame; tfs.word_wrap = True; tfs.margin_left = tfs.margin_top = 0
    p = tfs.paragraphs[0]; p.text = "⚡ რა ხდება სინამდვილეში?"
    p.font.name = FONT_ARIAL; p.font.size = Pt(14); p.font.bold = True; p.font.color.rgb = COLOR_DARK
    p = tfs.add_paragraph()
    p.text = "Self-Attention არ კითხულობს სიტყვებს იზოლირებულად. ის ყოველ სიტყვას აძლევს დინამიკურ მნიშვნელობას იმის მიხედვით, თუ რომელი სხვა სიტყვები იმყოფებიან მის გარშემო."
    p.font.name = FONT_ARIAL; p.font.size = Pt(13); p.font.color.rgb = COLOR_TEXT_MUTED; p.space_before = Pt(3)

    # -------------------------------------------------------------------------
    # SLIDE 5: Next-Token Prediction (How LLM actually works)
    # -------------------------------------------------------------------------
    s5 = created_slides[3]
    add_header(s5, "LLM-ის მთავარი ამოცანა: „მომდევნო სიტყვის გამოცნობა“", "პრინციპი • Next-Token Prediction")

    # Top Prompt box
    p_box = s5.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(1.0), Inches(2.35), Inches(17.0), Inches(2.0))
    p_box.fill.solid(); p_box.fill.fore_color.rgb = COLOR_CARD_RED
    p_box.line.color.rgb = COLOR_BORDER_RED; p_box.line.width = Pt(1.5)

    tb_p = s5.shapes.add_textbox(Inches(1.4), Inches(2.55), Inches(16.2), Inches(1.5))
    tf_p = tb_p.text_frame; tf_p.word_wrap = True; tf_p.margin_left = tf_p.margin_top = 0
    p = tf_p.paragraphs[0]; p.text = "ᲨᲔᲛᲐᲕᲐᲚᲘ ᲢᲔᲥᲡᲢᲘ (PROMPT)"
    p.font.name = FONT_ARIAL; p.font.size = Pt(12); p.font.bold = True; p.font.color.rgb = COLOR_RED

    p = tf_p.add_paragraph(); p.text = "„თბილისი არის საქართველოს [ ... ? ]“"
    p.font.name = FONT_ARIAL; p.font.size = Pt(26); p.font.bold = True; p.font.color.rgb = COLOR_DARK; p.space_before = Pt(4)

    p = tf_p.add_paragraph(); p.text = "მოდელი კითხულობს ამ ფრაზას და ითვლის ლექსიკონის ყველა შესაძლო სიტყვის გაგრძელების მათემატიკურ ალბათობას."
    p.font.name = FONT_ARIAL; p.font.size = Pt(13); p.font.color.rgb = COLOR_TEXT_MUTED; p.space_before = Pt(4)

    # Mid 3 Probability Cards
    pred_cards = [
        ("„დედაქალაქი“", "94.2% ალბათობა", "ყველაზე ლოგიკური და ხშირი კონტექსტური გაგრძელება", True),
        ("„უდიდესი ქალაქი“", "4.1% ალბათობა", "ფაქტობრივად მართალია, მაგრამ ნაკლებად გავრცელებული ფორმულირება", False),
        ("„ისტორიული ცენტრი“", "1.7% ალბათობა", "შესაძლოა გარკვეულ ისტორიულ კონტექსტში", False)
    ]
    card_w = Inches(5.4); card_gap = Inches(0.4); start_x = Inches(1.0); card_y = Inches(4.65)
    for i, (word, prob, desc, is_winner) in enumerate(pred_cards):
        cx = start_x + i * (card_w + card_gap)

        c = s5.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, cx, card_y, card_w, Inches(3.1))
        c.fill.solid()
        c.fill.fore_color.rgb = COLOR_GREEN_BG if is_winner else COLOR_BG_LIGHT
        c.line.color.rgb = COLOR_GREEN if is_winner else COLOR_BORDER
        c.line.width = Pt(1.8 if is_winner else 1.2)

        tb = s5.shapes.add_textbox(cx + Inches(0.4), card_y + Inches(0.3), card_w - Inches(0.8), Inches(2.5))
        tf = tb.text_frame; tf.word_wrap = True; tf.margin_left = tf.margin_top = 0

        p = tf.paragraphs[0]; p.text = "✅ ᲐᲠᲩᲔᲣᲚᲘ ᲢᲝᲙᲔᲜᲘ" if is_winner else "ᲐᲚᲢᲔᲠᲜᲐᲢᲘᲕᲐ"
        p.font.name = FONT_ARIAL; p.font.size = Pt(11); p.font.bold = True
        p.font.color.rgb = COLOR_GREEN if is_winner else COLOR_MUTED

        p = tf.add_paragraph(); p.text = word
        p.font.name = FONT_ARIAL; p.font.size = Pt(24); p.font.bold = True; p.font.color.rgb = COLOR_DARK; p.space_before = Pt(4)

        p = tf.add_paragraph(); p.text = prob
        p.font.name = FONT_ARIAL; p.font.size = Pt(15); p.font.bold = True
        p.font.color.rgb = COLOR_GREEN if is_winner else COLOR_MUTED; p.space_before = Pt(3)

        p = tf.add_paragraph(); p.text = desc
        p.font.name = FONT_ARIAL; p.font.size = Pt(12); p.font.color.rgb = COLOR_TEXT_MUTED; p.space_before = Pt(6)

    # Bottom takeaway
    bot_y = Inches(8.05)
    bb = s5.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(1.0), bot_y, Inches(17.0), Inches(1.95))
    bb.fill.solid(); bb.fill.fore_color.rgb = COLOR_WHITE
    bb.line.color.rgb = COLOR_BORDER; bb.line.width = Pt(1.2)

    tbb = s5.shapes.add_textbox(Inches(1.4), bot_y + Inches(0.25), Inches(16.2), Inches(1.4))
    tfb = tbb.text_frame; tfb.word_wrap = True; tfb.margin_left = tfb.margin_top = 0
    p = tfb.paragraphs[0]; p.text = "💡 რატომ არის ეს მნიშვნელოვანი?"
    p.font.name = FONT_ARIAL; p.font.size = Pt(14); p.font.bold = True; p.font.color.rgb = COLOR_RED
    p = tfb.add_paragraph()
    p.text = "ChatGPT, Claude და Gemini არ „გეგმავენ“ წინასწარ მთლიან წერილს. ისინი გენერირებას ახდენენ ტოკენ-ტოკენ (სიტყვა-სიტყვა), ყოველ ნაბიჯზე წინა კონტექსტის გათვალისწინებით. სწორედ ამიტომ ვხედავთ ეკრანზე ტექსტის ცოცხალ აკრეფას."
    p.font.name = FONT_ARIAL; p.font.size = Pt(13); p.font.color.rgb = COLOR_DARK; p.space_before = Pt(4)

    # -------------------------------------------------------------------------
    # SLIDE 6: The 3 Stages of Model Training (Overview Timeline)
    # -------------------------------------------------------------------------
    s6 = created_slides[4]
    add_header(s6, "როგორ იქმნება AI მოდელი: წვრთნის 3 საფეხურიანი გზა", "არქიტექტურა • ნედლი მონაცემებიდან ჭკვიან ასისტენტამდე")

    stages = [
        {
            "step": "ᲔᲢᲐᲞᲘ 01",
            "name": "Pre-training",
            "geo": "წინასწარი წვრთნა",
            "desc": "მთელი ინტერნეტის წაკითხვა (ტრილიონობით ტოკენი). მოდელი სწავლობს ენის გრამატიკას, სამყაროს ფაქტებსა და ლოგიკურ კავშირებს.",
            "status": "შედეგი: Base Model (ნედლი მოდელი)",
            "color": COLOR_CARD_RED, "bcolor": COLOR_BORDER_RED, "tcolor": COLOR_RED
        },
        {
            "step": "ᲔᲢᲐᲞᲘ 02",
            "name": "Fine-Tuning (SFT)",
            "geo": "სპეციალიზაცია და დიალოგი",
            "desc": "მოდელს ასწავლიან, როგორ იყოს ასისტენტი. ექსპერტების მიერ შედგენილი ათასობით ხარისხიანი „კითხვა-პასუხის“ დიალოგის გააზრება.",
            "status": "შედეგი: Instruct / Chat Model",
            "color": COLOR_BG_LIGHT, "bcolor": COLOR_BORDER, "tcolor": COLOR_MUTED
        },
        {
            "step": "ᲔᲢᲐᲞᲘ 03",
            "name": "RLHF",
            "geo": "ადამიანური უკუკავშირი",
            "desc": "ადამიანები აფასებენ და ადარებენ მოდელის პასუხებს. მოდელი სწავლობს ეთიკურ საზღვრებს, უსაფრთხოებასა და დამხმარე ტონს.",
            "status": "შედეგი: Production AI (ChatGPT, Claude)",
            "color": COLOR_BG_LIGHT, "bcolor": COLOR_BORDER, "tcolor": COLOR_MUTED
        }
    ]

    card_w = Inches(5.4); card_gap = Inches(0.4); start_x = Inches(1.0); card_y = Inches(2.35)
    for i, st in enumerate(stages):
        cx = start_x + i * (card_w + card_gap)

        c = s6.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, cx, card_y, card_w, Inches(7.7))
        c.fill.solid(); c.fill.fore_color.rgb = st["color"]
        c.line.color.rgb = st["bcolor"]; c.line.width = Pt(1.5)

        tb = s6.shapes.add_textbox(cx + Inches(0.4), card_y + Inches(0.4), card_w - Inches(0.8), Inches(6.9))
        tf = tb.text_frame; tf.word_wrap = True; tf.margin_left = tf.margin_top = 0

        p = tf.paragraphs[0]; p.text = st["step"]
        p.font.name = FONT_ARIAL; p.font.size = Pt(13); p.font.bold = True; p.font.color.rgb = st["tcolor"]

        p = tf.add_paragraph(); p.text = st["name"]
        p.font.name = FONT_ARIAL; p.font.size = Pt(26); p.font.bold = True; p.font.color.rgb = COLOR_DARK; p.space_before = Pt(4)

        p = tf.add_paragraph(); p.text = st["geo"]
        p.font.name = FONT_ARIAL; p.font.size = Pt(15); p.font.bold = True; p.font.color.rgb = COLOR_MUTED; p.space_before = Pt(2)

        p = tf.add_paragraph(); p.text = f"\n{st['desc']}"
        p.font.name = FONT_ARIAL; p.font.size = Pt(14); p.font.color.rgb = COLOR_TEXT_MUTED; p.space_before = Pt(8)

        # Status badge inside card
        p = tf.add_paragraph(); p.text = f"\n🎯 {st['status']}"
        p.font.name = FONT_ARIAL; p.font.size = Pt(13); p.font.bold = True; p.font.color.rgb = COLOR_DARK; p.space_before = Pt(12)

    # -------------------------------------------------------------------------
    # SLIDE 7: Stage 1 - Pre-training in Depth (Massive Compute & Web Scale)
    # -------------------------------------------------------------------------
    s7 = created_slides[5]
    add_header(s7, "ეტაპი 1: Pre-training — ინტერნეტის ბიბლიოთეკა და სუპერკომპიუტერები", "მასშტაბი • როგორ იბადება ნედლი მოდელი (Base Model)")

    pillars = [
        {
            "tag": "01 • ᲛᲝᲜᲐᲪᲔᲛᲔᲑᲘ",
            "title": "ტრილიონობით ტექსტი",
            "points": [
                "Wikipedia, ციფრული წიგნები, სამეცნიერო სტატიები და ღია ვებგვერდები.",
                "ათობით მილიარდი კოდის ხაზი (GitHub).",
                "მრავალენოვანი ლიტერატურა.",
                "მონაცემების ფილტრაცია და გასუფთავება უხარისხო სპამისგან."
            ]
        },
        {
            "tag": "02 • ᲒᲐᲛᲝᲗᲕᲚᲘᲗᲘ ᲡᲘᲛᲫᲚᲐᲕᲠᲔ",
            "title": "სუპერკომპიუტერების კლასტერი",
            "points": [
                "10,000-დან 50,000-მდე Nvidia H100 GPU ერთობლივ ქსელში.",
                "თვეების განმავლობაში უწყვეტი გამოთვლები.",
                "ათობით მილიონი დოლარის ელექტროენერგია და ინფრასტრუქტურა.",
                "პარამეტრების რაოდენობა: 8 მილიარდიდან 1 ტრილიონამდე."
            ]
        },
        {
            "tag": "03 • ᲨᲔᲓᲔᲒᲘ",
            "title": "Base Model (ნედლი მოდელი)",
            "points": [
                "იცის ფაქტები და გრამატიკა უმაღლეს დონეზე.",
                "❌ არ შეუძლია ადამიანური დიალოგი (მხოლოდ ტექსტს აგრძელებს).",
                "მაგალითად: კითხვაზე „რა არის მზე?“, პასუხის ნაცვლად შეიძლება დაწეროს „რა არის მთვარე?“ (როგორც ტესტის გაგრძელება).",
                "საჭიროებს შემდგომ დამუშავებას."
            ]
        }
    ]

    card_w = Inches(5.4); card_gap = Inches(0.4); start_x = Inches(1.0); card_y = Inches(2.35)
    for i, pil in enumerate(pillars):
        cx = start_x + i * (card_w + card_gap)

        c = s7.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, cx, card_y, card_w, Inches(7.7))
        c.fill.solid()
        c.fill.fore_color.rgb = COLOR_CARD_RED if i == 0 else COLOR_BG_LIGHT
        c.line.color.rgb = COLOR_BORDER_RED if i == 0 else COLOR_BORDER
        c.line.width = Pt(1.5)

        tb = s7.shapes.add_textbox(cx + Inches(0.4), card_y + Inches(0.4), card_w - Inches(0.8), Inches(6.9))
        tf = tb.text_frame; tf.word_wrap = True; tf.margin_left = tf.margin_top = 0

        p = tf.paragraphs[0]; p.text = pil["tag"]
        p.font.name = FONT_ARIAL; p.font.size = Pt(12); p.font.bold = True
        p.font.color.rgb = COLOR_RED if i == 0 else COLOR_MUTED

        p = tf.add_paragraph(); p.text = pil["title"]
        p.font.name = FONT_ARIAL; p.font.size = Pt(22); p.font.bold = True; p.font.color.rgb = COLOR_DARK; p.space_before = Pt(4)

        for pt in pil["points"]:
            p = tf.add_paragraph(); p.text = f"• {pt}"
            p.font.name = FONT_ARIAL; p.font.size = Pt(13); p.font.color.rgb = COLOR_TEXT_MUTED; p.space_before = Pt(8)

    # -------------------------------------------------------------------------
    # SLIDE 8: Stages 2 & 3 - Fine-Tuning & RLHF (Becoming an Assistant)
    # -------------------------------------------------------------------------
    s8 = created_slides[6]
    add_header(s8, "Fine-tuning და RLHF: როგორ ხდება მოდელი „ჭკვიანი ასისტენტი“?", "დახვეწა • ნედლი მოდელიდან ChatGPT-მდე")

    col_w = Inches(8.35); col_h = Inches(5.4); col_y = Inches(2.35)

    # Left: SFT
    b1 = s8.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(1.0), col_y, col_w, col_h)
    b1.fill.solid(); b1.fill.fore_color.rgb = COLOR_BG_LIGHT
    b1.line.color.rgb = COLOR_BORDER; b1.line.width = Pt(1.5)

    tb1 = s8.shapes.add_textbox(Inches(1.4), col_y + Inches(0.4), col_w - Inches(0.8), col_h - Inches(0.8))
    tf1 = tb1.text_frame; tf1.word_wrap = True; tf1.margin_left = tf1.margin_top = 0

    p = tf1.paragraphs[0]; p.text = "ᲔᲢᲐᲞᲘ 02 • SFT (SUPERVISED FINE-TUNING)"
    p.font.name = FONT_ARIAL; p.font.size = Pt(12); p.font.bold = True; p.font.color.rgb = COLOR_MUTED

    p = tf1.add_paragraph(); p.text = "დიალოგისა და ინსტრუქციის სწავლება"
    p.font.name = FONT_ARIAL; p.font.size = Pt(22); p.font.bold = True; p.font.color.rgb = COLOR_DARK; p.space_before = Pt(4)

    sft_desc = [
        ("როგორ მუშაობს?", "ადამიან ექსპერტებს ხელით შეჰყავთ ათასობით სამაგალითო დიალოგი: „მომხმარებელი: დამიწერე წერილი...“ -> „ასისტენტი: გამარჯობა, აი თქვენი წერილი...“."),
        ("რას ცვლის?", "მოდელი ხვდება, რომ ის არის დამხმარე ასისტენტი და კითხვას უნდა გასცეს ლოგიკური, ზრდილობიანი და სტრუქტურირებული პასუხი."),
        ("შედეგი:", "Instruct მოდელი (რომელიც ასრულებს დავალებებს და პასუხობს კითხვებს).")
    ]
    for h, d in sft_desc:
        p = tf1.add_paragraph(); p.text = f"\n📌 {h} {d}"
        p.font.name = FONT_ARIAL; p.font.size = Pt(13); p.font.color.rgb = COLOR_TEXT_MUTED; p.space_before = Pt(6)

    # Right: RLHF
    b2 = s8.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(9.65), col_y, col_w, col_h)
    b2.fill.solid(); b2.fill.fore_color.rgb = COLOR_CARD_RED
    b2.line.color.rgb = COLOR_BORDER_RED; b2.line.width = Pt(1.5)

    tb2 = s8.shapes.add_textbox(Inches(10.05), col_y + Inches(0.4), col_w - Inches(0.8), col_h - Inches(0.8))
    tf2 = tb2.text_frame; tf2.word_wrap = True; tf2.margin_left = tf2.margin_top = 0

    p = tf2.paragraphs[0]; p.text = "ᲔᲢᲐᲞᲘ 03 • RLHF (HUMAN FEEDBACK)"
    p.font.name = FONT_ARIAL; p.font.size = Pt(12); p.font.bold = True; p.font.color.rgb = COLOR_RED

    p = tf2.add_paragraph(); p.text = "ადამიანური შეფასება და ეთიკა"
    p.font.name = FONT_ARIAL; p.font.size = Pt(22); p.font.bold = True; p.font.color.rgb = COLOR_DARK; p.space_before = Pt(4)

    rlhf_desc = [
        ("როგორ მუშაობს?", "მოდელი აგენერირებს 2 განსხვავებულ პასუხს. ადამიანი-შემფასებელი ადარებს მათ და ირჩევს უკეთესს, უფრო უსაფრთხოს და ზუსტს."),
        ("ჯილდოს სისტემა (Reward):", "მოდელი იღებს მათემატიკურ „ქულას/ჯილდოს“ სწორი ქცევისთვის და ისჯება შეცდომების, ჰალუცინაციების ან აგრესიის შემთხვევაში."),
        ("შედეგი:", "უსაფრთხო, ეთიკური და სანდო პროდუქტი (ChatGPT, Claude, Gemini).")
    ]
    for h, d in rlhf_desc:
        p = tf2.add_paragraph(); p.text = f"\n⭐ {h} {d}"
        p.font.name = FONT_ARIAL; p.font.size = Pt(13); p.font.color.rgb = COLOR_TEXT_MUTED; p.space_before = Pt(6)

    # Bottom comparison summary
    bot_y = Inches(8.0)
    bb = s8.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(1.0), bot_y, Inches(17.0), Inches(2.0))
    bb.fill.solid(); bb.fill.fore_color.rgb = COLOR_WHITE
    bb.line.color.rgb = COLOR_BORDER; bb.line.width = Pt(1.2)

    tbb = s8.shapes.add_textbox(Inches(1.4), bot_y + Inches(0.3), Inches(16.2), Inches(1.4))
    tfb = tbb.text_frame; tfb.word_wrap = True; tfb.margin_left = tfb.margin_top = 0
    p = tfb.paragraphs[0]; p.text = "🔑 შეჯამება: რატომ არ არის საკმარისი მხოლოდ Pre-training?"
    p.font.name = FONT_ARIAL; p.font.size = Pt(14); p.font.bold = True; p.font.color.rgb = COLOR_RED
    p = tfb.add_paragraph()
    p.text = "Pre-training მოდელს აძლევს უსაზღვრო ცოდნას (ინტელექტს), ხოლო Fine-Tuning და RLHF ასწავლის მას კულტურას, უსაფრთხოებას და ადამიანთან ეფექტური კომუნიკაციის წესებს."
    p.font.name = FONT_ARIAL; p.font.size = Pt(13); p.font.color.rgb = COLOR_DARK; p.space_before = Pt(4)

    # -------------------------------------------------------------------------
    # SLIDE 9: Open-Weights vs Proprietary (Products & Architecture Landscape)
    # -------------------------------------------------------------------------
    s9 = created_slides[7]
    add_header(s9, "თანამედროვე LLM ეკოსისტემა: დახურული vs ღია მოდელები", "პროდუქტები • ვის რა მოდელი აქვს ბაზარზე")

    eco_columns = [
        {
            "cat": "01 • ᲓᲐᲮᲣᲠᲣᲚᲘ (PROPRIETARY)",
            "title": "Frontier Models",
            "sub": "გიგანტი კომპანიების უძლიერესი API-ები",
            "tools": [
                ("openai.png", "GPT-4o (OpenAI)", "მულტიმოდალური წამყვანი მოდელი"),
                ("claude.png", "Claude 3.5 Sonnet", "ლოგიკისა და კოდინგის ლიდერი"),
                ("gemini.png", "Gemini 1.5 Pro", "2 მილიონი ტოკენის გიგანტური კონტექსტი")
            ]
        },
        {
            "cat": "02 • ᲦᲘᲐ ᲬᲝᲜᲔᲑᲘ (OPEN WEIGHTS)",
            "title": "Open Models",
            "sub": "საკუთარ სერვერზე გაშვებადი მოდელები",
            "tools": [
                ("meta_llama.png", "Llama 3.1 (Meta)", "ღია სამყაროს ოქროს სტანდარტი"),
                ("mistral.png", "Mistral Large", "ევროპული ღია წამყვანი მოდელი"),
                ("deepseek.png", "DeepSeek V2.5 / V3", "მაღალეფექტური ღია არქიტექტურა")
            ]
        },
        {
            "cat": "03 • ᲚᲝᲙᲐᲚᲣᲠᲘ ᲒᲐᲨᲕᲔᲑᲐ",
            "title": "Local Runners",
            "sub": "მოდელების გაშვება საკუთარ ლეპტოპზე",
            "tools": [
                ("ollama.png", "Ollama", "მოდელების გაშვება 1 ბრძანებით"),
                ("lmstudio.png", "LM Studio", "ვიზუალური ინტერფეისი ლოკალური AI-სთვის")
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
        card_h = Inches(1.5) if num_tools == 3 else Inches(1.9)
        spacing = Inches(1.65) if num_tools == 3 else Inches(2.1)
        base_y = Inches(4.3) if num_tools == 3 else Inches(4.5)

        for j, (l_file, t_name, t_desc) in enumerate(col["tools"]):
            item_y = base_y + j * spacing

            t_card = s9.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, cx + Inches(0.3), item_y, col_w - Inches(0.6), card_h)
            t_card.fill.solid(); t_card.fill.fore_color.rgb = COLOR_WHITE
            t_card.line.color.rgb = COLOR_BORDER_RED if i == 0 else COLOR_BORDER
            t_card.line.width = Pt(1.2)

            # Square Logo 0.95" x 0.95" (or 1.1" if 2 tools)
            lp = os.path.join(logo_dir, l_file)
            logo_sz = Inches(0.95) if num_tools == 3 else Inches(1.1)
            logo_y = item_y + (card_h - logo_sz) / 2
            if os.path.exists(lp):
                s9.shapes.add_picture(lp, cx + Inches(0.5), logo_y, width=logo_sz, height=logo_sz)

            tb_item = s9.shapes.add_textbox(cx + Inches(1.6), item_y + Inches(0.2), col_w - Inches(2.0), card_h - Inches(0.4))
            tf = tb_item.text_frame; tf.word_wrap = True; tf.margin_left = tf.margin_top = 0

            p = tf.paragraphs[0]; p.text = t_name
            p.font.name = FONT_ARIAL; p.font.size = Pt(14); p.font.bold = True; p.font.color.rgb = COLOR_DARK

            p = tf.add_paragraph(); p.text = t_desc
            p.font.name = FONT_ARIAL; p.font.size = Pt(11); p.font.color.rgb = COLOR_TEXT_MUTED; p.space_before = Pt(3)

    # -------------------------------------------------------------------------
    # SAVE PPTX AND DISTRIBUTE TO TARGET DIRECTORIES
    # -------------------------------------------------------------------------
    filename = "2.2 - LLM არქიტექტურა - Transformer, Attention და მოდელის წვრთნა.pptx"
    out_master = os.path.join(r'd:\ხელოვნური ინტელექტის ინსტრუმენტები', filename)
    prs.save(out_master)
    print("Master PPTX for Lecture 2.2 generated successfully!")

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
    build_lecture_2_2()
