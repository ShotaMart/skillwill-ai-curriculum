import os, shutil
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.enum.shapes import MSO_SHAPE

def build_lecture_2_1():
    template_path = r'd:\ხელოვნური ინტელექტის ინსტრუმენტები\Skillwill Presentation Example .pptx'
    prs = Presentation(template_path)

    asset_dir = r'd:\ხელოვნური ინტელექტის ინსტრუმენტები\assets'
    logo_dir = os.path.join(asset_dir, 'logos')
    img_icon_title = os.path.join(r'd:\ხელოვნური ინტელექტის ინსტრუმენტები\.template_assets', 'image5.png')

    COLOR_RED = RGBColor(213, 43, 30)         # #D52B1E
    COLOR_DARK = RGBColor(17, 24, 39)          # #111827
    COLOR_MUTED = RGBColor(107, 114, 128)      # #6B7280
    COLOR_BORDER = RGBColor(226, 232, 240)     # #E2E8F0
    COLOR_WHITE = RGBColor(255, 255, 255)
    COLOR_BG_LIGHT = RGBColor(248, 250, 252)
    COLOR_CARD_RED = RGBColor(254, 242, 242)
    COLOR_BORDER_RED = RGBColor(254, 202, 202)
    COLOR_TEXT_MUTED = RGBColor(75, 85, 99)
    COLOR_DARK_BOX = RGBColor(15, 23, 42)
    COLOR_LIGHT_RED = RGBColor(252, 165, 165)
    COLOR_SLATE_LIGHT = RGBColor(148, 163, 184)
    COLOR_GREEN_BG = RGBColor(240, 253, 244)
    COLOR_GREEN_BORDER = RGBColor(187, 247, 208)

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
            p.text = "როგორ კითხულობს AI: ტოკენები, Embeddings და ჭკვიანი ძებნა"
            p.alignment = PP_ALIGN.CENTER
            p.font.name = FONT_ARIAL
            p.font.size = Pt(40)
            p.font.bold = True
            p.font.color.rgb = COLOR_RED
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
    def add_header(slide, title, category="ᲛᲝᲓᲣᲚᲘ 5 • ᲨᲔᲡᲐᲕᲐᲚᲘ AI"):
        if os.path.exists(img_icon_title):
            slide.shapes.add_picture(img_icon_title, Inches(1.0), Inches(0.85), width=Inches(0.42), height=Inches(0.45))

        tb = slide.shapes.add_textbox(Inches(1.55), Inches(0.6), Inches(17.0), Inches(1.2))
        tf = tb.text_frame
        tf.word_wrap = True
        tf.margin_left = tf.margin_top = tf.margin_right = tf.margin_bottom = 0

        p_cat = tf.paragraphs[0]
        p_cat.text = category.upper()
        p_cat.font.name = FONT_ARIAL
        p_cat.font.size = Pt(13)
        p_cat.font.bold = True
        p_cat.font.color.rgb = COLOR_RED

        p_t = tf.add_paragraph()
        p_t.text = title
        p_t.font.name = FONT_ARIAL
        p_t.font.size = Pt(30)
        p_t.font.bold = True
        p_t.font.color.rgb = COLOR_DARK
        p_t.space_before = Pt(4)

        line = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(1.0), Inches(1.95), Inches(18.0), Inches(0.02))
        line.fill.solid()
        line.fill.fore_color.rgb = COLOR_BORDER
        line.line.color.rgb = COLOR_BORDER

    # -------------------------------------------------------------------------
    # SLIDE 2: Why Computers Don't Understand Text
    # -------------------------------------------------------------------------
    add_header(s2, "რატომ არ ესმის კომპიუტერს ტექსტი?", "საფუძვლები • პარადიგმის ცვლილება")

    bl = s2.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(1.0), Inches(2.35), Inches(8.7), Inches(7.7))
    bl.fill.solid(); bl.fill.fore_color.rgb = COLOR_BG_LIGHT
    bl.line.color.rgb = COLOR_BORDER; bl.line.width = Pt(1.5)

    tb_l_head = s2.shapes.add_textbox(Inches(1.4), Inches(2.6), Inches(7.9), Inches(1.1))
    tf_lh = tb_l_head.text_frame; tf_lh.word_wrap = True; tf_lh.margin_left = tf_lh.margin_top = 0
    p = tf_lh.paragraphs[0]; p.text = "ᲐᲓᲐᲛᲘᲐᲜᲘᲡ ᲐᲦᲥᲛᲐ"
    p.font.name = FONT_ARIAL; p.font.size = Pt(13); p.font.bold = True; p.font.color.rgb = COLOR_MUTED
    p = tf_lh.add_paragraph(); p.text = "ტექსტური და ვიზუალური გააზრება"
    p.font.name = FONT_ARIAL; p.font.size = Pt(21); p.font.bold = True; p.font.color.rgb = COLOR_DARK; p.space_before = Pt(4)

    f_box1 = s2.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(1.4), Inches(3.8), Inches(7.9), Inches(1.2))
    f_box1.fill.solid(); f_box1.fill.fore_color.rgb = COLOR_WHITE
    f_box1.line.color.rgb = COLOR_BORDER; f_box1.line.width = Pt(1)
    tf_f1 = f_box1.text_frame; tf_f1.vertical_anchor = MSO_ANCHOR.MIDDLE
    p = tf_f1.paragraphs[0]; p.text = "სიტყვა („ვაშლი“)  ➔  გონებრივი ხატი / ასოციაცია"
    p.alignment = PP_ALIGN.CENTER; p.font.name = FONT_ARIAL; p.font.size = Pt(15); p.font.bold = True; p.font.color.rgb = COLOR_RED

    desc_l = s2.shapes.add_textbox(Inches(1.4), Inches(5.3), Inches(7.9), Inches(4.3))
    tf_dl = desc_l.text_frame; tf_dl.word_wrap = True; tf_dl.margin_left = tf_dl.margin_top = 0
    p = tf_dl.paragraphs[0]
    p.text = (
        "• ბუნებრივი ენა: ადამიანი კითხულობს სიტყვებს და მყისიერად გრძნობს შინაარსსა და ემოციას.\n\n"
        "• რეალური სამყარო: სიტყვა „ვაშლი“ ჩვენთვის არის წითელი, ტკბილი, წვნიანი ხილი.\n\n"
        "• პირდაპირი გაგება: ადამიანს არ სჭირდება სპეციალური კოდირება აზრის გასაგებად."
    )
    p.font.name = FONT_ARIAL; p.font.size = Pt(16); p.font.color.rgb = COLOR_TEXT_MUTED

    br = s2.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(10.3), Inches(2.35), Inches(8.7), Inches(7.7))
    br.fill.solid(); br.fill.fore_color.rgb = COLOR_CARD_RED
    br.line.color.rgb = COLOR_BORDER_RED; br.line.width = Pt(2)

    tb_r_head = s2.shapes.add_textbox(Inches(10.7), Inches(2.6), Inches(7.9), Inches(1.1))
    tf_rh = tb_r_head.text_frame; tf_rh.word_wrap = True; tf_rh.margin_left = tf_rh.margin_top = 0
    p = tf_rh.paragraphs[0]; p.text = "ᲙᲝᲛᲞᲘᲣᲢᲔᲠᲘᲡ / AI-Ს ᲐᲦᲥᲛᲐ"
    p.font.name = FONT_ARIAL; p.font.size = Pt(13); p.font.bold = True; p.font.color.rgb = COLOR_RED
    p = tf_rh.add_paragraph(); p.text = "მხოლოდ რიცხვების კალკულატორი"
    p.font.name = FONT_ARIAL; p.font.size = Pt(21); p.font.bold = True; p.font.color.rgb = COLOR_DARK; p.space_before = Pt(4)

    f_box2 = s2.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(10.7), Inches(3.8), Inches(7.9), Inches(1.2))
    f_box2.fill.solid(); f_box2.fill.fore_color.rgb = COLOR_WHITE
    f_box2.line.color.rgb = COLOR_BORDER_RED; f_box2.line.width = Pt(1.5)
    tf_f2 = f_box2.text_frame; tf_f2.vertical_anchor = MSO_ANCHOR.MIDDLE
    p = tf_f2.paragraphs[0]; p.text = "ტექსტი  ➔  ტოკენები  ➔  რიცხვები (Numbers)"
    p.alignment = PP_ALIGN.CENTER; p.font.name = FONT_ARIAL; p.font.size = Pt(15); p.font.bold = True; p.font.color.rgb = COLOR_RED

    desc_r = s2.shapes.add_textbox(Inches(10.7), Inches(5.3), Inches(7.9), Inches(4.3))
    tf_dr = desc_r.text_frame; tf_dr.word_wrap = True; tf_dr.margin_left = tf_dr.margin_top = 0
    p = tf_dr.paragraphs[0]
    p.text = (
        "• მათემატიკური მანქანა: კომპიუტერს მხოლოდ ციფრები ესმის (პროცესორი მუშაობს 0-ებსა და 1-ებზე).\n\n"
        "• ასოები არაფერს ნიშნავს: სიტყვა „ვაშლი“ AI-სთვის უბრალოდ უაზრო სიმბოლოების ჯაჭვია.\n\n"
        "• გარდაქმნის აუცილებლობა: სანამ AI რამეს „გაიფიქრებს“, ტექსტი აუცილებლად უნდა დაიჭრას და გადაიქცეს რიცხვებად!"
    )
    p.font.name = FONT_ARIAL; p.font.size = Pt(16); p.font.color.rgb = COLOR_TEXT_MUTED

    # -------------------------------------------------------------------------
    # SLIDE 3: What is a Token? Words as Lego Bricks
    # -------------------------------------------------------------------------
    s3 = created_slides[1]
    add_header(s3, "რა არის ტოკენი? — „სიტყვების ლეგო“", "ცნებები • ტოკენიზაცია")

    pr_box = s3.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(1.0), Inches(2.35), Inches(18.0), Inches(1.8))
    pr_box.fill.solid(); pr_box.fill.fore_color.rgb = COLOR_DARK_BOX
    pr_box.line.color.rgb = COLOR_RED; pr_box.line.width = Pt(2)

    tb_def = s3.shapes.add_textbox(Inches(1.4), Inches(2.55), Inches(17.2), Inches(1.4))
    tf = tb_def.text_frame; tf.word_wrap = True; tf.margin_left = tf.margin_top = 0
    p = tf.paragraphs[0]; p.text = "🎯 ᲫᲘᲠᲘᲗᲐᲓᲘ ᲒᲐᲜᲛᲐᲠᲢᲔᲑᲐ (LEVEL 0):"
    p.font.name = FONT_ARIAL; p.font.size = Pt(13); p.font.bold = True; p.font.color.rgb = COLOR_RED

    p = tf.add_paragraph()
    p.text = "AI ტექსტს არ კითხულობს მთლიანი სიტყვებით და არც ცალკეული ასოებით. ის სიტყვებს ჭრის ყველაზე ხშირ მარცვლებად და ნაწილებად — ტოკენებად (როგორც ლეგოს კუბიკები)."
    p.font.name = FONT_ARIAL; p.font.size = Pt(18); p.font.color.rgb = COLOR_WHITE; p.space_before = Pt(4)

    token_cards = [
        {
            "tag": "მაგალითი 01 • რთული სიტყვა",
            "title": "როგორ იჭრება სიტყვა?",
            "word": "„ხელოვნური“",
            "tokens": "[ხელ]  +  [ოვნ]  +  [ური]",
            "desc": "ქართული გრძელი სიტყვები იშლება ნაცნობ მარცვლებად (3 ტოკენი/კუბიკი)."
        },
        {
            "tag": "მაგალითი 02 • მარტივი სიტყვა",
            "title": "ხშირი ინგლისური სიტყვები",
            "word": "„Apple“ / „Dog“",
            "tokens": "[Apple]  ➔  1 ტოკენი",
            "desc": "ხშირად გამოყენებული ინგლისური სიტყვები ერთი მთლიანი კუბიკია."
        },
        {
            "tag": "მაგალითი 03 • ახალი ჟარგონი",
            "title": "უცნობი სიტყვების აწყობა",
            "word": "„დაგუგლვა“",
            "tokens": "[და]  +  [გუგლ]  +  [ვა]",
            "desc": "AI-ს არასდროს ეშინია ახალი სიტყვის — ის მას ნაცნობი კუბიკებისგან ააწყობს."
        }
    ]

    tc_w = Inches(5.8); tc_gap = Inches(0.3); tc_start_x = Inches(1.0)
    for i, tc in enumerate(token_cards):
        cx = tc_start_x + i * (tc_w + tc_gap)
        cbox = s3.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, cx, Inches(4.45), tc_w, Inches(5.6))
        cbox.fill.solid(); cbox.fill.fore_color.rgb = COLOR_BG_LIGHT
        cbox.line.color.rgb = COLOR_BORDER; cbox.line.width = Pt(1.5)

        tb = s3.shapes.add_textbox(cx + Inches(0.4), Inches(4.7), tc_w - Inches(0.8), Inches(5.1))
        tf = tb.text_frame; tf.word_wrap = True; tf.margin_left = tf.margin_top = 0

        p = tf.paragraphs[0]; p.text = tc["tag"]
        p.font.name = FONT_ARIAL; p.font.size = Pt(11); p.font.bold = True; p.font.color.rgb = COLOR_RED

        p = tf.add_paragraph(); p.text = tc["title"]
        p.font.name = FONT_ARIAL; p.font.size = Pt(19); p.font.bold = True; p.font.color.rgb = COLOR_DARK; p.space_before = Pt(3)

        p = tf.add_paragraph(); p.text = f"\nსიტყვა: {tc['word']}"
        p.font.name = FONT_ARIAL; p.font.size = Pt(15); p.font.bold = True; p.font.color.rgb = COLOR_DARK; p.space_before = Pt(6)

        p = tf.add_paragraph(); p.text = f"დაჭრა: {tc['tokens']}"
        p.font.name = FONT_ARIAL; p.font.size = Pt(14); p.font.bold = True; p.font.color.rgb = COLOR_RED; p.space_before = Pt(4)

        p = tf.add_paragraph(); p.text = f"\n{tc['desc']}"
        p.font.name = FONT_ARIAL; p.font.size = Pt(13); p.font.color.rgb = COLOR_TEXT_MUTED; p.space_before = Pt(6)

    # -------------------------------------------------------------------------
    # SLIDE 4: Token Economics - Georgian vs English
    # -------------------------------------------------------------------------
    s4 = created_slides[2]
    add_header(s4, "ტოკენების ეკონომიკა: ქართული vs ინგლისური", "პრაქტიკა • ენობრივი სხვაობა")

    b_en = s4.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(1.0), Inches(2.35), Inches(8.7), Inches(2.6))
    b_en.fill.solid(); b_en.fill.fore_color.rgb = COLOR_WHITE; b_en.line.color.rgb = COLOR_BORDER; b_en.line.width = Pt(1.5)

    tb_en = s4.shapes.add_textbox(Inches(1.4), Inches(2.55), Inches(7.9), Inches(2.2))
    tf = tb_en.text_frame; tf.word_wrap = True; tf.margin_left = tf.margin_top = 0
    p = tf.paragraphs[0]; p.text = "ᲘᲜᲒᲚᲘᲡᲣᲠᲘ ᲔᲜᲐ (ENGLISH)"
    p.font.name = FONT_ARIAL; p.font.size = Pt(12); p.font.bold = True; p.font.color.rgb = COLOR_MUTED
    p = tf.add_paragraph(); p.text = "„Artificial Intelligence“  ➔  ზუსტად 2 ტოკენი"
    p.font.name = FONT_ARIAL; p.font.size = Pt(18); p.font.bold = True; p.font.color.rgb = COLOR_DARK; p.space_before = Pt(4)
    p = tf.add_paragraph()
    p.text = "• სიტყვები ერთიან კუბიკებად აღიქმება.\n• მინიმალური ხარჯი და მაქსიმალური სისწრაფე."
    p.font.name = FONT_ARIAL; p.font.size = Pt(14); p.font.color.rgb = COLOR_TEXT_MUTED; p.space_before = Pt(6)

    b_ge = s4.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(10.3), Inches(2.35), Inches(8.7), Inches(2.6))
    b_ge.fill.solid(); b_ge.fill.fore_color.rgb = COLOR_CARD_RED; b_ge.line.color.rgb = COLOR_BORDER_RED; b_ge.line.width = Pt(2)

    tb_ge = s4.shapes.add_textbox(Inches(10.7), Inches(2.55), Inches(7.9), Inches(2.2))
    tf = tb_ge.text_frame; tf.word_wrap = True; tf.margin_left = tf.margin_top = 0
    p = tf.paragraphs[0]; p.text = "ᲥᲐᲠᲗᲣᲚᲘ ᲔᲜᲐ (GEORGIAN)"
    p.font.name = FONT_ARIAL; p.font.size = Pt(12); p.font.bold = True; p.font.color.rgb = COLOR_RED
    p = tf.add_paragraph(); p.text = "„ხელოვნური ინტელექტი“  ➔  8–10 ტოკენი!"
    p.font.name = FONT_ARIAL; p.font.size = Pt(18); p.font.bold = True; p.font.color.rgb = COLOR_RED; p.space_before = Pt(4)
    p = tf.add_paragraph()
    p.text = "• ქართული ასოები 2-2 ან 3-3 სიმბოლოდ იჭრება.\n• 3-4-ჯერ მეტი ტოკენი იმავე აზრის გადმოსაცემად."
    p.font.name = FONT_ARIAL; p.font.size = Pt(14); p.font.color.rgb = COLOR_DARK; p.space_before = Pt(6)

    eco_cards = [
        {
            "num": "01",
            "title": "რატომ ხდება ასე?",
            "desc": "AI მოდელები ძირითადად ინგლისურ ტექსტებზეა გაწვრთნილი. ქართული ანბანი ტოკენიზატორისთვის „იშვიათია“, ამიტომ სიტყვებს წვრილ ნაწილებად ჭრის."
        },
        {
            "num": "02",
            "title": "ხარჯების სხვაობა (API Pricing)",
            "desc": "OpenAI, Anthropic და Google თანხას ითხოვენ ტოკენების მიხედვით. ქართულ ენაზე მომუშავე აპლიკაცია ბიზნესისთვის 3-4-ჯერ უფრო ძვირი ჯდება."
        },
        {
            "num": "03",
            "title": "სისწრაფე და მეხსიერება",
            "desc": "მოდელს ერთ ჯერზე შეუძლია შეზღუდული რაოდენობის ტოკენების წაკითხვა. ქართული ტექსტით მოდელის „მეხსიერების ფანჯარა“ ბევრად უფრო სწრაფად ივსება."
        }
    ]

    ec_w = Inches(5.8); ec_gap = Inches(0.3); ec_start_x = Inches(1.0)
    for i, ec in enumerate(eco_cards):
        cx = ec_start_x + i * (ec_w + ec_gap)
        eb = s4.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, cx, Inches(5.25), ec_w, Inches(4.8))
        eb.fill.solid(); eb.fill.fore_color.rgb = COLOR_BG_LIGHT
        eb.line.color.rgb = COLOR_BORDER; eb.line.width = Pt(1.5)

        tb = s4.shapes.add_textbox(cx + Inches(0.4), Inches(5.5), ec_w - Inches(0.8), Inches(4.3))
        tf = tb.text_frame; tf.word_wrap = True; tf.margin_left = tf.margin_top = 0

        p = tf.paragraphs[0]; p.text = f"პუნქტი {ec['num']}"
        p.font.name = FONT_ARIAL; p.font.size = Pt(12); p.font.bold = True; p.font.color.rgb = COLOR_RED

        p = tf.add_paragraph(); p.text = ec["title"]
        p.font.name = FONT_ARIAL; p.font.size = Pt(18); p.font.bold = True; p.font.color.rgb = COLOR_DARK; p.space_before = Pt(3)

        p = tf.add_paragraph(); p.text = f"\n{ec['desc']}"
        p.font.name = FONT_ARIAL; p.font.size = Pt(14); p.font.color.rgb = COLOR_TEXT_MUTED; p.space_before = Pt(6)

    # -------------------------------------------------------------------------
    # SLIDE 5: What is an Embedding? The GPS Map of Meanings
    # -------------------------------------------------------------------------
    s5 = created_slides[3]
    add_header(s5, "რა არის Embedding? — „აზრების GPS რუკა“", "მოდელი • მნიშვნელობის სივრცე")

    emb_top = s5.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(1.0), Inches(2.35), Inches(18.0), Inches(1.7))
    emb_top.fill.solid(); emb_top.fill.fore_color.rgb = COLOR_CARD_RED; emb_top.line.color.rgb = COLOR_BORDER_RED; emb_top.line.width = Pt(1.5)

    tb = s5.shapes.add_textbox(Inches(1.4), Inches(2.55), Inches(17.2), Inches(1.3))
    tf = tb.text_frame; tf.word_wrap = True; tf.margin_left = tf.margin_top = 0
    p = tf.paragraphs[0]; p.text = "🧭 ᲕᲘᲖᲣᲐᲚᲣᲠᲘ ᲐᲜᲐᲚᲝᲒᲘᲐ (GPS MAP):"
    p.font.name = FONT_ARIAL; p.font.size = Pt(13); p.font.bold = True; p.font.color.rgb = COLOR_RED

    p = tf.add_paragraph()
    p.text = "წარმოიდგინეთ ქალაქის GPS რუკა: კაფეები დგანან ერთად, რესტორნები — გვერდით, ხოლო აეროპორტი — შორს. Embedding არის სიტყვებისა და აზრების GPS რუკა — ყოველ სიტყვას აქვს ზუსტი კოორდინატი!"
    p.font.name = FONT_ARIAL; p.font.size = Pt(17); p.font.bold = True; p.font.color.rgb = COLOR_DARK; p.space_before = Pt(3)

    map_cards = [
        {
            "tag": "წერტილი A • ახლო მეზობლები",
            "title": "მსგავსი მნიშვნელობა = ახლო მანძილი",
            "pair": "„ძაღლი“  ↔  „ლეკვი“",
            "desc": "რუკაზე დგანან ზუსტად გვერდიგვერდ. მათი კოორდინატები თითქმის ემთხვევა, რადგან ერთი და იგივე არსებაა."
        },
        {
            "tag": "წერტილი B • დაკავშირებული ცნებები",
            "title": "საერთო თემატური კატეგორია",
            "pair": "„კატა“  ↔  „ძაღლი“",
            "desc": "რუკაზე ახლოს არიან, რადგან ორივე შინაური ცხოველია, თუმცა აქვთ მცირე დისტანცია."
        },
        {
            "tag": "წერტილი C • სრულიად უცხო უბანი",
            "title": "შორეული განსხვავებული აზრი",
            "pair": "„ტრაქტორი“  ↔  „ძაღლი“",
            "desc": "დგანან რუკის სულ სხვადასხვა ბოლოში. მათ შორის კავშირი ნულის ტოლია (შორი მანძილი)."
        }
    ]

    mc_w = Inches(5.8); mc_gap = Inches(0.3); mc_start_x = Inches(1.0)
    for i, mc in enumerate(map_cards):
        cx = mc_start_x + i * (mc_w + mc_gap)
        mcb = s5.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, cx, Inches(4.35), mc_w, Inches(5.7))
        mcb.fill.solid(); mcb.fill.fore_color.rgb = COLOR_WHITE
        mcb.line.color.rgb = COLOR_BORDER; mcb.line.width = Pt(1.5)

        tb = s5.shapes.add_textbox(cx + Inches(0.4), Inches(4.6), mc_w - Inches(0.8), Inches(5.2))
        tf = tb.text_frame; tf.word_wrap = True; tf.margin_left = tf.margin_top = 0

        p = tf.paragraphs[0]; p.text = mc["tag"]
        p.font.name = FONT_ARIAL; p.font.size = Pt(11); p.font.bold = True; p.font.color.rgb = COLOR_RED

        p = tf.add_paragraph(); p.text = mc["title"]
        p.font.name = FONT_ARIAL; p.font.size = Pt(18); p.font.bold = True; p.font.color.rgb = COLOR_DARK; p.space_before = Pt(3)

        p = tf.add_paragraph(); p.text = f"\n{mc['pair']}"
        p.font.name = FONT_ARIAL; p.font.size = Pt(18); p.font.bold = True; p.font.color.rgb = COLOR_RED; p.space_before = Pt(8)

        p = tf.add_paragraph(); p.text = f"\n{mc['desc']}"
        p.font.name = FONT_ARIAL; p.font.size = Pt(14); p.font.color.rgb = COLOR_TEXT_MUTED; p.space_before = Pt(8)

    # -------------------------------------------------------------------------
    # SLIDE 6: Meaning Arithmetic (Vector Math Made Simple)
    # -------------------------------------------------------------------------
    s6 = created_slides[4]
    add_header(s6, "აზრების შეკრება და გამოკლება", "მექანიზმი • ვექტორული ლოგიკა")

    f_hero = s6.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(1.0), Inches(2.35), Inches(18.0), Inches(2.4))
    f_hero.fill.solid(); f_hero.fill.fore_color.rgb = COLOR_DARK_BOX
    f_hero.line.color.rgb = COLOR_RED; f_hero.line.width = Pt(2.5)

    tb_h = s6.shapes.add_textbox(Inches(1.4), Inches(2.55), Inches(17.2), Inches(2.0))
    tf = tb_h.text_frame; tf.word_wrap = True; tf.margin_left = tf.margin_top = 0

    p = tf.paragraphs[0]; p.text = "👑 ᲪᲜᲝᲑᲘᲚᲘ ᲛᲐᲗᲔᲛᲐᲢᲘᲙᲣᲠᲘ ᲡᲐᲝᲪᲠᲔᲑᲐ (LEVEL 0):"
    p.font.name = FONT_ARIAL; p.font.size = Pt(13); p.font.bold = True; p.font.color.rgb = COLOR_RED

    p = tf.add_paragraph()
    p.text = "„მეფე“   -   „კაცი“   +   „ქალი“   =   „დედოფალი“"
    p.alignment = PP_ALIGN.CENTER; p.font.name = FONT_ARIAL; p.font.size = Pt(28); p.font.bold = True; p.font.color.rgb = COLOR_WHITE; p.space_before = Pt(8)

    p = tf.add_paragraph()
    p.text = "AI-მ იცის, რომ მეფის კოორდინატს თუ გამოვაკლებთ „კაცს“ (სქესს) და დავუმატებთ „ქალს“, რუკაზე მივიღებთ ზუსტად დედოფალს!"
    p.alignment = PP_ALIGN.CENTER; p.font.name = FONT_ARIAL; p.font.size = Pt(15); p.font.color.rgb = COLOR_SLATE_LIGHT; p.space_before = Pt(6)

    v_cards = [
        {
            "tag": "მიმართულება 01 • გეოგრაფია",
            "title": "ქვეყნები და დედაქალაქები",
            "formula": "„პარიზი“  -  „საფრანგეთი“  +  „საქართველო“  =  „თბილისი“",
            "desc": "AI-ს ესმის „დედაქალაქობის“ აზრობრივი მიმართულება. მან იცის, რა კავშირიცაა პარიზსა და საფრანგეთს შორის, ზუსტად იგივე კავშირია თბილისსა და საქართველოს შორის."
        },
        {
            "tag": "მიმართულება 02 • ასაკობრივი კავშირები",
            "title": "ასაკი და ზრდა",
            "formula": "„ბავშვი“  -  „პატარა“  +  „დიდი“  =  „ზრდასრული“",
            "desc": "ცნებებს შორის ურთიერთობები (ასაკი, ზომა, სქესი, თანამდებობა) მათემატიკურად იზომება, როგორც რუკაზე გადაადგილების კონკრეტული ვექტორები."
        }
    ]

    vc_w = Inches(8.8); vc_gap = Inches(0.4); vc_start_x = Inches(1.0)
    for i, vc in enumerate(v_cards):
        cx = vc_start_x + i * (vc_w + vc_gap)
        vcb = s6.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, cx, Inches(5.05), vc_w, Inches(5.0))
        vcb.fill.solid(); vcb.fill.fore_color.rgb = COLOR_BG_LIGHT
        vcb.line.color.rgb = COLOR_BORDER; vcb.line.width = Pt(1.5)

        tb = s6.shapes.add_textbox(cx + Inches(0.4), Inches(5.3), vc_w - Inches(0.8), Inches(4.5))
        tf = tb.text_frame; tf.word_wrap = True; tf.margin_left = tf.margin_top = 0

        p = tf.paragraphs[0]; p.text = vc["tag"]
        p.font.name = FONT_ARIAL; p.font.size = Pt(12); p.font.bold = True; p.font.color.rgb = COLOR_RED

        p = tf.add_paragraph(); p.text = vc["title"]
        p.font.name = FONT_ARIAL; p.font.size = Pt(19); p.font.bold = True; p.font.color.rgb = COLOR_DARK; p.space_before = Pt(3)

        p = tf.add_paragraph(); p.text = f"\n{vc['formula']}"
        p.font.name = FONT_ARIAL; p.font.size = Pt(16); p.font.bold = True; p.font.color.rgb = COLOR_RED; p.space_before = Pt(8)

        p = tf.add_paragraph(); p.text = f"\n{vc['desc']}"
        p.font.name = FONT_ARIAL; p.font.size = Pt(14); p.font.color.rgb = COLOR_TEXT_MUTED; p.space_before = Pt(6)

    # -------------------------------------------------------------------------
    # SLIDE 7: Semantic Search vs Traditional Search
    # -------------------------------------------------------------------------
    s7 = created_slides[5]
    add_header(s7, "ჭკვიანი ძებნა (Semantic Search) vs ჩვეულებრივი ძებნა", "ტექნოლოგია • შედარება")

    sl = s7.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(1.0), Inches(2.35), Inches(8.7), Inches(7.7))
    sl.fill.solid(); sl.fill.fore_color.rgb = COLOR_BG_LIGHT
    sl.line.color.rgb = COLOR_BORDER; sl.line.width = Pt(1.5)

    tb_sl = s7.shapes.add_textbox(Inches(1.4), Inches(2.6), Inches(7.9), Inches(1.1))
    tf = tb_sl.text_frame; tf.word_wrap = True; tf.margin_left = tf.margin_top = 0
    p = tf.paragraphs[0]; p.text = "ᲢᲠᲐᲓᲘᲪᲘᲣᲚᲘ ᲫᲔᲑᲜᲐ (CTRL + F)"
    p.font.name = FONT_ARIAL; p.font.size = Pt(13); p.font.bold = True; p.font.color.rgb = COLOR_MUTED
    p = tf.add_paragraph(); p.text = "ზუსტი ასოობრივი დამთხვევა (Exact Match)"
    p.font.name = FONT_ARIAL; p.font.size = Pt(21); p.font.bold = True; p.font.color.rgb = COLOR_DARK; p.space_before = Pt(4)

    fb_sl = s7.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(1.4), Inches(3.8), Inches(7.9), Inches(1.4))
    fb_sl.fill.solid(); fb_sl.fill.fore_color.rgb = COLOR_WHITE; fb_sl.line.color.rgb = COLOR_BORDER
    tf = fb_sl.text_frame; tf.vertical_anchor = MSO_ANCHOR.MIDDLE
    p = tf.paragraphs[0]; p.text = "თქვენ ეძებთ: „ავტომობილის ყიდვა“\nდოკუმენტში წერია: „მანქანის შეძენა“"
    p.alignment = PP_ALIGN.CENTER; p.font.name = FONT_ARIAL; p.font.size = Pt(14); p.font.bold = True; p.font.color.rgb = COLOR_DARK

    p = tf.add_paragraph(); p.text = "შედეგი: ❌ ვერაფერი მოიძებნა!"
    p.alignment = PP_ALIGN.CENTER; p.font.name = FONT_ARIAL; p.font.size = Pt(16); p.font.bold = True; p.font.color.rgb = COLOR_RED; p.space_before = Pt(4)

    desc_sl = s7.shapes.add_textbox(Inches(1.4), Inches(5.5), Inches(7.9), Inches(4.3))
    tf = desc_sl.text_frame; tf.word_wrap = True; tf.margin_left = tf.margin_top = 0
    p = tf.paragraphs[0]
    p.text = (
        "• ბრმა ძებნა: ეძებს მხოლოდ ასოების იდენტურ კომბინაციას.\n\n"
        "• სინონიმების იგნორირება: სისტემამ არ იცის, რომ „ავტომობილი“ და „მანქანა“ ერთი და იგივეა.\n\n"
        "• შეზღუდვა: თუ მომხმარებელმა ზუსტი ტერმინი არ იცის, საჭირო ინფორმაციას ვერასდროს იპოვის."
    )
    p.font.name = FONT_ARIAL; p.font.size = Pt(15); p.font.color.rgb = COLOR_TEXT_MUTED

    sr = s7.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(10.3), Inches(2.35), Inches(8.7), Inches(7.7))
    sr.fill.solid(); sr.fill.fore_color.rgb = COLOR_CARD_RED
    sr.line.color.rgb = COLOR_BORDER_RED; sr.line.width = Pt(2)

    tb_sr = s7.shapes.add_textbox(Inches(10.7), Inches(2.6), Inches(7.9), Inches(1.1))
    tf = tb_sr.text_frame; tf.word_wrap = True; tf.margin_left = tf.margin_top = 0
    p = tf.paragraphs[0]; p.text = "ᲨᲔᲓᲔᲒᲘ AI-ᲡᲔᲛᲐᲜᲢᲘᲙᲣᲠᲘ ᲫᲔᲑᲜᲘᲗ"
    p.font.name = FONT_ARIAL; p.font.size = Pt(13); p.font.bold = True; p.font.color.rgb = COLOR_RED
    p = tf.add_paragraph(); p.text = "შინაარსობრივი დამთხვევა (Semantic Search)"
    p.font.name = FONT_ARIAL; p.font.size = Pt(21); p.font.bold = True; p.font.color.rgb = COLOR_DARK; p.space_before = Pt(4)

    fb_sr = s7.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(10.7), Inches(3.8), Inches(7.9), Inches(1.4))
    fb_sr.fill.solid(); fb_sr.fill.fore_color.rgb = COLOR_WHITE; fb_sr.line.color.rgb = COLOR_BORDER_RED; fb_sr.line.width = Pt(1.5)
    tf = fb_sr.text_frame; tf.vertical_anchor = MSO_ANCHOR.MIDDLE
    p = tf.paragraphs[0]; p.text = "თქვენ ეძებთ: „ავტომობილის ყიდვა“\nდოკუმენტში წერია: „მანქანის შეძენა“"
    p.alignment = PP_ALIGN.CENTER; p.font.name = FONT_ARIAL; p.font.size = Pt(14); p.font.bold = True; p.font.color.rgb = COLOR_DARK

    p = tf.add_paragraph(); p.text = "შედეგი: ✅ მყისიერად ნაპოვნია!"
    p.alignment = PP_ALIGN.CENTER; p.font.name = FONT_ARIAL; p.font.size = Pt(16); p.font.bold = True; p.font.color.rgb = COLOR_RED; p.space_before = Pt(4)

    desc_sr = s7.shapes.add_textbox(Inches(10.7), Inches(5.5), Inches(7.9), Inches(4.3))
    tf = desc_sr.text_frame; tf.word_wrap = True; tf.margin_left = tf.margin_top = 0
    p = tf.paragraphs[0]
    p.text = (
        "• აზრობრივი გაგება: ეძებს არა ასოებს, არამედ აზრს (რუკაზე ორივე ფრაზა ერთსა და იმავე წერტილშია).\n\n"
        "• სინონიმების ამოცნობა: ხვდება, რომ „ყიდვა“ და „შეძენა“ ერთი და იგივე ქმედებაა.\n\n"
        "• უპირატესობა: პოულობს პასუხს მაშინაც კი, როდესაც კითხვა და ტექსტი სრულიად განსხვავებული სიტყვებითაა დაწერილი!"
    )
    p.font.name = FONT_ARIAL; p.font.size = Pt(15); p.font.color.rgb = COLOR_TEXT_MUTED

    # -------------------------------------------------------------------------
    # SLIDE 8: RAG: How AI Reads Your Documents
    # -------------------------------------------------------------------------
    s8 = created_slides[6]
    add_header(s8, "RAG: როგორ კითხულობს AI თქვენს დოკუმენტებს?", "არქიტექტურა • პრაქტიკული გამოყენება")

    rag_top = s8.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(1.0), Inches(2.35), Inches(18.0), Inches(1.8))
    rag_top.fill.solid(); rag_top.fill.fore_color.rgb = COLOR_DARK_BOX
    rag_top.line.color.rgb = COLOR_RED; rag_top.line.width = Pt(2.5)

    tb = s8.shapes.add_textbox(Inches(1.4), Inches(2.55), Inches(17.2), Inches(1.4))
    tf = tb.text_frame; tf.word_wrap = True; tf.margin_left = tf.margin_top = 0
    p = tf.paragraphs[0]; p.text = "❓ ᲞᲠᲝᲑᲚᲔᲛᲐ & ᲒᲐᲓᲐᲬᲧᲕᲔᲢᲐ (LEVEL 0):"
    p.font.name = FONT_ARIAL; p.font.size = Pt(13); p.font.bold = True; p.font.color.rgb = COLOR_RED

    p = tf.add_paragraph()
    p.text = "ChatGPT-მ არ იცის თქვენი კომპანიის შიდა ფაილები და ხელშეკრულებები. RAG (Retrieval-Augmented Generation) არის მეთოდი, რომლითაც AI-ს ვაწვდით ზუსტად იმ მონაკვეთს, რაც კითხვაზე პასუხისთვის სჭირდება."
    p.font.name = FONT_ARIAL; p.font.size = Pt(17); p.font.color.rgb = COLOR_WHITE; p.space_before = Pt(4)

    steps = [
        {
            "num": "ნაბიჯი 01",
            "title": "1. დაჭრა და რუკაზე დატანა",
            "action": "დოკუმენტის დამუშავება",
            "desc": "კომპანიის დიდი PDF დოკუმენტები იჭრება პატარა აბზაცებად. ყოველ აბზაცს ენიჭება თავისი ადგილი „აზრების რუკაზე“ (Embedding)."
        },
        {
            "num": "ნაბიჯი 02",
            "title": "2. შესაბამისი ნაწილის პოვნა",
            "action": "ჭკვიანი ძებნა (Search)",
            "desc": "როდესაც მომხმარებელი სვამს კითხვას, სისტემა რუკაზე წამიერად პოულობს იმ კონკრეტულ აბზაცს, რომელიც კითხვის ყველაზე ახლო მეზობელია."
        },
        {
            "num": "ნაბიჯი 03",
            "title": "3. პასუხის გენერაცია",
            "action": "AI-ს მიერ შეჯამება",
            "desc": "ნაპოვნი აბზაცი გადაეცემა AI-ს მკაცრი ინსტრუქციით: „უპასუხე მომხმარებელს მხოლოდ ამ ოფიციალურ დოკუმენტზე დაყრდნობით“."
        }
    ]

    st_w = Inches(5.8); st_gap = Inches(0.3); st_start_x = Inches(1.0)
    for i, st in enumerate(steps):
        cx = st_start_x + i * (st_w + st_gap)
        sb = s8.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, cx, Inches(4.45), st_w, Inches(5.6))
        sb.fill.solid(); sb.fill.fore_color.rgb = COLOR_WHITE
        sb.line.color.rgb = COLOR_BORDER_RED if i == 2 else COLOR_BORDER
        sb.line.width = Pt(1.5)

        tb = s8.shapes.add_textbox(cx + Inches(0.4), Inches(4.75), st_w - Inches(0.8), Inches(5.1))
        tf = tb.text_frame; tf.word_wrap = True; tf.margin_left = tf.margin_top = 0

        p = tf.paragraphs[0]; p.text = st["num"]
        p.font.name = FONT_ARIAL; p.font.size = Pt(12); p.font.bold = True; p.font.color.rgb = COLOR_RED

        p = tf.add_paragraph(); p.text = st["title"]
        p.font.name = FONT_ARIAL; p.font.size = Pt(19); p.font.bold = True; p.font.color.rgb = COLOR_DARK; p.space_before = Pt(3)

        p = tf.add_paragraph(); p.text = f"\nფუნქცია: {st['action']}"
        p.font.name = FONT_ARIAL; p.font.size = Pt(14); p.font.bold = True; p.font.color.rgb = COLOR_RED; p.space_before = Pt(6)

        p = tf.add_paragraph(); p.text = f"\n{st['desc']}"
        p.font.name = FONT_ARIAL; p.font.size = Pt(14); p.font.color.rgb = COLOR_TEXT_MUTED; p.space_before = Pt(6)

    # -------------------------------------------------------------------------
    # SLIDE 9: Ecosystem Tools (Clean 512x512 Square Logos!)
    # -------------------------------------------------------------------------
    s9 = created_slides[7]
    add_header(s9, "ტექსტის გაციფრულების 3 მთავარი ინსტრუმენტი", "პროდუქტები • ვის რა ფუნქცია აქვს")

    eco_tools = [
        {
            "cat": "01 • დამჭრელები",
            "title": "Tokenizers",
            "sub": "ტექსტის დანაწევრება ლეგოს კუბიკებად",
            "tools": [
                ("openai.png", "OpenAI Tiktoken", "ChatGPT-ის ოფიციალური ტოკენიზატორი"),
                ("meta_llama.png", "Meta Llama Tokenizer", "ღია მოდელების ტოკენიზატორი")
            ]
        },
        {
            "cat": "02 • რუკის შემქმნელები",
            "title": "Embedding Models",
            "sub": "ტოკენების გადაქცევა GPS კოორდინატებად",
            "tools": [
                ("openai.png", "Text-Embedding-3", "OpenAI-ს უახლესი ემბედინგ-მოდელი"),
                ("google_cloud.png", "Vertex Embeddings", "Google Cloud-ის საწარმოო მოდელი")
            ]
        },
        {
            "cat": "03 • რუკის შემნახველები",
            "title": "Vector Databases",
            "sub": "კოორდინატების შენახვა და წამიერი ძებნა",
            "tools": [
                ("microsoft.png", "Azure AI Search", "Microsoft-ის საწარმოო ვექტორული ძებნა"),
                ("aws.png", "AWS Bedrock Search", "Amazon-ის ცოდნის ბაზების ინფრასტრუქტურა")
            ]
        }
    ]

    col_w = Inches(5.8); col_gap = Inches(0.3); start_x = Inches(1.0)
    for i, col in enumerate(eco_tools):
        cx = start_x + i * (col_w + col_gap)

        cb = s9.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, cx, Inches(2.35), col_w, Inches(7.7))
        cb.fill.solid()
        cb.fill.fore_color.rgb = COLOR_CARD_RED if i == 0 else COLOR_BG_LIGHT
        cb.line.color.rgb = COLOR_BORDER_RED if i == 0 else COLOR_BORDER
        cb.line.width = Pt(1.5)

        tb_head = s9.shapes.add_textbox(cx + Inches(0.4), Inches(2.6), col_w - Inches(0.8), Inches(1.8))
        tf = tb_head.text_frame; tf.word_wrap = True; tf.margin_left = tf.margin_top = 0

        p = tf.paragraphs[0]; p.text = col["cat"].upper()
        p.font.name = FONT_ARIAL; p.font.size = Pt(12); p.font.bold = True
        p.font.color.rgb = COLOR_RED if i == 0 else COLOR_MUTED

        p = tf.add_paragraph(); p.text = col["title"]
        p.font.name = FONT_ARIAL; p.font.size = Pt(22); p.font.bold = True; p.font.color.rgb = COLOR_DARK
        p.space_before = Pt(3)

        p = tf.add_paragraph(); p.text = col["sub"]
        p.font.name = FONT_ARIAL; p.font.size = Pt(13); p.font.color.rgb = COLOR_TEXT_MUTED
        p.space_before = Pt(4)

        tb_lbl = s9.shapes.add_textbox(cx + Inches(0.4), Inches(4.7), col_w - Inches(0.8), Inches(0.4))
        tf = tb_lbl.text_frame; tf.word_wrap = True; tf.margin_left = tf.margin_top = 0
        p = tf.paragraphs[0]; p.text = "ძირითადი ხელსაწყოები:"
        p.font.name = FONT_ARIAL; p.font.size = Pt(12); p.font.bold = True; p.font.color.rgb = COLOR_DARK

        for j, (l_file, t_name, t_desc) in enumerate(col["tools"]):
            item_y = Inches(5.25) + j * Inches(2.2)

            t_card = s9.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, cx + Inches(0.4), item_y, col_w - Inches(0.8), Inches(1.9))
            t_card.fill.solid(); t_card.fill.fore_color.rgb = COLOR_WHITE
            t_card.line.color.rgb = COLOR_BORDER_RED if i == 0 else COLOR_BORDER
            t_card.line.width = Pt(1.2)

            # 512x512 Square Logo 1.1" x 1.1"
            lp = os.path.join(logo_dir, l_file)
            if os.path.exists(lp):
                s9.shapes.add_picture(lp, cx + Inches(0.65), item_y + Inches(0.4), width=Inches(1.1), height=Inches(1.1))

            tb_item = s9.shapes.add_textbox(cx + Inches(2.0), item_y + Inches(0.3), col_w - Inches(2.6), Inches(1.3))
            tf = tb_item.text_frame; tf.word_wrap = True; tf.margin_left = tf.margin_top = 0

            p = tf.paragraphs[0]; p.text = t_name
            p.font.name = FONT_ARIAL; p.font.size = Pt(15); p.font.bold = True; p.font.color.rgb = COLOR_DARK

            p = tf.add_paragraph(); p.text = t_desc
            p.font.name = FONT_ARIAL; p.font.size = Pt(12); p.font.color.rgb = COLOR_TEXT_MUTED
            p.space_before = Pt(4)

    # -------------------------------------------------------------------------
    # SAVE MASTER PPTX AND COPY TO DRIVE
    # -------------------------------------------------------------------------
    filename = "როგორ კითხულობს AI - ტოკენები, Embeddings და ჭკვიანი ძებნა.pptx"
    out_master = os.path.join(r'd:\ხელოვნური ინტელექტის ინსტრუმენტები', filename)
    prs.save(out_master)
    print("Master PPTX for Lecture 2.1 generated successfully!")

    target_dirs = [
        r'G:\My Drive\შესავალი ხელოვნურ ინტელექტში',
        r'G:\My Drive\Skillwill\შესავალი ხელოვნურ ინტელექტში',
        r'D:\შესავალი ხელოვნურ ინტელექტში',
        r'D:\ხელოვნური ინტელექტის ინსტრუმენტები\შესავალი ხელოვნურ ინტელექტში'
    ]

    for d in target_dirs:
        try:
            os.makedirs(d, exist_ok=True)
            target_path = os.path.join(d, filename)
            shutil.copy2(out_master, target_path)
            print(f"Copied to Drive target: {target_path}")
        except Exception as e:
            print(f"Error copying to {d}: {e}")

build_lecture_2_1()
