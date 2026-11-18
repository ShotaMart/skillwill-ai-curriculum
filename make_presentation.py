"""
Skillwill Presentation Generator (ბრენდირებული პრეზენტაციების გენერატორი)
ავტორი: შოთა მარტყოფლიშვილი
პროგრამა: ხელოვნური ინტელექტის ინსტრუმენტები (06.7.1)
სტილი: მაღალ-ვიზუალური (ლოგოები, ბარათები, ნაკლები ტექსტი, მხოლოდ თეორია)
"""

import os
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN
from pptx.enum.shapes import MSO_SHAPE

# Skillwill Brand Palette
COLOR_RED = RGBColor(213, 43, 30)       # #D52B1E
COLOR_DARK = RGBColor(17, 24, 39)        # #111827
COLOR_MUTED = RGBColor(107, 114, 128)    # #6B7280
COLOR_LIGHT = RGBColor(248, 250, 252)    # #F8FAFC
COLOR_BORDER = RGBColor(226, 232, 240)   # #E2E8F0
COLOR_WHITE = RGBColor(255, 255, 255)
COLOR_CARD_RED = RGBColor(254, 242, 242)
COLOR_BORDER_RED = RGBColor(254, 202, 202)

FONT_TITLE = 'TBC Contractica CAPS Bold'
FONT_BODY = 'Segoe UI'

def create_visual_deck(output_path, title, subtitle, slides, template_path=None):
    """
    ქმნის სუფთა, მაღალ-ვიზუალურ პრეზენტაციას Skillwill-ის ბრენდინგით:
    - სატიტულო გვერდი: Skillwill ლოგო + კუთხის დეკორაციები + წითელი სათაური
    - შიდა გვერდები: მინიმალისტური ჰედერი/ფუტერი, ლოგოები, ბარათები, ნაკლები ტექსტი
    - ბოლო გვერდი: Skillwill მადლობის შაბლონი
    """
    if template_path is None:
        template_path = os.path.join(os.path.dirname(__file__), 'Skillwill Presentation Example .pptx')
    
    prs = Presentation(template_path)
    blank_layout = prs.slide_layouts[6]

    asset_dir = os.path.join(os.path.dirname(__file__), '.template_assets')
    img_corner_bl = os.path.join(asset_dir, 'image2.png')
    img_corner_tr = os.path.join(asset_dir, 'image3.png')
    img_icon_title = os.path.join(asset_dir, 'image5.png')

    # 1. Title Slide
    s1 = prs.slides[0]
    for shape in s1.shapes:
        if shape.has_text_frame and shape.text_frame.text.strip():
            tf = shape.text_frame
            tf.clear()
            p = tf.paragraphs[0]
            p.text = title
            p.alignment = PP_ALIGN.CENTER
            p.font.name = FONT_TITLE
            p.font.size = Pt(36)
            p.font.bold = True
            p.font.color.rgb = COLOR_RED

            if subtitle:
                p2 = tf.add_paragraph()
                p2.text = subtitle
                p2.alignment = PP_ALIGN.CENTER
                p2.font.name = FONT_BODY
                p2.font.size = Pt(17)
                p2.font.color.rgb = COLOR_MUTED
                p2.space_before = Pt(14)
            break

    # Drop template slide 2 & 3
    rId2 = prs.slides._sldIdLst[1].rId
    prs.part.drop_rel(rId2)
    del prs.slides._sldIdLst[1]

    rId3 = prs.slides._sldIdLst[1].rId
    prs.part.drop_rel(rId3)
    del prs.slides._sldIdLst[1]

    def add_chrome(slide, slide_title, category):
        if os.path.exists(img_corner_tr):
            slide.shapes.add_picture(img_corner_tr, Inches(18.3), Inches(0), width=Inches(1.7))
        if os.path.exists(img_icon_title):
            slide.shapes.add_picture(img_icon_title, Inches(1.2), Inches(0.95), width=Inches(0.42))

        tb = slide.shapes.add_textbox(Inches(1.75), Inches(0.7), Inches(16.0), Inches(1.1))
        tf = tb.text_frame
        tf.word_wrap = True
        tf.margin_left = tf.margin_top = tf.margin_right = tf.margin_bottom = 0

        p_cat = tf.paragraphs[0]
        p_cat.text = category.upper()
        p_cat.font.name = FONT_TITLE
        p_cat.font.size = Pt(11)
        p_cat.font.bold = True
        p_cat.font.color.rgb = COLOR_RED

        p_t = tf.add_paragraph()
        p_t.text = slide_title
        p_t.font.name = FONT_TITLE
        p_t.font.size = Pt(28)
        p_t.font.bold = True
        p_t.font.color.rgb = COLOR_DARK
        p_t.space_before = Pt(4)

        line = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(1.2), Inches(1.95), Inches(17.6), Inches(0.02))
        line.fill.solid()
        line.fill.fore_color.rgb = COLOR_BORDER
        line.line.color.rgb = COLOR_BORDER

        foot = slide.shapes.add_textbox(Inches(1.2), Inches(10.3), Inches(17.6), Inches(0.4))
        ftf = foot.text_frame
        ftf.word_wrap = True
        p_foot = ftf.paragraphs[0]
        p_foot.text = "SKILLWILL • ხელოვნური ინტელექტის ინსტრუმენტები | ლექტორი: შოთა მარტყოფლიშვილი"
        p_foot.font.name = FONT_BODY
        p_foot.font.size = Pt(11)
        p_foot.font.color.rgb = COLOR_MUTED

    # 2. Content Slides
    for s_info in slides:
        slide = prs.slides.add_slide(blank_layout)
        add_chrome(slide, s_info['title'], s_info.get('category', "SKILLWILL AI"))
        # Render handler based on slide type (custom or builder)
        if 'builder' in s_info:
            s_info['builder'](slide, s_info)

    # 3. Final Closing Slide
    s_final = prs.slides.add_slide(blank_layout)
    if os.path.exists(img_corner_bl):
        s_final.shapes.add_picture(img_corner_bl, Inches(0), Inches(8.75), width=Inches(2.52))
    if os.path.exists(img_corner_tr):
        s_final.shapes.add_picture(img_corner_tr, Inches(17.53), Inches(0), width=Inches(2.47))

    close_box = s_final.shapes.add_textbox(Inches(4.0), Inches(4.6), Inches(12.0), Inches(2.5))
    ctf = close_box.text_frame
    ctf.word_wrap = True

    p_close = ctf.paragraphs[0]
    p_close.text = "მადლობა ყურადღებისთვის!"
    p_close.alignment = PP_ALIGN.CENTER
    p_close.font.name = FONT_TITLE
    p_close.font.size = Pt(42)
    p_close.font.bold = True
    p_close.font.color.rgb = COLOR_DARK

    p_subclose = ctf.add_paragraph()
    p_subclose.text = "კითხვა-პასუხი & დისკუსია\nSkillwill • 2026"
    p_subclose.alignment = PP_ALIGN.CENTER
    p_subclose.font.name = FONT_BODY
    p_subclose.font.size = Pt(18)
    p_subclose.font.color.rgb = COLOR_RED
    p_subclose.space_before = Pt(16)

    prs.save(output_path)
    return output_path
