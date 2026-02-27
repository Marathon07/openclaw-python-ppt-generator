import json, sys
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.enum.text import PP_ALIGN
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_SHAPE

if len(sys.argv) < 3: sys.exit(1)
with open(sys.argv[1], 'r', encoding='utf-8') as f: slides_data = json.load(f)

prs = Presentation()
prs.slide_width, prs.slide_height = Inches(13.33), Inches(7.5)
BG, TITLE_COLOR, ACCENT, TEXT, BOX_BG = RGBColor(255,255,255), RGBColor(0,51,102), RGBColor(0,102,204), RGBColor(60,60,60), RGBColor(242,245,248)

def add_bullet(tf, text):
    p = tf.paragraphs[0] if not tf.text else tf.add_paragraph()
    p.level, p.line_spacing, p.space_after = 0, 1.3, Pt(12)
    colon = text.find("：") if "：" in text else text.find(": ")
    if colon != -1:
        r1, r2 = p.add_run(), p.add_run()
        r1.text, r1.font.name, r1.font.size, r1.font.bold, r1.font.color.rgb = text[:colon+1], 'Microsoft YaHei', Pt(15), True, TEXT
        r2.text, r2.font.name, r2.font.size, r2.font.color.rgb = text[colon+1:], 'Microsoft YaHei', Pt(15), TEXT
    else:
        r = p.add_run()
        r.text, r.font.name, r.font.size, r.font.color.rgb = text, 'Microsoft YaHei', Pt(15), TEXT

for data in slides_data:
    layout = data.get("layout", "two-column")
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    slide.background.fill.solid(); slide.background.fill.fore_color.rgb = BG
    
    if layout == "cover":
        txBox = slide.shapes.add_textbox(Inches(1), Inches(2.5), Inches(11.33), Inches(2))
        tf = txBox.text_frame; tf.word_wrap = True; tf.auto_size = 2
        p = tf.paragraphs[0]
        p.text, p.font.name, p.font.size, p.font.bold, p.font.color.rgb, p.alignment = data.get("title", ""), 'Microsoft YaHei', Pt(40), True, TITLE_COLOR, PP_ALIGN.CENTER
        if "subtitle" in data:
            p2 = tf.add_paragraph()
            p2.text, p2.font.name, p2.font.size, p2.font.color.rgb, p2.alignment, p2.space_before = data["subtitle"], 'Microsoft YaHei', Pt(24), ACCENT, PP_ALIGN.CENTER, Pt(20)
        continue

    title_box = slide.shapes.add_textbox(Inches(0.5), Inches(0.4), Inches(12.33), Inches(0.8))
    tf_title = title_box.text_frame; tf_title.word_wrap = True; tf_title.auto_size = 2
    p = tf_title.paragraphs[0]
    p.text, p.font.name, p.font.size, p.font.bold, p.font.color.rgb = data.get("action_title", ""), 'Microsoft YaHei', Pt(20), True, TITLE_COLOR
    slide.shapes.add_connector(1, Inches(0.5), Inches(1.3), Inches(12.83), Inches(1.3)).line.color.rgb = ACCENT
    
    content_bottom = 6.8
    if "takeaway" in data:
        content_bottom = 5.9
        shape = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(0.5), Inches(6.1), Inches(12.33), Inches(0.9))
        shape.fill.solid(); shape.fill.fore_color.rgb = BOX_BG
        shape.line.color.rgb, shape.line.width = ACCENT, Pt(1.5)
        tf_ta = shape.text_frame; tf_ta.word_wrap = True; tf_ta.auto_size = 2
        p_ta = tf_ta.paragraphs[0]
        p_ta.text, p_ta.font.name, p_ta.font.size, p_ta.font.bold, p_ta.font.color.rgb = data["takeaway"], 'Microsoft YaHei', Pt(16), True, TITLE_COLOR

    if layout in ["two-column", "three-column"]:
        cols = data.get("columns", [])
        n_cols = len(cols)
        if n_cols > 0:
            spacing, total_width = 0.4, 12.33
            col_width = (total_width - (n_cols - 1) * spacing) / n_cols
            for i, col in enumerate(cols):
                x, y, h = 0.5 + i * (col_width + spacing), 1.5, content_bottom - 1.5
                if "title" in col:
                    ch_box = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(x), Inches(y), Inches(col_width), Inches(0.45))
                    ch_box.fill.solid(); ch_box.fill.fore_color.rgb = TITLE_COLOR
                    tf_ch = ch_box.text_frame; tf_ch.vertical_anchor = 3
                    p_ch = tf_ch.paragraphs[0]
                    p_ch.text, p_ch.font.name, p_ch.font.size, p_ch.font.bold, p_ch.font.color.rgb, p_ch.alignment = col["title"], 'Microsoft YaHei', Pt(16), True, BG, PP_ALIGN.CENTER
                    y += 0.55; h -= 0.55
                body_box = slide.shapes.add_textbox(Inches(x), Inches(y), Inches(col_width), Inches(h))
                tf_body = body_box.text_frame; tf_body.word_wrap = True; tf_body.auto_size = 2
                for bullet in col.get("bullets", []): add_bullet(tf_body, "• " + bullet)
    elif layout == "timeline":
        steps = data.get("steps", [])
        width, spacing, start_x = 2.2, 0.2, 0.6
        for i, step in enumerate(steps):
            x = start_x + i * (width + spacing)
            shape = slide.shapes.add_shape(MSO_SHAPE.CHEVRON, Inches(x), Inches(2.2), Inches(width), Inches(0.8))
            shape.fill.solid(); shape.fill.fore_color.rgb = ACCENT if i == 2 else RGBColor(230, 230, 230)
            shape.line.fill.background()
            tf = shape.text_frame; tf.word_wrap = True
            p = tf.paragraphs[0]
            p.text, p.font.size, p.font.bold, p.font.color.rgb, p.alignment = step.get("title", ""), Pt(14), True, BG if i == 2 else TEXT, PP_ALIGN.CENTER
            tb = slide.shapes.add_textbox(Inches(x), Inches(3.2), Inches(width), Inches(3))
            tb.text_frame.word_wrap = True; tb.text_frame.auto_size = 2
            p2 = tb.text_frame.paragraphs[0]
            p2.text, p2.font.size, p2.font.color.rgb = step.get("desc", ""), Pt(13), RGBColor(80,80,80)
    elif layout == "matrix":
        quads = data.get("quadrants", [])
        coords = [(1.5, 1.8), (7, 1.8), (1.5, 4.0), (7, 4.0)]
        for i, quad in enumerate(quads):
            if i >= len(coords): break
            x, y = coords[i]
            shape = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(x), Inches(y), Inches(5.0), Inches(2.0))
            shape.fill.solid(); shape.fill.fore_color.rgb = BOX_BG; shape.line.color.rgb = ACCENT
            tf = shape.text_frame; tf.word_wrap = True; tf.auto_size = 2
            p = tf.paragraphs[0]
            p.text, p.font.bold, p.font.size, p.font.color.rgb = quad.get("title", ""), True, Pt(16), TITLE_COLOR
            p2 = tf.add_paragraph()
            p2.text, p2.font.size, p2.font.color.rgb = "\n" + quad.get("desc", ""), Pt(14), TEXT

prs.save(sys.argv[2])
