import json
import sys
import os
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.enum.text import PP_ALIGN
from pptx.dml.color import RGBColor

if len(sys.argv) < 3:
    print("Usage: python generate_ppt.py <input.json> <output.pptx>")
    sys.exit(1)

input_file = sys.argv[1]
output_file = sys.argv[2]

with open(input_file, 'r', encoding='utf-8') as f:
    slides_data = json.load(f)

prs = Presentation()
prs.slide_width = Inches(13.33)
prs.slide_height = Inches(7.5) # 16:9

BG_COLOR = RGBColor(248, 249, 250)
TITLE_COLOR = RGBColor(0, 51, 102)
ACCENT_COLOR = RGBColor(0, 102, 204)
TEXT_COLOR = RGBColor(51, 51, 51)
CITATION_COLOR = RGBColor(128, 128, 128)

def set_slide_bg(slide):
    background = slide.background
    fill = background.fill
    fill.solid()
    fill.fore_color.rgb = BG_COLOR

for data in slides_data:
    if data.get("is_cover"):
        slide = prs.slides.add_slide(prs.slide_layouts[6]) # Blank
        set_slide_bg(slide)
        
        txBox = slide.shapes.add_textbox(Inches(1), Inches(2.5), Inches(11.33), Inches(2))
        tf = txBox.text_frame
        tf.word_wrap = True
        
        p = tf.paragraphs[0]
        p.text = data.get("title", "")
        p.font.name = 'Microsoft YaHei'
        p.font.size = Pt(40)
        p.font.bold = True
        p.font.color.rgb = TITLE_COLOR
        p.alignment = PP_ALIGN.CENTER
        
        if "subtitle" in data:
            p2 = tf.add_paragraph()
            p2.text = data["subtitle"]
            p2.font.name = 'Microsoft YaHei'
            p2.font.size = Pt(24)
            p2.font.color.rgb = ACCENT_COLOR
            p2.alignment = PP_ALIGN.CENTER
            p2.space_before = Pt(20)
    else:
        slide = prs.slides.add_slide(prs.slide_layouts[6])
        set_slide_bg(slide)
        
        # Title
        title_box = slide.shapes.add_textbox(Inches(0.8), Inches(0.5), Inches(11.73), Inches(1))
        tf_title = title_box.text_frame
        p = tf_title.paragraphs[0]
        p.text = data.get("title", "")
        p.font.name = 'Microsoft YaHei'
        p.font.size = Pt(32)
        p.font.bold = True
        p.font.color.rgb = TITLE_COLOR
        
        slide.shapes.add_connector(1, Inches(0.8), Inches(1.5), Inches(12.53), Inches(1.5)).line.color.rgb = ACCENT_COLOR
        
        # Body
        body_box = slide.shapes.add_textbox(Inches(0.8), Inches(1.8), Inches(11.73), Inches(5.2))
        tf_body = body_box.text_frame
        tf_body.word_wrap = True
        
        first = True
        for bullet in data.get("bullets", []):
            if first:
                p = tf_body.paragraphs[0]
                first = False
            else:
                p = tf_body.add_paragraph()
                
            p.level = 0
            p.line_spacing = 1.3
            p.space_after = Pt(16)
            
            # Extract citation if exists
            citation_text = ""
            main_text = bullet
            for tag in ["[出处：", "[Citation:"]:
                if tag in bullet:
                    parts = bullet.split(tag)
                    main_text = parts[0]
                    citation_text = tag + parts[1]
                    break
            
            # Bold parsing (colon split)
            colon_found = False
            for colon in ["：", ": "]:
                if colon in main_text:
                    sub_parts = main_text.split(colon, 1)
                    run_bold = p.add_run()
                    run_bold.text = sub_parts[0] + colon
                    run_bold.font.name = 'Microsoft YaHei'
                    run_bold.font.size = Pt(22)
                    run_bold.font.bold = True
                    run_bold.font.color.rgb = TEXT_COLOR
                    
                    run_norm = p.add_run()
                    run_norm.text = sub_parts[1]
                    run_norm.font.name = 'Microsoft YaHei'
                    run_norm.font.size = Pt(22)
                    run_norm.font.color.rgb = TEXT_COLOR
                    colon_found = True
                    break
            
            if not colon_found:
                run_norm = p.add_run()
                run_norm.text = main_text
                run_norm.font.name = 'Microsoft YaHei'
                run_norm.font.size = Pt(22)
                run_norm.font.color.rgb = TEXT_COLOR
                
            if citation_text:
                run_cit = p.add_run()
                run_cit.text = " " + citation_text
                run_cit.font.name = 'Microsoft YaHei'
                run_cit.font.size = Pt(16)
                run_cit.font.italic = True
                run_cit.font.color.rgb = CITATION_COLOR

prs.save(output_file)
print(f"Presentation saved to {output_file}")
