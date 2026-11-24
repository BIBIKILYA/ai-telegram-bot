import sys
import json
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.enum.text import PP_ALIGN
from pptx.dml.color import RGBColor

# Получаем данные из аргументов командной строки
data = json.loads(sys.argv[1])
output_path = sys.argv[2]

# Создаем презентацию (широкоформатная 16:9)
prs = Presentation()
prs.slide_width = Inches(13.333)
prs.slide_height = Inches(7.5)

# === ТИТУЛЬНЫЙ СЛАЙД ===
title_slide_layout = prs.slide_layouts[0]
slide = prs.slides.add_slide(title_slide_layout)
title = slide.shapes.title
subtitle = slide.placeholders[1]

title.text = data.get("title", "Презентация")
title.text_frame.paragraphs[0].font.size = Pt(44)
title.text_frame.paragraphs[0].font.bold = True
title.text_frame.paragraphs[0].font.color.rgb = RGBColor(0, 51, 102)

subtitle.text = data.get("subtitle", "")
subtitle.text_frame.paragraphs[0].font.size = Pt(24)

# === СЛАЙДЫ С КОНТЕНТОМ ===
bullet_slide_layout = prs.slide_layouts[1]

for slide_data in data.get("slides", []):
    slide = prs.slides.add_slide(bullet_slide_layout)
    shapes = slide.shapes
    
    # Заголовок слайда
    title_shape = shapes.title
    title_shape.text = slide_data.get("title", "")
    title_shape.text_frame.paragraphs[0].font.size = Pt(36)
    title_shape.text_frame.paragraphs[0].font.bold = True
    title_shape.text_frame.paragraphs[0].font.color.rgb = RGBColor(0, 51, 102)
    
    # Тело слайда (буллеты)
    body_shape = shapes.placeholders[1]
    tf = body_shape.text_frame
    tf.clear()
    
    bullets = slide_data.get("bullets", [])
    if bullets:
        tf.text = bullets[0]
        tf.paragraphs[0].font.size = Pt(20)
        
        for bullet in bullets[1:]:
            p = tf.add_paragraph()
            p.text = bullet
            p.font.size = Pt(20)
            p.level = 0

# === ФИНАЛЬНЫЙ СЛАЙД ===
final_slide_layout = prs.slide_layouts[5]  # Пустой слайд
slide = prs.slides.add_slide(final_slide_layout)

# Добавляем текст "Спасибо за внимание!"
left = Inches(2)
top = Inches(3)
width = Inches(9)
height = Inches(2)
txBox = slide.shapes.add_textbox(left, top, width, height)
tf = txBox.text_frame
p = tf.paragraphs[0]
p.text = "Спасибо за внимание!"
p.font.size = Pt(48)
p.font.bold = True
p.font.color.rgb = RGBColor(0, 51, 102)
p.alignment = PP_ALIGN.CENTER

# Сохраняем
prs.save(output_path)
print(f"SUCCESS:{output_path}")