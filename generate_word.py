import sys
import json
from docx import Document
from docx.shared import Pt, Inches, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH

# Получаем данные из аргументов командной строки
data = json.loads(sys.argv[1])
output_path = sys.argv[2]

# Создаем документ
doc = Document()

# Настройка стилей
style = doc.styles['Normal']
font = style.font
font.name = 'Times New Roman'
font.size = Pt(12)

# === ТИТУЛЬНАЯ СТРАНИЦА ===
doc.add_paragraph()  # Пустая строка
doc.add_paragraph()

# Название доклада
title = doc.add_paragraph()
title.alignment = WD_ALIGN_PARAGRAPH.CENTER
run = title.add_run(data.get("title", "Доклад"))
run.font.size = Pt(24)
run.font.bold = True
run.font.color.rgb = RGBColor(0, 51, 102)

# Подзаголовок
if data.get("subtitle"):
    subtitle = doc.add_paragraph()
    subtitle.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run = subtitle.add_run(data.get("subtitle"))
    run.font.size = Pt(16)
    run.font.color.rgb = RGBColor(100, 100, 100)

doc.add_paragraph()
doc.add_paragraph()

# === ОГЛАВЛЕНИЕ ===
doc.add_heading('Оглавление', level=1)
for i, section in enumerate(data.get("sections", []), 1):
    doc.add_paragraph(f"{i}. {section.get('title', '')}", style='List Number')

doc.add_page_break()

# === ОСНОВНОЙ КОНТЕНТ ===
for section in data.get("sections", []):
    # Заголовок раздела
    doc.add_heading(section.get("title", ""), level=1)
    
    # Введение к разделу
    if section.get("intro"):
        doc.add_paragraph(section.get("intro"))
    
    # Подразделы
    for subsection in section.get("subsections", []):
        doc.add_heading(subsection.get("title", ""), level=2)
        
        # Текст подраздела
        if subsection.get("content"):
            doc.add_paragraph(subsection.get("content"))
        
        # Ключевые пункты
        if subsection.get("bullets"):
            for bullet in subsection.get("bullets", []):
                doc.add_paragraph(bullet, style='List Bullet')
    
    # Выводы раздела
    if section.get("conclusion"):
        doc.add_paragraph()
        conclusion = doc.add_paragraph()
        run = conclusion.add_run("Вывод: ")
        run.bold = True
        conclusion.add_run(section.get("conclusion"))
    
    doc.add_paragraph()  # Отступ между разделами

# === ЗАКЛЮЧЕНИЕ ===
if data.get("conclusion"):
    doc.add_page_break()
    doc.add_heading('Заключение', level=1)
    doc.add_paragraph(data.get("conclusion"))

# === ИСТОЧНИКИ ===
if data.get("sources"):
    doc.add_paragraph()
    doc.add_heading('Источники информации', level=2)
    for source in data.get("sources", []):
        doc.add_paragraph(source, style='List Number')

# Сохраняем документ
doc.save(output_path)
print(f"SUCCESS:{output_path}")