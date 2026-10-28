import os
import re
import io
import base64
import glob
import requests
import docx
from docx.shared import Inches, Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.oxml import OxmlElement
from docx.oxml.ns import qn

def set_cell_background(cell, fill_hex):
    tcPr = cell._tc.get_or_add_tcPr()
    shd = OxmlElement('w:shd')
    shd.set(qn('w:val'), 'clear')
    shd.set(qn('w:color'), 'auto')
    shd.set(qn('w:fill'), fill_hex)
    tcPr.append(shd)

def set_cell_margins(cell, top=100, bottom=100, left=150, right=150):
    tcPr = cell._tc.get_or_add_tcPr()
    tcMar = OxmlElement('w:tcMar')
    for m, val in [('top', top), ('bottom', bottom), ('left', left), ('right', right)]:
        node = OxmlElement(f'w:{m}')
        node.set(qn('w:w'), str(val))
        node.set(qn('w:type'), 'dxa')
        tcMar.append(node)
    tcPr.append(tcMar)

def set_table_borders(table, color="000000", sz="4", val="single"):
    tblPr = table._tbl.tblPr
    tblBorders = OxmlElement('w:tblBorders')
    for border_name in ['top', 'left', 'bottom', 'right', 'insideH', 'insideV']:
        border = OxmlElement(f'w:{border_name}')
        border.set(qn('w:val'), val)
        border.set(qn('w:sz'), sz)
        border.set(qn('w:space'), '0')
        border.set(qn('w:color'), color)
        tblBorders.append(border)
    tblPr.append(tblBorders)

def add_page_number_to_run(run):
    fldChar1 = OxmlElement('w:fldChar')
    fldChar1.set(qn('w:fldCharType'), 'begin')
    instrText = OxmlElement('w:instrText')
    instrText.set(qn('xml:space'), 'preserve')
    instrText.text = "PAGE"
    fldChar2 = OxmlElement('w:fldChar')
    fldChar2.set(qn('w:fldCharType'), 'separate')
    fldChar3 = OxmlElement('w:fldChar')
    fldChar3.set(qn('w:fldCharType'), 'end')
    r = run._r
    r.append(fldChar1)
    r.append(instrText)
    r.append(fldChar2)
    r.append(fldChar3)

def create_gost_document():
    doc = docx.Document()
    section = doc.sections[0]
    
    # Поля по ГОСТ 7.32-2017: Левое - 30 мм, Правое - 15 мм, Верхнее - 20 мм, Нижнее - 20 мм
    section.page_width = Inches(8.27)
    section.page_height = Inches(11.69)
    section.top_margin = Inches(0.79)     # 20 мм
    section.bottom_margin = Inches(0.79)  # 20 мм
    section.left_margin = Inches(1.18)    # 30 мм
    section.right_margin = Inches(0.59)   # 15 мм
    
    # Номер страницы: на первой странице не выводится
    section.different_first_page_header_footer = True
    footer = section.footer
    footer_p = footer.paragraphs[0]
    footer_p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    footer_run = footer_p.add_run()
    footer_run.font.name = 'Times New Roman'
    footer_run.font.size = Pt(12)
    add_page_number_to_run(footer_run)
    
    # Базовый стиль текста
    style_normal = doc.styles['Normal']
    font = style_normal.font
    font.name = 'Times New Roman'
    font.size = Pt(14)
    font.color.rgb = RGBColor(0, 0, 0)
    
    return doc

def add_title_page(doc, lab_num, lab_title):
    p_header = doc.add_paragraph()
    p_header.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p_header.paragraph_format.line_spacing = 1.0
    p_header.paragraph_format.space_after = Pt(2)
    p_header.paragraph_format.first_line_indent = Inches(0)
    
    r1 = p_header.add_run("МИНИСТЕРСТВО НАУКИ И ВЫСШЕГО ОБРАЗОВАНИЯ РОССИЙСКОЙ ФЕДЕРАЦИИ\n")
    r1.bold = True
    r1.font.size = Pt(11)
    r2 = p_header.add_run("ФЕДЕРАЛЬНОЕ ГОСУДАРСТВЕННОЕ АВТОНОМНОЕ ОБРАЗОВАТЕЛЬНОЕ УЧРЕЖДЕНИЕ\nВЫСШЕГО ОБРАЗОВАНИЯ\n")
    r2.font.size = Pt(10)
    r3 = p_header.add_run("«СЕВЕРО-КАВКАЗСКИЙ ФЕДЕРАЛЬНЫЙ УНИВЕРСИТЕТ»\n\n")
    r3.bold = True
    r3.font.size = Pt(12)
    
    r4 = p_header.add_run("Институт цифрового развития\n")
    r4.font.size = Pt(12)
    r5 = p_header.add_run("Кафедра информационных технологий\n")
    r5.font.size = Pt(12)
    
    p_spacer = doc.add_paragraph()
    p_spacer.paragraph_format.space_before = Pt(40)
    p_spacer.paragraph_format.first_line_indent = Inches(0)
    
    p_title = doc.add_paragraph()
    p_title.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p_title.paragraph_format.line_spacing = 1.5
    p_title.paragraph_format.first_line_indent = Inches(0)
    
    r_rep = p_title.add_run("ОТЧЕТ\n")
    r_rep.bold = True
    r_rep.font.size = Pt(18)
    
    r_prac = p_title.add_run(f"по практической работе № {lab_num}\n")
    r_prac.bold = True
    r_prac.font.size = Pt(14)
    
    r_disc = p_title.add_run("по дисциплине «Искусственный интеллект в профессиональной сфере»\n\n")
    r_disc.font.size = Pt(14)
    
    r_topic_lbl = p_title.add_run("Тема: ")
    r_topic_lbl.font.size = Pt(14)
    r_topic = p_title.add_run(f"«{lab_title}»\n")
    r_topic.bold = True
    r_topic.font.size = Pt(14)
    
    p_sign_spacer = doc.add_paragraph()
    p_sign_spacer.paragraph_format.space_before = Pt(48)
    p_sign_spacer.paragraph_format.first_line_indent = Inches(0)
    
    tbl = doc.add_table(rows=1, cols=2)
    tbl.alignment = WD_TABLE_ALIGNMENT.CENTER
    cell_left = tbl.cell(0, 0)
    cell_right = tbl.cell(0, 1)
    cell_left.width = Inches(2.3)
    cell_right.width = Inches(4.2)
    
    p_sign = cell_right.paragraphs[0]
    p_sign.alignment = WD_ALIGN_PARAGRAPH.LEFT
    p_sign.paragraph_format.line_spacing = 1.15
    p_sign.paragraph_format.first_line_indent = Inches(0)
    
    r_st = p_sign.add_run("Выполнил:\n")
    r_st.bold = True
    r_st.font.size = Pt(12)
    p_sign.add_run("студент 3 курса группы ПИН-б-о-24-1(2)\n"
                   "очной формы обучения\n"
                   "направления подготовки:\n"
                   "09.03.03 «Прикладная информатика»\n"
                   "Мухамеджанов Эльёр Тимурович\n\n")
    
    r_tch = p_sign.add_run("Проверил:\n")
    r_tch.bold = True
    r_tch.font.size = Pt(12)
    p_sign.add_run("доцент кафедры ИТ,\n"
                   "к.т.н., доцент\n"
                   "Березина Виктория Андреевна\n")
    
    p_city_spacer = doc.add_paragraph()
    p_city_spacer.paragraph_format.space_before = Pt(64)
    p_city_spacer.paragraph_format.first_line_indent = Inches(0)
    
    p_city = doc.add_paragraph()
    p_city.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p_city.paragraph_format.first_line_indent = Inches(0)
    r_city = p_city.add_run("Ставрополь, 2026 г.")
    r_city.font.size = Pt(12)
    
    doc.add_page_break()

def add_heading_1(doc, text):
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p.paragraph_format.first_line_indent = Inches(0)
    p.paragraph_format.space_before = Pt(18)
    p.paragraph_format.space_after = Pt(12)
    p.paragraph_format.keep_with_next = True
    run = p.add_run(text.upper())
    run.bold = True
    run.font.name = 'Times New Roman'
    run.font.size = Pt(14)
    run.font.color.rgb = RGBColor(0, 0, 0)
    return p

def add_heading_2(doc, text):
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.LEFT
    p.paragraph_format.first_line_indent = Inches(0.49)
    p.paragraph_format.space_before = Pt(14)
    p.paragraph_format.space_after = Pt(6)
    p.paragraph_format.keep_with_next = True
    run = p.add_run(text)
    run.bold = True
    run.font.name = 'Times New Roman'
    run.font.size = Pt(14)
    run.font.color.rgb = RGBColor(0, 0, 0)
    return p

def add_heading_3(doc, text):
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.LEFT
    p.paragraph_format.first_line_indent = Inches(0.49)
    p.paragraph_format.space_before = Pt(10)
    p.paragraph_format.space_after = Pt(4)
    p.paragraph_format.keep_with_next = True
    run = p.add_run(text)
    run.bold = True
    run.font.name = 'Times New Roman'
    run.font.size = Pt(14)
    run.font.color.rgb = RGBColor(0, 0, 0)
    return p

def clean_inline_math(text):
    return (text.replace('\\text', '')
                .replace('{', '')
                .replace('}', '')
                .replace('\\le', '≤')
                .replace('\\ge', '≥')
                .replace('\\approx', '≈')
                .replace('\\times', '×')
                .replace('\\sum', '∑')
                .replace('\\Delta', 'Δ')
                .replace('\\mu', 'μ')
                .replace('\\eta', 'η')
                .replace('\\alpha', 'α')
                .replace('\\beta', 'β')
                .replace('\\in', '∈')
                .replace('\\to', '→')
                .replace('\\rightarrow', '→')
                .replace('\\neq', '≠')
                .replace('\\cdot', '·')
                .replace('\\land', ' ∧ ')
                .replace('\\lor', ' ∨ ')
                .replace('\\quad', ' ')
                .replace('\\implies', ' ⟹ '))

def format_text_runs(paragraph, text, base_font_size=14):
    tokens = re.split(r'(\*\*.*?\*\*|\*.*?\*|`.*?`|\$\$.*?\$\$|\$.*?\$)', text)
    for token in tokens:
        if not token:
            continue
        if token.startswith('**') and token.endswith('**'):
            r = paragraph.add_run(token[2:-2])
            r.bold = True
            r.font.name = 'Times New Roman'
            r.font.size = Pt(base_font_size)
        elif token.startswith('*') and token.endswith('*'):
            r = paragraph.add_run(token[1:-1])
            r.italic = True
            r.font.name = 'Times New Roman'
            r.font.size = Pt(base_font_size)
        elif token.startswith('`') and token.endswith('`'):
            r = paragraph.add_run(token[1:-1])
            r.font.name = 'Consolas'
            r.font.size = Pt(base_font_size - 1)
        elif (token.startswith('$$') and token.endswith('$$')) or (token.startswith('$') and token.endswith('$')):
            math_content = token.strip('$')
            clean_m = clean_inline_math(math_content)
            r = paragraph.add_run(clean_m)
            r.italic = True
            r.font.name = 'Times New Roman'
            r.font.size = Pt(base_font_size)
        else:
            r = paragraph.add_run(token)
            r.font.name = 'Times New Roman'
            r.font.size = Pt(base_font_size)

def add_paragraph_with_inline(doc, text):
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
    p.paragraph_format.first_line_indent = Inches(0.49)
    p.paragraph_format.line_spacing = 1.5
    p.paragraph_format.space_before = Pt(0)
    p.paragraph_format.space_after = Pt(0)
    format_text_runs(p, text, base_font_size=14)
    return p

def add_list_item(doc, text, is_numbered=False, num_index=1):
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
    p.paragraph_format.first_line_indent = Inches(0.49)
    p.paragraph_format.line_spacing = 1.5
    p.paragraph_format.space_before = Pt(0)
    p.paragraph_format.space_after = Pt(0)
    
    prefix = f"{num_index}. " if is_numbered else "– "
    r_pref = p.add_run(prefix)
    r_pref.bold = is_numbered
    r_pref.font.name = 'Times New Roman'
    r_pref.font.size = Pt(14)
    
    format_text_runs(p, text, base_font_size=14)
    return p

def add_table_to_doc(doc, table_data, table_num, table_caption=""):
    if not table_data:
        return
    
    p_cap = doc.add_paragraph()
    p_cap.alignment = WD_ALIGN_PARAGRAPH.LEFT
    p_cap.paragraph_format.first_line_indent = Inches(0.49)
    p_cap.paragraph_format.space_before = Pt(12)
    p_cap.paragraph_format.space_after = Pt(4)
    p_cap.paragraph_format.keep_with_next = True
    
    r_tab = p_cap.add_run(f"Таблица {table_num}")
    r_tab.bold = True
    r_tab.font.name = 'Times New Roman'
    r_tab.font.size = Pt(14)
    if table_caption:
        r_txt = p_cap.add_run(f" — {table_caption}")
        r_txt.font.name = 'Times New Roman'
        r_txt.font.size = Pt(14)
    
    # Нормализация колонок
    num_rows = len(table_data)
    num_cols = max(len(row) for row in table_data) if num_rows > 0 else 0
    if num_rows == 0 or num_cols == 0:
        return
        
    normalized_data = []
    for r in table_data:
        if len(r) < num_cols:
            r = r + [''] * (num_cols - len(r))
        elif len(r) > num_cols:
            r = r[:num_cols-1] + [' '.join(r[num_cols-1:])]
        normalized_data.append(r)
        
    table = doc.add_table(rows=num_rows, cols=num_cols)
    table.alignment = WD_TABLE_ALIGNMENT.CENTER
    set_table_borders(table, color="000000", sz="4", val="single")
    
    for row_idx, row in enumerate(normalized_data):
        for col_idx, cell_text in enumerate(row):
            cell = table.cell(row_idx, col_idx)
            set_cell_margins(cell, top=80, bottom=80, left=120, right=120)
            p = cell.paragraphs[0]
            p.paragraph_format.line_spacing = 1.0
            p.paragraph_format.space_before = Pt(2)
            p.paragraph_format.space_after = Pt(2)
            p.paragraph_format.first_line_indent = Inches(0)
            
            cleaned_text = cell_text.replace('<br>', '\n').replace('\\n', '\n')
            
            if row_idx == 0:
                p.alignment = WD_ALIGN_PARAGRAPH.CENTER
                set_cell_background(cell, "F2F2F2")
                format_text_runs(p, cleaned_text, base_font_size=12)
                for run in p.runs:
                    run.bold = True
            else:
                p.alignment = WD_ALIGN_PARAGRAPH.LEFT
                format_text_runs(p, cleaned_text, base_font_size=12)

def add_figure_to_doc(doc, content, fig_num, fig_caption=""):
    is_mermaid = 'flowchart' in content or 'graph ' in content or 'xychart' in content
    img_added = False
    
    if is_mermaid:
        try:
            b64 = base64.b64encode(content.encode('utf8')).decode('ascii')
            url = 'https://mermaid.ink/img/' + b64
            resp = requests.get(url, timeout=7)
            if resp.status_code == 200:
                img_stream = io.BytesIO(resp.content)
                p_img = doc.add_paragraph()
                p_img.alignment = WD_ALIGN_PARAGRAPH.CENTER
                p_img.paragraph_format.space_before = Pt(12)
                p_img.paragraph_format.space_after = Pt(4)
                p_img.paragraph_format.first_line_indent = Inches(0)
                run_img = p_img.add_run()
                run_img.add_picture(img_stream, width=Inches(5.5))
                img_added = True
        except Exception:
            img_added = False
            
    if not img_added:
        tbl = doc.add_table(rows=1, cols=1)
        tbl.alignment = WD_TABLE_ALIGNMENT.CENTER
        set_table_borders(tbl, color="CCCCCC", sz="4", val="single")
        cell = tbl.cell(0, 0)
        cell.width = Inches(6.5)
        set_cell_background(cell, "F9FBFD")
        set_cell_margins(cell, top=100, bottom=100, left=150, right=150)
        
        p = cell.paragraphs[0]
        p.alignment = WD_ALIGN_PARAGRAPH.LEFT
        p.paragraph_format.line_spacing = 1.0
        p.paragraph_format.space_before = Pt(2)
        p.paragraph_format.space_after = Pt(2)
        p.paragraph_format.first_line_indent = Inches(0)
        
        run = p.add_run(content.strip())
        run.font.name = 'Consolas'
        run.font.size = Pt(10)
    
    p_cap = doc.add_paragraph()
    p_cap.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p_cap.paragraph_format.first_line_indent = Inches(0)
    p_cap.paragraph_format.space_before = Pt(6)
    p_cap.paragraph_format.space_after = Pt(12)
    
    r_fig = p_cap.add_run(f"Рисунок {fig_num}")
    r_fig.bold = True
    r_fig.font.name = 'Times New Roman'
    r_fig.font.size = Pt(14)
    if fig_caption:
        r_c = p_cap.add_run(f" — {fig_caption}")
        r_c.font.name = 'Times New Roman'
        r_c.font.size = Pt(14)

def add_standard_bibliography(doc):
    add_heading_1(doc, "СПИСОК ИСПОЛЬЗОВАННЫХ ИСТОЧНИКОВ")
    sources = [
        "Станкевич, Л. А. Интеллектуальные системы и технологии : учебник и практикум для бакалавриата и магистратуры / Л. А. Станкевич. — Москва : Издательство Юрайт, 2018. — 397 с.",
        "Кудрявцев, В. Б. Интеллектуальные системы : учебник и практикум для вузов / В. Б. Кудрявцев, Э. Э. Гасанов, А. С. Подколзин. — 2-е изд., испр. и доп. — Москва : Издательство Юрайт, 2020. — 165 с.",
        "Рассел, С. Искусственный интеллект: современный подход / С. Рассел, П. Норвиг ; пер. с англ. — Москва : Издательский дом «Вильямс», 2021. — 704 с.",
        "Калан, Р. Основные концепции нейронных сетей / Р. Калан ; пер. с англ. — Москва : Издательский дом «Вильямс», 2018. — 287 с.",
        "Джексон, П. Введение в экспертные системы : учебное пособие / П. Джексон ; пер. с англ. и ред. В. Т. Тертышного. — Москва : Вильямс, 2021. — 622 с.",
        "Березина, В. А. Искусственный интеллект в профессиональной сфере : методические указания по выполнению практических работ / В. А. Березина. — Ставрополь : Изд-во СКФУ, 2023. — 34 с.",
        "ГОСТ 7.32-2017. Система стандартов по информации, библиотечному и издательскому делу. Отчет о научно-исследовательской работе. Структура и правила оформления. — Москва : Стандартинформ, 2018. — 28 с."
    ]
    for idx, src in enumerate(sources, start=1):
        add_list_item(doc, src, is_numbered=True, num_index=idx)

def process_markdown_file(md_path, docx_path):
    with open(md_path, 'r', encoding='utf-8') as f:
        lines = f.readlines()
        
    doc = create_gost_document()
    
    # Извлечение номера и темы
    lab_num = "1"
    lab_title = "Тема"
    
    for line in lines[:10]:
        m1 = re.search(r'Практическая работа №\s*(\d+)', line)
        if m1:
            lab_num = m1.group(1)
        m2 = re.search(r'##\s*(.+)', line)
        if m2:
            lab_title = m2.group(1).strip()
            
    # Добавление титульного листа
    add_title_page(doc, lab_num, lab_title)
    
    # Парсер содержимого
    in_code_block = False
    code_lang = ""
    code_lines = []
    
    in_table = False
    table_lines = []
    
    fig_counter = 1
    table_counter = 1
    current_task_title = ""
    
    # Пропуск титульных метаданных до разделителя '---'
    body_lines = []
    dash_count = 0
    for line in lines:
        if line.strip() == '---':
            dash_count += 1
            if dash_count == 1:
                continue
        if dash_count >= 1:
            body_lines.append(line)
            
    i = 0
    while i < len(body_lines):
        line = body_lines[i]
        stripped = line.strip()
        
        # Блоки кода / диаграмм
        if stripped.startswith('```'):
            if not in_code_block:
                in_code_block = True
                code_lang = stripped[3:].strip()
                code_lines = []
            else:
                in_code_block = False
                block_content = "\n".join(code_lines)
                cap = current_task_title or f"Схема к практической работе № {lab_num}"
                add_figure_to_doc(doc, block_content, fig_counter, cap)
                fig_counter += 1
                code_lines = []
            i += 1
            continue
            
        if in_code_block:
            code_lines.append(line.rstrip('\r\n'))
            i += 1
            continue
            
        # Таблицы Markdown
        if stripped.startswith('|') and stripped.endswith('|'):
            if not in_table:
                in_table = True
                table_lines = []
            table_lines.append(stripped)
            i += 1
            continue
        else:
            if in_table:
                in_table = False
                # Парсинг таблицы
                parsed_rows = []
                for tline in table_lines:
                    # Пропуск строки-разделителя |---|---|
                    if re.match(r'^\|[\s\-:|]+\|$', tline):
                        continue
                    clean_line = tline.replace(r'\|', '‖')
                    cells = [c.strip().replace('‖', '|') for c in clean_line.strip('|').split('|')]
                    parsed_rows.append(cells)
                if parsed_rows:
                    cap = current_task_title or f"Данные к практической работе № {lab_num}"
                    add_table_to_doc(doc, parsed_rows, table_counter, cap)
                    table_counter += 1
                table_lines = []
                
        if not stripped or stripped == '---':
            i += 1
            continue
            
        # Заголовки
        if stripped.startswith('### '):
            h_text = stripped[4:].strip()
            if 'цель' in h_text.lower():
                add_heading_1(doc, "ЦЕЛЬ ЗАНЯТИЯ")
            elif 'вопросы' in h_text.lower():
                add_heading_1(doc, "ВОПРОСЫ, ВЫНОСИМЫЕ НА ОБСУЖДЕНИЕ")
            elif 'вывод' in h_text.lower():
                add_heading_1(doc, "ВЫВОДЫ")
            else:
                add_heading_2(doc, h_text)
                current_task_title = h_text
            i += 1
            continue
            
        if stripped.startswith('#### '):
            h_text = stripped[5:].strip()
            add_heading_3(doc, h_text)
            current_task_title = h_text
            i += 1
            continue
            
        # Списки
        m_num = re.match(r'^(\d+)\.\s+(.*)', stripped)
        if m_num:
            add_list_item(doc, m_num.group(2), is_numbered=True, num_index=int(m_num.group(1)))
            i += 1
            continue
            
        m_bullet = re.match(r'^[*\-–]\s+(.*)', stripped)
        if m_bullet:
            add_list_item(doc, m_bullet.group(1), is_numbered=False)
            i += 1
            continue
            
        # Обычный абзац текста
        add_paragraph_with_inline(doc, stripped)
        i += 1
        
    # Добавление списка литературы
    add_standard_bibliography(doc)
    
    doc.save(docx_path)
    print(f"Generated: {docx_path}")

def main():
    os.makedirs("Отчеты_Word_Мухамеджанов", exist_ok=True)
    md_files = sorted(glob.glob("docs/practice_*.md"))
    
    print(f"Starting generation of {len(md_files)} DOCX reports by GOST 7.32-2017...")
    for md_file in md_files:
        basename = os.path.basename(md_file)
        m = re.search(r'practice_(\d+)', basename)
        if m:
            num = m.group(1)
            docx_name = f"ПР_{num}_Мухамеджанов_ИИ_в_проф_сфере.docx"
            docx_path = os.path.join("Отчеты_Word_Мухамеджанов", docx_name)
            process_markdown_file(md_file, docx_path)
            
    print("ALL REPORTS GENERATED SUCCESSFULLY!")

if __name__ == '__main__':
    main()
