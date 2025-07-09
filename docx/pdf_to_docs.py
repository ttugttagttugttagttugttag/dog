import fitz, os, re
from docx import Document
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.table import Table
from docx.text.paragraph import Paragraph
from docx.shared import Pt, RGBColor, Inches, Cm
from sentence_transformers import SentenceTransformer, util
from docx.enum.section import WD_ORIENT
from docx.enum.table import WD_ROW_HEIGHT_RULE

pdf_path = '/Users/kjb/Desktop/python/opensource/docx/pdf'

pdf_files = sorted([f for f in os.listdir(pdf_path) if f.endswith('.pdf')])
results = []
for pdf_file in pdf_files:
    pdf_file_path = os.path.join(pdf_path, pdf_file)
    doc = fitz.open(pdf_file_path)
    pdf_text = []
    # 각 페이지의 텍스트 추출
    for page in doc:
        page_dict = page.get_text("dict")
        for block in page_dict["blocks"]:
            for line in block.get("lines", []):
                line_text = " ".join([span["text"] for span in line["spans"]])
                pdf_text.append(line_text)
    results.append({
        "file_name": pdf_file,
        "text": pdf_text
    })


doc = Document("/Users/kjb/Desktop/python/opensource/docx/template_docx/보고서.docx")

def get_grid_span(cell):
    tc = cell._tc
    grid_span = tc.xpath('.//w:gridSpan')
    if grid_span:
        return int(grid_span[0].get(qn('w:val')))
    return 1

def get_vmerge_type(cell):
    tc = cell._tc
    v_merge = tc.xpath('.//w:vMerge')
    if v_merge:
        val = v_merge[0].get(qn('w:val'))
        return val if val else "continue"
    return None

def get_table_border_info(table):
    tblPr = table._element.find(qn('w:tblPr'))
    border_info = {}
    if tblPr is not None:
        tblBorders = tblPr.find(qn('w:tblBorders'))
        if tblBorders is not None:
            for side in ['top', 'bottom', 'left', 'right', 'insideH', 'insideV']:
                el = tblBorders.find(qn(f'w:{side}'))
                if el is not None:
                    border_info[side] = {
                        'val': el.get(qn('w:val')) or el.get('w:val'),
                        'sz': el.get(qn('w:sz')) or el.get('w:sz'),
                        'color': el.get(qn('w:color')) or el.get('w:color'),
                        'space': el.get(qn('w:space')) or el.get('w:space'),
                    }
    return border_info if border_info else None

def get_cell_border_info(cell, table_border_info=None):
    tc = cell._tc
    tcPr = tc.get_or_add_tcPr()
    border_info = {}
    borders = tcPr.find(qn('w:tcBorders'))
    if borders is not None:
        for side in ['top', 'bottom', 'left', 'right']:
            el = borders.find(qn(f'w:{side}'))
            if el is not None:
                border_info[side] = {
                    'val': el.get(qn('w:val')) or el.get('w:val'),
                    'sz': el.get(qn('w:sz')) or el.get('w:sz'),
                    'color': el.get(qn('w:color')) or el.get('w:color'),
                    'space': el.get(qn('w:space')) or el.get('w:space'),
                }
    # 셀에 테두리 정보 없으면 표 전체 테두리 정보로 fallback
    if not border_info and table_border_info:
        for side in ['top', 'bottom', 'left', 'right']:
            if side in table_border_info:
                border_info[side] = table_border_info[side]
    return border_info if border_info else None

def get_cell_style_info(cell, table_border_info=None):
    tc = cell._tc
    tcPr = tc.get_or_add_tcPr()
    width_info = {}
    tcW = tcPr.find(qn('w:tcW'))
    if tcW is not None:
        width_info = {
            'width': tcW.get(qn('w:w')) or tcW.get('w:w'),
            'type': tcW.get(qn('w:type')) or tcW.get('w:type'),
        }
    tr = tc.getparent()
    trPr = tr.find(qn('w:trPr')) if tr is not None else None
    height_info = {}
    if trPr is not None:
        trHeight = trPr.find(qn('w:trHeight'))
        if trHeight is not None:
            height_info = {
                'height': trHeight.get(qn('w:val')) or trHeight.get('w:val'),
                'type': trHeight.get(qn('w:hRule')) or trHeight.get('w:hRule'),
            }
    border_info = get_cell_border_info(cell, table_border_info)
    return {
        'border_info': border_info if border_info else None,
        'width_info': width_info if width_info else None,
        'height_info': height_info if height_info else None,
    }

def parse_paragraph(p):
    para_style = {
        'text': '',  # 나중에 run 합쳐서 넣을 거야
        'alignment': p.alignment,
        'runs': [],
        'source': 'text'
    }
    full_text = ''
    for run in p.runs:
        run_style = {
            'text': run.text,
            'font_name': run.font.name,
            'font_size': run.font.size.pt if run.font.size else None,
            'bold': run.bold,
            'italic': run.italic,
            'underline': run.underline,
            'color': str(run.font.color.rgb) if run.font.color else None
        }
        full_text += run.text
        para_style['runs'].append(run_style)
    para_style['text'] = full_text
    text_keys.append(full_text)
    return para_style

def parse_table(table, table_index):
    table_entries = []
    table_border_info = get_table_border_info(table)
    for row_index, row in enumerate(table.rows):
        for col_index, cell in enumerate(row.cells):
            cell_style = get_cell_style_info(cell, table_border_info)
            cell_info = {
                'row': row_index,
                'col': col_index,
                'grid_span': get_grid_span(cell),
                'vmerge': get_vmerge_type(cell),
                'border_info': cell_style.get('border_info'),
                'width_info': cell_style.get('width_info'),
                'height_info': cell_style.get('height_info'),
                'paragraphs': []
            }

            # 셀 전체 텍스트 합치기용
            merged_text = ''

            for p in cell.paragraphs:
                para_data = {
                    'text': '',
                    'alignment': p.alignment,
                    'runs': []
                }
                para_text = ''
                for run in p.runs:
                    run_style = {
                        'text': run.text,
                        'font_name': run.font.name,
                        'font_size': run.font.size.pt if run.font.size else None,
                        'bold': run.bold,
                        'italic': run.italic,
                        'underline': run.underline,
                        'color': str(run.font.color.rgb) if run.font.color else None
                    }
                    para_text += run.text
                    para_data['runs'].append(run_style)

                para_data['text'] = para_text
                merged_text += para_text  # 모든 문단 텍스트 합치기
                cell_info['paragraphs'].append(para_data)

            # 합쳐진 텍스트로 등록
            if merged_text.strip():
                table_keys.append(merged_text.strip())

            table_entries.append(cell_info)
    return table_entries


# 페이지(섹션) 단위 템플릿 구조 생성
template_styles = []
text_keys = []
table_keys = []

# 모든 페이지(섹션) 기준으로 반복
for page_index, section in enumerate(doc.sections):
    page_style = {
        'source': 'page',
        'page_index': page_index,
        'page_settings': {
            'page_width_cm': round(section.page_width.cm, 2),
            'page_height_cm': round(section.page_height.cm, 2),
            'orientation': 'landscape' if section.orientation == 1 else 'portrait',
            'top_margin_cm': round(section.top_margin.cm, 2),
            'bottom_margin_cm': round(section.bottom_margin.cm, 2),
            'left_margin_cm': round(section.left_margin.cm, 2),
            'right_margin_cm': round(section.right_margin.cm, 2),
            'header_distance_cm': round(section.header_distance.cm, 2),
            'footer_distance_cm': round(section.footer_distance.cm, 2),
            'gutter_cm': round(section.gutter.cm, 2),
        },
        'content': []
    }

    # 문서 전체에서 본문 요소 순회
    tbl_idx = 0
    for block in doc.element.body:
        if block.tag == qn('w:p'):
            p = Paragraph(block, doc)
            para = parse_paragraph(p)
            page_style['content'].append(para)
        elif block.tag == qn('w:tbl'):
            t = Table(block, doc)
            table_data = {
                'source': 'table',
                'table_index': tbl_idx,
                'table_border_info': get_table_border_info(t),
                'cells': parse_table(t, tbl_idx)
            }
            page_style['content'].append(table_data)
            tbl_idx += 1

    # 이 페이지를 전체 리스트에 추가
    template_styles.append(page_style)
text_keys = list(dict.fromkeys(text_keys))
table_keys = list(dict.fromkeys(table_keys))

def split_meaningful(text):
    # "항목 : 값" 패턴에서 의미 단위로 분리
    parts = re.split(r'\s*(:)\s*', text)
    return parts
model = SentenceTransformer("jhgan/ko-sbert-nli")
keywords = text_keys + table_keys
keywords = [
    part for keyword in keywords
    for part in re.split(r'\s*[:：]\s*', keyword)
    if part
]
keyword_emb = model.encode(keywords, convert_to_tensor=True)

updated_texts = []

for text in results[0]['text']:
    words = split_meaningful(text)
    new_words = []
    for word in words:
        word_emb = model.encode(word, convert_to_tensor=True)
        cos_scores = util.cos_sim(word_emb, keyword_emb)[0]
        max_score = cos_scores.max().item()
        best_idx = cos_scores.argmax().item()
        
        if max_score >= 0.7:
            mapped_word = keywords[best_idx]
            new_words.append(mapped_word)
        else:
            new_words.append(word)
    
    updated_text = ' '.join(new_words)
    updated_texts.append(updated_text)
    
# 결과 출력
for original, updated in zip(results[0]['text'], updated_texts):
    print(f"원문: {original}")
    print(f"수정: {updated}")
results[0]['text'] = updated_texts

def set_section_settings(section, page_settings):
    # page_settings 예시: {'page_width_cm': 21.0, ...}
    if 'orientation' in page_settings:
        if page_settings['orientation'] == 'landscape':
            section.orientation = WD_ORIENT.LANDSCAPE
        else:
            section.orientation = WD_ORIENT.PORTRAIT
    if 'page_width_cm' in page_settings:
        section.page_width = Cm(page_settings['page_width_cm'])
    if 'page_height_cm' in page_settings:
        section.page_height = Cm(page_settings['page_height_cm'])
    if 'top_margin_cm' in page_settings:
        section.top_margin = Cm(page_settings['top_margin_cm'])
    if 'bottom_margin_cm' in page_settings:
        section.bottom_margin = Cm(page_settings['bottom_margin_cm'])
    if 'left_margin_cm' in page_settings:
        section.left_margin = Cm(page_settings['left_margin_cm'])
    if 'right_margin_cm' in page_settings:
        section.right_margin = Cm(page_settings['right_margin_cm'])
    if 'header_distance_cm' in page_settings:
        section.header_distance = Cm(page_settings['header_distance_cm'])
    if 'footer_distance_cm' in page_settings:
        section.footer_distance = Cm(page_settings['footer_distance_cm'])
    if 'gutter_cm' in page_settings:
        section.gutter = Cm(page_settings['gutter_cm'])

# --- 스타일 및 레이아웃 적용 함수 ---
def apply_table_borders(table, border_info):
    if not border_info:
        return
    tbl = table._tbl
    tblPr = tbl.find(qn('w:tblPr'))
    if tblPr is None:
        tblPr = OxmlElement('w:tblPr')
        tbl.insert(0, tblPr)
    tblBorders = tblPr.find(qn('w:tblBorders'))
    if tblBorders is None:
        tblBorders = OxmlElement('w:tblBorders')
        tblPr.append(tblBorders)
    for side, style in border_info.items():
        if style:
            side_el = tblBorders.find(qn(f'w:{side}'))
            if side_el is None:
                side_el = OxmlElement(f'w:{side}')
                tblBorders.append(side_el)
            side_el.set(qn('w:val'), style.get('val', 'single'))
            side_el.set(qn('w:sz'), style.get('sz', '4'))
            side_el.set(qn('w:color'), style.get('color', '000000'))
            if style.get('space'):
                side_el.set(qn('w:space'), style['space'])

def apply_cell_border(cell, border_info):
    if not border_info:
        return
    tc = cell._tc
    tcPr = tc.get_or_add_tcPr()
    borders = tcPr.find(qn('w:tcBorders'))
    if borders is None:
        borders = OxmlElement('w:tcBorders')
        tcPr.append(borders)
    for side, style in border_info.items():
        if style:
            side_el = borders.find(qn(f'w:{side}'))
            if side_el is None:
                side_el = OxmlElement(f'w:{side}')
                borders.append(side_el)
            side_el.set(qn('w:val'), style.get('val', 'single'))
            side_el.set(qn('w:sz'), style.get('sz', '4'))
            side_el.set(qn('w:color'), style.get('color', '000000'))
            if style.get('space'):
                side_el.set(qn('w:space'), style['space'])

def apply_column_widths(table, col_widths):
    for col_idx, width in enumerate(col_widths):
        if width == 0:
            continue
        pt_width = int(width) / 20
        inch_width = pt_width / 72
        try:
            table.columns[col_idx].width = Inches(inch_width)
        except:
            pass
        for row in table.rows:
            cell = row.cells[col_idx]
            tc = cell._tc
            tcPr = tc.get_or_add_tcPr()
            tcW = tcPr.find(qn('w:tcW'))
            if tcW is None:
                tcW = OxmlElement('w:tcW')
                tcPr.append(tcW)
            tcW.set(qn('w:w'), str(int(width)))
            tcW.set(qn('w:type'), 'dxa')

def apply_row_heights(table, row_heights):
    for row_idx, height in enumerate(row_heights):
        if height is None or height == 0:
            continue
        row = table.rows[row_idx]
        pt_height = int(height) / 20
        row.height = Pt(pt_height)
        row.height_rule = 1  # exact
        tr = row._tr
        trPr = tr.get_or_add_trPr()
        trHeight = trPr.find(qn('w:trHeight'))
        if trHeight is None:
            trHeight = OxmlElement('w:trHeight')
            trPr.append(trHeight)
        trHeight.set(qn('w:val'), str(int(height)))
        trHeight.set(qn('w:hRule'), 'exact')

def set_cell_vertical_margins(cell, margin_top=0, margin_bottom=0):
    tc = cell._tc
    tcPr = tc.get_or_add_tcPr()
    tcMar = tcPr.find(qn('w:tcMar'))
    if tcMar is None:
        tcMar = OxmlElement('w:tcMar')
        tcPr.append(tcMar)
    for side, value in (('top', margin_top), ('bottom', margin_bottom)):
        el = tcMar.find(qn(f'w:{side}'))
        if el is None:
            el = OxmlElement(f'w:{side}')
            tcMar.append(el)
        el.set(qn('w:w'), str(value))
        el.set(qn('w:type'), 'dxa')

# --- 문서 생성 및 텍스트 채우기 ---
def restore_doc_from_template_and_ocr(template_styles, doc, lines):
    skip_cells = set()  # 표 셀 병합시 중복 방지

    for item in template_styles['content']:
        if item['source'] == 'text':
            para = doc.add_paragraph()
            para.alignment = item['alignment']
            cell_key = item['text']
            print(f"source == text")
            print(f"cell_key = {cell_key}")
            matched_line = next((line for line in lines if cell_key and cell_key in line), None)
            print(f"matched_line = {matched_line}")
            final_text = matched_line if matched_line else cell_key
            print(f"final_text = {final_text}")

            run = para.add_run(final_text)
            if item['runs']:
                print(f"형식 존재 O {final_text}\n\n")
                run.font.name = item['runs'][0]['font_name']
                if item['runs'][0]['font_size']:
                    run.font.size = Pt(item['runs'][0]['font_size'])
                run.bold = item['runs'][0]['bold']
                run.italic = item['runs'][0]['italic']
                run.underline = item['runs'][0]['underline']
                if item['runs'][0]['color'] and item['runs'][0]['color'] != 'None':
                    run.font.color.rgb = RGBColor.from_string(item['runs'][0]['color'])
                rPr = run._element.get_or_add_rPr()
                rFonts = rPr.find(qn('w:rFonts'))
                if rFonts is None:
                    rFonts = OxmlElement('w:rFonts')
                    rPr.append(rFonts)
                if item['runs'][0]['font_name']:
                    rFonts.set(qn('w:eastAsia'), item['runs'][0]['font_name'])
            else:
                print(f"형식 존재 X {final_text}\n\n")
                run.font.name = '맑은 고딕'
                run.font.size = Pt(10.5)
                run.bold = False
                run.italic = False
                run.underline = False

        elif item['source'] == 'table':
            print(f"source == table")
            cells = item['cells']
            max_row = max(cell['row'] for cell in cells) + 1
            max_col = max(cell['col'] for cell in cells) + 1
            table = doc.add_table(rows=max_row, cols=max_col)
            apply_table_borders(table, item.get('table_border_info'))

            # 열 너비, 행 높이 설정
            col_widths = [0] * max_col
            row_heights = [0] * max_row
            for cell in cells:
                if cell['width_info'] and 'width' in cell['width_info']:
                    col_widths[cell['col']] = max(col_widths[cell['col']], float(cell['width_info']['width']))
                if cell['height_info'] and 'height' in cell['height_info']:
                    row_heights[cell['row']] = max(row_heights[cell['row']], float(cell['height_info']['height']))

            apply_column_widths(table, col_widths)
            apply_row_heights(table, row_heights)

            for idx, cell in enumerate(cells):
                row, col = cell['row'], cell['col']
                if (row, col) in skip_cells:
                    continue
                tcell = table.cell(row, col)
                apply_cell_border(tcell, cell.get('border_info'))
                set_cell_vertical_margins(tcell)

                grid_span = cell.get('grid_span', 1)
                if grid_span > 1:
                    try:
                        tcell.merge(table.cell(row, col + grid_span - 1))
                        for k in range(1, grid_span):
                            skip_cells.add((row, col + k))
                    except:
                        pass

                if cell.get('vmerge') == 'restart':
                    merged_tcell = tcell
                    for k in range(1, 20):
                        next_row = row + k
                        match_entry = next(
                            (e for e in cells if e.get('row') == next_row and e.get('col') == col and e.get('vmerge') == 'continue'),
                            None
                        )
                        if not match_entry:
                            break
                        try:
                            merged_tcell = merged_tcell.merge(table.cell(next_row, col))
                            skip_cells.add((next_row, col))
                        except:
                            break
                if cell.get('vmerge') == 'continue':
                    continue

                raw_text = cell['paragraphs'][0]['text']
                cell_key = re.sub(r'[^\w\sㄱ-ㅎ가-힣]', '', raw_text)
                print(f"raw_text = {raw_text}\ncell_key = {cell_key}")
                matched_line = next((line for line in lines if cell_key and cell_key in line), None)
                if matched_line:
                    lines.remove(matched_line)
                print(f"matched_line = {matched_line}")
                if matched_line:
                    match = re.match(rf"{re.escape(raw_text)}\s*[:：]\s*(.*)", matched_line)
                    print(f"match = {match}")
                    if match:
                        next_filled_value = match.group(1).strip()
                        current_text = raw_text
                        item['cells'][idx]['filled_value'] = current_text
                        print(f"next_filled_value = {next_filled_value}")
                        if idx + 1 < len(item['cells']):
                            next_cell = item['cells'][idx + 1]
                            if next_cell['grid_span'] == 1 and next_cell['paragraphs'][0]['text'] == '' and next_cell['row'] == item['cells'][idx]['row']:
                                next_cell['filled_value'] = next_filled_value
                                row = next_cell['row']
                                col = next_cell['col']
                                cell_pos = (row, col)
                                if cell_pos in skip_cells:
                                    skip_cells.remove(cell_pos)
                            else:
                                for k in range(idx + 1, len(item['cells'])):
                                    find_flag = False
                                    print(f"k = {k}")
                                    if(item['cells'][k]['grid_span'] != 1 and item['cells'][k]['paragraphs'][0]['text'] == current_text): continue
                                    if(item['cells'][k]['paragraphs'][0]['text'] == ''):
                                        for j in range(k, len(item['cells'])):
                                            if(item['cells'][j]['col'] == item['cells'][idx]['col'] and item['cells'][j]['paragraphs'][0]['text'] == ''):
                                                item['cells'][j]['filled_value'] = next_filled_value
                                                print(f"({item['cells'][j]['row']},{item['cells'][j]['col']})")
                                                row = item['cells'][j]['row']
                                                col = item['cells'][j]['col']
                                                cell_pos = (row, col)
                                                if cell_pos in skip_cells:
                                                    skip_cells.remove(cell_pos)
                                                find_flag = True
                                                break
                                    if(find_flag): break
                    else:
                        current_text = matched_line
                        item['cells'][idx]['filled_value'] = current_text
                else:
                    if 'filled_value' not in item['cells'][idx]:
                        print(f"no filled value")
                        item['cells'][idx]['filled_value'] = raw_text

                print(f"filled_value = {item['cells'][idx]['filled_value']}\n\n")
                final_text = item['cells'][idx]['filled_value']
                if item['cells'][idx]['paragraphs'][0]['runs']:
                    para = tcell.paragraphs[0] if tcell.paragraphs else tcell.add_paragraph()
                    para.alignment = item['cells'][idx]['paragraphs'][0]['alignment']
                    run_data = item['cells'][idx]['paragraphs'][0]['runs'][0]
                    run = para.add_run(final_text)
                    print(f"({row}, {col})에작성중 ... {final_text}\n\n")
                    run.font.name = run_data['font_name']
                    if run_data['font_size']:
                        run.font.size = Pt(run_data['font_size'])
                    run.bold = run_data['bold']
                    run.italic = run_data['italic']
                    run.underline = run_data['underline']
                    if run_data['color'] and run_data['color'] != 'None':
                        run.font.color.rgb = RGBColor.from_string(run_data['color'])
                else:
                    para = tcell.paragraphs[0] if tcell.paragraphs else tcell.add_paragraph
                    run = para.add_run(final_text)
                    print(f"else({row}, {col})에작성중 ...{final_text}\n")
                    run.font.name = '맑은 고딕'
                    run.font.size = Pt(10.5)
                    run.bold = False
                    run.italic = False
                    run.underline = False
    doc.add_page_break()

if __name__ == "__main__":
    doc = Document()

    for page_idx, result in enumerate(results):
        print(f"\n📄 OCR 페이지 {page_idx + 1} 시작")
        for template_idx, styles in enumerate(template_styles):
            if styles['source'] != 'page':
                continue

            # 각 템플릿별 섹션 추가
            if page_idx == 0 and template_idx == 0:
                section = doc.sections[0]
            else:
                section = doc.add_section(0)

            print(f"➤ 템플릿 {template_idx + 1} 적용")
            set_section_settings(section, styles['page_settings'])
            print(f"{result['text']}")
            restore_doc_from_template_and_ocr(styles, doc, result['text'])

    doc.save("pdf_to_docx.docx")

