import requests
from bs4 import BeautifulSoup, NavigableString, Tag
from docx import Document
from docx.shared import Inches
from io import BytesIO
from PIL import Image as PILImage
import re
import os

def clean_text(text):
    """주어진 텍스트에서 특수 문자를 제거하고 원하는 문자열로 치환합니다."""
    if text:
        # NULL 바이트 및 제어 문자 제거
        text = re.sub(r'[\x00-\x1F\x7F]', '', text).replace('â', ' ').strip()
        # 문자열 치환 추가
        text = text.replace('â', '○').replace('â', '✕')
        return text
    return ""

def fetch_and_convert(urls):
    """주어진 URL 목록에서 내용을 가져와 Word 문서로 변환합니다."""
    doc = Document()
    
    for url in urls:
        print(f"Processing {url}...")
        response = requests.get(url)
        response.encoding = 'utf-8'  # 인코딩 설정
        soup = BeautifulSoup(response.text, 'html.parser')

        # theme-doc-markdown 요소 찾기
        content = soup.find(class_='theme-doc-markdown')
        
        # content를 raw HTML 형태로 txt 파일로 저장
        if content:
            with open('content.txt', 'w', encoding='utf-8') as f:
                f.write(str(content))
                f.write('\n\n')

        if content:
            # 콘텐츠를 재귀적으로 순회하여 모든 요소를 처리
            parse_element(content, doc)

    doc.save('output.docx')  # 최종 DOCX 파일 저장

def parse_element(element, doc):
    """HTML 요소를 재귀적으로 순회하여 Word 문서에 추가합니다."""
    if isinstance(element, NavigableString):
        # 텍스트 노드 처리
        text = clean_text(str(element))
        if text:
            # 현재 단락이 있으면 이어서 추가
            if doc.paragraphs and doc.paragraphs[-1].text:
                doc.paragraphs[-1].add_run(text)
            else:
                doc.add_paragraph(text)
        return

    if element.name == 'hr':
        # hr 요소는 처리하지 않고 건너뜁니다.
        return

    if element.name in ['h1', 'h2', 'h3', 'h4', 'h5']:
        level = int(element.name[1])
        heading_text = clean_text(element.get_text(strip=True))
        doc.add_heading(heading_text, level=level)
    elif element.name == 'p':
        paragraph_text = clean_text(element.get_text(strip=True))
        if paragraph_text:
            doc.add_paragraph(paragraph_text)
    elif element.name == 'ul':
        # 리스트 처리
        for li in element.find_all('li', recursive=False):
            parse_element(li, doc)
    elif element.name == 'li':
        li_text = clean_text(element.get_text(strip=True))
        if li_text:
            doc.add_paragraph(li_text, style='ListBullet')
    elif element.name == 'table':
        # 테이블 처리
        process_table(element, doc)
    elif element.name == 'div' and 'theme-admonition' in element.get('class', []):
        # 경고 메시지 처리
        admonition_title = element.find(class_=lambda c: c and 'admonitionHeading' in c)
        admonition_content = element.find(class_=lambda c: c and 'admonitionContent' in c)
        if admonition_title and admonition_content:
            # 제목 처리
            doc.add_paragraph(f"[{clean_text(admonition_title.get_text(strip=True))}]")
            # 내용 처리
            for child in admonition_content.children:
                parse_element(child, doc)
    elif element.name == 'details':
        # details 태그 처리
        summary = element.find('summary')
        if summary:
            summary_text = clean_text(summary.get_text(strip=True))
            doc.add_paragraph(f"Details: {summary_text}", style='Intense Quote')
            details_content = element.find_all('p')
            for p in details_content:
                parse_element(p, doc)
    elif element.name == 'img':
        # 이미지 처리
        img_src = element.get('src')
        if img_src:
            if not img_src.startswith('http'):
                img_src = 'https://docs.whatap.io' + img_src
            try:
                img_data = requests.get(img_src).content
                img = PILImage.open(BytesIO(img_data))
                img.save("temp_image.png")
                doc.add_picture("temp_image.png", width=Inches(5))
                os.remove("temp_image.png")
            except Exception as e:
                print(f"Failed to load image {img_src}: {e}")
    else:
        # 기타 요소의 경우 하위 요소를 재귀적으로 처리
        for child in element.children:
            parse_element(child, doc)

def process_table(table_element, doc):
    """HTML 테이블 요소를 Word 문서에 추가합니다."""
    rows = table_element.find_all('tr')
    if not rows:
        return

    # 테이블의 최대 열 수 계산
    max_cols = 0
    for row in rows:
        col_count = 0
        for cell in row.find_all(['th', 'td']):
            colspan = int(cell.get('colspan', 1))
            col_count += colspan
        if col_count > max_cols:
            max_cols = col_count

    table = doc.add_table(rows=0, cols=max_cols)
    table.style = 'Table Grid'

    grid = []
    row_idx = 0
    for row in rows:
        cells = row.find_all(['th', 'td'])
        while len(grid) <= row_idx:
            grid.append([])
        grid_row = grid[row_idx]
        col_idx = 0
        for cell in cells:
            # 이미 병합된 위치는 건너뜀
            while col_idx < len(grid_row) and grid_row[col_idx] == 'skip':
                col_idx += 1

            rowspan = int(cell.get('rowspan', 1))
            colspan = int(cell.get('colspan', 1))

            # 셀 내부의 HTML을 파싱하여 이모지 변환 및 텍스트 추출
            cell_html = str(cell)
            cell_soup = BeautifulSoup(cell_html, 'html.parser')

            # 이모지 변환 처리
            for emoji_span in cell_soup.find_all('span', class_='emoji-ok'):
                emoji_span.replace_with('○')
            for emoji_span in cell_soup.find_all('span', class_='emoji-no'):
                emoji_span.replace_with('✕')

            # 추가적인 문자열 치환
            cell_text = clean_text(cell_soup.get_text(strip=True))

            # 셀에 텍스트 추가
            if len(table.rows) <= row_idx:
                table.add_row()
            while len(table.rows[row_idx].cells) < max_cols:
                table.rows[row_idx].add_cell('')

            table_cell = table.cell(row_idx, col_idx)
            table_cell.text = cell_text

            # 셀 병합 처리
            if rowspan > 1 or colspan > 1:
                ensure_table_size(table, row_idx + rowspan - 1, col_idx + colspan - 1)
                merge_cell(table, row_idx, col_idx, rowspan, colspan)

            # 병합된 셀 위치 표시
            for i in range(rowspan):
                while len(grid) <= row_idx + i:
                    grid.append([])
                grid_row = grid[row_idx + i]
                while len(grid_row) < col_idx:
                    grid_row.append(None)
                for j in range(colspan):
                    while len(grid_row) <= col_idx + j:
                        grid_row.append(None)
                    if i == 0 and j == 0:
                        grid_row[col_idx + j] = cell_text
                    else:
                        grid_row[col_idx + j] = 'skip'
            col_idx += colspan
        row_idx += 1

def ensure_table_size(table, row_idx, col_idx):
    """테이블의 크기를 조정하여 특정 행과 열에 접근 가능하도록 합니다."""
    while len(table.rows) <= row_idx:
        table.add_row()
    for row in table.rows:
        while len(row.cells) <= col_idx:
            row.add_cell('')

def merge_cell(table, row_idx, col_idx, rowspan, colspan):
    """Word 테이블에서 셀 병합을 처리합니다."""
    start_cell = table.cell(row_idx, col_idx)
    end_row = row_idx + rowspan - 1
    end_col = col_idx + colspan - 1

    ensure_table_size(table, end_row, end_col)

    end_cell = table.cell(end_row, end_col)
    start_cell.merge(end_cell)

# URL 목록 읽기
with open('urls.txt', 'r') as file:
    urls = [line.strip() for line in file if line.strip()]

fetch_and_convert(urls)  # 함수 호출
