import requests
from bs4 import BeautifulSoup
from docx import Document
from docx.shared import Inches  # 인치 단위로 크기 설정
from io import BytesIO
from PIL import Image as PILImage  # 이미지 크기 조정용

def clean_text(text):
    """주어진 텍스트에서 특수 문자를 제거하고 여백을 정리합니다."""
    return text.replace('â', ' ').strip()  # 필요에 따라 수정

def fetch_and_convert(urls):
    """주어진 URL 목록에서 내용을 가져와 Word 문서로 변환합니다."""
    doc = Document()
    
    for url in urls:
        print(f"Processing {url}...")  # 현재 처리 중인 URL 출력
        response = requests.get(url)
        soup = BeautifulSoup(response.text, 'html.parser')

        # theme-doc-markdown 요소 찾기
        content = soup.find(class_='theme-doc-markdown')
        
        # content를 raw HTML 형태로 txt 파일로 저장
        if content:
            with open('content.txt', 'w', encoding='utf-8') as f:
                f.write(str(content))  # 원본 HTML을 문자열로 변환하여 저장
                f.write('\n\n')  # 각 URL의 내용을 구분하기 위해 두 줄 추가

        if content:
            for section in content.find_all(['h1', 'h2', 'h3', 'h4', 'h5']):  # h1 ~ h5 태그 포함
                level = int(section.name[1])  # h1, h2, h3에 따라 레벨 지정
                doc.add_heading(clean_text(section.get_text(strip=True)), level=level)

                # 섹션 아래의 내용을 추가
                for sibling in section.find_next_siblings():
                    if sibling.name in ['h1', 'h2', 'h3', 'h4', 'h5']:
                        break  # 다음 섹션 헤딩까지 추가 중지
                    
                    if sibling.name == 'p':
                        # 단락 추가
                        paragraph_text = []

                        # 단락 내의 모든 요소에서 텍스트를 수집
                        for elem in sibling.children:
                            if elem.name == 'a':
                                # 링크 처리: 링크 텍스트를 수집
                                paragraph_text.append(clean_text(elem.get_text(strip=True)))
                            elif elem.name == 'img':
                                # 이미지 처리
                                img_url = elem['src']
                                # src가 https로 시작하지 않으면 앞에 기본 URL 추가
                                if not img_url.startswith('https://'):
                                    img_url = f'https://docs.whatap.io{img_url}'
                                try:
                                    img_response = requests.get(img_url)
                                    img_response.raise_for_status()  # 요청이 성공했는지 확인

                                    # 이미지 열기
                                    img = PILImage.open(BytesIO(img_response.content))
                                    img.save("temp_image.png")  # 임시로 파일 저장
                                    doc.add_picture("temp_image.png", width=Inches(5))  # 크기를 조정하여 DOCX에 추가

                                except Exception as e:
                                    print(f"Failed to load image {img_url}: {e}")
                            elif elem.string:  # 일반 텍스트 처리
                                paragraph_text.append(clean_text(elem.string))
                                paragraph_text.append(' ')  # 일반 텍스트 뒤에 공백 추가

                        # 최종 문자열을 공백으로 조인하여 단락 추가
                        doc.add_paragraph(''.join(paragraph_text).strip())

                    elif sibling.name == 'ul':
                        # 리스트 추가
                        for li in sibling.find_all('li'):
                            doc.add_paragraph(clean_text(li.get_text(strip=True)), style='ListBullet')  # 글머리 기호 스타일로 추가

                    elif sibling.name == 'table':
                        # 테이블 생성
                        table = doc.add_table(rows=0, cols=len(sibling.find_all('th')))
                        table.style = 'Table Grid'  # 테이블 스타일 설정

                        # 테이블 헤더 추가
                        hdr_cells = table.add_row().cells
                        for i, header in enumerate(sibling.find_all('th')):
                            hdr_cells[i].text = clean_text(header.get_text(strip=True))

                        # 테이블 내용 추가
                        for row in sibling.find_all('tr')[1:]:  # 첫 번째 행은 헤더이므로 건너뜀
                            cells = row.find_all(['td', 'th'])
                            new_row = table.add_row().cells
                            for i, cell in enumerate(cells):
                                new_row[i].text = clean_text(cell.get_text(strip=True))

                    elif sibling.name == 'div' and 'theme-admonition' in sibling.get('class', []):
                        # 경고 메시지 처리
                        admonition_title = sibling.find(class_=lambda c: c and 'admonitionHeading' in c)
                        admonition_content = sibling.find(class_=lambda c: c and 'admonitionContent' in c)
                        if admonition_title and admonition_content:
                            doc.add_paragraph(f"[{clean_text(admonition_title.get_text(strip=True))}] {clean_text(admonition_content.get_text(strip=True))}")

    doc.save('output.docx')  # 최종 DOCX 파일 저장

# URL 목록 읽기
with open('urls.txt', 'r') as file:
    urls = [line.strip() for line in file if line.strip()]

fetch_and_convert(urls)  # 함수 호출
