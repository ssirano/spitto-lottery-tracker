from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options
from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from datetime import datetime
import pandas as pd
import os
import time
import re

def setup_driver():
    options = Options()
    options.add_argument('--headless')
    options.add_argument('--no-sandbox')
    options.add_argument('--disable-dev-shm-usage')
    service = Service(ChromeDriverManager().install())
    return webdriver.Chrome(service=service, options=options)

def extract_prize_info(text_block):
    lines = text_block.split('\n')
    result = {'기준일': '', '1등금액': '', '2등금액': '', '3등금액': '', 
              '1등잔여': '', '2등잔여': '', '3등잔여': '', '판매점입고율': ''}
    
    # 현재 처리 중인 등수
    current_rank = None
    found_amounts = False
    매_counts = []
    
    for i, line in enumerate(lines):
        # 기준일 찾기
        if '기준' in line and not result['기준일']:
            result['기준일'] = line.strip()
        
        # 당첨금액 찾기
        if '억원' in line or '천만원' in line or '백만원' in line or '만원' in line or '천원' in line:
            if '10억원' in line or '5억원' in line or '2억원' in line:
                result['1등금액'] = line.strip()
            elif '1억원' in line or '2천만원' in line or '1백만원' in line:
                result['2등금액'] = line.strip()
            elif '1천만원' in line or '1만원' in line or '5천원' in line:
                result['3등금액'] = line.strip()
            found_amounts = True
        
        # 잔여매수 찾기
        if found_amounts and '매' in line:
            매_counts.append(line.replace('매', '').replace(',', '').strip())
        
        # 판매점 입고율 찾기
        if '%' in line:
            result['판매점입고율'] = line.replace('%', '').strip()
    
    # 잔여매수 할당
    if len(매_counts) >= 3:
        result['1등잔여'] = 매_counts[0]
        result['2등잔여'] = 매_counts[1]
        result['3등잔여'] = 매_counts[2]
    
    return result

def get_spitto_data():
    driver = setup_driver()
    wait = WebDriverWait(driver, 10)
    all_data = {}
    seen_data = set()
    
    try:
        driver.get('https://dhlottery.co.kr/common.do?method=main')
        time.sleep(3)
        
        while True:
            spitto_section = wait.until(EC.presence_of_element_located((By.CLASS_NAME, 'speetto-new')))
            section_text = spitto_section.text
            
            if section_text in seen_data:
                break
            
            seen_data.add(section_text)
            
            # 스피또 게임 정보 추출 (500 제외)
            games = re.findall(r'스피또 (?:1000|2000) \d+회', section_text)
            for game in games:
                if game not in all_data:
                    game_text_start = section_text.find(game)
                    next_game_start = float('inf')
                    for next_game in games:
                        if next_game != game:
                            pos = section_text.find(next_game)
                            if pos > game_text_start and pos < next_game_start:
                                next_game_start = pos
                    
                    game_text = section_text[game_text_start:] if next_game_start == float('inf') else section_text[game_text_start:next_game_start]
                    prize_info = extract_prize_info(game_text)
                    all_data[game] = prize_info
            
            next_button = driver.find_element(By.CLASS_NAME, 'slick-next')
            driver.execute_script("arguments[0].click();", next_button)
            time.sleep(2)
        
        results = []
        for game_name, info in all_data.items():
            game_type = game_name.split()[1]
            game_round = game_name.split()[2].replace('회', '')
            
            results.append({
                '게임명': f'스피또{game_type}',
                '회차': game_round,
                '기준일': info['기준일'],
                '1등당첨금': info['1등금액'],
                '2등당첨금': info['2등금액'],
                '3등당첨금': info['3등금액'],
                '1등잔여매수': info['1등잔여'],
                '2등잔여매수': info['2등잔여'],
                '3등잔여매수': info['3등잔여'],
                '판매점입고율': info['판매점입고율']
            })
        
        return results
        
    except Exception as e:
        print(f'오류 발생: {str(e)}')
        return []
    finally:
        driver.quit()

def save_to_excel(results):
    if not results:
        return
    
    # 데이터 포맷팅
    formatted_results = []
    for row in results:
        formatted_row = row.copy()
        formatted_row['1등잔여매수'] = f"{row['1등잔여매수']}매"
        formatted_row['2등잔여매수'] = f"{row['2등잔여매수']}매"
        formatted_row['3등잔여매수'] = f"{row['3등잔여매수']}매"
        formatted_row['판매점입고율'] = f"{row['판매점입고율']}%"
        formatted_results.append(formatted_row)
    
    current_time = datetime.now().strftime('%Y%m%d_%H%M%S')
    output_dir = 'spitto_results'
    if not os.path.exists(output_dir):
        os.makedirs(output_dir)
    
    df = pd.DataFrame(formatted_results)
    excel_path = os.path.join(output_dir, f'스피또_현황_{current_time}.xlsx')
    
    writer = pd.ExcelWriter(excel_path, engine='xlsxwriter')
    df.to_excel(writer, index=False, sheet_name='스피또 현황')
    
    worksheet = writer.sheets['스피또 현황']
    workbook = writer.book
    
    # 포맷 정의
    header_format = workbook.add_format({
        'bold': True,
        'bg_color': '#D9E1F2',
        'border': 1,
        'align': 'center',
        'text_wrap': True,
        'valign': 'vcenter',
        'font_size': 11
    })
    
    cell_format = workbook.add_format({
        'border': 1,
        'align': 'center',
        'valign': 'vcenter',
        'font_size': 11
    })
    
    # A4 가로 페이지에 맞추기 위한 설정
    worksheet.set_landscape()  # 가로 방향
    worksheet.set_paper(9)     # A4 용지
    worksheet.set_margins(left=0.5, right=0.5, top=0.5, bottom=0.5)  # 여백 설정
    
    # 열 너비 설정 (A4 가로 기준으로 조정)
    column_widths = {
        '게임명': 12,
        '회차': 8,
        '기준일': 15,
        '1등당첨금': 12,
        '2등당첨금': 12,
        '3등당첨금': 12,
        '1등잔여매수': 12,
        '2등잔여매수': 12,
        '3등잔여매수': 12,
        '판매점입고율': 12
    }
    
    # 열 설정 적용
    for idx, col in enumerate(df.columns):
        width = column_widths.get(col, 10)
        worksheet.set_column(idx, idx, width, cell_format)
        worksheet.write(0, idx, col, header_format)
    
    # 행 높이 설정
    worksheet.set_default_row(25)
    worksheet.set_row(0, 30)  # 헤더 행 높이
    
    # 테두리와 그리드 라인 추가
    worksheet.print_gridlines = True
    
    # 인쇄 영역 설정
    worksheet.print_area(0, 0, len(results), len(df.columns)-1)
    
    # 페이지 나누기 설정
    worksheet.fit_to_pages(1, 1)  # 1페이지에 모두 출력
    
    writer.close()
    print(f'결과가 저장되었습니다: {excel_path}')

def main():
    results = get_spitto_data()
    if results:
        save_to_excel(results)
        print(f"수집된 데이터: {results}")
    else:
        print("수집된 데이터가 없습니다.")

if __name__ == '__main__':
    main()


