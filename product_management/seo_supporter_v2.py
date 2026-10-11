
from pathlib import Path as _ConfigPath
import sys as _config_sys
_config_sys.path.insert(0, str(_ConfigPath(__file__).resolve().parents[1]))
from commerce_config import configured_url, output_dir
_GODOMALL_BASE_URL = configured_url('godomall_base_url')

import re
import time
import traceback
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import NoSuchElementException
from tkinter import Tk, Button, Entry
import threading

# 브라우저 설정
options = webdriver.ChromeOptions()
options.add_argument('--disable-gpu')
options.add_argument('--no-sandbox')
options.add_argument('--disable-popup-blocking')
driver = webdriver.Chrome(options=options)

# URL 설정
login_url = _GODOMALL_BASE_URL + '/base/login.php'

def show_login_popup():
    def on_confirm():
        print("사용자가 로그인 완료를 확인했습니다.")
        root.destroy()

    root = Tk()
    root.title("로그인 요청")
    root.geometry("300x100")
    root.resizable(False, False)
    Button(root, text="로그인을 완료한 후 이 버튼을 눌러주세요.", command=on_confirm).pack(pady=30)
    root.mainloop()

def show_search_popup():
    def on_confirm():
        global search_url
        search_url = driver.current_url  # 현재 URL을 자동으로 수집
        root.destroy()

    root = Tk()
    root.title("상품검색 요청")
    root.geometry("300x100")
    root.resizable(False, False)
    Button(root, text="상품검색을 완료한 후 '확인' 버튼을 눌러주세요.", command=on_confirm).pack(pady=30)
    root.mainloop()

def clean_and_remove_duplicates(text_list):
    """입력된 텍스트 리스트에서 기호를 치환하고 중복 단어를 제거합니다."""
    cleaned_keywords = []
    for text in text_list:
        # 텍스트에서 치환 작업 수행
        text = text.replace('/', ', ')
        text = text.replace('>', ', ')
        text = text.replace('(', ', ')
        text = text.replace(')', ', ')
        text = text.replace('[', ', ')  # 대괄호 열기 치환
        text = text.replace(']', ', ')  # 대괄호 닫기 치환
        text = text.replace('_', ', ')  # 언더스코어도 치환
        cleaned_keywords.extend(text.split(', '))  # 치환된 텍스트를 분리하여 추가

    # 중복 제거 및 정렬
    unique_keywords = list(set(cleaned_keywords))  # 중복 제거
    unique_keywords = [kw.strip() for kw in unique_keywords if kw.strip()]  # 공백 제거 및 빈 값 제거
    unique_keywords.sort()  # (선택사항) 정렬
    return unique_keywords

def extract_and_update_product(i):
    try:
        row_index = 1 + (i * 2)  # 상품 행의 XPATH 계산
        product_xpath = f"/html/body/table[2]/tbody/tr[1]/td[3]/table/tbody/tr/td/table/tbody/tr[{row_index}]/td[3]/div/a[3]"
        product_link = WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.XPATH, product_xpath)))
        product_link.click()

        # 새 창으로 전환
        driver.switch_to.window(driver.window_handles[-1])

        # 카테고리, 키워드, 상품명, 브랜드명 추출
        # 브랜드 이름 기본값 설정
        brand_name = '고캠프'  # 기본값 설정
        # goodsnm_element에서 브랜드 이름 추출 시도
        try:
            goodsnm_element = driver.find_element(By.XPATH, '//*[@id="goodsnm"]')
            product_name = goodsnm_element.get_attribute('value')  # 요소의 value 값을 가져옴

            # 대괄호 안의 브랜드 이름 추출
            brand_match = re.search(r'\[(.*?)\]', product_name)
            if brand_match:
                brand_name = brand_match.group(1)
            else:
                # 2. 대괄호가 없으면 첫 단어를 브랜드로 추출
                brand_name = product_name.split()[0] if product_name else brand_name    
            

        except Exception as e:
            print(f"브랜드 이름 추출 중 오류 발생: {e}, 기본 브랜드명 사용 '{brand_name}'.")

        
        # 3. 카테고리 데이터 수집
        filter_keywords = ["할인제외상품", "네이버", "신세계", "제휴업체이벤트", "백패킹죤"]  # 필터링 단어 리스트
        category_elements = driver.find_elements(By.XPATH, '//*[@id="objCategory"]/tbody/tr/td[1]')
        filtered_category = [element.text.strip() for element in category_elements if element.text.strip() not in filter_keywords]

        # 4. 상품명 치환 작업
        product_name_cleaned = product_name.replace('[', ', ').replace(']', ', ')  # 대괄호 치환

        # 5. 키워드 조합 및 중복 제거
        raw_keywords = filtered_category + [product_name_cleaned, brand_name]
        final_keywords = clean_and_remove_duplicates(raw_keywords)  # 중복 제거 및 정리
        unique_keywords_2 = list(set(final_keywords))  # 중복 제거
        final_keywords_text = ', '.join(unique_keywords_2)  # 문자열로 변환


        # 입력 필드에 값 설정
        seo_title_element = driver.find_element(By.XPATH, '//*[@id="seo_title"]')
        seo_title_element.clear()
        seo_title_element.send_keys(product_name)

        seo_auothor_element = driver.find_element(By.XPATH, '//*[@id="seo_author"]')
        seo_auothor_element.clear()
        seo_auothor_element.send_keys(brand_name)

        seo_keywords_element = driver.find_element(By.XPATH, '//*[@id="seo_keywords"]')
        seo_keywords_element.clear()
        seo_keywords_element.send_keys(final_keywords_text)
    
        # 저장 버튼 클릭
        driver.find_element(By.XPATH, "//*[@id='enamoo-anchor-helper']/ul/li[1]/button").click()
        time.sleep(2)
        print(f"Product {i + 1} updated.")

        # 창 닫기 및 원래 창으로 돌아가기
        driver.close()
        driver.switch_to.window(driver.window_handles[0])

    except Exception as e:
        print(f"상품 {i + 1} 처리 중 오류 발생: {e}")
        traceback.print_exc()
        # 현재 창 닫기 및 원래 창으로 복귀
        driver.close()
        driver.switch_to.window(driver.window_handles[0])

def process_products_on_page(current_page, search_url):
    """한 페이지의 모든 상품을 처리"""
    while True:
        try:
            product_list_url = search_url + f"&page={current_page}"
            driver.get(product_list_url)
            print(f"페이지 {current_page}로 이동했습니다.")

            for i in range(100):  # 페이지당 최대 100개 상품
                try:
                    extract_and_update_product(i)
                except NoSuchElementException:
                    print("더 이상 상품이 없습니다. 페이지 처리 종료.")
                    break
        except Exception as e:
            print(f"상품 처리 중 오류 발생: {e}")
            traceback.print_exc()

        # 페이지 번호 증가
        current_page += 1

def main():
    try:
        # 1. 로그인 페이지로 이동
        driver.get(login_url)
        print("로그인 페이지로 이동했습니다.")

        # 2. 로그인 팝업 표시
        popup_thread = threading.Thread(target=show_login_popup)
        popup_thread.start()
        popup_thread.join()

        # 로그인 확인
        WebDriverWait(driver, timeout=300).until(lambda d: "adm_login_cert.php" not in d.current_url)
        print("로그인 완료")

        # 3. 상품 리스트 페이지로 이동
        driver.get(_GODOMALL_BASE_URL + '/goods/goods_list.php')  # 직접 상품리스트 페이지로 이동
        show_search_popup()

        # 4. 페이지 순차적으로 처리
        current_page = 1
        while True:
            url = f"{search_url}?page={current_page}"
            driver.get(url)

            process_products_on_page(current_page, url)
            
            # 더 이상 상품이 없다면 종료
            if not driver.find_elements(By.XPATH, '//*[@id="objList"]/tbody/tr'):  # 상품 목록이 없다면
                print("현재 페이지 처리를 완료했습니다.")
                break

    finally:
        driver.quit()

if __name__ == "__main__":
    main()
