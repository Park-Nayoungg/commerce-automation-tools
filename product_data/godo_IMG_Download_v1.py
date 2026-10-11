
from pathlib import Path as _ConfigPath
import sys as _config_sys
_config_sys.path.insert(0, str(_ConfigPath(__file__).resolve().parents[1]))
from commerce_config import configured_url, output_dir
_GODOMALL_BASE_URL = configured_url('godomall_base_url')

import re
import time
import traceback
import os
import mimetypes
from urllib.parse import urlparse, parse_qs, urlencode
from selenium import webdriver
from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
from selenium.common.exceptions import NoSuchElementException
from selenium.webdriver.support import expected_conditions as EC
from tkinter import Tk, messagebox, Button  # 팝업 창 띄우기 위한 라이브러리
from bs4 import BeautifulSoup
from PIL import Image
from io import BytesIO
import time
import requests
import threading  # Tkinter와 Selenium 병렬 실행을 위해 사용


#로그인요청 함수
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

#검색요청 함수
def show_search_popup():
    def on_confirm():
        global search_url
        search_url = driver.current_url  # 현재 URL을 자동으로 수집
        root.destroy()

    root = Tk()
    root.title("상품검색 요청")
    root.geometry("300x100")
    root.resizable(False, False)
    Button(root, text="상품검색을 완료한 후 이 버튼을 눌러주세요.", command=on_confirm).pack(pady=30)
    root.mainloop()

# 페이지 번호를 증가시키며 다운로드 작업 수행하는 함수
def process_pages(driver, save_dir):
    page_num = 1  # 첫 번째 페이지부터 시작
    while True:
        # 현재 URL 정보 가져오기 및 분리
        current_url = driver.current_url
        parsed_url = urlparse(current_url)
        base_url = f"{parsed_url.scheme}://{parsed_url.netloc}{parsed_url.path}"

        # 기존 쿼리 파라미터 파싱
        query_params = parse_qs(parsed_url.query)
        # 딕셔너리 값이 리스트 형태이므로 단일 값으로 정리
        query_params = {k: v[0] for k, v in query_params.items()}

        # 쿼리 파라미터에 page 추가
        query_params = {'page': page_num, **query_params}
        query_string = urlencode(query_params)

        # 최종 페이지 URL 구성
        page_url = f"{base_url}?{query_string}"

        print(f"페이지 {page_num} 로딩: {page_url}")
        driver.get(page_url)
        time.sleep(2)  # 페이지 로딩 대기

        # 각 페이지에서 이미지 다운로드 작업 수행
        if not download_imgs(driver, save_dir):
            print("상품이 존재하지 않는 페이지를 만났습니다. 프로그램을 종료합니다.")
            break  # 상품이 없는 페이지가 나오면 종료

        # 페이지 번호 증가
        page_num += 1
        


def download_imgs(driver, save_dir):
    # 각 페이지에 있는 100개의 상품에 대해 작업 수행
    for i in range(0, 100):  # 100개의 상품
        try:
            # tr[i]의 값은 1, 3, 5, 7, 9로 증가
            row_index = 1 + (i * 2)
            product_xpath = f"/html/body/div[1]/div/div[2]/div[2]/div/form[2]/div[2]/table/tbody/tr[{row_index}]/td[5]/div[1]/a"
            
            # tr 요소가 존재하지 않을 경우 NoSuchElementException 발생
            product_link = driver.find_element(By.XPATH, product_xpath)

            # 새 창에서 열기 버튼 클릭
            product_link.click()
            time.sleep(2)
            
            # 현재 창 핸들 저장
            original_window_handle = driver.current_window_handle

            # 새 창으로 전환
            new_window_handle = driver.window_handles[-1]
            driver.switch_to.window(new_window_handle)

            img_idx = 1
            while True:
                try:
                    # img 태그의 XPath 설정
                    img_xpath = f'/html/body/div[1]/div/form/div[47]/div/div[3]/table/tbody/tr[3]/td/table/tbody/tr[1]/td/div[2]/div[{img_idx}]/span/img'
                    img_element = WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH, img_xpath)))
                
                    
                    # 이미지 URL 가져오기
                    img_url = img_element.get_attribute('src')
                    if not img_url:
                        print(f"이미지 URL이 없습니다: {img_xpath}")
                        break

                    # URL에서 파일 이름 추출
                    file_name = os.path.basename(urlparse(img_url).path)
                    if not file_name:
                        print(f"파일 이름을 추출할 수 없습니다: {img_url}")
                        break

                    # 파일 확장자 확인 및 추가
                    file_extension = file_name.split('.')[-1].lower()
                    if file_extension in ['png']:
                        file_name = os.path.splitext(file_name)[0] + '.png'
                    elif file_extension in ['jpg', 'jpeg']:
                        file_name = os.path.splitext(file_name)[0] + '.jpg'
                    elif file_extension in ['gif']:
                        file_name = os.path.splitext(file_name)[0] + '.gif'

                    # 다운로드할 이미지의 최종 URL 생성
                    base_url = "https://cdn-saas-web-219-28.cdn-nhncommerce.com/nuree2152_godomall_com/data/goods/"
                    file_extension = file_name.split('.')[-1].lower()
                    base_name = os.path.splitext(file_name)[0]
                    possible_extensions = [file_extension] + [ext for ext in ['jpg', 'png', 'gif'] if ext != file_extension]
                    final_img_url = None
                    for ext in possible_extensions:
                        test_url = f"{base_url}{base_name}.{ext}"
                        response = requests.head(test_url)
                        if response.status_code == 200:
                            final_img_url = test_url
                            file_name = f"{base_name}.{ext}"  # 실제 사용 확장자 반영
                            break

                    if final_img_url:
                        driver.execute_script(f"window.open('{final_img_url}', '_blank');")
                        temp_window_handle = driver.window_handles[-1]
                        driver.switch_to.window(temp_window_handle)
                        time.sleep(2)

                    try:
                        # 이미지 요소 찾기
                        img_in_new_window = WebDriverWait(driver, 10).until(
                            EC.presence_of_element_located((By.XPATH, '/html/body/div/img'))
                        )
                        img_src = img_in_new_window.get_attribute('src')

                        if not img_src:
                            print(f"새 창에서 이미지 URL을 찾을 수 없습니다: {final_img_url}")
                        else:
                            # 파일 경로 설정
                            file_path = os.path.join(save_dir, file_name)

                            # 이미지 다운로드
                            img_data = requests.get(img_src).content
                            with open(file_path, 'wb') as handler:
                                handler.write(img_data)

                            print(f'{file_name} 다운로드 완료')

                    except Exception as e:
                        print(f"새 창에서 이미지 다운로드 오류: {e}")

                    # 새 창 닫기 및 원래 창으로 돌아가기
                    driver.close()
                    driver.switch_to.window(new_window_handle)
                    print(f"Switched back to original new window: {new_window_handle}")

                    # 다음 이미지로 이동
                    img_idx += 1

                except Exception as e:
                    # 더 이상 이미지가 없으면 루프 종료
                    print(f"이미지 다운로드 종료: {e}")
                    break

            # 작업 완료 후 창 닫기 (현재 창 닫기)
            driver.close()
            driver.switch_to.window(original_window_handle)
            print(f"Switched back to original window: {original_window_handle}")

        except Exception as e:
            print(f"상품 {i + 1} 처리 중 오류 발생: {e}")
            traceback.print_exc()
            continue

# 실행 로직
if __name__ == "__main__":
    # 브라우저 설정
    options = webdriver.ChromeOptions()
    options.add_argument("--disable-gpu")
    options.add_argument("--no-sandbox")
    options.add_argument("--remote-debugging-port=9222")  # 이거 꼭 추가
    options.add_argument("--disable-popup-blocking")
    #options.add_argument("profile-directory=Default")

    driver = webdriver.Chrome(options=options)

    try:
        # 사용자 로그인 요청
        login_url = _GODOMALL_BASE_URL + '/base/login.php'
        driver.get(login_url)
        show_login_popup()

        # 3. 상품 리스트 페이지로 이동
        driver.get(_GODOMALL_BASE_URL + '/goods/goods_list.php')  # 직접 상품리스트 페이지로 이동
        show_search_popup()

        # 이미지 저장 디렉토리 설정
        save_directory = str(output_dir('downloads'))
        if not os.path.exists(save_directory):
            os.makedirs(save_directory)

        # 상품별 이미지 다운로드 작업 수행
        process_pages(driver, save_directory)

    finally:
        driver.quit()
