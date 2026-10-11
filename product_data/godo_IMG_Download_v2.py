
from pathlib import Path as _ConfigPath
import sys as _config_sys
_config_sys.path.insert(0, str(_ConfigPath(__file__).resolve().parents[1]))
from commerce_config import configured_url, output_dir
_LEGACY_SHOP_BASE_URL = configured_url('legacy_shop_base_url')

import re
import time
import traceback
import os
import mimetypes
from urllib.parse import urlparse
from selenium import webdriver
from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
from selenium.common.exceptions import NoSuchElementException
from selenium.webdriver.support import expected_conditions as EC
from tkinter import Tk, messagebox, Button
from bs4 import BeautifulSoup
from PIL import Image
from io import BytesIO
import requests
import threading

# 로그인 요청 함수
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

# 검색 요청 함수
def show_search_popup():
    def on_confirm():
        global search_url
        search_url = driver.current_url
        root.destroy()

    root = Tk()
    root.title("상품검색 요청")
    root.geometry("300x100")
    root.resizable(False, False)
    Button(root, text="상품검색을 완료한 후 이 버튼을 눌러주세요.", command=on_confirm).pack(pady=30)
    root.mainloop()

# 페이지 처리 함수
def process_pages(driver, save_dir, max_pages=1000):
    page_num = 1
    while True:
        page_url = f"{driver.current_url}&page={page_num}"
        print(f"페이지 {page_num} 로딩: {page_url}")
        driver.get(page_url)
        time.sleep(2)

        if not download_imgs(driver, save_dir):
            print("상품이 존재하지 않는 페이지를 만났습니다. 프로그램을 종료합니다.")
            break

        page_num += 1

# 이미지 다운로드 함수
def download_imgs(driver, save_dir):
    for i in range(100):
        try:
            row_index = 1 + (i * 2)
            product_xpath = f"/html/body/table[2]/tbody/tr[1]/td[3]/table/tbody/tr/td/table/tbody/tr[{row_index}]/td[3]/div/a[3]"
            product_link = driver.find_element(By.XPATH, product_xpath)
            product_link.click()
            time.sleep(2)

            original_window_handle = driver.current_window_handle
            new_window_handle = driver.window_handles[-1]
            driver.switch_to.window(new_window_handle)

            img_idx = 1
            while True:
                try:
                    img_xpath = f'//*[@id="tb_l"]/tbody/tr[{img_idx}]/td[3]/img'
                    img_element = WebDriverWait(driver, 10).until(
                        EC.presence_of_element_located((By.XPATH, img_xpath))
                    )

                    img_url = img_element.get_attribute('src')
                    if not img_url:
                        print(f"이미지 URL이 없습니다: {img_xpath}")
                        break

                    file_name = os.path.basename(urlparse(img_url).path)
                    file_base, file_ext = os.path.splitext(file_name)
                    file_ext = file_ext.lower()

                    if file_ext not in ['.png', '.jpg', '.jpeg', '.gif']:
                        print(f"지원하지 않는 확장자: {file_ext}")
                        break

                    final_img_url = img_url  # 원본 그대로 사용

                    driver.execute_script(f"window.open('{final_img_url}', '_blank');")
                    temp_window_handle = driver.window_handles[-1]
                    driver.switch_to.window(temp_window_handle)
                    time.sleep(2)

                    try:
                        img_in_new_window = WebDriverWait(driver, 10).until(
                            EC.presence_of_element_located((By.XPATH, '/html/body/div/img'))
                        )
                        img_src = img_in_new_window.get_attribute('src')

                        if not img_src:
                            print(f"새 창에서 이미지 URL을 찾을 수 없습니다: {final_img_url}")
                        else:
                            file_path = os.path.join(save_dir, file_name)
                            img_data = requests.get(img_src).content
                            with open(file_path, 'wb') as handler:
                                handler.write(img_data)

                            print(f'{file_name} 다운로드 완료')

                    except Exception as e:
                        print(f"새 창에서 이미지 다운로드 오류: {e}")

                    driver.close()
                    driver.switch_to.window(new_window_handle)
                    print(f"Switched back to original new window: {new_window_handle}")

                    img_idx += 1

                except Exception as e:
                    print(f"이미지 다운로드 종료: {e}")
                    break

            driver.close()
            driver.switch_to.window(original_window_handle)
            print(f"Switched back to original window: {original_window_handle}")

        except Exception as e:
            print(f"상품 {i + 1} 처리 중 오류 발생: {e}")
            traceback.print_exc()
            continue

    return True

# 실행 메인
if __name__ == "__main__":
    options = webdriver.ChromeOptions()
    options.add_argument("--disable-gpu")
    options.add_argument("--no-sandbox")
    options.add_argument("--disable-popup-blocking")

    driver = webdriver.Chrome(options=options)

    try:
        login_url = _LEGACY_SHOP_BASE_URL + '/shop/admin/login/adm_login_cert.php'
        driver.get(login_url)
        show_login_popup()

        driver.get(_LEGACY_SHOP_BASE_URL + '/shop/admin/goods/adm_goods_list.php?')
        show_search_popup()

        save_directory = str(output_dir('downloads'))
        if not os.path.exists(save_directory):
            os.makedirs(save_directory)

        process_pages(driver, save_directory)

    finally:
        driver.quit()
