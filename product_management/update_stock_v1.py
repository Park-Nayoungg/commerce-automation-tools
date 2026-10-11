
from pathlib import Path as _ConfigPath
import sys as _config_sys
_config_sys.path.insert(0, str(_ConfigPath(__file__).resolve().parents[1]))
from commerce_config import configured_url, output_dir
_LEGACY_SHOP_BASE_URL = configured_url('legacy_shop_base_url')

import os
import time
import traceback
import threading
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import NoSuchElementException
from tkinter import Tk, messagebox, Button
from openpyxl import Workbook
from tkinter import filedialog

# 브라우저 설정
options = webdriver.ChromeOptions()
options.add_argument('--disable-gpu')
options.add_argument('--no-sandbox')
options.add_argument('--disable-popup-blocking')
driver = webdriver.Chrome(options=options)

# URL 설정
login_url = _LEGACY_SHOP_BASE_URL + '/shop/admin/login/adm_login_cert.php'



# 로그인 팝업
def show_login_popup(): 
    def on_confirm():
        root.destroy()
    root = Tk()
    root.title("로그인 요청")
    root.geometry("300x100")
    Button(root, text="로그인을 완료한 후 이 버튼을 눌러주세요.", command=on_confirm).pack(pady=30)
    root.mainloop()

# 검색 팝업
def show_search_popup():
    def on_confirm():
        global search_url
        search_url = driver.current_url
        root.destroy()
    root = Tk()
    root.title("상품검색 요청")
    root.geometry("300x100")
    Button(root, text="상품검색을 완료한 후 이 버튼을 눌러주세요.", command=on_confirm).pack(pady=30)
    root.mainloop()

# 계속할지 묻는 팝업
def ask_continue():
    root = Tk()
    root.withdraw()
    result = messagebox.askyesno("계속 여부", "계속해서 다른 상품을 검색하시겠습니까?")
    root.destroy()
    return result

#재고 수정
def update_stock(i):
    try:
        row_index = 1 + (i * 2)
        product_xpath = f"/html/body/table[2]/tbody/tr[1]/td[3]/table/tbody/tr/td/table/tbody/tr[{row_index}]/td[3]/div/a[3]"
        product_link = WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.XPATH, product_xpath)))
        product_link.click()
        driver.switch_to.window(driver.window_handles[-1])

        totstock = driver.find_element(By.ID, 'totstock')

        if not totstock.get_attribute('disabled'):  # totstock이 수정 가능할 때
            tot_value = int(totstock.get_attribute('value') or 0)
            if tot_value == 0:
                print(f"[{i+1}번 상품] totstock 0 -> 수정없이 취소")
                cancel_button = driver.find_element(By.XPATH, '//*[@id="enamoo-anchor-helper"]/ul/li[2]/button')
                cancel_button.click()
            else:
                realstock = int(driver.find_element(By.ID, 'realstock').get_attribute('value') or 0)
                new_stock = 100 + realstock
                totstock.clear()
                totstock.send_keys(str(new_stock))
                print(f"[{i+1}번 상품] totstock 수정: {new_stock}")
                driver.find_element(By.XPATH, '//*[@id="enamoo-anchor-helper"]/ul/li[1]/button').click()

        else:  # totstock이 disabled일 때 옵션별 수정
            option_rows = driver.find_elements(By.CSS_SELECTOR, '#el-option-list tbody tr')
            print(f"[{i+1}번 상품] 옵션 {len(option_rows)}개 확인 중")

            for idx, row in enumerate(option_rows):
                try:
                    stock_input = row.find_element(By.NAME, f'option_stock[{idx}]')
                    realstock_input = row.find_element(By.NAME, f'option_realstock[{idx}]')

                    option_stock = int(stock_input.get_attribute('value') or 0)
                    option_realstock = int(realstock_input.get_attribute('value') or 0)

                    if option_stock <= 0:
                        print(f" - 옵션 {idx+1}: stock {option_stock} → 수정 없이 패스")
                        continue

                    new_stock = option_realstock + 100
                    stock_input.clear()
                    stock_input.send_keys(str(new_stock))
                    print(f" - 옵션 {idx+1}: stock {option_stock} → {new_stock} 변경 (realstock {option_realstock} 기준)")

                except Exception as e:
                    print(f" - 옵션 {idx+1}: 오류 발생 → {e}")

            # 옵션 수정 후 저장
            driver.find_element(By.XPATH, '//*[@id="enamoo-anchor-helper"]/ul/li[1]/button').click()

        time.sleep(2)
        print(f"상품 {i+1} 수정 완료")

        driver.close()
        driver.switch_to.window(driver.window_handles[0])

    except Exception as e:
        print(f"[{i+1}번 상품 오류]: {e}")
        traceback.print_exc()
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
                    update_stock(i)
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
        driver.get(_LEGACY_SHOP_BASE_URL + '/shop/admin/goods/adm_goods_list.php?')  # 직접 상품리스트 페이지로 이동
        show_search_popup()

        # 4. 페이지 순차적으로 처리
        current_page = 1
        while True:
            url = f"{search_url}&page={current_page}"
            print(f"페이지 {current_page} 시작")
            driver.get(url)

            process_products_on_page(current_page, url)
            
            # 더 이상 상품이 없다면 종료
            if not ask_continue():
                break

    finally:
        driver.quit()

if __name__ == "__main__":
    main()
