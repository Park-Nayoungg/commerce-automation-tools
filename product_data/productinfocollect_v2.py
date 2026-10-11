
from pathlib import Path as _ConfigPath
import sys as _config_sys
_config_sys.path.insert(0, str(_ConfigPath(__file__).resolve().parents[1]))
from commerce_config import configured_url, output_dir
_LEGACY_SHOP_BASE_URL = configured_url('legacy_shop_base_url')

import os
import time
import traceback
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from tkinter import Tk, messagebox, Button
from openpyxl import Workbook
from tkinter import filedialog


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

# 저장 폴더 선택 함수
def ask_save_directory():
    root = Tk()
    root.withdraw()  # 메인 윈도우 숨기기
    folder_selected = filedialog.askdirectory(title="엑셀 파일을 저장할 폴더를 선택하세요")
    root.destroy()
    return folder_selected

#실재고 추출 개선 코드
def get_option_stock(driver):
    option_texts = []
    tables = driver.find_elements(By.ID, "el-option-list")

    for table in tables:
        try:
            tbody = table.find_element(By.TAG_NAME, "tbody")
            rows = tbody.find_elements(By.TAG_NAME, "tr")

            for tr in rows:
                tds = tr.find_elements(By.TAG_NAME, "td")
                if not tds:
                    continue

                option_parts = []
                realstock = ""

                for td in tds:
                    inputs = td.find_elements(By.TAG_NAME, "input")
                    for inp in inputs:
                        name_attr = inp.get_attribute("name") or ""
                        value = inp.get_attribute("value") or ""

                        if "opt[" in name_attr:
                            option_parts.append(value)
                        elif "option_realstock" in name_attr:
                            realstock = value

                if option_parts and realstock != "":
                    option_label = "/".join(option_parts)
                    option_texts.append(f"{option_label}: {realstock}개")

        except Exception as e:
            continue  # 에러 발생 시 해당 테이블은 무시

    return option_texts

# 상품정보 수집
def collect_product_data(driver, max_pages=78):
    data = []
    main_window = driver.current_window_handle
    page_num = 1

    while page_num <= max_pages:
        page_url = f"{search_url}&page={page_num}"
        driver.get(page_url)
        time.sleep(2)

        if "등록된 상품이 없습니다" in driver.page_source:
            break

        for i in range(0, 100):
            try:
                row_index = 1 + (i * 2)
                detail_row_index = row_index + 1

                name_xpath = f'/html/body/table[2]/tbody/tr[1]/td[3]/table/tbody/tr/td/table/tbody/tr[{row_index}]/td[3]/div/a[2]'
                price_xpath = f'/html/body/table[2]/tbody/tr[1]/td[3]/table/tbody/tr/td/table/tbody/tr[{row_index}]/td[4]'
                stock_xpath = f'/html/body/table[2]/tbody/tr[1]/td[3]/table/tbody/tr/td/table/tbody/tr[{row_index}]/td[5]'
                realstock_xpath = f'/html/body/table[2]/tbody/tr[1]/td[3]/table/tbody/tr/td/table/tbody/tr[{row_index}]/td[7]/input'
                detail_tbody_xpath = f'/html/body/table[2]/tbody/tr[1]/td[3]/table/tbody/tr/td/table/tbody/tr[{detail_row_index}]/td/div/table/tbody'

                name = driver.find_element(By.XPATH, name_xpath).text
                price_text = driver.find_element(By.XPATH, price_xpath).text.strip().split('\n')[0]
                stock = driver.find_element(By.XPATH, stock_xpath).text.strip()

                tbody_elem = driver.find_element(By.XPATH, detail_tbody_xpath)
                option_exists = "옵션" in tbody_elem.text

                if option_exists:
                    # 옵션창 열기
                    option_btn_xpath = f'/html/body/table[2]/tbody/tr[1]/td[3]/table/tbody/tr/td/table/tbody/tr[{row_index}]/td[3]/div/a[3]'
                    driver.find_element(By.XPATH, option_btn_xpath).click()

                    # 새 창으로 전환
                    WebDriverWait(driver, 10).until(EC.number_of_windows_to_be(2))
                    new_window = [w for w in driver.window_handles if w != main_window][0]
                    driver.switch_to.window(new_window)

                    # 옵션 정보 수집
                    time.sleep(1)  # 옵션창 로딩 대기
                    option_texts = get_option_stock(driver)  # → 개선된 함수 호출
                    realstock_combined = "\n".join(option_texts)

                    data.append([name, price_text, stock, realstock_combined])

                    # 옵션창 닫기 및 복귀
                    driver.close()
                    driver.switch_to.window(main_window)

                else:
                    # 일반 상품 처리
                    realstock_value = driver.find_element(By.XPATH, realstock_xpath).get_attribute('value')
                    data.append([name, price_text, stock, realstock_value])

            except Exception as e:
                print(f"[페이지 {page_num}, 상품 {i+1}] 처리 중 오류 발생: {e}")
                break

        page_num += 1

    return data

# 엑셀 저장
def save_to_excel(data, file_path):
    wb = Workbook()
    ws = wb.active
    ws.append(["상품명", "가격", "재고", "실재고"])  # 헤더
    for row in data:
        ws.append(row)
    wb.save(file_path)

# 메인 실행
if __name__ == "__main__":
    options = webdriver.ChromeOptions()
    options.add_argument("--disable-gpu")
    options.add_argument("--no-sandbox")
    options.add_argument("--disable-popup-blocking")

    driver = webdriver.Chrome(options=options)

    try:
        while True:
            driver.get(_LEGACY_SHOP_BASE_URL + '/shop/admin/login/adm_login_cert.php')
            show_login_popup()

            driver.get(_LEGACY_SHOP_BASE_URL + '/shop/admin/goods/adm_goods_list.php?')
            show_search_popup()

            product_data = collect_product_data(driver)

            # 저장할 폴더 묻기
            save_dir = ask_save_directory()
            if not save_dir:  # 사용자가 취소 누르면 기본 경로 또는 종료 처리 가능
                print("저장할 폴더를 선택하지 않아 기본 경로에 저장합니다.")
                save_dir = str(output_dir('exports'))

            os.makedirs(save_dir, exist_ok=True)

            timestamp = time.strftime("%Y%m%d_%H%M%S")
            file_name = f"상품정보_{timestamp}.xlsx"
            file_path = os.path.join(save_dir, file_name)
            save_to_excel(product_data, file_path)

            print(f"[완료] 엑셀 저장 위치: {file_path}")

            if not ask_continue():
                break

    except Exception as e:
        print("오류 발생:", e)
        traceback.print_exc()
    finally:
        driver.quit()
