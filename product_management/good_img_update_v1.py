
from pathlib import Path as _ConfigPath
import sys as _config_sys
_config_sys.path.insert(0, str(_ConfigPath(__file__).resolve().parents[1]))
from commerce_config import configured_url, output_dir
_GODOMALL_BASE_URL = configured_url('godomall_base_url')

import time
from urllib.parse import urlparse, parse_qs, urlencode
from selenium import webdriver
from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
from selenium.common.exceptions import NoSuchElementException
from selenium.webdriver.support import expected_conditions as EC
from tkinter import Tk, messagebox, Button  # 팝업 창 띄우기 위한 라이브러리

import time
import threading  # Tkinter와 Selenium 병렬 실행을 위해 사용



# 전역 종료 플래그
stop_flag = False

def stop_program():
    global stop_flag
    stop_flag = True
    messagebox.showinfo("알림", "작업이 중간에 종료됩니다.")

def create_control_window():
    root = Tk()
    root.title("작업 제어")
    root.geometry("250x100")

    btn = Button(root, text="▶ 작업 중간 종료", command=stop_program, fg="red")
    btn.pack(pady=20)

    root.mainloop()

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


def process_pages(driver, wait, save_dir=None):
    page_num = 1  # 첫 번째 페이지부터 시작

    while True:
        # 현재 URL 가져오기
        current_url = driver.current_url
        parsed_url = urlparse(current_url)
        base_url = f"{parsed_url.scheme}://{parsed_url.netloc}{parsed_url.path}"

        # 기존 쿼리 파라미터 정리
        query_params = parse_qs(parsed_url.query)
        query_params = {k: v[0] for k, v in query_params.items()}  # 리스트 → 값

        # page 값 덮어쓰기
        query_params['page'] = page_num
        query_string = urlencode(query_params)

        page_url = f"{base_url}?{query_string}"
        print(f"페이지 {page_num} 로딩: {page_url}")
        driver.get(page_url)
        time.sleep(2)

        # tbody 확인 (상품이 없으면 중단)
        try:
            tbody_xpath = "/html/body/div[1]/div/div[2]/div[2]/div/form[2]/div[2]/table/tbody"
            tbody = wait.until(EC.presence_of_element_located((By.XPATH, tbody_xpath)))
            rows = tbody.find_elements(By.TAG_NAME, "tr")
            total_rows = len(rows)
            if total_rows == 0:
                print("상품이 없는 페이지 → 종료")
                break
        except:
            print("상품이 없는 페이지 → 종료")
            break

        # 페이지 내 row 처리 (1,3,5,... 만)
        for i in range((total_rows + 1) // 2):
            row_index = 1 + (i * 2)
            if row_index > total_rows:
                break

            try:
                origin = driver.current_window_handle

                link_xpath = f"{tbody_xpath}/tr[{row_index}]/td[5]/div[1]/a"
                link = wait.until(EC.element_to_be_clickable((By.XPATH, link_xpath)))
                driver.execute_script("arguments[0].scrollIntoView({block:'center'});", link)
                try:
                    link.click()
                except:
                    driver.execute_script("arguments[0].click();", link)

                wait.until(EC.number_of_windows_to_be(2))
                new_handle = [h for h in driver.window_handles if h != origin][0]
                driver.switch_to.window(new_handle)
                time.sleep(7)

                radio_xpath = '//*[@id="depth-toggle-layer-relation"]/table/tbody/tr[1]/td/label[2]/input'
                radio = wait.until(EC.element_to_be_clickable((By.XPATH, radio_xpath)))
                driver.execute_script("arguments[0].scrollIntoView({block:'center'});", radio)
                try:
                    radio.click()
                except:
                    driver.execute_script("arguments[0].click();", radio)
                time.sleep(0.2)

                # 버튼 클릭
                btn_xpath = "/html/body/div[1]/div/form/div[86]/div[3]/input"
                btn = wait.until(EC.element_to_be_clickable((By.XPATH, btn_xpath)))
                btn.click()

                time.sleep(5)

                driver.close()
                driver.switch_to.window(origin)
                wait.until(EC.number_of_windows_to_be(1))

                print(f"{page_num}페이지 - {row_index}행 처리 완료")

            except Exception as e:
                print(f"{page_num}페이지 - {row_index}행 오류: {e}")
                continue

        # 페이지 증가
        page_num += 1
        


if __name__ == "__main__":
    # ===== 브라우저 설정 =====
    options = webdriver.ChromeOptions()
    options.add_argument("--disable-gpu")
    options.add_argument("--no-sandbox")
    options.add_argument("--remote-debugging-port=9222")  # 요청: 꼭 추가
    options.add_argument("--disable-popup-blocking")
    # options.add_argument("profile-directory=Default")

    driver = webdriver.Chrome(options=options)
    wait = WebDriverWait(driver, 20)

    # ===== 상시 종료 버튼 GUI 실행(데몬 스레드) =====
    # create_control_window() 내부에서 stop_flag를 True로 바꾸는 버튼이 있어야 함
    threading.Thread(target=create_control_window, daemon=True).start()

    user_wants_quit = None  # 최종 종료 의사 기록용

    try:
        # 1) 로그인 페이지 → 로그인 완료 팝업 대기
        login_url = _GODOMALL_BASE_URL + '/base/login.php'
        driver.get(login_url)
        show_login_popup()  # 사용자가 닫으면 다음 단계 진행

        # 2) 상품 리스트 페이지 → 검색/필터 설정 팝업 대기
        driver.get(_GODOMALL_BASE_URL + '/goods/goods_list.php')
        show_search_popup()

        # 3) 페이지 단위 순회 + 각 페이지 내 row(1,3,5,...) 처리
        #    process_pages는 내부에서 stop_flag를 수시로 확인하도록 구현되어 있어야 함
        process_pages(driver=driver, wait=wait, save_dir=None)

        # 4) 완료 처리 (중간 종료가 아닌 경우만 종료 여부 질문)
        if not stop_flag:
            root = Tk()
            root.withdraw()
            result = messagebox.askyesno("작업 완료", "모든 작업이 완료되었습니다.\n드라이버를 종료하시겠습니까?")
            user_wants_quit = bool(result)
            if user_wants_quit:
                driver.quit()
                messagebox.showinfo("종료됨", "드라이버가 종료되었습니다.")
            else:
                messagebox.showinfo("유지됨", "드라이버는 계속 실행 상태로 유지됩니다.")
            root.destroy()
        else:
            # 중간 종료 버튼으로 멈춘 경우
            root = Tk()
            root.withdraw()
            messagebox.showinfo("중단됨", "사용자 요청으로 작업이 중단되었습니다.\n브라우저는 그대로 유지됩니다.")
            root.destroy()

    except Exception as e:
        print(f"[치명오류] {e}")
        # 필요 시 여기서 에러 팝업만 안내하고 드라이버는 그대로 유지할 수 있음
        try:
            root = Tk()
            root.withdraw()
            messagebox.showerror("오류", f"예기치 못한 오류가 발생했습니다.\n{e}")
            root.destroy()
        except:
            pass

    # 중요: finally에서 driver.quit()으로 강제 종료하지 않음
    # 사용자가 '예'를 선택했을 때만 위에서 종료
