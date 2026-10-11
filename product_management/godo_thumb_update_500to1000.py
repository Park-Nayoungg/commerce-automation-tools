
from pathlib import Path as _ConfigPath
import sys as _config_sys
_config_sys.path.insert(0, str(_ConfigPath(__file__).resolve().parents[1]))
from commerce_config import configured_url, output_dir
_GODOMALL_BASE_URL = configured_url('godomall_base_url')

import os
import time
import threading
import tempfile
from io import BytesIO
from urllib.parse import urlsplit, urlunsplit, parse_qsl, urlencode, urlparse
from urllib.request import urlopen, Request
from tkinter import Tk, messagebox, Button

from PIL import Image
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import TimeoutException, NoSuchElementException


stop_flag = False
search_url = None

LIST_TBODY_XPATH = "/html/body/div[1]/div/div[2]/div[2]/div/form[2]/div[2]/table/tbody"
ROW_LINK_XPATH = "/html/body/div[1]/div/div[2]/div[2]/div/form[2]/div[2]/table/tbody/tr[{i}]/td[5]/div[1]/a"
SAVE_BTN_XPATH = "/html/body/div[1]/div/form/div[86]/div[3]/input"


def stop_program():
    global stop_flag
    stop_flag = True

    root = Tk()
    root.withdraw()
    root.attributes("-topmost", True)
    messagebox.showinfo("알림", "작업이 중간에 종료됩니다.")
    root.destroy()


def create_control_window():
    root = Tk()
    root.title("작업 제어")
    root.geometry("250x100")
    root.attributes("-topmost", True)
    btn = Button(root, text="▶ 작업 중간 종료", command=stop_program, fg="red")
    btn.pack(pady=20)
    root.mainloop()


def show_login_popup():
    def on_confirm():
        print("사용자가 로그인 완료를 확인했습니다.")
        root.destroy()

    root = Tk()
    root.title("로그인 요청")
    root.geometry("300x100")
    root.resizable(False, False)
    root.attributes("-topmost", True)
    Button(root, text="로그인을 완료한 후 이 버튼을 눌러주세요.", command=on_confirm).pack(pady=30)
    root.mainloop()


def show_search_popup():
    def on_confirm():
        global search_url
        search_url = driver.current_url
        print(f"검색 URL 저장: {search_url}")
        root.destroy()

    root = Tk()
    root.title("상품검색 요청")
    root.geometry("300x100")
    root.resizable(False, False)
    root.attributes("-topmost", True)
    Button(root, text="상품검색을 완료한 후 이 버튼을 눌러주세요.", command=on_confirm).pack(pady=30)
    root.mainloop()


def ask_research_popup():
    root = Tk()
    root.withdraw()
    root.attributes("-topmost", True)
    result = messagebox.askyesno("작업 완료", "더 이상 수정할 상품이 없습니다.\n다시 상품검색을 하시겠습니까?")
    root.destroy()
    return bool(result)


def ask_driver_quit_popup():
    root = Tk()
    root.withdraw()
    root.attributes("-topmost", True)
    result = messagebox.askyesno("종료 확인", "드라이버를 종료하시겠습니까?")
    root.destroy()
    return bool(result)


def safe_click(driver, element):
    driver.execute_script("arguments[0].scrollIntoView({block:'center'});", element)
    time.sleep(0.2)
    try:
        element.click()
    except:
        driver.execute_script("arguments[0].click();", element)


def build_page_url(base_url: str, page_num: int) -> str:
    parts = urlsplit(base_url)
    query_items = parse_qsl(parts.query, keep_blank_values=True)
    query_items = [(k, v) for (k, v) in query_items if k.lower() != "page"]

    if page_num >= 2:
        query_items.insert(0, ("page", str(page_num)))

    new_query = urlencode(query_items, doseq=True)
    return urlunsplit((parts.scheme, parts.netloc, parts.path, new_query, parts.fragment))


def wait_until_list_ready(driver, timeout=60):
    wait = WebDriverWait(driver, timeout)
    try:
        wait.until(EC.presence_of_element_located((By.XPATH, LIST_TBODY_XPATH)))
        return True
    except TimeoutException:
        return False


def list_has_first_row(driver) -> bool:
    try:
        driver.find_element(By.XPATH, ROW_LINK_XPATH.format(i=1))
        return True
    except NoSuchElementException:
        return False


def download_image_bytes(url):
    req = Request(url, headers={"User-Agent": "Mozilla/5.0"})
    with urlopen(req, timeout=30) as response:
        return response.read()


def resize_to_square_keep_ratio(image_bytes, save_path, size=1000, bg_color=(255, 255, 255)):
    img = Image.open(BytesIO(image_bytes)).convert("RGB")

    original_w, original_h = img.size
    if original_w <= 0 or original_h <= 0:
        raise Exception("원본 이미지 크기가 올바르지 않습니다.")

    scale = min(size / original_w, size / original_h)
    resized_w = int(original_w * scale)
    resized_h = int(original_h * scale)

    resized = img.resize((resized_w, resized_h), Image.LANCZOS)

    canvas = Image.new("RGB", (size, size), bg_color)
    offset_x = (size - resized_w) // 2
    offset_y = (size - resized_h) // 2
    canvas.paste(resized, (offset_x, offset_y))
    canvas.save(save_path, format="JPEG", quality=95)


def make_resized_file_from_url(image_url):
    image_bytes = download_image_bytes(image_url)
    temp_dir = tempfile.mkdtemp(prefix="godomall_resize_")
    filename = os.path.basename(urlparse(image_url).path)
    name, _ = os.path.splitext(filename)
    save_path = os.path.abspath(os.path.join(temp_dir, f"{name}_1000.jpg"))
    resize_to_square_keep_ratio(image_bytes, save_path, size=1000)
    return save_path


def force_file_input_visible(driver, file_input):
    driver.execute_script("""
        arguments[0].style.display = 'block';
        arguments[0].style.visibility = 'visible';
        arguments[0].style.opacity = '1';
        arguments[0].style.position = 'fixed';
        arguments[0].style.left = '0';
        arguments[0].style.top = '0';
        arguments[0].style.width = '1px';
        arguments[0].style.height = '1px';
        arguments[0].style.clip = 'auto';
        arguments[0].removeAttribute('disabled');
    """, file_input)


def upload_file_to_slot(driver, slot_div, resized_file_path):
    file_inputs = slot_div.find_elements(By.XPATH, ".//input[@type='file']")
    if not file_inputs:
        raise Exception("업로드용 file input을 찾지 못했습니다.")

    file_input = file_inputs[0]
    force_file_input_visible(driver, file_input)
    file_input.send_keys(resized_file_path)
    time.sleep(0.5)


def get_image_size_from_url(image_url):
    image_bytes = download_image_bytes(image_url)
    img = Image.open(BytesIO(image_bytes))
    return img.size


def get_original_image_url_from_preview(slot_div):
    preview_imgs = slot_div.find_elements(By.XPATH, ".//img[@src]")
    if not preview_imgs:
        raise Exception("미리보기 이미지(img src)를 찾지 못했습니다.")

    for img in preview_imgs:
        src = (img.get_attribute("src") or "").strip()
        if src and ("godomall-storage" in src or src.startswith("http")):
            return src

    raise Exception("다운로드 가능한 이미지 src를 찾지 못했습니다.")


def is_first_magnify_image_already_1000(driver):
    try:
        section = driver.find_element(By.ID, "imageMagnify")
        slot_divs = section.find_elements(By.XPATH, "./div[starts-with(@id, 'imageMagnify')]")
        if not slot_divs:
            print("[검사] 확대 이미지 슬롯이 없습니다.")
            return False

        first_slot = slot_divs[0]
        original_url = get_original_image_url_from_preview(first_slot)
        width, height = get_image_size_from_url(original_url)

        print(f"[검사] 첫 확대 이미지 크기: {width}x{height}")
        return width == 1000 and height == 1000
    except Exception as e:
        print(f"[검사 오류] 첫 확대 이미지 크기 확인 실패: {e}")
        return False


def process_image_section(driver, section_id):
    try:
        section = driver.find_element(By.ID, section_id)
    except:
        print(f"[건너뜀] 섹션 없음: {section_id}")
        return 0

    slot_divs = section.find_elements(By.XPATH, f"./div[starts-with(@id, '{section_id}')]")
    processed_count = 0

    for idx, slot_div in enumerate(slot_divs, start=1):
        if stop_flag:
            break

        try:
            original_url = get_original_image_url_from_preview(slot_div)
            print(f"[원본 URL] {section_id} {idx}: {original_url}")

            resized_file_path = make_resized_file_from_url(original_url)
            upload_file_to_slot(driver, slot_div, resized_file_path)

            processed_count += 1
            print(f"[완료] {section_id} {idx}번 슬롯 업로드 완료")
        except Exception as e:
            print(f"[오류] {section_id} {idx}번 슬롯 처리 실패: {e}")
            continue

    return processed_count


def click_save_button(driver, wait):
    save_button = wait.until(EC.element_to_be_clickable((By.XPATH, SAVE_BTN_XPATH)))
    safe_click(driver, save_button)
    time.sleep(2)

    try:
        alert = driver.switch_to.alert
        print(f"[알림창] {alert.text}")
        alert.accept()
        time.sleep(1)
    except:
        pass


def process_product_detail(driver, wait):
    if is_first_magnify_image_already_1000(driver):
        print("[건너뜀] 첫 확대 이미지가 이미 1000x1000이므로 이 상품은 전체 작업을 생략합니다.")
        return "SKIPPED"

    total_uploaded = 0
    total_uploaded += process_image_section(driver, "imageMagnify")

    if total_uploaded == 0:
        print("[안내] 처리된 이미지가 없어서 저장하지 않습니다.")
        return "NO_CHANGE"

    click_save_button(driver, wait)
    print(f"[저장 완료] 총 {total_uploaded}개 확대 이미지 업로드")
    return "UPDATED"


def process_pages(driver, wait):
    global stop_flag, search_url

    if not search_url:
        raise Exception("검색 URL이 저장되지 않았습니다. 검색 후 팝업 버튼을 눌러주세요.")

    page_num = 4
    consecutive_empty_pages = 0

    while True:
        if stop_flag:
            print("사용자 중단 요청으로 종료합니다.")
            return "STOPPED"

        page_url = build_page_url(search_url, page_num)
        print(f"페이지 {page_num} 로딩: {page_url}")
        driver.get(page_url)

        if not wait_until_list_ready(driver, timeout=60):
            print(f"[목록 로딩 실패] page={page_num}")
            return "NO_MORE_PRODUCTS"

        if not list_has_first_row(driver):
            consecutive_empty_pages += 1
            print(f"[빈 페이지] page={page_num}, 연속 빈페이지={consecutive_empty_pages}")

            if consecutive_empty_pages >= 2:
                print("더 이상 수정할 상품이 없는 것으로 판단합니다.")
                return "NO_MORE_PRODUCTS"

            page_num += 1
            continue

        consecutive_empty_pages = 0

        try:
            tbody = wait.until(EC.presence_of_element_located((By.XPATH, LIST_TBODY_XPATH)))
            rows = tbody.find_elements(By.TAG_NAME, "tr")
            total_rows = len(rows)
        except Exception as e:
            print(f"[목록 분석 실패] page={page_num}, error={e}")
            return "NO_MORE_PRODUCTS"

        page_processed_count = 0

        for i in range((total_rows + 1) // 2):
            if stop_flag:
                print("사용자 중단 요청으로 행 처리 종료")
                return "STOPPED"

            row_index = 1 + (i * 2)
            if row_index > total_rows:
                break

            origin = driver.current_window_handle

            try:
                link_xpath = ROW_LINK_XPATH.format(i=row_index)
                link = wait.until(EC.element_to_be_clickable((By.XPATH, link_xpath)))
                safe_click(driver, link)

                wait.until(EC.number_of_windows_to_be(2))
                new_handle = [h for h in driver.window_handles if h != origin][0]
                driver.switch_to.window(new_handle)
                time.sleep(5)

                detail_result = process_product_detail(driver, wait)
                print(f"[상품 결과] {detail_result}")

                driver.close()
                driver.switch_to.window(origin)
                wait.until(EC.number_of_windows_to_be(1))

                page_processed_count += 1
                print(f"{page_num}페이지 - {row_index}행 처리 완료 ({detail_result})")

            except Exception as e:
                print(f"{page_num}페이지 - {row_index}행 오류: {e}")
                try:
                    if len(driver.window_handles) > 1 and driver.current_window_handle != origin:
                        driver.close()
                except:
                    pass

                try:
                    driver.switch_to.window(origin)
                except:
                    pass
                continue

        if page_processed_count == 0:
            print("현재 페이지에서 처리 가능한 상품이 없습니다.")
            return "NO_MORE_PRODUCTS"

        page_num += 1


if __name__ == "__main__":
    options = webdriver.ChromeOptions()
    options.add_argument("--disable-gpu")
    options.add_argument("--no-sandbox")
    options.add_argument("--remote-debugging-port=9222")
    options.add_argument("--disable-popup-blocking")

    driver = webdriver.Chrome(options=options)
    wait = WebDriverWait(driver, 20)

    threading.Thread(target=create_control_window, daemon=True).start()

    try:
        login_url = _GODOMALL_BASE_URL + '/base/login.php'
        driver.get(login_url)
        show_login_popup()

        while True:
            if stop_flag:
                root = Tk()
                root.withdraw()
                root.attributes("-topmost", True)
                messagebox.showinfo("중단됨", "사용자 요청으로 작업이 중단되었습니다.\n브라우저는 그대로 유지됩니다.")
                root.destroy()
                break

            driver.get(_GODOMALL_BASE_URL + '/goods/goods_list.php')
            show_search_popup()

            result = process_pages(driver=driver, wait=wait)

            if result == "STOPPED":
                root = Tk()
                root.withdraw()
                root.attributes("-topmost", True)
                messagebox.showinfo("중단됨", "사용자 요청으로 작업이 중단되었습니다.\n브라우저는 그대로 유지됩니다.")
                root.destroy()
                break

            if result == "NO_MORE_PRODUCTS":
                retry_search = ask_research_popup()

                if retry_search:
                    continue

                quit_driver = ask_driver_quit_popup()
                if quit_driver:
                    driver.quit()
                    root = Tk()
                    root.withdraw()
                    root.attributes("-topmost", True)
                    messagebox.showinfo("종료됨", "드라이버가 종료되었습니다.")
                    root.destroy()
                else:
                    root = Tk()
                    root.withdraw()
                    root.attributes("-topmost", True)
                    messagebox.showinfo("유지됨", "드라이버는 계속 실행 상태로 유지됩니다.")
                    root.destroy()
                break

    except Exception as e:
        print(f"[치명오류] {e}")
        try:
            root = Tk()
            root.withdraw()
            root.attributes("-topmost", True)
            messagebox.showerror("오류", f"예기치 못한 오류가 발생했습니다.\n{e}")
            root.destroy()
        except:
            pass
