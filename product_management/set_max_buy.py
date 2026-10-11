
from pathlib import Path as _ConfigPath
import sys as _config_sys
_config_sys.path.insert(0, str(_ConfigPath(__file__).resolve().parents[1]))
from commerce_config import configured_url, output_dir
_GODOMALL_BASE_URL = configured_url('godomall_base_url')

import time
from urllib.parse import urlsplit, urlunsplit, parse_qsl, urlencode

import tkinter as tk
from tkinter import messagebox

from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait, Select
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import TimeoutException, NoSuchElementException
from selenium.webdriver.common.keys import Keys


BASE_LIST_URL = _GODOMALL_BASE_URL + '/goods/goods_list.php?detailSearch=&delFl=n&scmFl=all&key=goodsNm&keyword=&searchDateFl=regDt&searchDate%5B%5D=&searchDate%5B%5D=&searchPeriod=-1&cateGoods%5B%5D=&cateGoods%5B%5D=&cateGoods%5B%5D=&cateGoods%5B%5D=&displayTheme%5B%5D=&displayTheme%5B%5D=&goodsPrice%5B%5D=&goodsPrice%5B%5D=&mileage%5B%5D=&mileage%5B%5D=&mileageFl=&stock%5B%5D=&stock%5B%5D=&optionFl=&optionTextFl=&addGoodsFl=&goodsDisplayFl=&goodsSellFl=&goodsDisplayMobileFl=&goodsSellMobileFl=&stockFl=&soldOut=&goodsIconCdPeriod=&goodsIconCd=&goodsDeliveryFl=&goodsDeliveryFixFl%5B%5D=all&event_text=&eventThemeSno=&sort=g.goodsNo+desc&pageNum=500&searchFl=y&applyPath=%2Fgoods%2Fgoods_list.php'

ROW_LINK_XPATH = "/html/body/div[1]/div/div[2]/div[2]/div/form[2]/div[2]/table/tbody/tr[{i}]/td[5]/div[1]/a"

FIXED_ORDER_SELECT = (By.ID, "fixedOrderCnt")
MIN_ORDER_INPUT = (By.NAME, "minOrderCnt")
MAX_ORDER_INPUT = (By.NAME, "maxOrderCnt")
SAVE_BTN_XPATH = "/html/body/div[1]/div/form/div[86]/div[3]/input"
RADIO_ENABLE_XPATH = "/html/body/div[1]/div/form/div[29]/table/tbody/tr[5]/td/div[1]/label[2]/input"  # 제한 사용 라디오
MIN_INPUT_XPATH    = "/html/body/div[1]/div/form/div[29]/table/tbody/tr[5]/td/div[1]/label[3]/input"  # minOrderCnt
MAX_INPUT_XPATH    = "/html/body/div[1]/div/form/div[29]/table/tbody/tr[5]/td/div[1]/label[4]/input"  # maxOrderCnt
DESC_SAME_LABEL_XPATH = "/html/body/div[1]/div/form/div[50]/div/div[2]/ul/li[3]/label"
DESC_SAME_INPUT_XPATH = "/html/body/div[1]/div/form/div[50]/div/div[2]/ul/li[3]/label/input"

# =========================
# UI: 종료 확인 팝업
# =========================
def ask_stop_no_more_goods(consecutive_empty_pages: int, page: int) -> bool:
    """
    True면 종료, False면 계속
    """
    root = tk.Tk()
    root.withdraw()  # 메인 창 숨김
    root.attributes("-topmost", True)

    msg = (
        f"상품이 없는 페이지가 연속 {consecutive_empty_pages}번 확인되었습니다.\n"
        f"(현재 페이지: {page})\n\n"
        "더 이상 상품이 없을 수 있습니다.\n"
        "종료하시겠습니까?"
    )
    stop = messagebox.askyesno("확인", msg)
    root.destroy()
    return stop


def make_driver():
    options = webdriver.ChromeOptions()
    options.add_argument("--start-maximized")

    # (권장) 로그인 세션 유지용 크롬 프로필

    return webdriver.Chrome(options=options)

def show_login_required_popup():
    root = tk.Tk()
    root.withdraw()
    root.attributes("-topmost", True)

    messagebox.showinfo(
        "로그인 필요",
        "관리자 페이지 로그인이 필요합니다.\n\n브라우저에서 로그인 후 확인을 눌러주세요."
    )
    root.destroy()

def build_page_url(base_url: str, page: int) -> str:
    parts = urlsplit(base_url)
    q = parse_qsl(parts.query, keep_blank_values=True)
    q = [(k, v) for (k, v) in q if k.lower() != "page"]

    if page >= 2:
        q.insert(0, ("page", str(page)))

    new_query = urlencode(q, doseq=True)
    return urlunsplit((parts.scheme, parts.netloc, parts.path, new_query, parts.fragment))


def wait_until_list_ready(driver, timeout=60) -> bool:
    wait = WebDriverWait(driver, timeout)
    try:
        wait.until(EC.presence_of_element_located(
            (By.XPATH, "/html/body/div[1]/div/div[2]/div[2]/div/form[2]/div[2]/table/tbody")
        ))
        return True
    except TimeoutException:
        return False


def list_has_first_row(driver) -> bool:
    try:
        driver.find_element(By.XPATH, ROW_LINK_XPATH.format(i=1))
        return True
    except NoSuchElementException:
        return False


def open_popup_by_row(driver, row_i: int) -> str:
    link = driver.find_element(By.XPATH, ROW_LINK_XPATH.format(i=row_i))
    before = set(driver.window_handles)
    link.click()

    WebDriverWait(driver, 10).until(lambda d: len(d.window_handles) > len(before))
    after = set(driver.window_handles)
    return list(after - before)[0]


# =========================
# 로딩 대기(3~4초 감안)
# =========================
def wait_goods_edit_loaded(driver, timeout=20):
    """
    상품 수정 화면에서 필수 요소가 뜰 때까지 기다립니다.
    3~4초 정도 걸려도 여기서 흡수됩니다.
    """
    wait = WebDriverWait(driver, timeout)
    wait.until(EC.presence_of_element_located(FIXED_ORDER_SELECT))
    wait.until(EC.presence_of_element_located(MIN_ORDER_INPUT))
    wait.until(EC.presence_of_element_located(MAX_ORDER_INPUT))
    # DOM은 떴는데 값 세팅이 늦는 경우가 있어 짧게 보정
    time.sleep(0.7)


def _clear_and_type(el, text: str):
    """
    GodoMall 입력칸은 clear()가 가끔 덜 지워질 때가 있어
    Ctrl+A → Backspace 방식으로 확실히 비웁니다.
    """
    el.click()
    el.send_keys(Keys.CONTROL, "a")
    el.send_keys(Keys.BACKSPACE)
    el.send_keys(text)

def ensure_order_limit_inputs_ready(driver):
    """
    1) maxOrderChk 라디오가 'n(제한 안함)' 체크 상태면 '제한 사용' 라디오(label[2]/input) 클릭
    2) min/max input이 enabled 될 때까지 대기
    """
    wait = WebDriverWait(driver, 15)

    # n 라디오가 체크인지 확인(있으면)
    try:
        radio_n = driver.find_element(By.CSS_SELECTOR, "input[type='radio'][name='maxOrderChk'][value='n']")
        if radio_n.is_selected():
            # 제한 사용 라디오 클릭(이 클릭이 disabled_switch를 실행)
            driver.find_element(By.XPATH, RADIO_ENABLE_XPATH).click()
    except NoSuchElementException:
        # 라디오 구조가 다르면, 아래에서 그냥 입력 enable 대기 실패로 잡히게 두는 편이 안전합니다.
        pass

    # min/max input이 disabled가 풀릴 때까지 대기
    wait.until(lambda d: d.find_element(By.NAME, "minOrderCnt").get_attribute("disabled") is None)
    wait.until(lambda d: d.find_element(By.NAME, "maxOrderCnt").get_attribute("disabled") is None)

def set_min_max_by_label_xpaths(driver, min_val="1", max_val="5"):
    """
    사용자께서 지정하신 label[3], label[4] input을 직접 클릭해서 값 입력
    """
    min_el = driver.find_element(By.XPATH, MIN_INPUT_XPATH)
    max_el = driver.find_element(By.XPATH, MAX_INPUT_XPATH)

    cur_min = (min_el.get_attribute("value") or "").strip()
    if cur_min != min_val:
        _clear_and_type(min_el, min_val)

    cur_max = (max_el.get_attribute("value") or "").strip()
    if cur_max != max_val:
        _clear_and_type(max_el, max_val)

def set_input_value(driver, locator, desired: str):
    el = driver.find_element(*locator)
    cur = (el.get_attribute("value") or "").strip()
    if cur == desired:
        return
    el.clear()
    el.send_keys(desired)


def ensure_desc_same_checked(driver):
    cb = driver.find_element(By.XPATH, DESC_SAME_INPUT_XPATH)

    if not cb.is_selected():
        # input 대신 label 클릭 (이벤트 안정성 ↑)
        driver.find_element(By.XPATH, DESC_SAME_LABEL_XPATH).click()

        # 상태 반영 대기 (안정성)
        WebDriverWait(driver, 5).until(
            lambda d: d.find_element(By.XPATH, DESC_SAME_INPUT_XPATH).is_selected()
        )


def click_save_and_handle_alert(driver):
    driver.find_element(By.XPATH, SAVE_BTN_XPATH).click()

    # 저장 알럿 처리
    try:
        WebDriverWait(driver, 3).until(EC.alert_is_present())
        driver.switch_to.alert.accept()
    except TimeoutException:
        pass

    # 저장 처리/리로드 여유
    time.sleep(1.0)


def process_one_goods_popup(driver):
    wait_goods_edit_loaded(driver, timeout=20)  # 기존 로딩 대기

    # ✅ 라디오 클릭으로 제한 입력란 활성화
    ensure_order_limit_inputs_ready(driver)

    # ✅ label[3]/label[4] input을 클릭해서 1/5 입력
    set_min_max_by_label_xpaths(driver, "1", "5")

    # 상세정보 동일 체크
    ensure_desc_same_checked(driver)

    # 저장
    click_save_and_handle_alert(driver)


def process_one_page(driver) -> int:
    main_handle = driver.current_window_handle
    processed = 0

    i = 1
    while True:
        driver.switch_to.window(main_handle)

        try:
            driver.find_element(By.XPATH, ROW_LINK_XPATH.format(i=i))
        except NoSuchElementException:
            break

        try:
            new_handle = open_popup_by_row(driver, i)
            driver.switch_to.window(new_handle)

            process_one_goods_popup(driver)

            driver.close()
            driver.switch_to.window(main_handle)

            processed += 1
            i += 2

        except Exception as e:
            print(f"[ERROR] i={i} 처리 실패: {e}")

            try:
                if driver.current_window_handle != main_handle:
                    driver.close()
            except Exception:
                pass

            driver.switch_to.window(main_handle)
            i += 2

    return processed


def run():
    driver = make_driver()
    driver.get(BASE_LIST_URL)

    show_login_required_popup()
    
    # 최초 로그인 대기(수동)
    if not wait_until_list_ready(driver, timeout=300):
        print("로그인/목록 로딩 확인 실패(300초). 종료합니다.")
        driver.quit()
        return

    page = 1
    consecutive_empty_pages = 0

    while True:
        page_url = build_page_url(BASE_LIST_URL, page)
        driver.get(page_url)

        if not wait_until_list_ready(driver, timeout=60):
            print(f"[STOP] page={page} 목록 로딩 실패")
            break

        # 페이지에 상품이 없으면(첫 행 링크 없음)
        if not list_has_first_row(driver):
            consecutive_empty_pages += 1
            print(f"[EMPTY] page={page} 연속 빈페이지={consecutive_empty_pages}")

            # ✅ 연속 3번이면 사용자에게 종료 여부 확인
            if consecutive_empty_pages >= 3:
                stop = ask_stop_no_more_goods(consecutive_empty_pages, page)
                if stop:
                    print("[DONE] 사용자가 종료를 선택했습니다.")
                    break
                else:
                    # 계속 진행(다음 페이지로)
                    consecutive_empty_pages = 0

            page += 1
            continue

        # 상품이 있으면 빈 페이지 연속 카운트 리셋
        consecutive_empty_pages = 0

        processed = process_one_page(driver)
        print(f"[PAGE] page={page} processed={processed}")

        # 안전장치: 상품이 있다고 했는데 처리 0이면 이상 케이스
        if processed == 0:
            print(f"[WARN] page={page} 에서 processed=0 (DOM 구조 변경 가능). 종료합니다.")
            break

        page += 1

    driver.quit()


if __name__ == "__main__":
    run()
