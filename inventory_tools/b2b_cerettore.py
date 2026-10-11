import re
import html
import time
import threading
import requests
import tkinter as tk
from tkinter import filedialog, messagebox
from bs4 import BeautifulSoup
from openpyxl import Workbook
from openpyxl.styles import Alignment
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.action_chains import ActionChains

BASE_URL = "http://demand.aferp.co.kr"
OUTPUT_FILE = "aferp_soldout_result.xlsx"

class CerrotorreApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Cerrotorre 품절 조사기")
        self.root.geometry("760x520")
        self.all_rows = []
        self.processed_codes = set()
        self.is_running = False
        
        self.driver = None
        self.output_path = tk.StringVar(value=OUTPUT_FILE)

        top_frame = tk.Frame(root)
        top_frame.pack(fill="x", padx=10, pady=10)

        tk.Label(top_frame, text="저장 파일").pack(side="left")
        tk.Entry(top_frame, textvariable=self.output_path, width=60).pack(side="left", padx=6)
        tk.Button(top_frame, text="저장 위치 선택", command=self.choose_save_path).pack(side="left")

        button_frame = tk.Frame(root)
        button_frame.pack(fill="x", padx=10, pady=5)

        guide_frame = tk.Frame(root)
        guide_frame.pack(fill="x", padx=10, pady=5)

        guide_text = (
            "사용 방법\n"
            "1. 엑셀파일을 저장할 경로를 설정하세요.\n"
            "2. '브라우저 열기' 버튼을 누르세요.\n"
            "3. 열린 브라우저에서 직접 로그인하세요.\n"
            "4. 주문등록 페이지로 이동한 뒤 원하는 상품을 검색하세요.\n"
            "5. 검색 결과가 보이는 상태에서 '조사 시작' 버튼을 누르세요.\n"
            "6. 프로그램이 현재 페이지부터 마지막 페이지까지 상품 옵션 품절 여부를 조사합니다.\n"
            "7. 종료 버튼을 누르면 현재까지 조사한 결과를 저장하고 종료합니다."
        )

        self.guide_label = tk.Label(
            guide_frame,
            text=guide_text,
            justify="left",
            anchor="w",
            bg="#f4f4f4",
            relief="solid",
            bd=1,
            padx=8,
            pady=8
        )
        self.guide_label.pack(fill="x")

        self.open_button = tk.Button(button_frame, text="브라우저 열기", width=18, command=self.open_browser)
        self.open_button.pack(side="left", padx=4)

        self.start_button = tk.Button(button_frame, text="조사 시작", width=18, command=self.start_work)
        self.start_button.pack(side="left", padx=4)

        self.close_button = tk.Button(button_frame, text="종료", width=18, command=self.close_app)
        self.close_button.pack(side="left", padx=4)

        self.log_text = tk.Text(root, wrap="word")
        self.log_text.pack(fill="both", expand=True, padx=10, pady=10)

    def log(self, message):
        self.root.after(0, self._append_log, message)

    def _append_log(self, message):
        self.log_text.insert("end", f"{message}\n")
        self.log_text.see("end")

    def choose_save_path(self):
        path = filedialog.asksaveasfilename(
            title="엑셀 저장 위치 선택",
            defaultextension=".xlsx",
            filetypes=[("Excel files", "*.xlsx")]
        )
        if path:
            self.output_path.set(path)

    def create_driver(self):
        options = webdriver.ChromeOptions()
        options.add_experimental_option("detach", True)
        driver = webdriver.Chrome(options=options)
        driver.maximize_window()
        return driver

    def open_browser(self):
        try:
            if self.driver is not None:
                self.log("브라우저가 이미 열려 있습니다.")
                self.log("이미 열린 브라우저에서 로그인/검색을 계속 진행하신 뒤 조사 시작을 눌러주세요.")
                return

            self.driver = self.create_driver()
            self.driver.get(f"{BASE_URL}/cerrotorre")

            self.log("브라우저를 열었습니다.")
            self.log("다음 순서대로 진행해주세요.")
            self.log("1) 로그인합니다.")
            self.log("2) 주문등록 페이지로 이동합니다.")
            self.log("3) 원하는 상품을 검색합니다.")
            self.log("4) 검색 결과가 화면에 보이면 '조사 시작' 버튼을 누릅니다.")
        except Exception as e:
            messagebox.showerror("오류", f"브라우저 실행 실패\n{e}")

    def start_work(self):
        if self.driver is None:
            messagebox.showwarning("안내", "먼저 '브라우저 열기' 버튼을 눌러주세요.")
            return

        if self.is_running:
            messagebox.showwarning("안내", "이미 조사가 진행 중입니다.")
            return

        if not self.output_path.get().strip() or self.output_path.get().strip() == OUTPUT_FILE:
            path = filedialog.asksaveasfilename(
                title="엑셀 저장 위치 선택",
                defaultextension=".xlsx",
                filetypes=[("Excel files", "*.xlsx")]
            )
            if not path:
                messagebox.showwarning("안내", "저장 위치를 선택해야 조사를 시작할 수 있습니다.")
                return
            self.output_path.set(path)

        self.log("조사를 시작합니다.")
        self.log("현재 검색 결과 페이지의 상품부터 순서대로 조사합니다.")
        self.log("조사 중에는 브라우저를 임의로 닫지 않는 것을 권장합니다.")
        self.log(f"엑셀 저장 경로: {self.output_path.get()}")

        self.start_button.config(state="disabled")
        self.is_running = True
        threading.Thread(target=self.run_main_task, daemon=True).start()


    def parse_product_cards(self):
        products = []
        seen = set()
        cards = self.driver.find_elements(By.CSS_SELECTOR, "td.dxdvItem_MetropolisBlue")

        for card_idx, card in enumerate(cards):
            try:
                card_html = card.get_attribute("innerHTML")
            except Exception:
                continue

            m = re.search(r"OnShowSTYLEDetail\('([^']+)'\s*,\s*'([^']+)'\)", card_html)
            if not m:
                continue

            product_code = m.group(1).strip()
            deli_price = m.group(2).strip()
            key = (product_code, deli_price)

            if key in seen:
                continue
            seen.add(key)

            product_name = product_code
            try:
                name_el = card.find_element(By.CSS_SELECTOR, "a.dxeHyperlink_MetropolisBlue")
                product_name = (name_el.text or name_el.get_attribute("title") or product_code).strip()
            except Exception:
                pass

            products.append({
                "product_name": product_name,
                "product_code": product_code,
                "deli_price": deli_price,
                "card_index": card_idx,
            })

        return products

    def force_open_dropdown(self, button_id, popup_id, item_table_selector):
        btn = WebDriverWait(self.driver, 10).until(
            EC.presence_of_element_located((By.ID, button_id))
        )

        self.driver.execute_script("arguments[0].scrollIntoView({block:'center'});", btn)
        time.sleep(0.3)

        try:
            ActionChains(self.driver).move_to_element(btn).pause(0.2).click_and_hold(btn).pause(0.2).release(btn).perform()
        except Exception:
            try:
                self.driver.execute_script("""
                    var el = arguments[0];
                    var ev = new MouseEvent('mousedown', {bubbles:true, cancelable:true, view:window});
                    el.dispatchEvent(ev);
                """, btn)
            except Exception:
                self.driver.execute_script("arguments[0].click();", btn)

        WebDriverWait(self.driver, 10).until(
            lambda d: d.find_element(By.ID, popup_id).is_displayed()
        )

        WebDriverWait(self.driver, 10).until(
            lambda d: len(d.find_elements(By.CSS_SELECTOR, item_table_selector)) > 0
        )

        time.sleep(0.5)

    def open_color_dropdown(self):
        popup_id = "ctl00_RootContentHolder_ASPxFormLayout1_CD_COLOR_ComboBox_DDD_PW-1"
        item_selector = "#ctl00_RootContentHolder_ASPxFormLayout1_CD_COLOR_ComboBox_DDD_L_LBT tr.dxeListBoxItemRow_MetropolisBlue"
        button_id = "ctl00_RootContentHolder_ASPxFormLayout1_CD_COLOR_ComboBox_B-1"

        try:
            popup = self.driver.find_element(By.ID, popup_id)
            if popup.is_displayed():
                rows = self.driver.find_elements(By.CSS_SELECTOR, item_selector)
                if len(rows) > 0:
                    return
        except Exception:
            pass

        combo_btn = WebDriverWait(self.driver, 10).until(
            EC.presence_of_element_located((By.ID, button_id))
        )

        self.driver.execute_script("arguments[0].scrollIntoView({block:'center'});", combo_btn)
        time.sleep(0.2)

        opened = False

        try:
            ActionChains(self.driver).move_to_element(combo_btn).pause(0.1).click(combo_btn).perform()
            opened = True
        except Exception:
            pass

        if not opened:
            try:
                self.driver.execute_script("arguments[0].click();", combo_btn)
                opened = True
            except Exception:
                pass

        if not opened:
            try:
                self.driver.execute_script("""
                    ASPx.DDDropDown('ctl00_RootContentHolder_ASPxFormLayout1_CD_COLOR_ComboBox', null);
                """)
                opened = True
            except Exception:
                pass

        if not opened:
            raise RuntimeError("색상 드롭다운 열기 실패")

        WebDriverWait(self.driver, 10).until(
            lambda d: d.find_element(By.ID, popup_id).is_displayed()
        )
        WebDriverWait(self.driver, 10).until(
            lambda d: len(d.find_elements(By.CSS_SELECTOR, item_selector)) > 0
        )
        time.sleep(0.3)
        
    def get_color_options_from_dom(self):
        rows = self.driver.find_elements(
            By.CSS_SELECTOR,
            "#ctl00_RootContentHolder_ASPxFormLayout1_CD_COLOR_ComboBox_DDD_L_LBT tr.dxeListBoxItemRow_MetropolisBlue"
        )

        items = []
        for idx, row in enumerate(rows):
            try:
                td = row.find_element(By.CSS_SELECTOR, "td.dxeListBoxItem_MetropolisBlue")
                text = (td.text or "").replace("\xa0", " ").strip()
                if not text:
                    continue

                items.append({
                    "index": idx,
                    "color_name": text
                })
            except Exception:
                continue

        return items
        
    def wait_until_color_applied(self, expected_text, old_value):
        input_id = "ctl00_RootContentHolder_ASPxFormLayout1_CD_COLOR_ComboBox_I"

        def _normalized(text):
            return re.sub(r"\s+", " ", (text or "").replace("\xa0", " ")).strip()

        expected_norm = _normalized(expected_text)
        old_norm = _normalized(old_value)
        expected_head = expected_norm.split("(")[0].strip().lower()

        WebDriverWait(self.driver, 10).until(
            lambda d: _normalized(d.find_element(By.ID, input_id).get_attribute("value")) != old_norm
        )

        WebDriverWait(self.driver, 10).until(
            lambda d: expected_head in _normalized(
                d.find_element(By.ID, input_id).get_attribute("value")
            ).lower()
        )

        time.sleep(1.0)

    def select_color_by_index(self, color_index):
        self.log(f"[DEBUG] select_color_by_index 진입: index={color_index}")

        self.open_color_dropdown()
        self.log("[DEBUG] 색상 드롭다운 열림 확인 완료")

        input_id = "ctl00_RootContentHolder_ASPxFormLayout1_CD_COLOR_ComboBox_I"
        old_value = (
            self.driver.find_element(By.ID, input_id).get_attribute("value") or ""
        ).strip()
        self.log(f"[DEBUG] 기존 색상 input 값: {old_value}")

        rows = WebDriverWait(self.driver, 10).until(
            lambda d: d.find_elements(
                By.CSS_SELECTOR,
                "#ctl00_RootContentHolder_ASPxFormLayout1_CD_COLOR_ComboBox_DDD_L_LBT tr.dxeListBoxItemRow_MetropolisBlue"
            )
        )
        self.log(f"[DEBUG] 현재 드롭다운 행 수: {len(rows)}")

        if color_index >= len(rows):
            raise RuntimeError(f"색상 index 범위 초과: {color_index}")

        row = rows[color_index]
        td = row.find_element(By.CSS_SELECTOR, "td.dxeListBoxItem_MetropolisBlue")
        color_name = (td.text or "").replace("\xa0", " ").strip()

        self.log(f"[DEBUG] 색상 클릭 대상 index: {color_index}")
        self.log(f"[DEBUG] 색상 클릭 시도: {color_name}")

        self.driver.execute_script("arguments[0].scrollIntoView({block:'nearest'});", td)
        time.sleep(0.2)

        clicked = False

        try:
            ActionChains(self.driver).move_to_element(td).pause(0.1).click(td).perform()
            clicked = True
        except Exception:
            pass

        if not clicked:
            try:
                ActionChains(self.driver).move_to_element(row).pause(0.1).click(row).perform()
                clicked = True
            except Exception:
                pass

        if not clicked:
            try:
                self.driver.execute_script("""
                    var el = arguments[0];
                    el.dispatchEvent(new MouseEvent('mousedown', {bubbles:true, cancelable:true, view:window}));
                    el.dispatchEvent(new MouseEvent('mouseup', {bubbles:true, cancelable:true, view:window}));
                    el.dispatchEvent(new MouseEvent('click', {bubbles:true, cancelable:true, view:window}));
                """, td)
                clicked = True
            except Exception:
                pass

        if not clicked:
            raise RuntimeError(f"색상 클릭 실패: {color_name}")

        self.wait_until_color_applied(color_name, old_value)

        current_color_value = (
            self.driver.find_element(By.ID, input_id).get_attribute("value") or ""
        ).strip()

        self.log(f"[DEBUG] 색상 클릭 후 input 값: {current_color_value}")

        return color_name

    def open_size_dropdown(self):
        self.force_open_dropdown(
            "ctl00_RootContentHolder_ASPxFormLayout1_CD_SIZE_CallbackPanel_CD_SIZE_DropDownEdit_B-1",
            "ctl00_RootContentHolder_ASPxFormLayout1_CD_SIZE_CallbackPanel_CD_SIZE_DropDownEdit_DDD_PW-1",
            "#ctl00_RootContentHolder_ASPxFormLayout1_CD_SIZE_CallbackPanel_CD_SIZE_DropDownEdit_DDD_DDTC_CD_SIZE_ListBox_LBT td.dxeT"
        )

    def get_size_items_from_dom(self):
        rows = self.driver.find_elements(
            By.CSS_SELECTOR,
            "#ctl00_RootContentHolder_ASPxFormLayout1_CD_SIZE_CallbackPanel_CD_SIZE_DropDownEdit_DDD_DDTC_CD_SIZE_ListBox_LBT tr"
        )

        items = []
        for row in rows:
            try:
                text_td = row.find_element(By.CSS_SELECTOR, "td.dxeT")
                text = text_td.text.strip()
                if not text:
                    continue

                sold_out = "재고부족" in text
                size_name = text.replace("-재고부족-", "").replace("재고부족", "").replace("\xa0", " ")
                size_name = re.sub(r"\s+", " ", size_name).strip()

                items.append({
                    "size_name": size_name,
                    "sold_out": sold_out,
                    "raw_text": text,
                })
            except Exception:
                continue

        return items
    
    def wait_until_size_area_ready(self):
        size_btn_id = "ctl00_RootContentHolder_ASPxFormLayout1_CD_SIZE_CallbackPanel_CD_SIZE_DropDownEdit_B-1"

        WebDriverWait(self.driver, 10).until(
            EC.presence_of_element_located((By.ID, size_btn_id))
        )

        WebDriverWait(self.driver, 10).until(
            lambda d: d.find_element(By.ID, size_btn_id).is_displayed()
        )

        time.sleep(0.8)

    def check_product_options(self, product):
        all_lines = []

        original_window = self.driver.current_window_handle
        before_handles = self.driver.window_handles[:]

        cards = self.driver.find_elements(By.CSS_SELECTOR, "td.dxdvItem_MetropolisBlue")
        target_card = cards[product["card_index"]]

        clicked = False

        try:
            link_el = target_card.find_element(By.CSS_SELECTOR, "a.dxeHyperlink_MetropolisBlue")
            self.driver.execute_script("arguments[0].click();", link_el)
            clicked = True
        except Exception:
            pass

        if not clicked:
            try:
                img_el = target_card.find_element(By.CSS_SELECTOR, "img.dxeImage_MetropolisBlue")
                self.driver.execute_script("arguments[0].click();", img_el)
                clicked = True
            except Exception:
                pass

        if not clicked:
            raise RuntimeError("상품 클릭 실패")

        WebDriverWait(self.driver, 10).until(
            lambda d: len(d.window_handles) > len(before_handles)
        )

        new_handles = [h for h in self.driver.window_handles if h not in before_handles]
        if not new_handles:
            raise RuntimeError("새 탭이 열리지 않았습니다")

        detail_handle = new_handles[0]
        self.driver.switch_to.window(detail_handle)

        try:
            WebDriverWait(self.driver, 20).until(
                lambda d: "W0101020.aspx" in d.current_url
            )

            self.log(f"[DEBUG] 상세 탭 진입: {self.driver.current_url}")

            self.open_color_dropdown()
            color_options = self.get_color_options_from_dom()
            self.log(f"[DEBUG] 색상 수: {len(color_options)}")

            if not color_options:
                return "옵션없음 [확인필요]"

            for color in color_options:
                try:
                    color_name = self.select_color_by_index(color["index"])
                    self.log(f"[DEBUG] 색상 클릭 완료: {color_name}")

                    current_color_value = (
                        self.driver.find_element(
                            By.ID,
                            "ctl00_RootContentHolder_ASPxFormLayout1_CD_COLOR_ComboBox_I"
                        ).get_attribute("value") or ""
                    ).strip()
                    self.log(f"[DEBUG] 현재 색상 입력값: {current_color_value}")

                    self.wait_until_size_area_ready()
                    self.open_size_dropdown()

                    popup_visible = self.driver.find_element(
                        By.ID,
                        "ctl00_RootContentHolder_ASPxFormLayout1_CD_SIZE_CallbackPanel_CD_SIZE_DropDownEdit_DDD_PW-1"
                    ).is_displayed()
                    self.log(f"[DEBUG] 사이즈 팝업 표시 여부: {popup_visible}")

                    size_items = self.get_size_items_from_dom()
                    self.log(f"[DEBUG] 사이즈 raw 개수: {len(size_items)}")
                    for item in size_items:
                        self.log(f"[DEBUG] 사이즈 텍스트: {item['raw_text']}")

                    if not size_items:
                        all_lines.append(f"{color_name} - 옵션없음 [확인필요]")
                        continue

                    for size_item in size_items:
                        status_text = "품절" if size_item["sold_out"] else "판매가능"
                        all_lines.append(f"{color_name} - {size_item['size_name']} [{status_text}]")

                except Exception as e:
                    self.log(f"[DEBUG] 색상 조사 실패: {color['color_name']} / {type(e).__name__}: {repr(e)}")
                    all_lines.append(f"{color['color_name']} - 옵션없음 [확인필요]")

            if not all_lines:
                return "옵션없음 [확인필요]"

            return "\n".join(all_lines)

        finally:
            try:
                self.driver.close()
            finally:
                self.driver.switch_to.window(original_window)

    def has_next_page(self):
        try:
            next_btn = self.driver.find_element(By.CSS_SELECTOR, "#ctl00_ctl00_RootContentHolder_MainContentHolder_TopPager a[onclick*=\"'PBN'\"]")
            return next_btn.is_displayed() and next_btn.is_enabled()
        except Exception:
            return False

    def go_next_page(self):
        pager = WebDriverWait(self.driver, 10).until(
            EC.presence_of_element_located((By.ID, "ctl00_ctl00_RootContentHolder_MainContentHolder_TopPager"))
        )
        old_html = pager.get_attribute("innerHTML")

        next_btn = self.driver.find_element(By.CSS_SELECTOR, "#ctl00_ctl00_RootContentHolder_MainContentHolder_TopPager a[onclick*=\"'PBN'\"]")
        self.driver.execute_script("arguments[0].click();", next_btn)

        WebDriverWait(self.driver, 20).until(
            lambda d: d.find_element(By.ID, "ctl00_ctl00_RootContentHolder_MainContentHolder_TopPager").get_attribute("innerHTML") != old_html
        )

    def save_to_excel(self, rows, filename):
        self.log(f"엑셀 저장 시작: {filename}")

        wb = Workbook()
        ws = wb.active
        ws.title = "품절조사"

        ws.append(["상품명", "상품코드", "옵션별 품절 여부"])

        for row in rows:
            ws.append([
                row["product_name"],
                row["product_code"],
                row["option_status"],
            ])

        for col in ("A", "B", "C"):
            for cell in ws[col]:
                cell.alignment = Alignment(wrap_text=True, vertical="top")

        ws.column_dimensions["A"].width = 30
        ws.column_dimensions["B"].width = 18
        ws.column_dimensions["C"].width = 60

        wb.save(filename)
        self.log(f"엑셀 저장 완료: {filename}")

    def run_main_task(self):
        try:
            self.log("조사를 시작합니다.")
            self.log("현재 브라우저에 열려 있는 검색 결과 페이지를 기준으로 조사합니다.")
            self.log("상품 검색 결과가 바르게 보이는 상태인지 확인해주세요.")
            self.log(f"[DEBUG] 조사 시작 URL: {self.driver.current_url}")
            self.log(f"[DEBUG] 조사 시작 title: {self.driver.title}")

            page_index = 1 if not self.all_rows else 1

            while True:
                WebDriverWait(self.driver, 30).until(
                    lambda d: len(d.find_elements(By.CSS_SELECTOR, "td.dxdvItem_MetropolisBlue")) > 0
                )

                cards = self.driver.find_elements(By.CSS_SELECTOR, "td.dxdvItem_MetropolisBlue")
                self.log(f"[DEBUG] 페이지 {page_index} 카드 수: {len(cards)}")

                products = self.parse_product_cards()

                self.log(f"[페이지 {page_index}] 상품 수: {len(products)}")

                for idx, product in enumerate(products, start=1):
                    if product["product_code"] in self.processed_codes:
                        continue

                    try:
                        option_status = self.check_product_options(product)
                    except Exception as e:
                        option_status = f"조사 실패 [확인필요] - {e}"

                    self.all_rows.append({
                        "product_name": product["product_name"],
                        "product_code": product["product_code"],
                        "option_status": option_status,
                    })
                    self.processed_codes.add(product["product_code"])
                    self.log(f"{idx}/{len(products)} 완료 - {product['product_name']}")

                if not self.has_next_page():
                    break

                self.log("다음 페이지로 이동합니다.")
                self.go_next_page()
                page_index += 1
                time.sleep(1)

            self.save_to_excel(self.all_rows, self.output_path.get())
            self.log("엑셀 저장 완료")
            self.root.after(0, lambda: messagebox.showinfo("완료", f"엑셀 저장 완료\n{self.output_path.get()}"))

        except Exception as e:
            self.log(f"오류 발생: {e}")
            self.root.after(0, lambda: messagebox.showerror("오류", str(e)))
        finally:
            self.is_running = False
            self.root.after(0, lambda: self.start_button.config(state="normal"))

    def close_app(self):
        try:
            if self.all_rows:
                should_save = messagebox.askyesno(
                    "종료 확인",
                    "지금까지 조사한 결과를 저장한 뒤 종료하시겠습니까?"
                )
                if should_save:
                    save_path = self.output_path.get().strip()

                    if not save_path or save_path == OUTPUT_FILE:
                        path = filedialog.asksaveasfilename(
                            title="엑셀 저장 위치 선택",
                            defaultextension=".xlsx",
                            filetypes=[("Excel files", "*.xlsx")]
                        )
                        if not path:
                            cancel_exit = messagebox.askyesno(
                                "저장 취소",
                                "저장 위치를 선택하지 않았습니다.\n저장하지 않고 종료하시겠습니까?"
                            )
                            if not cancel_exit:
                                return
                        else:
                            save_path = path
                            self.output_path.set(path)

                    if save_path and save_path != OUTPUT_FILE:
                        try:
                            self.save_to_excel(self.all_rows, save_path)
                            messagebox.showinfo("저장 완료", f"엑셀 파일이 저장되었습니다.\n{save_path}")
                        except Exception as e:
                            messagebox.showerror("저장 실패", f"엑셀 저장 중 오류가 발생했습니다.\n{e}")
                            return
            else:
                exit_now = messagebox.askyesno(
                    "종료 확인",
                    "아직 저장할 조사 결과가 없습니다.\n그대로 종료하시겠습니까?"
                )
                if not exit_now:
                    return

            if self.driver is not None:
                self.driver.quit()
        except Exception:
            pass

        self.root.destroy()

if __name__ == "__main__":
    root = tk.Tk()
    app = CerrotorreApp(root)
    root.mainloop()
