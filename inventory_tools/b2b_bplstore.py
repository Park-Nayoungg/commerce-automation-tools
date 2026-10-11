import os
import re
import json
import time
from datetime import datetime
from typing import Any, Optional
from urllib.parse import urlparse, parse_qs

import requests
from bs4 import BeautifulSoup

import tkinter as tk
from tkinter import filedialog, messagebox

from openpyxl import Workbook, load_workbook
from openpyxl.styles import Alignment


BASE = "https://www.bplstore.com"
CATEGORY_NO = 378

LIST_URL = BASE + "/product/list.html?cate_no={cate_no}&page={page}"
DETAIL_URL = BASE + "/product/detail.html?product_no={product_no}&cate_no={cate_no}&display_group=1"

HEADERS = {
    "User-Agent": "Mozilla/5.0",
}

SAVE_EVERY_PAGES = 2
REQUEST_SLEEP_SEC = 0.4


def fetch(url: str) -> str:
    r = requests.get(url, headers=HEADERS, timeout=30)
    r.raise_for_status()
    return r.text


def parse_product_nos_from_list(html: str) -> list[int]:
    soup = BeautifulSoup(html, "html.parser")
    nos: set[int] = set()

    for a in soup.select('a[href*="/product/detail.html?product_no="]'):
        href = a.get("href") or ""
        m = re.search(r"product_no=(\d+)", href)
        if m:
            nos.add(int(m.group(1)))

    return sorted(nos)


def parse_last_page_from_list(html: str) -> Optional[int]:
    soup = BeautifulSoup(html, "html.parser")
    a = soup.select_one(".xans-product-normalpaging p.last a[href*='page=']")
    if not a:
        return None

    href = a.get("href") or ""
    try:
        qs = parse_qs(urlparse(href).query)
        p = qs.get("page", [None])[0]
        return int(p) if p is not None else None
    except Exception:
        return None

def _parse_price_to_int(text: str) -> str:
    # "498,000원" -> "498000"
    s = (text or "").strip()
    s = s.replace(",", "").replace("원", "").strip()
    m = re.search(r"\d+", s)
    return m.group(0) if m else ""


def parse_products_from_list(html: str) -> list[dict]:
    """
    반환 예:
    [
        {
            "product_no": 4069,
            "product_name": "...",   # 리스트에 있는 상품명(선택)
            "sale_price": "498000",
            "consumer_price": "498000",
            "brand": "MYSTERY RANCH",
            "detail_url": "https://www.bplstore.com/product/detail.html?product_no=4069&cate_no=378&display_group=1",
        },
        ...
    ]
    """
    soup = BeautifulSoup(html, "html.parser")
    out: list[dict] = []

    for li in soup.select('ul.prdList > li[id^="anchorBoxId_"]'):
        li_id = li.get("id") or ""
        m = re.search(r"anchorBoxId_(\d+)", li_id)
        if not m:
            continue

        product_no = int(m.group(1))

        a = li.select_one('a[href*="/product/detail.html?product_no="]')
        href = (a.get("href") if a else "") or ""
        detail_url = BASE + href if href.startswith("/") else href

        a_tag = li.select_one("p.name > a")
        product_name = ""

        if a_tag:
            # 실제 상품명: strong.title 다음에 오는 span
            real_name_span = a_tag.select_one("strong.title + span")
            if real_name_span:
                product_name = real_name_span.get_text(" ", strip=True).replace("\n", " ").strip()
            else:
                # 방어: 마지막 span을 상품명으로
                spans = a_tag.select("> span")
                if spans:
                    product_name = spans[-1].get_text(" ", strip=True).replace("\n", " ").strip()

        # 혹시 라벨만 잡힌 경우 방어
        if product_name == "상품명":
            product_name = ""

        sale_price = ""
        consumer_price = ""
        brand = ""

        # 리스트 아이템 영역에서 "판매가/소비자가/브랜드" 라벨을 찾아 값 span을 읽음
        for item_li in li.select("ul.xans-product-listitem > li"):
            label_el = item_li.select_one("strong span")
            value_el = item_li.select_one("strong ~ span")

            label = label_el.get_text(" ", strip=True) if label_el else ""
            value = value_el.get_text(" ", strip=True) if value_el else ""

            if "판매가" in label:
                sale_price = _parse_price_to_int(value)
            elif "소비자가" in label:
                consumer_price = _parse_price_to_int(value)
            elif "브랜드" in label:
                brand = value.strip()

        out.append({
            "product_no": product_no,
            "product_name": product_name,
            "sale_price": sale_price,
            "consumer_price": consumer_price,
            "brand": brand,
            "detail_url": detail_url,
        })

    # 중복 제거(페이지 내 중복 방지)
    seen = set()
    uniq = []
    for p in out:
        n = p["product_no"]
        if n in seen:
            continue
        seen.add(n)
        uniq.append(p)

    return uniq

def extract_product_name(detail_html: str) -> str:
    soup = BeautifulSoup(detail_html, "html.parser")

    # 1) 사용자가 확인한 구조: detailArea 안의 h3
    h3 = soup.select_one("div.detailArea h3")
    if h3:
        # 해외배송 문구 등 displaynone 제거
        for sp in h3.select(".displaynone"):
            sp.decompose()
        return h3.get_text(" ", strip=True)

    # 2) 혹시 다른 스킨에서 h2를 쓰는 경우 대비
    h2 = soup.select_one("div.detailArea h2")
    if h2:
        for sp in h2.select(".displaynone"):
            sp.decompose()
        return h2.get_text(" ", strip=True)

    # 3) 최후의 fallback
    title = soup.title.get_text(" ", strip=True) if soup.title else ""
    return title.strip()


def extract_option_stock_data(detail_html: str) -> dict[str, Any]:
    patterns = [
        r"option_stock_data\s*=\s*'([^']*)'\s*;",
        r'option_stock_data\s*=\s*"([^"]*)"\s*;',
    ]

    raw = None
    for p in patterns:
        m = re.search(p, detail_html)
        if m:
            raw = m.group(1)
            break

    if raw is None:
        return {}

    raw1 = raw.replace(r"\/", "/")

    # 1) 바로 JSON 파싱 시도
    try:
        data = json.loads(raw1)
        if isinstance(data, dict):
            return data
    except Exception:
        pass

    # 2) unicode_escape 후 재시도
    try:
        raw2 = raw1.encode("utf-8").decode("unicode_escape")
        data = json.loads(raw2)
        if isinstance(data, dict):
            return data
        return {}
    except Exception:
        return {}

def extract_single_stock(detail_html: str) -> tuple[str, bool]:
    """
    옵션 없는 상품(단품) 재고 추출
    반환:
        stock_text: "10" 같은 숫자 문자열(못 찾으면 "")
        soldout: 품절 여부(재고 0이면 True)
    우선순위:
        1) single_option_stock_data JSON 안의 stock_number
        2) var stock_number = '...'
    """
    # 1) single_option_stock_data 우선
    m = re.search(r"single_option_stock_data\s*=\s*'([^']*)'\s*;", detail_html)
    if not m:
        m = re.search(r'single_option_stock_data\s*=\s*"([^"]*)"\s*;', detail_html)

    if m:
        raw = m.group(1)

        # 역슬래시 이스케이프가 섞여 있으니 2단계로 시도
        # - 1차: \/ 치환
        # - 2차: unicode_escape 디코드
        raw1 = raw.replace(r"\/", "/")

        for candidate in (raw1, raw1.encode("utf-8").decode("unicode_escape")):
            try:
                j = json.loads(candidate)
                if isinstance(j, dict):
                    sn = j.get("stock_number", "")
                    try:
                        n = int(sn)
                        return str(n), (n == 0)
                    except Exception:
                        s = str(sn).strip()
                        if s.isdigit():
                            return s, (s == "0")
                        return s, (s == "0")
            except Exception:
                pass

    # 2) var stock_number = '0' fallback
    m2 = re.search(r"var\s+stock_number\s*=\s*'(\d+)'\s*;", detail_html)
    if m2:
        s = m2.group(1).strip()
        return s, (s == "0")

    return "", False


def ask_save_xlsx_path() -> str:
    root = tk.Tk()
    root.withdraw()
    root.attributes("-topmost", True)

    path = filedialog.asksaveasfilename(
        title="BPL 재고 엑셀(.xlsx) 저장 위치 선택",
        defaultextension=".xlsx",
        filetypes=[("Excel 파일", "*.xlsx")],
        initialfile="bpl_cate378_stock.xlsx",
    )

    root.destroy()
    return path

def ask_exit_after_done(out_path: str) -> None:
    root = tk.Tk()
    root.withdraw()
    root.attributes("-topmost", True)

    answer = messagebox.askyesno(
        "작업 완료",
        f"재고 수집이 완료되었습니다.\n\n저장 위치:\n{out_path}\n\n프로그램을 종료할까요?"
    )

    root.destroy()

    if answer:
        os._exit(0)  # 완전 종료


def open_or_create_workbook(out_path: str) -> tuple[Workbook, Any, set[int]]:
    """
    Products 시트:
      - product_no
      - product_name
      - option_summary  (줄바꿈)
      - detail_url
      - fetched_at
    """
    expected = [ "product_no", "product_name", "sale_price", "consumer_price", "brand", "option_summary", "detail_url", "fetched_at", ]
    if os.path.exists(out_path):
        wb = load_workbook(out_path)
        ws = wb["Products"] if "Products" in wb.sheetnames else wb.active
        ws.title = "Products"

        header = [c.value for c in ws[1]] if ws.max_row >= 1 else []
        if header != expected:
            raise RuntimeError(f"기존 파일 헤더가 예상과 다릅니다.\n기존: {header}\n예상: {expected}")

        existing: set[int] = set()
        for row in ws.iter_rows(min_row=2, values_only=True):
            if not row:
                continue
            v = row[0]
            if v is None:
                continue
            try:
                existing.add(int(v))
            except Exception:
                continue

        return wb, ws, existing

    wb = Workbook()
    ws = wb.active
    ws.title = "Products"
    ws.append(expected)

    ws.column_dimensions["F"].width = 60
    ws.column_dimensions["B"].width = 45 

    return wb, ws, set()


def build_option_summary(option_data: dict[str, Any]) -> str:
    """
    옵션 1개당 한 줄:
      - "옵션값 - 재고(n)" 또는 "옵션값 [품절]" 또는 "옵션값 - 재고(미확인)"
    """
    if not option_data:
        return ""

    lines: list[str] = []

    # 옵션 표시는 사람이 보기 좋게 option_value 기준 정렬(없으면 item_code)
    sortable = []
    for item_code, v in option_data.items():
        if not isinstance(v, dict):
            continue
        opt_val = str(v.get("option_value", "")).strip()
        sortable.append((opt_val, item_code, v))

    sortable.sort(key=lambda x: (x[0] == "", x[0], x[1]))

    for opt_val, item_code, v in sortable:
        name = opt_val if opt_val else str(item_code)

        stock_raw = v.get("stock_number", "")
        is_selling = str(v.get("is_selling", "")).strip()  # "T"/"F"
        is_display = str(v.get("is_display", "")).strip()  # "T"/"F"
        use_stock = v.get("use_stock", None)               # bool 가능

        stock_text = ""
        stock_int: Optional[int] = None

        try:
            stock_int = int(stock_raw)
            stock_text = str(stock_int)
        except Exception:
            stock_text = str(stock_raw).strip()

        # 품절 판단(원하시면 기준 조정 가능)
        soldout = False
        if stock_int is not None and stock_int == 0:
            soldout = True
        if is_selling and is_selling != "T":
            soldout = True
        if is_display and is_display != "T":
            soldout = True

        if soldout:
            lines.append(f"{name} [품절]")
        else:
            if stock_text != "":
                # use_stock가 False여도 재고 숫자가 있으면 표시
                lines.append(f"{name} - 재고({stock_text})")
            else:
                lines.append(f"{name} - 재고(미확인)")

    return "\n".join(lines).strip()


def set_wrap_alignment(ws, row_idx: int, col_idx: int) -> None:
    cell = ws.cell(row=row_idx, column=col_idx)
    cell.alignment = Alignment(wrap_text=True, vertical="top")


def main() -> None:
    out_path = ask_save_xlsx_path()
    if not out_path:
        messagebox.showinfo("취소됨", "저장 위치 선택이 취소되어 작업을 종료합니다.")
        return

    wb, ws, existing_products = open_or_create_workbook(out_path)

    # 총 페이지 추정
    first_html = fetch(LIST_URL.format(cate_no=CATEGORY_NO, page=1))
    last_page = parse_last_page_from_list(first_html)
    if last_page is None:
        last_page = 999999  # 마지막 페이지를 못 찾으면 빈 페이지 나올 때까지

    fetched_at = datetime.now().strftime("%Y-%m-%d %H:%M:%S")

    saved_pages = 0
    total_new_products = 0

    page = 1
    while page <= last_page:
        list_url = LIST_URL.format(cate_no=CATEGORY_NO, page=page)
        list_html = fetch(list_url)

        products = parse_products_from_list(list_html)
        if not products:
            wb.save(out_path)
            print(f"[DONE] page={page} 상품 없음. 저장 후 종료 -> {out_path}")
            messagebox.showinfo("완료", f"작업 완료:\n{out_path}")
            return

        print(f"[PAGE] {page}/{last_page if last_page != 999999 else '?'}  products={len(products)}")

        new_on_page = 0
        for p in products:
            product_no = int(p["product_no"])
            if product_no in existing_products:
                continue

            durl = p.get("detail_url") or DETAIL_URL.format(product_no=product_no, cate_no=CATEGORY_NO)

            detail_html = fetch(durl)

            # 상품명은 리스트 우선, 비면 상세에서 보완
            product_name = (p.get("product_name") or "").strip()
            if (not product_name) or (product_name == "상품명"):
                product_name = extract_product_name(detail_html)

            option_data = extract_option_stock_data(detail_html)

            if option_data:
                option_summary = build_option_summary(option_data)
            else:
                single_stock, single_soldout = extract_single_stock(detail_html)
                if single_stock != "":
                    if single_soldout:
                        option_summary = f"단품 [품절]"
                    else:
                        option_summary = f"단품 - 재고({single_stock})"
                else:
                    option_summary = ""

            ws.append([
                product_no,
                product_name,
                p.get("sale_price", ""),
                p.get("consumer_price", ""),
                p.get("brand", ""),
                option_summary,
                durl,
                fetched_at,
            ])

            # option_summary(6열) 줄바꿈 표시
            set_wrap_alignment(ws, ws.max_row, 6)

            existing_products.add(product_no)
            total_new_products += 1
            new_on_page += 1

            if total_new_products % 10 == 0:
                print(f"  [PROGRESS] new_products={total_new_products} (latest product_no={product_no})")

            time.sleep(REQUEST_SLEEP_SEC)

        print(f"[PAGE DONE] page={page} new={new_on_page} total_new={total_new_products}")

        saved_pages += 1
        if saved_pages >= SAVE_EVERY_PAGES:
            wb.save(out_path)
            saved_pages = 0
            print(f"[SAVE] every {SAVE_EVERY_PAGES} pages -> {out_path}")

        page += 1

    wb.save(out_path)
    ask_exit_after_done(out_path)
    return


if __name__ == "__main__":
    main()
