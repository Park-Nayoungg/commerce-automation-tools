# 관리자 상품 일괄 수정

| 파일 | 기능 |
|---|---|
| [godomall_option_stock_sync.py](godomall_option_stock_sync.py) | 엑셀과 상품·옵션을 매칭하고 관리자 재고에 반영 |
| [godo_thumb_update_500to1000.py](godo_thumb_update_500to1000.py) | 상품 이미지를 1000px 정사각형으로 변환·업로드 |
| [good_img_update_v1.py](good_img_update_v1.py) | 관리자 상품 이미지 설정 일괄 수정 |
| [seo_supporter_v2.py](seo_supporter_v2.py) | 상품 SEO 관련 항목 정리·수정 |
| [set_max_buy.py](set_max_buy.py) | 상품 최소·최대 구매 수량 설정 |
| [update_stock_v1.py](update_stock_v1.py) | 구형 관리자 페이지 재고 일괄 수정 |

## 설치와 실행

저장소 루트에서 실행합니다. Windows 데스크톱 Python과 Tkinter를 기준으로 합니다.

```powershell
python -m pip install -r product_management/requirements.txt
python run.py product_management/godomall_option_stock_sync.py
```

## 설정

- `godomall_option_stock_sync.py`: `godomall_base_url`
- `godo_thumb_update_500to1000.py`: `godomall_base_url`
- `good_img_update_v1.py`: `godomall_base_url`
- `seo_supporter_v2.py`: `godomall_base_url`
- `set_max_buy.py`: `godomall_base_url`
- `update_stock_v1.py`: `legacy_shop_base_url`

사이트 설정은 루트의 `config.example.json`을 참고해 `config.local.json`에 입력합니다. 사이트별 DOM과 업무 규칙을 전제로 작성된 도구이므로 주소만 바꿔 임의의 쇼핑몰에서 사용할 수 있는 범용 프로그램은 아닙니다.

## 실제 변경이 발생하는 도구

이 폴더의 도구는 관리자 화면에서 상품·이미지·재고 값을 저장합니다. 단순 수집 도구와 구분하여, 백업한 데이터와 소수의 테스트 상품으로 대상 및 업무 규칙을 확인한 뒤 사용합니다. 공통 dry-run 기능은 없습니다.

특히 옵션 재고 동기화는 엑셀의 상품·옵션 매칭과 재고 변환 규칙을 사용합니다. 적용할 사이트의 옵션명·품절 정책과 코드의 규칙이 맞는지 확인해야 합니다. 이번 정리에서는 실제 상품을 변경하지 않았습니다.
