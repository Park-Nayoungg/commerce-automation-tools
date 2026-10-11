# 상품 데이터·이미지 수집

| 파일 | 기능 |
|---|---|
| [crawl_gocamp_prod.py](crawl_gocamp_prod.py) | 고도몰 관리자 상품 목록 수집 및 엑셀 저장 |
| [crawl_northpeak_prod.py](crawl_northpeak_prod.py) | Northpeak 상품명 수집 및 엑셀 저장 |
| [productinfocollect_v2.py](productinfocollect_v2.py) | 구형 관리자 페이지 상품·옵션 정보 수집 |
| [godo_IMG_Download_v1.py](godo_IMG_Download_v1.py) | 고도몰 관리자 상품 이미지 다운로드 |
| [godo_IMG_Download_v2.py](godo_IMG_Download_v2.py) | 구형 관리자 페이지 상품 이미지 다운로드 |

## 설치와 실행

저장소 루트에서 실행합니다. Windows 데스크톱 Python과 Tkinter를 기준으로 합니다.

```powershell
python -m pip install -r product_data/requirements.txt
python run.py product_data/crawl_gocamp_prod.py
```

## 설정

- `crawl_gocamp_prod.py`: `godomall_base_url`
- `crawl_northpeak_prod.py`: 공통 사이트 설정 없음
- `productinfocollect_v2.py`: `legacy_shop_base_url`, `output_directory`
- `godo_IMG_Download_v1.py`: `godomall_base_url`, `output_directory`
- `godo_IMG_Download_v2.py`: `legacy_shop_base_url`, `output_directory`

사이트 설정은 루트의 `config.example.json`을 참고해 `config.local.json`에 입력합니다. 사이트별 DOM과 업무 규칙을 전제로 작성된 도구이므로 주소만 바꿔 임의의 쇼핑몰에서 사용할 수 있는 범용 프로그램은 아닙니다.

## 입력과 출력

브라우저에서 직접 로그인하고 상품을 검색한 뒤 안내 창에서 진행합니다. 목록 수집 결과는 선택한 위치에 엑셀로 저장합니다. 이미지 다운로드 도구의 기본 출력은 `data/downloads/`입니다. 상품 정보 저장 위치를 취소한 경우의 기본 경로는 `data/exports/`입니다.

`godo_IMG_Download_v1.py`는 고도몰 관리자, `godo_IMG_Download_v2.py`는 구형 관리자 페이지를 대상으로 합니다. 버전 번호만 보고 동일 기능의 이전·최신 버전으로 취급하지 않습니다.
