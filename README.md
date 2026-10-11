# Commerce Automation Tools

쇼핑몰 운영 중 반복되던 이미지 가공, 상품 정보 수집, 재고 확인, 관리자 상품 수정 업무를 개선하기 위해 만든 Python 도구 모음입니다. 업무별 단일 스크립트를 기능별 폴더로 정리했습니다.

## 해결한 문제

| 업무 문제 | 접근 | 실제 업무에서의 결과 |
|---|---|---|
| 상품 등록마다 이미지 크기를 수동 수정 | 이미지 리사이징 프로그램 제작·활용 | 상품당 등록 시간 5분 이상 → 2~3분 |
| B2B 재고 확인과 물량 확보 | 사이트별 재고 파악 도구 제작 | 다른 업체보다 먼저 물량을 확보하고 매출 확대에 기여 |
| 상품을 하나씩 수정하는 반복 업무 | 관리자 화면의 반복 작업 자동화 | 1만 개 이상 상품 수정 작업에 자동화 적용 |

성과는 작성자의 실제 업무 경험을 바탕으로 기술했습니다. 리사이징의 약 40~60% 시간 단축은 기존 5분 기준의 경험적 추정이며 정식 벤치마크가 아닙니다. 1만 개 이상은 업무 적용 규모이며 개별 스크립트의 처리량 보장이 아닙니다. 구체적인 매출 증가율은 기재하지 않았습니다.

## 구성

| 폴더 | 역할 | 도구 수 |
|---|---|---|
| [image_tools](image_tools/README.md) | 리사이징·용량 축소·형식 변환 | 4 |
| [inventory_tools](inventory_tools/README.md) | 재고 수집·변화 감시·가격 비교 | 8 |
| [product_data](product_data/README.md) | 상품 정보·이미지 수집 | 5 |
| [product_management](product_management/README.md) | 관리자 상품·이미지·재고 수정 | 6 |

총 23개 도구와 2개의 선택적 PyInstaller 빌드 설정을 포함합니다. 전체 도구 목록은 `python run.py --list`로 확인할 수 있습니다.

## 시작하기

전체 의존성 설치에는 Python 3.11 이상이 필요하며, 이번 로컬 검증은 Python 3.13에서 수행했습니다. Tkinter가 포함된 Windows 데스크톱 환경을 기준으로 합니다. 브라우저 자동화 도구는 Chrome 및 로그인 가능한 계정이 필요합니다. 이미지만 처리하는 도구는 브라우저가 필요하지 않습니다.

아래 명령은 저장소 루트에서 실행합니다. 가상환경 활성화 없이 실행 파일을 직접 지정하므로 PowerShell 실행 정책 변경이 필요하지 않습니다.

```powershell
python -m venv .venv
.\.venv\Scripts\python.exe -m pip install -r requirements.txt
.\.venv\Scripts\python.exe run.py --list
.\.venv\Scripts\python.exe run.py image_tools/img_resizing_v1.py
```

이미지 도구만 사용할 경우 `requirements.txt` 대신 `image_tools/requirements.txt`로 설치합니다. 의존성 버전은 정리 시점의 검증 환경에 설치된 버전을 기준으로 기록했습니다. 모든 Python 버전 조합에서 설치·동작을 보증하는 잠금 파일은 아닙니다.

## 사이트와 저장 경로 설정

공개 저장소에는 실제 운영 설정을 포함하지 않습니다. 새로 내려받았다면 예시 파일을 복사합니다. 이미 개인 설정 파일이 있다면 덮어쓰지 않습니다.

```powershell
Copy-Item config.example.json config.local.json
```

| 설정 키 | 값 |
|---|---|
| `godomall_base_url` | 본인 고도몰 관리자 기본 주소. 경로·검색 조건 없이 입력 |
| `legacy_shop_base_url` | `/shop/admin/` 경로를 사용하는 기존 쇼핑몰의 기본 주소 |
| `outdoor_start_url` | B2B Outdoor에서 사용할 전체 검색·목록 URL |
| `output_directory` | 기본값 `data`. 상대 경로는 저장소 루트 기준 |

계정과 비밀번호는 설정 파일에 저장하지 않습니다. 해당 도구의 안내에 따라 열린 브라우저에서 직접 로그인합니다. 공급사별 공개 기본 주소·페이지 선택자는 코드에 남아 있으며, 사이트 개편에 따라 수정이 필요할 수 있습니다. `COMMERCE_CONFIG` 환경변수로 다른 설정 파일 경로를 지정할 수도 있습니다.

## AI 활용과 작성자 역할

ChatGPT와 Codex를 활용해 필요한 자동화 프로그램을 만들고 실제 업무에 사용했습니다. 이 저장소는 AI를 활용한 구현 경험, 업무별 사용 목적과 실제 적용 결과를 함께 정리한 것입니다.

## 검증 및 사용 범위

- Python 문법 검사와 로컬 이미지 처리·설정 로딩 테스트를 수행합니다. 상세 결과는 [검증 기록](docs/VALIDATION.md)에 있습니다.
- 실제 사이트 로그인, 실시간 재고의 정확도, 운영 상품 변경, 브라우저 GUI 전체 흐름과 EXE 빌드는 별도 현장 검증 대상입니다.
- `product_management`는 실제 값을 저장하는 도구입니다. 실행 전에 대상 상품과 옵션·재고 규칙을 확인합니다.
- 운영 설정, 엑셀 결과, 다운로드, 재고 스냅샷, 가상환경과 실행파일은 Git에서 제외합니다.

## 문서

- [GitHub 최초 커밋·푸시](docs/GITHUB_UPLOAD.md)
- [공개용 정리 내역](docs/PREPARATION.md)
- [로컬 검증 기록](docs/VALIDATION.md)
