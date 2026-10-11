# 재고·가격 확인

| 파일 | 기능 |
|---|---|
| [b2b_bplstore.py](b2b_bplstore.py) | BPL 상품·옵션 재고 수집 및 엑셀 저장 |
| [b2b_cerettore.py](b2b_cerettore.py) | Cerrotorre 옵션별 품절 조사 GUI |
| [b2b_exped.py](b2b_exped.py) | Exped 상품·옵션 재고 수집 및 기존 엑셀 갱신 |
| [b2b_exped_comfort.py](b2b_exped_comfort.py) | Exped 상품·재고 정보를 수집하는 별도 구현 |
| [b2b_jdo.py](b2b_jdo.py) | JDO 재고 수집·키워드 검색·변화 감시 |
| [b2b_parkers.py](b2b_parkers.py) | Parkers 재고 수집·키워드 검색·변화 감시 |
| [b2boutdoorbridgecrawl.py](b2boutdoorbridgecrawl.py) | B2B Outdoor 상품별 옵션·재고 수집 |
| [price_monitor.py](price_monitor.py) | 관리자 상품 가격과 엑셀 가격을 비교해 결과 표시 |

## 설치와 실행

저장소 루트에서 실행합니다. Windows 데스크톱 Python과 Tkinter를 기준으로 합니다.

```powershell
python -m pip install -r inventory_tools/requirements.txt
python run.py inventory_tools/b2b_bplstore.py
```

## 설정

- `b2b_bplstore.py`: 공통 사이트 설정 없음
- `b2b_cerettore.py`: 공통 사이트 설정 없음
- `b2b_exped.py`: 공통 사이트 설정 없음
- `b2b_exped_comfort.py`: 공통 사이트 설정 없음
- `b2b_jdo.py`: 공통 사이트 설정 없음
- `b2b_parkers.py`: 공통 사이트 설정 없음
- `b2boutdoorbridgecrawl.py`: `outdoor_start_url`
- `price_monitor.py`: `godomall_base_url`

사이트 설정은 루트의 `config.example.json`을 참고해 `config.local.json`에 입력합니다. 사이트별 DOM과 업무 규칙을 전제로 작성된 도구이므로 주소만 바꿔 임의의 쇼핑몰에서 사용할 수 있는 범용 프로그램은 아닙니다.

## 입력과 출력

사이트별 로그인·검색·저장 위치 안내를 따릅니다. JDO와 Parkers 도구는 재고 수집 외에 키워드 검색과 재고 변화 감시 기능을 포함합니다. 감시 기록은 `*_snapshot.json`, 결과는 엑셀에 저장하며 Git에서 제외합니다. 가격 비교 도구는 엑셀 파일과 상품명·가격 열을 선택해 관리자 화면의 가격과 비교합니다.

`b2b_exped.py`와 `b2b_exped_comfort.py`는 서로 다른 구현을 보존한 것입니다. 실행 대상은 필요한 기능과 해당 사이트 상태에 맞춰 선택합니다.

## 선택 사항: EXE 제작

루트에서 빌드 도구를 설치한 후 이 폴더에서 빌드합니다. EXE는 이번 정리에서 빌드·검증하지 않았습니다.

```powershell
python -m pip install -r requirements-build.txt
cd inventory_tools
python -m PyInstaller b2b_cerettore.spec
python -m PyInstaller b2b_parkers.spec
```

PyInstaller가 현재 사용 폴더에서 실행되도록 `.spec`를 원본 스크립트 옆에 보관했습니다.
