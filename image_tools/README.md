# 이미지 처리

| 파일 | 기능 |
|---|---|
| [img_resizing_v1.py](img_resizing_v1.py) | 폴더 내 이미지를 지정 너비로 비율 유지 리사이징 |
| [image_size_reducer_gui_percent.py](image_size_reducer_gui_percent.py) | 픽셀 크기를 유지하며 이미지 용량을 줄이는 GUI |
| [convertIMG.py](convertIMG.py) | WebP 이미지를 JPG로 변환하고 투명 배경 처리 |
| [gif_to_jpg_v1.py](gif_to_jpg_v1.py) | GIF의 첫 프레임을 JPG로 변환 |

## 설치와 실행

저장소 루트에서 실행합니다. Windows 데스크톱 Python과 Tkinter를 기준으로 합니다.

```powershell
python -m pip install -r image_tools/requirements.txt
python run.py image_tools/img_resizing_v1.py
```

## 설정

- `img_resizing_v1.py`: 공통 사이트 설정 없음
- `image_size_reducer_gui_percent.py`: 공통 사이트 설정 없음
- `convertIMG.py`: 공통 사이트 설정 없음
- `gif_to_jpg_v1.py`: 공통 사이트 설정 없음

사이트 설정은 루트의 `config.example.json`을 참고해 `config.local.json`에 입력합니다. 사이트별 DOM과 업무 규칙을 전제로 작성된 도구이므로 주소만 바꿔 임의의 쇼핑몰에서 사용할 수 있는 범용 프로그램은 아닙니다.

## 이미지 도구 사용법

- 리사이징: 입력 폴더, 출력 폴더, 목표 너비를 선택합니다. 원본과 다른 출력 폴더를 사용하고 너비에는 양의 정수를 입력합니다.
- 용량 줄이기: 목표 용량 비율과 저장 형식을 선택합니다. 지정한 비율은 목표이며 이미지 특성·설정에 따라 실제 결과가 다릅니다. 픽셀 크기는 변경하지 않습니다.
- WebP 변환: 입력·출력 폴더와 JPG 품질을 설정합니다. 투명 영역은 선택한 배경색으로 합성합니다.
- GIF 변환: 첫 프레임만 정지 이미지로 변환합니다. 애니메이션은 보존하지 않습니다.

```powershell
python run.py image_tools/gif_to_jpg_v1.py data/input data/gif-output
```

## 업무 사례

상품 등록 전에 반복하던 이미지 리사이징을 도구로 처리했습니다. 작성자의 경험상 상품당 전체 등록 시간은 5분 이상에서 2~3분으로 줄었습니다. 약 40~60% 단축은 5분을 기준으로 한 추정이며, 정식 벤치마크나 모든 도구의 성능 보장은 아닙니다.
