"""Convert the first frame of each GIF without a hard-coded local folder."""
import argparse
from pathlib import Path
from PIL import Image

def convert_folder(input_directory, output_directory):
    source = Path(input_directory)
    target = Path(output_directory)
    if not source.is_dir():
        raise ValueError("입력 폴더가 존재하지 않습니다.")
    target.mkdir(parents=True, exist_ok=True)
    count = 0
    for path in sorted(source.iterdir()):
        if path.is_file() and path.suffix.lower() == ".gif":
            destination = target / (path.stem + ".jpg")
            if destination.exists():
                raise FileExistsError(f"기존 출력 파일이 있습니다: {destination.name}")
            with Image.open(path) as image:
                image.seek(0)
                image.convert("RGB").save(destination, "JPEG")
            count += 1
    return count

def main():
    parser = argparse.ArgumentParser(description="GIF 첫 프레임을 JPG로 변환")
    parser.add_argument("input_directory")
    parser.add_argument("output_directory")
    args = parser.parse_args()
    print(f"변환 완료: {convert_folder(args.input_directory, args.output_directory)}개")

if __name__ == "__main__":
    main()
