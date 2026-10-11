"""List and launch a single tool from the repository root."""
import argparse
import json
import os
from pathlib import Path
import runpy
import sys

ROOT = Path(__file__).resolve().parent

def main():
    parser = argparse.ArgumentParser(description="쇼핑몰 업무 자동화 도구 실행")
    parser.add_argument("tool", nargs="?", help="예: image_tools/img_resizing_v1.py")
    parser.add_argument("--list", action="store_true", help="도구 목록만 표시")
    parser.add_argument("tool_args", nargs=argparse.REMAINDER)
    args = parser.parse_args()
    tools = json.loads((ROOT / "tools.json").read_text(encoding="utf-8"))
    by_path = {item["path"]: item for item in tools}
    if args.list or not args.tool:
        for item in tools:
            print(f'{item["path"]}  —  {item["description"]}')
        return
    key = args.tool.replace("\\", "/")
    if key not in by_path:
        parser.error("목록에 있는 도구 경로를 지정하세요. python run.py --list")
    from commerce_config import configured_url
    try:
        for setting in by_path[key]["configuration"]:
            if setting.endswith("url"):
                configured_url(setting)
    except RuntimeError as error:
        parser.error(str(error))
    os.chdir(ROOT)
    sys.path.insert(0, str(ROOT))
    sys.argv = [str(ROOT / key), *args.tool_args]
    runpy.run_path(str(ROOT / key), run_name="__main__")

if __name__ == "__main__":
    main()
