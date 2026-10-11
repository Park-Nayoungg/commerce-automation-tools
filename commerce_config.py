"""Local runtime settings; config.local.json is deliberately excluded from Git."""
import json
import os
from pathlib import Path
from urllib.parse import urlsplit

ROOT = Path(__file__).resolve().parent

def settings():
    path = Path(os.environ.get("COMMERCE_CONFIG", ROOT / "config.local.json"))
    if not path.is_file():
        return {}
    try:
        value = json.loads(path.read_text(encoding="utf-8-sig"))
    except (OSError, ValueError) as error:
        raise RuntimeError("config.local.json 파일의 JSON 형식을 확인하세요.") from error
    if not isinstance(value, dict):
        raise RuntimeError("설정 파일은 JSON 객체여야 합니다.")
    return value

def configured_url(key):
    value = settings().get(key, "")
    parts = urlsplit(value) if isinstance(value, str) else None
    if not parts or parts.scheme not in ("http", "https") or not parts.hostname or parts.username or parts.password:
        raise RuntimeError(f"config.local.json의 {key}에 유효한 사이트 주소를 입력하세요. config.example.json을 참고하세요.")
    if parts.hostname.endswith(".invalid") or parts.hostname in ("example.com", "example.org"):
        raise RuntimeError(f"{key}의 예시 주소를 실제 사용 가능한 주소로 변경하세요.")
    if key.endswith("base_url"):
        if parts.path not in ("", "/") or parts.query or parts.fragment:
            raise RuntimeError(f"{key}에는 경로 없이 사이트의 기본 주소만 입력하세요.")
        return value.rstrip("/")
    return value

def output_dir(name):
    directory = Path(settings().get("output_directory", "data"))
    if not directory.is_absolute():
        directory = ROOT / directory
    directory = directory / name
    directory.mkdir(parents=True, exist_ok=True)
    return directory
