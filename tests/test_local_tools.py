import importlib.util
import json
import os
from pathlib import Path
import sys
import tempfile
import unittest
from unittest.mock import patch

from PIL import Image

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT))
import commerce_config


def load_tool(relative):
    name = Path(relative).stem
    spec = importlib.util.spec_from_file_location(name, ROOT / relative)
    module = importlib.util.module_from_spec(spec)
    sys.modules[name] = module
    spec.loader.exec_module(module)
    return module


class LocalToolsTests(unittest.TestCase):
    def test_resize_preserves_aspect_ratio_and_source(self):
        tool = load_tool("image_tools/img_resizing_v1.py")
        with tempfile.TemporaryDirectory() as directory:
            root = Path(directory)
            source = root / "input.png"
            output = root / "output"
            output.mkdir()
            Image.new("RGB", (120, 60), "red").save(source)
            before = source.read_bytes()
            success, _ = tool.resize_image(str(source), str(output), 60)
            self.assertTrue(success)
            with Image.open(output / source.name) as result:
                self.assertEqual(result.size, (60, 30))
            self.assertEqual(source.read_bytes(), before)

    def test_webp_transparency_becomes_white_jpeg_background(self):
        tool = load_tool("image_tools/convertIMG.py")
        with tempfile.TemporaryDirectory() as directory:
            source = Path(directory) / "input.webp"
            output = Path(directory) / "output.jpg"
            Image.new("RGBA", (20, 10), (255, 0, 0, 0)).save(source, "WEBP", lossless=True)
            tool.convert_one_webp_to_jpg(source, output, 90, "white")
            with Image.open(output) as result:
                self.assertEqual(result.mode, "RGB")
                self.assertTrue(all(value > 245 for value in result.getpixel((0, 0))))

    def test_compression_keeps_dimensions_and_reports_actual_bytes(self):
        tool = load_tool("image_tools/image_size_reducer_gui_percent.py")
        with tempfile.TemporaryDirectory() as directory:
            source = Path(directory) / "input.jpg"
            Image.new("RGB", (120, 80), "blue").save(source, quality=100)
            output, before, after = tool.compress_one(source, Path(directory) / "output", None, True, False)
            with Image.open(output) as result:
                self.assertEqual(result.size, (120, 80))
            self.assertEqual(before, source.stat().st_size)
            self.assertEqual(after, output.stat().st_size)
            self.assertEqual(tool.human(2048), "2.0KB")
            self.assertEqual(tool.human(3 * 1024 * 1024), "3.0MB")

    def test_gif_conversion_and_existing_output_protection(self):
        tool = load_tool("image_tools/gif_to_jpg_v1.py")
        with tempfile.TemporaryDirectory() as directory:
            source = Path(directory) / "input"
            output = Path(directory) / "output"
            source.mkdir()
            Image.new("RGB", (30, 20), "green").save(source / "test.GIF", "GIF")
            self.assertEqual(tool.convert_folder(source, output), 1)
            before = (output / "test.jpg").read_bytes()
            with self.assertRaises(FileExistsError):
                tool.convert_folder(source, output)
            self.assertEqual((output / "test.jpg").read_bytes(), before)

    def test_site_configuration_rejects_missing_and_invalid_values(self):
        with tempfile.TemporaryDirectory() as directory:
            config = Path(directory) / "settings.json"
            with patch.dict(os.environ, {"COMMERCE_CONFIG": str(config)}):
                for value in ["", "https://example.invalid", "https://user:password@shop.test", "https://:", "file:///tmp/test", "https://shop.test/path"]:
                    config.write_text(json.dumps({"godomall_base_url": value}), encoding="utf-8")
                    with self.assertRaises(RuntimeError):
                        commerce_config.configured_url("godomall_base_url")
                config.write_text(json.dumps({"godomall_base_url": "https://shop.test/"}), encoding="utf-8")
                self.assertEqual(commerce_config.configured_url("godomall_base_url"), "https://shop.test")

    def test_output_directory_uses_configured_location(self):
        with tempfile.TemporaryDirectory() as directory:
            config = Path(directory) / "settings.json"
            config.write_text(json.dumps({"output_directory": str(Path(directory) / "data")}), encoding="utf-8")
            with patch.dict(os.environ, {"COMMERCE_CONFIG": str(config)}):
                result = commerce_config.output_dir("downloads")
                self.assertEqual(result, Path(directory) / "data" / "downloads")
                self.assertTrue(result.is_dir())

    def test_price_comparison_module_loads_without_browser(self):
        with tempfile.TemporaryDirectory() as directory:
            config = Path(directory) / "settings.json"
            config.write_text(json.dumps({"godomall_base_url": "https://shop.test"}), encoding="utf-8")
            with patch.dict(os.environ, {"COMMERCE_CONFIG": str(config)}):
                tool = load_tool("inventory_tools/price_monitor.py")
                self.assertEqual(tool.to_int_price("12,300원"), 12300)
                self.assertEqual(tool.Config().goods_list_url, "https://shop.test/goods/goods_list.php")


if __name__ == "__main__":
    unittest.main()
