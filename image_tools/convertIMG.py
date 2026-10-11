#!/usr/bin/env python3
# -*- coding: utf-8 -*-

"""
webp_to_jpg_gui.py
폴더 단위로 WebP -> JPG 변환하는 GUI 도구
- 입력: WebP(.webp)
- 출력: JPG(.jpg)
- 원본 파일은 수정/삭제하지 않음
- 투명(알파) 포함 WebP는 JPG 특성상 배경색(흰/검)으로 합성
- (선택) 하위 폴더까지 재귀 처리 + 폴더 구조 유지
"""

import os
from pathlib import Path
from typing import List, Tuple, Optional

from PIL import Image
import tkinter as tk
from tkinter import filedialog, messagebox


SUPPORTED_EXTS = (".webp",)


def safe_int(s: str, default: int) -> int:
    try:
        return int(str(s).strip())
    except Exception:
        return default


def convert_one_webp_to_jpg(
    src: Path,
    dst: Path,
    quality: int,
    background: str = "white",  # "white" | "black"
) -> None:
    """
    src webp를 dst jpg로 저장.
    알파(투명) 존재 시 배경색 합성.
    """
    with Image.open(src) as im:
        im.load()

        # JPG는 알파 지원 X → 투명 있으면 배경 합성
        has_alpha = (
            im.mode in ("RGBA", "LA")
            or (im.mode == "P" and "transparency" in im.info)
        )

        if has_alpha:
            rgba = im.convert("RGBA")
            bg_color = (255, 255, 255) if background == "white" else (0, 0, 0)
            bg = Image.new("RGB", rgba.size, bg_color)
            bg.paste(rgba, mask=rgba.split()[-1])
            out = bg
        else:
            out = im.convert("RGB")

        dst.parent.mkdir(parents=True, exist_ok=True)
        out.save(
            dst,
            "JPEG",
            quality=int(quality),
            optimize=True,
            progressive=True,
        )


def collect_webp_files(input_dir: Path, recursive: bool) -> List[Path]:
    pattern = "**/*.webp" if recursive else "*.webp"
    return [p for p in input_dir.glob(pattern) if p.is_file()]


def build_output_path(
    src: Path,
    input_dir: Path,
    output_dir: Path,
    keep_structure: bool,
) -> Path:
    """
    keep_structure=True이면 input_dir 기준 상대경로를 유지한 채 확장자만 .jpg로 변경.
    False이면 output_dir 바로 아래에 파일명만 저장(중복 위험 존재).
    """
    if keep_structure:
        rel = src.relative_to(input_dir)
        return (output_dir / rel).with_suffix(".jpg")
    else:
        return (output_dir / src.name).with_suffix(".jpg")


class App(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("WEBP → JPG 변환기")
        self.geometry("700x360")

        self.input_dir = tk.StringVar()
        self.output_dir = tk.StringVar()

        self.quality = tk.StringVar(value="92")
        self.recursive = tk.BooleanVar(value=True)
        self.keep_structure = tk.BooleanVar(value=True)

        self.background = tk.StringVar(value="white")  # white | black

        self.create_widgets()

    def create_widgets(self):
        pad = {"padx": 10, "pady": 6}

        # 입력 폴더
        tk.Label(self, text="입력 폴더(WebP 위치):").grid(row=0, column=0, sticky="e", **pad)
        tk.Entry(self, textvariable=self.input_dir, width=55).grid(row=0, column=1, **pad)
        tk.Button(self, text="선택", command=self.choose_input).grid(row=0, column=2, **pad)

        # 출력 폴더
        tk.Label(self, text="출력 폴더(JPG 저장):").grid(row=1, column=0, sticky="e", **pad)
        tk.Entry(self, textvariable=self.output_dir, width=55).grid(row=1, column=1, **pad)
        tk.Button(self, text="선택", command=self.choose_output).grid(row=1, column=2, **pad)

        # 품질
        tk.Label(self, text="JPG 품질(1~100):").grid(row=2, column=0, sticky="e", **pad)
        tk.Entry(self, textvariable=self.quality, width=10).grid(row=2, column=1, sticky="w", **pad)
        tk.Label(self, text="예: 92 (높을수록 화질↑ 용량↑)").grid(row=2, column=1, sticky="e", **pad)

        # 옵션 프레임
        opt = tk.LabelFrame(self, text="옵션")
        opt.grid(row=3, column=0, columnspan=3, sticky="we", padx=10, pady=6)

        tk.Checkbutton(opt, text="하위 폴더까지 포함(재귀)", variable=self.recursive).grid(
            row=0, column=0, sticky="w", padx=10, pady=6
        )
        tk.Checkbutton(opt, text="폴더 구조 유지(권장)", variable=self.keep_structure).grid(
            row=0, column=1, sticky="w", padx=10, pady=6
        )

        bg_frame = tk.Frame(opt)
        bg_frame.grid(row=1, column=0, columnspan=3, sticky="w", padx=10, pady=6)
        tk.Label(bg_frame, text="투명 배경 처리:").pack(side="left")
        tk.Radiobutton(bg_frame, text="흰색", variable=self.background, value="white").pack(side="left", padx=8)
        tk.Radiobutton(bg_frame, text="검정", variable=self.background, value="black").pack(side="left", padx=8)

        # 실행 버튼
        tk.Button(self, text="변환 실행", command=self.run).grid(row=4, column=0, columnspan=3, pady=14)

        # 도움말
        help_text = (
            "• 입력 폴더의 .webp만 .jpg로 변환합니다(원본 webp는 유지).\n"
            "• JPG는 투명(알파)을 지원하지 않아, 투명 영역은 선택한 배경색으로 합성됩니다.\n"
            "• 재귀 ON + 폴더 구조 유지 ON이면 하위 폴더 구조를 그대로 출력 폴더에 재현합니다.\n"
            "• 입력/출력 폴더를 동일하게 지정해도 동작은 가능하지만(원본 유지), 파일이 섞여 관리가 어려울 수 있습니다."
        )
        tk.Label(self, text=help_text, justify="left", fg="#444").grid(
            row=5, column=0, columnspan=3, padx=12, pady=4, sticky="w"
        )

    def choose_input(self):
        d = filedialog.askdirectory(title="입력 폴더 선택(WebP 폴더)")
        if d:
            self.input_dir.set(d)

    def choose_output(self):
        d = filedialog.askdirectory(title="출력 폴더 선택(JPG 저장 폴더)")
        if d:
            self.output_dir.set(d)

    def run(self):
        in_dir = self.input_dir.get().strip()
        out_dir = self.output_dir.get().strip()

        if not in_dir or not os.path.isdir(in_dir):
            messagebox.showerror("오류", "유효한 입력 폴더를 선택하세요.")
            return
        if not out_dir:
            messagebox.showerror("오류", "출력 폴더를 선택하세요.")
            return

        input_dir = Path(in_dir)
        output_dir = Path(out_dir)

        # 품질 검증
        q = safe_int(self.quality.get(), 92)
        if q < 1 or q > 100:
            messagebox.showerror("입력 오류", "JPG 품질은 1~100 사이 정수여야 합니다.")
            return

        # 입력/출력 동일 경고
        try:
            if input_dir.resolve() == output_dir.resolve():
                ok = messagebox.askyesno(
                    "주의",
                    "입력 폴더와 출력 폴더가 같습니다.\n"
                    "원본 WEBP는 삭제되지 않고 JPG만 추가 생성됩니다.\n"
                    "같은 폴더에 생성하시겠습니까?"
                )
                if not ok:
                    return
        except Exception:
            # resolve 실패 시에도 진행은 가능
            pass

        files = collect_webp_files(input_dir, self.recursive.get())
        if not files:
            messagebox.showerror("오류", "입력 폴더에 .webp 파일이 없습니다.")
            return

        output_dir.mkdir(parents=True, exist_ok=True)

        ok_cnt = 0
        fail_cnt = 0
        failed_samples: List[Tuple[str, str]] = []

        for src in files:
            try:
                dst = build_output_path(
                    src=src,
                    input_dir=input_dir,
                    output_dir=output_dir,
                    keep_structure=self.keep_structure.get(),
                )
                convert_one_webp_to_jpg(
                    src=src,
                    dst=dst,
                    quality=q,
                    background=self.background.get(),
                )
                ok_cnt += 1
            except Exception as e:
                fail_cnt += 1
                if len(failed_samples) < 5:
                    failed_samples.append((str(src), str(e)))

        msg = f"완료: {ok_cnt}개, 실패: {fail_cnt}개"
        if failed_samples:
            msg += "\n\n실패 예시(최대 5개):"
            for p, err in failed_samples:
                msg += f"\n- {p}\n  {err}"

        if fail_cnt > 0:
            messagebox.showwarning("변환 결과", msg)
        else:
            messagebox.showinfo("변환 결과", msg)


if __name__ == "__main__":
    App().mainloop()
