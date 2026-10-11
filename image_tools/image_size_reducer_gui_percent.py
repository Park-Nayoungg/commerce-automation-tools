#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
image_size_reducer_gui_percent.py
폴더 단위로 이미지 '용량만' 줄이는 GUI 도구 (리사이즈 없음, 퍼센트 기반)
- 입력: JPG/JPEG/PNG/WebP
- 출력: 원본 포맷 유지 또는 WebP로 변환 선택 가능
- 메타데이터 제거(최대한)
- 목표 비율(%)로: JPEG/WebP는 이진 탐색으로 원본 대비 목표 바이트에 가깝게 저장
- PNG는 optimize + (선택) 팔레트 축소로 용량 축소(정확히 목표치로 맞추기는 어려움)
- 픽셀 크기(가로x세로) 변경 없음

필요 패키지:
  pip install pillow
"""

import os
import io
from pathlib import Path
from typing import Optional, Tuple

from PIL import Image
import tkinter as tk
from tkinter import filedialog, messagebox

SUPPORTED_EXTS = ('.jpg', '.jpeg', '.png', '.webp')

def human(n: int) -> str:
    for unit in ['B','KB','MB','GB']:
        if n < 1024 or unit == 'GB':
            return f"{n:.0f}{unit}" if unit=='B' else f"{n:.1f}{unit}"
        n /= 1024

def save_jpeg(im: Image.Image, out_path: Path, quality: int, progressive=True, optimize=True, subsampling="4:2:0") -> int:
    if im.mode not in ('RGB', 'L'):
        im = im.convert('RGB')
    buf = io.BytesIO()
    im.save(buf, format='JPEG', quality=int(quality), progressive=progressive, optimize=optimize, subsampling=subsampling)
    data = buf.getvalue()
    out_path.write_bytes(data)
    return len(data)

def save_webp(im: Image.Image, out_path: Path, quality: int, lossless=False, method=6) -> int:
    buf = io.BytesIO()
    im.save(buf, format='WEBP', quality=int(quality), lossless=bool(lossless), method=int(method))
    data = buf.getvalue()
    out_path.write_bytes(data)
    return len(data)

def save_png(im: Image.Image, out_path: Path, optimize=True, palette_quantize=True, colors=256, lossless_only=False) -> int:
    im2 = im
    if not lossless_only:
        # 팔레트 축소(근사 무손실). 알파가 없는 경우에만 효과가 큼.
        if im.mode not in ('P',) and im.mode != 'RGBA':
            try:
                im2 = im.convert('P', palette=Image.Palette.ADAPTIVE, colors=colors)
            except Exception:
                im2 = im
    buf = io.BytesIO()
    im2.save(buf, format='PNG', optimize=optimize)
    data = buf.getvalue()
    out_path.write_bytes(data)
    return len(data)

def binary_search_quality(save_func, im: Image.Image, out_path: Path, target_bytes: int,
                          qmin=40, qmax=95, guess=82) -> Tuple[int, int]:
    """
    품질을 이진탐색해서 target_bytes 이하로 가깝게 맞춤.
    반환: (최종품질, 최종파일크기)
    """
    best_q = None
    best_size = None

    # 상/하한 빠른 체크
    size_hi = save_func(im, out_path, qmax)
    if size_hi <= target_bytes:
        return qmax, size_hi
    size_lo = save_func(im, out_path, qmin)
    if size_lo > target_bytes:
        # 최소 품질로도 목표보다 큼
        return qmin, size_lo

    lo, hi = qmin, qmax
    while lo <= hi:
        q = (lo + hi) // 2
        size = save_func(im, out_path, q)
        if size <= target_bytes:
            best_q, best_size = q, size
            lo = q + 1
        else:
            hi = q - 1

    if best_q is not None:
        save_func(im, out_path, best_q)  # best 결과로 저장
        return best_q, best_size

    return q, size  # fallback

def compress_one(in_path: Path, out_dir: Path, target_bytes: Optional[int],
                 keep_format: bool, to_webp: bool,
                 jpeg_min_quality=60, jpeg_max_quality=90,
                 webp_min_quality=60, webp_max_quality=90,
                 png_lossless_only=False) -> Tuple[Path, int, int]:
    """
    단일 이미지 처리. 픽셀 크기 변경 없음.
    """
    out_dir.mkdir(parents=True, exist_ok=True)
    ext = in_path.suffix.lower()
    with Image.open(in_path) as im:
        im.load()        # 데이터 로드
        im = im.copy()   # 메타/프로필 최대 제거 효과

        orig_size = in_path.stat().st_size

        # 출력 포맷 결정
        if to_webp:
            fmt = 'webp'
        else:
            if keep_format:
                fmt = ext.lstrip('.')
                if fmt == 'jpg':
                    fmt = 'jpeg'
            else:
                fmt = 'jpeg' if ext in ('.jpg', '.jpeg') else 'png'

        out_name = in_path.stem + ('.webp' if fmt=='webp' else '.jpg' if fmt=='jpeg' else '.png')
        out_path = out_dir / out_name

        # 포맷별 저장 로직
        if fmt == 'jpeg':
            if target_bytes:
                q, final = binary_search_quality(lambda img, p, q: save_jpeg(img, p, q, progressive=True, optimize=True, subsampling="4:2:0"),
                                                 im, out_path, target_bytes,
                                                 qmin=jpeg_min_quality, qmax=jpeg_max_quality, guess=82)
            else:
                q = jpeg_max_quality
                final = save_jpeg(im, out_path, q, progressive=True, optimize=True, subsampling="4:2:0")
        elif fmt == 'webp':
            if target_bytes:
                q, final = binary_search_quality(lambda img, p, q: save_webp(img, p, q, lossless=False, method=6),
                                                 im, out_path, target_bytes,
                                                 qmin=webp_min_quality, qmax=webp_max_quality, guess=82)
            else:
                q = webp_max_quality
                final = save_webp(im, out_path, q, lossless=False, method=6)
        elif fmt == 'png':
            # PNG는 목표 바이트 정확히 맞추기 어려움 -> 최적화 중심
            final = save_png(im, out_path, optimize=True, palette_quantize=not png_lossless_only,
                             colors=256, lossless_only=png_lossless_only)
        else:
            raise ValueError("지원하지 않는 포맷")

        return out_path, orig_size, final

class App(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("이미지 용량 줄이기 (퍼센트, 리사이즈 없음)")
        self.geometry("660x330")

        self.input_dir = tk.StringVar()
        self.output_dir = tk.StringVar()
        self.target_percent = tk.StringVar(value="70")  # 기본 70%
        self.keep_format = tk.BooleanVar(value=True)
        self.to_webp = tk.BooleanVar(value=False)
        self.png_lossless_only = tk.BooleanVar(value=False)

        self.create_widgets()

    def create_widgets(self):
        pad = {'padx':10, 'pady':6}

        # 입력 폴더
        tk.Label(self, text="입력 폴더:").grid(row=0, column=0, sticky="e", **pad)
        tk.Entry(self, textvariable=self.input_dir, width=55).grid(row=0, column=1, **pad)
        tk.Button(self, text="선택", command=self.choose_input).grid(row=0, column=2, **pad)

        # 출력 폴더
        tk.Label(self, text="출력 폴더:").grid(row=1, column=0, sticky="e", **pad)
        tk.Entry(self, textvariable=self.output_dir, width=55).grid(row=1, column=1, **pad)
        tk.Button(self, text="선택", command=self.choose_output).grid(row=1, column=2, **pad)

        # 목표 비율(%)
        tk.Label(self, text="목표 용량 (% of 원본):").grid(row=2, column=0, sticky="e", **pad)
        tk.Entry(self, textvariable=self.target_percent, width=10).grid(row=2, column=1, sticky="w", **pad)
        tk.Label(self, text="예: 70 → 원본 파일 크기의 70%로 압축 (JPEG/WebP)").grid(row=2, column=1, sticky="e", **pad)

        # 포맷 옵션
        fmt_frame = tk.LabelFrame(self, text="출력 포맷 옵션")
        fmt_frame.grid(row=3, column=0, columnspan=3, sticky="we", padx=10, pady=6)
        tk.Checkbutton(fmt_frame, text="원본 포맷 유지", variable=self.keep_format,
                       command=self._on_keep_format_change).grid(row=0, column=0, sticky="w", padx=10, pady=6)
        tk.Checkbutton(fmt_frame, text="모두 WebP로 변환(권장)", variable=self.to_webp,
                       command=self._on_webp_change).grid(row=0, column=1, sticky="w", padx=10, pady=6)
        tk.Checkbutton(fmt_frame, text="PNG 완전 무손실(최적화만)", variable=self.png_lossless_only).grid(row=0, column=2, sticky="w", padx=10, pady=6)

        # 실행 버튼
        tk.Button(self, text="압축 실행", command=self.run).grid(row=4, column=0, columnspan=3, pady=14)

        # 도움말
        help_text = (
            "• 리사이즈 없음(픽셀 크기 유지). 메타데이터 제거.\n"
            "• JPEG/WebP: 원본 크기의 N% 목표로 품질 자동 조절(이진탐색).\n"
            "• PNG: 무손실 최적화 또는 팔레트 축소로 용량 절감(정확한 % 맞춤은 어려움).\n"
            "• '모두 WebP'는 동일 용량 대비 화질이 유리하고, 대체로 더 작은 파일이 됩니다."
        )
        tk.Label(self, text=help_text, justify="left", fg="#444").grid(row=5, column=0, columnspan=3, padx=12, pady=4, sticky="w")

    def _on_keep_format_change(self):
        if self.keep_format.get():
            self.to_webp.set(False)

    def _on_webp_change(self):
        if self.to_webp.get():
            self.keep_format.set(False)

    def choose_input(self):
        d = filedialog.askdirectory(title="이미지 폴더 선택")
        if d:
            self.input_dir.set(d)

    def choose_output(self):
        d = filedialog.askdirectory(title="저장 폴더 선택")
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

        # 퍼센트 입력 확인
        if self.target_percent.get().strip():
            try:
                percent = int(self.target_percent.get().strip())
                if percent <= 0 or percent > 100:
                    raise ValueError
            except Exception:
                messagebox.showerror("입력 오류", "목표 비율(%)은 1~100 사이 정수여야 합니다.")
                return
        else:
            percent = None  # 비율 미지정 시, 품질 상한으로만 저장

        files = [f for f in os.listdir(in_dir) if f.lower().endswith(SUPPORTED_EXTS)]
        if not files:
            messagebox.showerror("오류", "지원되는 이미지가 없습니다 (jpg, jpeg, png, webp).")
            return

        os.makedirs(out_dir, exist_ok=True)

        ok, fail = 0, 0
        total_before = 0
        total_after = 0
        for name in files:
            in_path = Path(in_dir) / name
            try:
                # 원본 크기와 목표 바이트 계산 (JPEG/WebP에서만 정확히 반영)
                target_bytes = None
                if percent is not None:
                    orig_size = in_path.stat().st_size
                    target_bytes = int(orig_size * percent / 100)

                out_path, before, after = compress_one(
                    in_path,
                    Path(out_dir),
                    target_bytes=target_bytes,
                    keep_format=self.keep_format.get(),
                    to_webp=self.to_webp.get(),
                    jpeg_min_quality=60, jpeg_max_quality=90,
                    webp_min_quality=60, webp_max_quality=90,
                    png_lossless_only=self.png_lossless_only.get()
                )
                ok += 1
                total_before += before
                total_after += after
            except Exception as e:
                fail += 1

        msg = f"완료: {ok}개, 실패: {fail}개\n"
        if total_before:
            msg += f"총 용량: {human(total_before)} → {human(total_after)}  ({(total_after/total_before)*100:.1f}% of original)"
        messagebox.showinfo("작업 결과", msg)

if __name__ == "__main__":
    App().mainloop()
