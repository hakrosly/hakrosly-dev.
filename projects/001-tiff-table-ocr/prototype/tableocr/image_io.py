"""TIFF (マルチページ対応) の読み込み。"""

from __future__ import annotations

from dataclasses import dataclass

import cv2
import numpy as np
from PIL import Image, ImageSequence

DEFAULT_DPI = 300


@dataclass
class Page:
    index: int  # 1 始まり
    gray: np.ndarray  # uint8 のグレースケール画像
    dpi: int


def _to_gray(frame: Image.Image) -> np.ndarray:
    if frame.mode in ("1", "L", "P", "RGB", "RGBA", "CMYK", "LA"):
        return np.array(frame.convert("L"))
    # 16bit などはそのまま配列化して 0-255 に正規化する
    arr = np.array(frame)
    if arr.ndim == 3:
        arr = arr[..., :3].mean(axis=2)
    return cv2.normalize(arr.astype(np.float32), None, 0, 255, cv2.NORM_MINMAX).astype(np.uint8)


def _dpi_of(frame: Image.Image) -> int:
    dpi = frame.info.get("dpi")
    if dpi:
        try:
            value = int(round(float(dpi[0])))
            if value >= 50:
                return value
        except (TypeError, ValueError):
            pass
    return DEFAULT_DPI


def load_pages(path: str) -> list[Page]:
    pages = []
    with Image.open(path) as im:
        for i, frame in enumerate(ImageSequence.Iterator(im), start=1):
            pages.append(Page(index=i, gray=_to_gray(frame), dpi=_dpi_of(frame)))
    if not pages:
        raise ValueError(f"画像を読み込めませんでした: {path}")
    return pages
