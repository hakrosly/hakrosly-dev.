"""OpenCV で罫線を検出し、表とセルの位置を求める。"""

from __future__ import annotations

from dataclasses import dataclass, field

import cv2
import numpy as np


@dataclass
class Cell:
    row: int
    col: int
    row_span: int
    col_span: int
    x0: int
    y0: int
    x1: int
    y1: int


@dataclass
class Table:
    page: int
    x: int
    y: int
    w: int
    h: int
    row_lines: list[int]  # ページ座標の y
    col_lines: list[int]  # ページ座標の x
    cells: list[Cell] = field(default_factory=list)

    @property
    def n_rows(self) -> int:
        return len(self.row_lines) - 1

    @property
    def n_cols(self) -> int:
        return len(self.col_lines) - 1


@dataclass
class LineMasks:
    binary: np.ndarray  # 文字・罫線 = 255
    horizontal: np.ndarray
    vertical: np.ndarray


def binarize(gray: np.ndarray) -> np.ndarray:
    blur = cv2.GaussianBlur(gray, (3, 3), 0)
    return cv2.adaptiveThreshold(
        blur, 255, cv2.ADAPTIVE_THRESH_MEAN_C, cv2.THRESH_BINARY_INV, 31, 15
    )


def estimate_skew(gray: np.ndarray, max_angle: float = 5.0) -> float:
    """横罫線の傾き (度) を推定する。見つからなければ 0。"""
    binary = binarize(gray)
    h, w = binary.shape
    kernel = cv2.getStructuringElement(cv2.MORPH_RECT, (max(20, w // 50), 1))
    horiz = cv2.morphologyEx(binary, cv2.MORPH_OPEN, kernel)
    lines = cv2.HoughLinesP(
        horiz, 1, np.pi / 1800, threshold=100, minLineLength=max(50, w // 8), maxLineGap=20
    )
    if lines is None:
        return 0.0
    angles = []
    for x1, y1, x2, y2 in lines.reshape(-1, 4):
        angle = np.degrees(np.arctan2(y2 - y1, x2 - x1))
        if abs(angle) <= max_angle:
            angles.append(angle)
    if not angles:
        return 0.0
    return float(np.median(angles))


def deskew(gray: np.ndarray, min_angle: float = 0.05) -> tuple[np.ndarray, float]:
    angle = estimate_skew(gray)
    if abs(angle) < min_angle:
        return gray, 0.0
    h, w = gray.shape
    matrix = cv2.getRotationMatrix2D((w / 2, h / 2), angle, 1.0)
    rotated = cv2.warpAffine(
        gray, matrix, (w, h), flags=cv2.INTER_LINEAR, borderMode=cv2.BORDER_CONSTANT, borderValue=255
    )
    return rotated, angle


def _keep_touching(mask: np.ndarray, other: np.ndarray) -> np.ndarray:
    """mask の連結成分のうち、other (膨張済み) に接するものだけ残す。

    文字の縦棒・横棒は罫線に接しないので、これで取り除ける。
    """
    n, labels = cv2.connectedComponents(mask)
    if n <= 1:
        return mask
    touching = np.unique(labels[(other > 0) & (labels > 0)])
    keep = np.zeros(n, dtype=bool)
    keep[touching] = True
    return np.where(keep[labels], 255, 0).astype(np.uint8)


def line_masks(gray: np.ndarray, dpi: int) -> LineMasks:
    binary = binarize(gray)
    h, w = binary.shape
    # 罫線とみなす最小の長さ。0.15 インチ (300dpi で 45px) を下限にする
    min_len = max(15, int(dpi * 0.15))
    h_len = max(min_len, w // 60)
    v_len = min_len

    horiz = cv2.morphologyEx(
        binary, cv2.MORPH_OPEN, cv2.getStructuringElement(cv2.MORPH_RECT, (h_len, 1))
    )
    vert = cv2.morphologyEx(
        binary, cv2.MORPH_OPEN, cv2.getStructuringElement(cv2.MORPH_RECT, (1, v_len))
    )
    # かすれた罫線の途切れを少しつなぐ
    gap = max(3, dpi // 100)
    horiz = cv2.morphologyEx(
        horiz, cv2.MORPH_CLOSE, cv2.getStructuringElement(cv2.MORPH_RECT, (gap * 2 + 1, 1))
    )
    vert = cv2.morphologyEx(
        vert, cv2.MORPH_CLOSE, cv2.getStructuringElement(cv2.MORPH_RECT, (1, gap * 2 + 1))
    )

    grow = cv2.getStructuringElement(cv2.MORPH_RECT, (gap * 2 + 1, gap * 2 + 1))
    horiz_f = _keep_touching(horiz, cv2.dilate(vert, grow))
    vert_f = _keep_touching(vert, cv2.dilate(horiz, grow))
    return LineMasks(binary=binary, horizontal=horiz_f, vertical=vert_f)


def _line_positions(profile: np.ndarray, min_count: int, merge_gap: int) -> list[int]:
    """射影プロファイルから罫線の中心位置を求める。"""
    idx = np.flatnonzero(profile >= min_count)
    if idx.size == 0:
        return []
    groups = [[idx[0]]]
    for i in idx[1:]:
        if i - groups[-1][-1] <= merge_gap:
            groups[-1].append(i)
        else:
            groups.append([i])
    return [int(round((g[0] + g[-1]) / 2)) for g in groups]


def _coverage(segment: np.ndarray) -> float:
    """罫線が通っている割合 (segment の各行に 1 画素でも線があれば通っているとみなす)。"""
    if segment.size == 0:
        return 0.0
    return float(np.count_nonzero(segment.any(axis=1))) / len(segment)


def _build_cells(
    masks: LineMasks, rows: list[int], cols: list[int], tol: int
) -> list[Cell]:
    n_rows, n_cols = len(rows) - 1, len(cols) - 1
    parent = list(range(n_rows * n_cols))

    def find(a: int) -> int:
        while parent[a] != a:
            parent[a] = parent[parent[a]]
            a = parent[a]
        return a

    def union(a: int, b: int) -> None:
        ra, rb = find(a), find(b)
        if ra != rb:
            parent[max(ra, rb)] = min(ra, rb)

    vert, horiz = masks.vertical, masks.horizontal
    for r in range(n_rows):
        y0, y1 = rows[r] + tol, rows[r + 1] - tol
        for c in range(n_cols - 1):
            x = cols[c + 1]
            seg = vert[max(y0, 0):max(y1, 0), max(x - tol, 0):x + tol + 1]
            if _coverage(seg) < 0.5:
                union(r * n_cols + c, r * n_cols + c + 1)
    for c in range(n_cols):
        x0, x1 = cols[c] + tol, cols[c + 1] - tol
        for r in range(n_rows - 1):
            y = rows[r + 1]
            seg = horiz[max(y - tol, 0):y + tol + 1, max(x0, 0):max(x1, 0)].T
            if _coverage(seg) < 0.5:
                union(r * n_cols + c, (r + 1) * n_cols + c)

    groups: dict[int, list[tuple[int, int]]] = {}
    for r in range(n_rows):
        for c in range(n_cols):
            groups.setdefault(find(r * n_cols + c), []).append((r, c))

    cells = []
    for members in groups.values():
        r0 = min(m[0] for m in members)
        r1 = max(m[0] for m in members)
        c0 = min(m[1] for m in members)
        c1 = max(m[1] for m in members)
        cells.append(
            Cell(
                row=r0,
                col=c0,
                row_span=r1 - r0 + 1,
                col_span=c1 - c0 + 1,
                x0=cols[c0],
                y0=rows[r0],
                x1=cols[c1 + 1],
                y1=rows[r1 + 1],
            )
        )
    cells.sort(key=lambda cell: (cell.row, cell.col))
    return cells


def detect_tables(
    gray: np.ndarray,
    dpi: int,
    page: int = 1,
    roi: tuple[int, int, int, int] | None = None,
    masks: LineMasks | None = None,
) -> tuple[list[Table], LineMasks]:
    """ページ内の罫線付きの表を上から順に返す。"""
    if masks is None:
        masks = line_masks(gray, dpi)
    grid = cv2.bitwise_or(masks.horizontal, masks.vertical)
    if roi is not None:
        rx, ry, rw, rh = roi
        limited = np.zeros_like(grid)
        limited[ry:ry + rh, rx:rx + rw] = grid[ry:ry + rh, rx:rx + rw]
        grid = limited

    tol = max(3, dpi // 60)
    grid_d = cv2.dilate(grid, cv2.getStructuringElement(cv2.MORPH_RECT, (tol, tol)))
    contours, _ = cv2.findContours(grid_d, cv2.RETR_EXTERNAL, cv2.CHAIN_APPROX_SIMPLE)

    min_side = max(20, int(dpi * 0.2))
    tables = []
    for contour in contours:
        x, y, w, h = cv2.boundingRect(contour)
        if w < min_side or h < min_side:
            continue
        h_roi = masks.horizontal[y:y + h, x:x + w]
        v_roi = masks.vertical[y:y + h, x:x + w]
        min_len = max(15, int(dpi * 0.15))
        rows = _line_positions(np.count_nonzero(h_roi, axis=1), min_len, tol)
        cols = _line_positions(np.count_nonzero(v_roi, axis=0), min_len, tol)
        # 外枠が無い表でも端を区切りとして扱う
        if not rows or rows[0] > tol * 2:
            rows.insert(0, 0)
        if rows[-1] < h - 1 - tol * 2:
            rows.append(h - 1)
        if not cols or cols[0] > tol * 2:
            cols.insert(0, 0)
        if cols[-1] < w - 1 - tol * 2:
            cols.append(w - 1)
        if len(rows) < 2 or len(cols) < 2:
            continue
        if len(rows) == 2 and len(cols) == 2:
            continue  # ただの枠線は表とみなさない
        rows = [y + r for r in rows]
        cols = [x + c for c in cols]
        table = Table(page=page, x=x, y=y, w=w, h=h, row_lines=rows, col_lines=cols)
        table.cells = _build_cells(masks, rows, cols, tol)
        tables.append(table)

    # 読む順番 (上から下、左から右) に並べる
    line_h = max(10, dpi // 10)
    tables.sort(key=lambda t: (t.y // line_h, t.x))
    return tables, masks


def draw_tables(gray: np.ndarray, tables: list[Table], numbers: list[int] | None = None) -> np.ndarray:
    """確認用に表とセルを色付きで描いた画像を返す。"""
    out = cv2.cvtColor(gray, cv2.COLOR_GRAY2BGR)
    for i, table in enumerate(tables):
        for cell in table.cells:
            merged = cell.row_span > 1 or cell.col_span > 1
            color = (0, 140, 255) if merged else (0, 180, 0)
            cv2.rectangle(out, (cell.x0 + 2, cell.y0 + 2), (cell.x1 - 2, cell.y1 - 2), color, 2)
        cv2.rectangle(out, (table.x, table.y), (table.x + table.w, table.y + table.h), (255, 0, 0), 3)
        label = f"#{numbers[i] if numbers else i + 1}"
        cv2.putText(
            out, label, (table.x, max(30, table.y - 10)), cv2.FONT_HERSHEY_SIMPLEX, 1.2, (255, 0, 0), 3
        )
    return out
