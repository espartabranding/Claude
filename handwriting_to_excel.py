#!/usr/bin/env python3
"""
Handwriting Recognition to Excel Tool - v2
============================================
High-capability tool for recognizing messy/disorganized handwriting
from JPG images and exporting structured data to Excel spreadsheets.

v2: Grid-aware engine that detects table lines first, crops each cell
individually, and runs OCR per-cell for dramatically better accuracy
on structured forms and handwritten tables.

Usage:
    python handwriting_to_excel.py input.jpg -o output.xlsx
    python handwriting_to_excel.py input.jpg --engine both --enhance aggressive
"""

import argparse
import sys
import os
import re
import logging
from pathlib import Path
from dataclasses import dataclass

import cv2
import numpy as np
from PIL import Image, ImageEnhance, ImageFilter
import openpyxl
from openpyxl.styles import Font, Alignment, Border, Side, PatternFill

logging.basicConfig(level=logging.INFO, format="%(levelname)s: %(message)s")
logger = logging.getLogger(__name__)


# ---------------------------------------------------------------------------
# Data structures
# ---------------------------------------------------------------------------

@dataclass
class TextBlock:
    """Represents a detected text region with position and content."""
    text: str
    x: int
    y: int
    width: int
    height: int
    confidence: float = 0.0

    @property
    def center_x(self):
        return self.x + self.width // 2

    @property
    def center_y(self):
        return self.y + self.height // 2

    @property
    def right(self):
        return self.x + self.width

    @property
    def bottom(self):
        return self.y + self.height


@dataclass
class CellBBox:
    """Bounding box of a single table cell."""
    row: int
    col: int
    x: int
    y: int
    w: int
    h: int


# ---------------------------------------------------------------------------
# Image Preprocessor
# ---------------------------------------------------------------------------

class ImagePreprocessor:
    """Advanced image preprocessing for messy handwriting recognition."""

    def __init__(self, enhance_level="standard"):
        self.enhance_level = enhance_level

    def load_image(self, image_path: str) -> np.ndarray:
        """Load image from file path."""
        img = cv2.imread(image_path)
        if img is None:
            raise FileNotFoundError(f"Could not load image: {image_path}")
        logger.info(f"Loaded image: {img.shape[1]}x{img.shape[0]} pixels")
        return img

    def preprocess(self, img: np.ndarray) -> np.ndarray:
        """Full preprocessing pipeline."""
        img = self._normalize_size(img)
        gray = cv2.cvtColor(img, cv2.COLOR_BGR2GRAY)
        gray = self._deskew(gray)
        gray = self._denoise(gray)
        gray = self._enhance_contrast(gray)
        binary = self._binarize(gray)
        binary = self._morphological_cleanup(binary)
        return binary

    def preprocess_for_ocr(self, img: np.ndarray) -> list[np.ndarray]:
        """Generate multiple preprocessed versions for OCR ensemble."""
        versions = []
        versions.append(self.preprocess(img))
        gray = cv2.cvtColor(img, cv2.COLOR_BGR2GRAY)
        gray = self._deskew(gray)
        gray = self._denoise(gray)
        clahe = cv2.createCLAHE(clipLimit=3.0, tileGridSize=(8, 8))
        versions.append(clahe.apply(gray))
        versions.append(cv2.bitwise_not(versions[0]))
        if self.enhance_level == "aggressive":
            kernel = np.ones((2, 2), np.uint8)
            dilated = cv2.dilate(versions[0], kernel, iterations=1)
            versions.append(dilated)
        return versions

    def preprocess_cell(self, cell_img: np.ndarray) -> np.ndarray:
        """Lightweight preprocessing for an individual cell crop."""
        if len(cell_img.shape) == 3:
            gray = cv2.cvtColor(cell_img, cv2.COLOR_BGR2GRAY)
        else:
            gray = cell_img.copy()

        # Upscale small cells for better OCR
        h, w = gray.shape
        if max(h, w) < 100:
            scale = 100 / max(h, w)
            gray = cv2.resize(gray, None, fx=scale, fy=scale,
                              interpolation=cv2.INTER_CUBIC)

        # Denoise
        gray = cv2.fastNlMeansDenoising(gray, h=8, templateWindowSize=7,
                                         searchWindowSize=21)
        # Contrast
        clahe = cv2.createCLAHE(clipLimit=3.0, tileGridSize=(4, 4))
        gray = clahe.apply(gray)

        # Binarize
        binary = cv2.adaptiveThreshold(
            gray, 255, cv2.ADAPTIVE_THRESH_GAUSSIAN_C,
            cv2.THRESH_BINARY, 15, 8
        )
        return binary

    def _normalize_size(self, img: np.ndarray) -> np.ndarray:
        h, w = img.shape[:2]
        if max(h, w) > 4000:
            scale = 4000 / max(h, w)
            img = cv2.resize(img, None, fx=scale, fy=scale,
                             interpolation=cv2.INTER_AREA)
        elif max(h, w) < 1000:
            scale = 1000 / max(h, w)
            img = cv2.resize(img, None, fx=scale, fy=scale,
                             interpolation=cv2.INTER_CUBIC)
        return img

    def _deskew(self, gray: np.ndarray) -> np.ndarray:
        edges = cv2.Canny(gray, 50, 150, apertureSize=3)
        lines = cv2.HoughLinesP(edges, 1, np.pi / 180, threshold=100,
                                minLineLength=gray.shape[1] // 8, maxLineGap=20)
        if lines is None or len(lines) < 3:
            return gray
        angles = []
        for line in lines:
            x1, y1, x2, y2 = line[0]
            angle = np.degrees(np.arctan2(y2 - y1, x2 - x1))
            if abs(angle) < 30:
                angles.append(angle)
        if not angles:
            return gray
        median_angle = np.median(angles)
        if abs(median_angle) < 0.5:
            return gray
        logger.info(f"Deskewing by {median_angle:.1f} degrees")
        h, w = gray.shape
        center = (w // 2, h // 2)
        matrix = cv2.getRotationMatrix2D(center, median_angle, 1.0)
        return cv2.warpAffine(gray, matrix, (w, h), flags=cv2.INTER_CUBIC,
                              borderMode=cv2.BORDER_REPLICATE)

    def _denoise(self, gray: np.ndarray) -> np.ndarray:
        h_val = 15 if self.enhance_level == "aggressive" else 10
        return cv2.fastNlMeansDenoising(gray, h=h_val, templateWindowSize=7,
                                         searchWindowSize=21)

    def _enhance_contrast(self, gray: np.ndarray) -> np.ndarray:
        if self.enhance_level == "aggressive":
            clahe = cv2.createCLAHE(clipLimit=4.0, tileGridSize=(4, 4))
        else:
            clahe = cv2.createCLAHE(clipLimit=2.0, tileGridSize=(8, 8))
        return clahe.apply(gray)

    def _binarize(self, gray: np.ndarray) -> np.ndarray:
        return cv2.adaptiveThreshold(gray, 255, cv2.ADAPTIVE_THRESH_GAUSSIAN_C,
                                     cv2.THRESH_BINARY, 21, 10)

    def _morphological_cleanup(self, binary: np.ndarray) -> np.ndarray:
        kernel = np.ones((2, 2), np.uint8)
        cleaned = cv2.morphologyEx(binary, cv2.MORPH_OPEN, kernel)
        cleaned = cv2.morphologyEx(cleaned, cv2.MORPH_CLOSE, kernel)
        return cleaned


# ---------------------------------------------------------------------------
# Grid Detector - NEW: detects table lines and extracts cell bounding boxes
# ---------------------------------------------------------------------------

class GridDetector:
    """Detect grid lines in an image and extract individual cell regions."""

    def __init__(self, min_line_length_ratio=0.15):
        self.min_line_length_ratio = min_line_length_ratio

    def detect_cells(self, img: np.ndarray) -> list[CellBBox] | None:
        """Detect grid and return cell bounding boxes, or None if no grid found."""
        if len(img.shape) == 3:
            gray = cv2.cvtColor(img, cv2.COLOR_BGR2GRAY)
        else:
            gray = img.copy()

        h, w = gray.shape

        # Binarize (invert so lines are white)
        _, binary = cv2.threshold(gray, 0, 255,
                                  cv2.THRESH_BINARY_INV + cv2.THRESH_OTSU)

        # --- Detect horizontal lines ---
        horiz_kernel_len = max(int(w * self.min_line_length_ratio), 40)
        horiz_kernel = cv2.getStructuringElement(cv2.MORPH_RECT,
                                                  (horiz_kernel_len, 1))
        horiz = cv2.morphologyEx(binary, cv2.MORPH_OPEN, horiz_kernel, iterations=2)

        # --- Detect vertical lines ---
        vert_kernel_len = max(int(h * self.min_line_length_ratio), 40)
        vert_kernel = cv2.getStructuringElement(cv2.MORPH_RECT,
                                                 (1, vert_kernel_len))
        vert = cv2.morphologyEx(binary, cv2.MORPH_OPEN, vert_kernel, iterations=2)

        # Combine
        grid_mask = cv2.add(horiz, vert)

        # Find contours on combined grid
        contours, _ = cv2.findContours(grid_mask, cv2.RETR_TREE,
                                        cv2.CHAIN_APPROX_SIMPLE)

        if len(contours) < 4:
            logger.info("No grid detected (too few line segments)")
            return None

        # Extract horizontal and vertical line positions
        h_lines = self._extract_line_positions(horiz, axis="horizontal")
        v_lines = self._extract_line_positions(vert, axis="vertical")

        logger.info(f"Grid detected: {len(h_lines)} horizontal, {len(v_lines)} vertical lines")

        if len(h_lines) < 2 or len(v_lines) < 2:
            logger.info("Not enough lines to form a grid")
            return None

        # Build cells from line intersections
        cells = self._build_cells_from_lines(h_lines, v_lines)
        logger.info(f"Extracted {len(cells)} cells from grid")
        return cells

    def _extract_line_positions(self, line_mask: np.ndarray,
                                 axis: str) -> list[int]:
        """Extract sorted unique line positions from a binary line mask."""
        if axis == "horizontal":
            # Project horizontally: sum along columns
            projection = np.sum(line_mask, axis=1)
        else:
            # Project vertically: sum along rows
            projection = np.sum(line_mask, axis=0)

        # Find positions where projection is significant
        threshold = np.max(projection) * 0.3 if np.max(projection) > 0 else 0
        positions = np.where(projection > threshold)[0]

        if len(positions) == 0:
            return []

        # Cluster nearby positions (lines can be a few pixels thick)
        clustered = []
        cluster_start = positions[0]
        prev = positions[0]

        for pos in positions[1:]:
            if pos - prev > 5:
                clustered.append((cluster_start + prev) // 2)
                cluster_start = pos
            prev = pos
        clustered.append((cluster_start + prev) // 2)

        return clustered

    def _build_cells_from_lines(self, h_lines: list[int],
                                 v_lines: list[int]) -> list[CellBBox]:
        """Build cell bounding boxes from horizontal and vertical line positions."""
        h_lines = sorted(h_lines)
        v_lines = sorted(v_lines)

        cells = []
        for row_idx in range(len(h_lines) - 1):
            for col_idx in range(len(v_lines) - 1):
                y1 = h_lines[row_idx]
                y2 = h_lines[row_idx + 1]
                x1 = v_lines[col_idx]
                x2 = v_lines[col_idx + 1]

                # Skip very small cells (likely noise)
                cell_w = x2 - x1
                cell_h = y2 - y1
                if cell_w < 10 or cell_h < 10:
                    continue

                cells.append(CellBBox(
                    row=row_idx, col=col_idx,
                    x=x1, y=y1, w=cell_w, h=cell_h,
                ))

        return cells


# ---------------------------------------------------------------------------
# Tally Mark Detector - reads Brazilian square-based tally marks
# ---------------------------------------------------------------------------

class TallyMarkDetector:
    """Detect Brazilian square-based tally marks (quadradinhos) in cell images.

    Notation system:
        |       1 side  = 1 unit
        |_      2 sides (L shape) = 2 units
        |_|     3 sides (inverted U) = 3 units
        |_|     4 sides (complete square) = 4 units
         ‾
    Multiple squares are summed. E.g.: one complete square + L = 4+2 = 6.
    """

    def detect(self, cell_img: np.ndarray) -> int | None:
        """Try to read tally marks from a cell image.

        Returns the count if tally marks detected, None otherwise.
        """
        if len(cell_img.shape) == 3:
            gray = cv2.cvtColor(cell_img, cv2.COLOR_BGR2GRAY)
        else:
            gray = cell_img.copy()

        h, w = gray.shape
        if h < 5 or w < 5:
            return None

        # Binarize (ink = white)
        _, binary = cv2.threshold(gray, 0, 255,
                                  cv2.THRESH_BINARY_INV + cv2.THRESH_OTSU)

        # Check if cell has any ink at all
        ink_ratio = np.sum(binary > 0) / binary.size
        if ink_ratio < 0.01:
            return 0  # Empty cell = 0
        if ink_ratio > 0.5:
            return None  # Too much ink - probably not tally marks

        # Find connected components (each tally group is one component)
        num_labels, labels, stats, centroids = cv2.connectedComponentsWithStats(
            binary, connectivity=8
        )

        total_count = 0
        valid_groups = 0

        for label_id in range(1, num_labels):  # Skip background (0)
            comp_x = stats[label_id, cv2.CC_STAT_LEFT]
            comp_y = stats[label_id, cv2.CC_STAT_TOP]
            comp_w = stats[label_id, cv2.CC_STAT_WIDTH]
            comp_h = stats[label_id, cv2.CC_STAT_HEIGHT]
            comp_area = stats[label_id, cv2.CC_STAT_AREA]

            # Skip tiny noise
            if comp_area < 15 or comp_w < 3 or comp_h < 3:
                continue

            # Skip components that are too large relative to cell
            if comp_w > w * 0.9 and comp_h > h * 0.9:
                continue

            # Extract this component
            component_mask = (labels[comp_y:comp_y + comp_h,
                                     comp_x:comp_x + comp_w] == label_id).astype(np.uint8) * 255

            sides = self._count_sides(component_mask, comp_w, comp_h)

            if sides is not None and 1 <= sides <= 4:
                total_count += sides
                valid_groups += 1

        if valid_groups == 0:
            return None  # No valid tally marks found

        return total_count

    def _count_sides(self, mask: np.ndarray, w: int, h: int) -> int | None:
        """Count how many sides of a square are drawn in this component.

        Checks the 4 edges of the bounding box for ink presence.
        """
        if w < 3 or h < 3:
            return None

        # Define edge regions (thin strips along each border)
        strip = max(2, min(w, h) // 5)

        # Top edge: top strip of pixels, middle portion
        top_region = mask[0:strip, :]
        # Bottom edge
        bottom_region = mask[h - strip:h, :]
        # Left edge
        left_region = mask[:, 0:strip]
        # Right edge
        right_region = mask[:, w - strip:w]

        # Minimum ink threshold for a side to be considered "present"
        min_fill = 0.15

        sides = 0

        # Check each edge
        top_fill = np.sum(top_region > 0) / max(top_region.size, 1)
        bottom_fill = np.sum(bottom_region > 0) / max(bottom_region.size, 1)
        left_fill = np.sum(left_region > 0) / max(left_region.size, 1)
        right_fill = np.sum(right_region > 0) / max(right_region.size, 1)

        if top_fill > min_fill:
            sides += 1
        if bottom_fill > min_fill:
            sides += 1
        if left_fill > min_fill:
            sides += 1
        if right_fill > min_fill:
            sides += 1

        # Sanity check: a single straight line (horizontal or vertical)
        # should count as 1
        aspect = w / max(h, 1)
        if sides == 0:
            # Maybe it's a thin mark that doesn't fill edges well
            if aspect > 2.5:
                return 1  # Horizontal line
            elif aspect < 0.4:
                return 1  # Vertical line

        return sides if sides > 0 else None

    def looks_like_tally(self, cell_img: np.ndarray) -> bool:
        """Quick check: does this cell likely contain tally marks rather than text?

        Tally marks tend to have:
        - Few connected components (1-5 squares max)
        - Relatively low ink density
        - Components that are roughly square-ish
        """
        if len(cell_img.shape) == 3:
            gray = cv2.cvtColor(cell_img, cv2.COLOR_BGR2GRAY)
        else:
            gray = cell_img.copy()

        _, binary = cv2.threshold(gray, 0, 255,
                                  cv2.THRESH_BINARY_INV + cv2.THRESH_OTSU)

        ink_ratio = np.sum(binary > 0) / binary.size

        # Tally marks are sparse (not dense text)
        if ink_ratio > 0.35:
            return False
        if ink_ratio < 0.005:
            return False  # Empty

        num_labels, _, stats, _ = cv2.connectedComponentsWithStats(
            binary, connectivity=8
        )

        # Count significant components
        significant = 0
        for i in range(1, num_labels):
            if stats[i, cv2.CC_STAT_AREA] > 15:
                significant += 1

        # Tally marks: typically 1-8 components (1-8 square groups)
        # Text: usually many more small components
        return significant <= 10


# ---------------------------------------------------------------------------
# OCR Engine
# ---------------------------------------------------------------------------

class OCREngine:
    """Unified OCR engine supporting Tesseract and EasyOCR."""

    def __init__(self, engine="both", languages=None):
        self.engine = engine
        self.languages = languages or ["pt", "en"]
        self._tesseract_available = False
        self._easyocr_reader = None
        self._init_engines()

    def _init_engines(self):
        if self.engine in ("tesseract", "both"):
            try:
                import pytesseract
                pytesseract.get_tesseract_version()
                self._tesseract_available = True
                logger.info("Tesseract OCR initialized")
            except Exception as e:
                logger.warning(f"Tesseract not available: {e}")
                if self.engine == "tesseract":
                    raise RuntimeError("Tesseract requested but not available")

        if self.engine in ("easyocr", "both"):
            try:
                import easyocr
                self._easyocr_reader = easyocr.Reader(
                    self.languages, gpu=False, verbose=False
                )
                logger.info("EasyOCR initialized")
            except Exception as e:
                logger.warning(f"EasyOCR not available: {e}")
                if self.engine == "easyocr":
                    raise RuntimeError("EasyOCR requested but not available")

    def ocr_cell(self, cell_img: np.ndarray) -> str:
        """Run OCR on a single cell image and return best text."""
        results = []

        if self._tesseract_available:
            text = self._tesseract_cell(cell_img)
            if text.strip():
                results.append(text.strip())

        if self._easyocr_reader is not None:
            text = self._easyocr_cell(cell_img)
            if text.strip():
                results.append(text.strip())

        if not results:
            return ""

        # Return the longest result (usually more complete)
        return max(results, key=len)

    def recognize(self, images: list[np.ndarray]) -> list[TextBlock]:
        """Run OCR on full images (fallback for non-grid mode)."""
        all_blocks = []
        for i, img in enumerate(images):
            if self._tesseract_available:
                blocks = self._run_tesseract_full(img)
                all_blocks.extend(blocks)
            if self._easyocr_reader is not None and i == 0:
                blocks = self._run_easyocr_full(img)
                all_blocks.extend(blocks)
        merged = self._merge_results(all_blocks)
        return merged

    def _tesseract_cell(self, img: np.ndarray) -> str:
        """Run Tesseract on a single cell with optimized settings."""
        import pytesseract
        lang_str = "+".join(self._tesseract_languages())
        # PSM 7 = single line, PSM 6 = uniform block
        # Try PSM 7 first (most cells are single line)
        for psm in [7, 6, 13]:
            config = f"--oem 3 --psm {psm} -l {lang_str}"
            text = pytesseract.image_to_string(img, config=config).strip()
            if text:
                return text
        return ""

    def _easyocr_cell(self, img: np.ndarray) -> str:
        """Run EasyOCR on a single cell."""
        results = self._easyocr_reader.readtext(img, paragraph=True)
        texts = [text for _, text, conf in results if conf > 0.1]
        return " ".join(texts)

    def _run_tesseract_full(self, img: np.ndarray) -> list[TextBlock]:
        import pytesseract
        lang_str = "+".join(self._tesseract_languages())
        config = f"--oem 3 --psm 6 -l {lang_str} -c preserve_interword_spaces=1"
        data = pytesseract.image_to_data(img, config=config,
                                          output_type=pytesseract.Output.DICT)
        blocks = []
        for i in range(len(data["text"])):
            text = data["text"][i].strip()
            conf = int(data["conf"][i])
            if text and conf > 20:
                blocks.append(TextBlock(
                    text=text, x=data["left"][i], y=data["top"][i],
                    width=data["width"][i], height=data["height"][i],
                    confidence=conf / 100.0,
                ))
        return blocks

    def _run_easyocr_full(self, img: np.ndarray) -> list[TextBlock]:
        results = self._easyocr_reader.readtext(img, paragraph=False)
        blocks = []
        for bbox, text, conf in results:
            if not text.strip():
                continue
            xs = [p[0] for p in bbox]
            ys = [p[1] for p in bbox]
            x, y = int(min(xs)), int(min(ys))
            blocks.append(TextBlock(
                text=text.strip(), x=x, y=y,
                width=int(max(xs) - x), height=int(max(ys) - y),
                confidence=conf,
            ))
        return blocks

    def _tesseract_languages(self) -> list[str]:
        mapping = {"pt": "por", "en": "eng", "es": "spa", "fr": "fra", "de": "deu"}
        return [mapping.get(lang, lang) for lang in self.languages]

    def _merge_results(self, blocks: list[TextBlock]) -> list[TextBlock]:
        if not blocks:
            return []
        blocks.sort(key=lambda b: (b.y, b.x))
        merged = []
        used = set()
        for i, block in enumerate(blocks):
            if i in used:
                continue
            best = block
            for j in range(i + 1, len(blocks)):
                if j in used:
                    continue
                other = blocks[j]
                if self._blocks_overlap(best, other):
                    if other.confidence > best.confidence:
                        used.add(i)
                        best = other
                    else:
                        used.add(j)
            merged.append(best)
        return merged

    def _blocks_overlap(self, a: TextBlock, b: TextBlock, threshold=0.5) -> bool:
        x_overlap = max(0, min(a.right, b.right) - max(a.x, b.x))
        y_overlap = max(0, min(a.bottom, b.bottom) - max(a.y, b.y))
        overlap_area = x_overlap * y_overlap
        area_a = a.width * a.height
        area_b = b.width * b.height
        min_area = min(area_a, area_b) if min(area_a, area_b) > 0 else 1
        return (overlap_area / min_area) > threshold


# ---------------------------------------------------------------------------
# Table Structure Detector (fallback for images without grid lines)
# ---------------------------------------------------------------------------

class TableStructureDetector:
    """Detects table structure from scattered text blocks (no grid lines)."""

    def __init__(self, row_tolerance=None, col_tolerance=None):
        self.row_tolerance = row_tolerance
        self.col_tolerance = col_tolerance

    def detect(self, blocks: list[TextBlock]) -> list[list[str]]:
        if not blocks:
            return [[]]
        if self.row_tolerance is None or self.col_tolerance is None:
            self._auto_tolerances(blocks)
        rows = self._group_into_rows(blocks)
        col_positions = self._detect_columns(rows)
        table = self._build_table(rows, col_positions)
        return table

    def _auto_tolerances(self, blocks):
        if len(blocks) < 2:
            self.row_tolerance = 20
            self.col_tolerance = 40
            return
        heights = [b.height for b in blocks]
        median_height = sorted(heights)[len(heights) // 2]
        self.row_tolerance = max(int(median_height * 0.7), 10)
        self.col_tolerance = max(int(median_height * 1.5), 20)

    def _group_into_rows(self, blocks):
        sorted_blocks = sorted(blocks, key=lambda b: b.center_y)
        rows = []
        current_row = [sorted_blocks[0]]
        current_y = sorted_blocks[0].center_y
        for block in sorted_blocks[1:]:
            if abs(block.center_y - current_y) <= self.row_tolerance:
                current_row.append(block)
                current_y = sum(b.center_y for b in current_row) / len(current_row)
            else:
                current_row.sort(key=lambda b: b.x)
                rows.append(current_row)
                current_row = [block]
                current_y = block.center_y
        if current_row:
            current_row.sort(key=lambda b: b.x)
            rows.append(current_row)
        return rows

    def _detect_columns(self, rows):
        all_x = [b.x for row in rows for b in row]
        if not all_x:
            return [0]
        sorted_x = sorted(all_x)
        columns = [sorted_x[0]]
        for x in sorted_x[1:]:
            if x - columns[-1] > self.col_tolerance:
                columns.append(x)
        return columns

    def _build_table(self, rows, col_positions):
        table = []
        for row_blocks in rows:
            row = [""] * len(col_positions)
            for block in row_blocks:
                col_idx = min(range(len(col_positions)),
                              key=lambda i: abs(block.x - col_positions[i]))
                if row[col_idx]:
                    row[col_idx] += " " + block.text
                else:
                    row[col_idx] = block.text
            table.append(row)
        if table:
            max_cols = max(len(r) for r in table)
            for row in table:
                while len(row) < max_cols:
                    row.append("")
        return table


# ---------------------------------------------------------------------------
# Line Extractor (for non-tabular content)
# ---------------------------------------------------------------------------

class LineExtractor:
    def extract_lines(self, blocks: list[TextBlock]) -> list[list[str]]:
        if not blocks:
            return [[]]
        sorted_blocks = sorted(blocks, key=lambda b: b.center_y)
        heights = [b.height for b in sorted_blocks]
        median_h = sorted(heights)[len(heights) // 2]
        tolerance = max(int(median_h * 0.6), 10)
        lines = []
        current_line = [sorted_blocks[0]]
        current_y = sorted_blocks[0].center_y
        for block in sorted_blocks[1:]:
            if abs(block.center_y - current_y) <= tolerance:
                current_line.append(block)
                current_y = sum(b.center_y for b in current_line) / len(current_line)
            else:
                current_line.sort(key=lambda b: b.x)
                lines.append([" ".join(b.text for b in current_line)])
                current_line = [block]
                current_y = block.center_y
        if current_line:
            current_line.sort(key=lambda b: b.x)
            lines.append([" ".join(b.text for b in current_line)])
        return lines


# ---------------------------------------------------------------------------
# Excel Exporter
# ---------------------------------------------------------------------------

class ExcelExporter:
    def __init__(self):
        self.header_fill = PatternFill(start_color="4472C4", end_color="4472C4",
                                        fill_type="solid")
        self.header_font = Font(name="Calibri", bold=True, color="FFFFFF", size=11)
        self.cell_font = Font(name="Calibri", size=11)
        self.border = Border(
            left=Side(style="thin"), right=Side(style="thin"),
            top=Side(style="thin"), bottom=Side(style="thin"),
        )
        self.alignment = Alignment(wrap_text=True, vertical="center")

    def export(self, table_data: list[list[str]], output_path: str,
               first_row_header: bool = True, sheet_name: str = "Dados Reconhecidos"):
        wb = openpyxl.Workbook()
        ws = wb.active
        ws.title = sheet_name

        if not table_data or (len(table_data) == 1 and not any(table_data[0])):
            wb.save(output_path)
            return

        for row_idx, row in enumerate(table_data, start=1):
            for col_idx, value in enumerate(row, start=1):
                cell = ws.cell(row=row_idx, column=col_idx,
                               value=self._clean_text(value))
                cell.font = self.cell_font
                cell.border = self.border
                cell.alignment = self.alignment

        if first_row_header and len(table_data) > 1:
            for col_idx in range(1, len(table_data[0]) + 1):
                cell = ws.cell(row=1, column=col_idx)
                cell.fill = self.header_fill
                cell.font = self.header_font

        max_cols = max(len(r) for r in table_data)
        for col_idx in range(1, max_cols + 1):
            max_length = 0
            col_letter = openpyxl.utils.get_column_letter(col_idx)
            for row_idx in range(1, len(table_data) + 1):
                cell = ws.cell(row=row_idx, column=col_idx)
                if cell.value:
                    max_length = max(max_length, len(str(cell.value)))
            ws.column_dimensions[col_letter].width = min(max(max_length + 4, 10), 50)

        if first_row_header:
            ws.freeze_panes = "A2"

        if table_data:
            ws.auto_filter.ref = (
                f"A1:{openpyxl.utils.get_column_letter(max_cols)}{len(table_data)}"
            )

        wb.save(output_path)
        logger.info(f"Excel saved: {output_path} "
                     f"({len(table_data)} rows x {max_cols} cols)")

    def _clean_text(self, text: str) -> str:
        if not text:
            return ""
        text = re.sub(r"\s+", " ", text).strip()
        text = "".join(c for c in text if c.isprintable() or c in "\n\t")
        return text


# ---------------------------------------------------------------------------
# Main Pipeline - v2 with grid-aware cell-by-cell OCR
# ---------------------------------------------------------------------------

class HandwritingToExcel:
    """Main pipeline with grid-aware cell-by-cell OCR."""

    def __init__(self, engine="both", languages=None, enhance_level="standard",
                 mode="auto"):
        self.preprocessor = ImagePreprocessor(enhance_level=enhance_level)
        self.ocr = OCREngine(engine=engine, languages=languages)
        self.grid_detector = GridDetector()
        self.tally_detector = TallyMarkDetector()
        self.table_detector = TableStructureDetector()
        self.line_extractor = LineExtractor()
        self.exporter = ExcelExporter()
        self.mode = mode
        self.enhance_level = enhance_level

    def convert(self, image_path: str, output_path: str,
                first_row_header: bool = True) -> list[list[str]]:
        """Full pipeline: image -> grid detect -> cell OCR -> Excel."""

        # Step 1: Load image
        img = self.preprocessor.load_image(image_path)
        img = self.preprocessor._normalize_size(img)

        # Step 2: Try grid detection first (for structured forms)
        if self.mode in ("auto", "table"):
            cells = self.grid_detector.detect_cells(img)

            if cells:
                logger.info("Using GRID mode: cell-by-cell OCR")
                table_data = self._ocr_grid_cells(img, cells)

                # Export
                self.exporter.export(table_data, output_path,
                                     first_row_header=first_row_header)
                return table_data

        # Step 3: Fallback to full-image OCR
        logger.info("Using FULL-IMAGE mode (no grid detected)")
        processed_versions = self.preprocessor.preprocess_for_ocr(img)
        text_blocks = self.ocr.recognize(processed_versions)

        if not text_blocks:
            logger.warning("No text detected!")
            self.exporter.export([[]], output_path)
            return [[]]

        if self.mode == "lines":
            table_data = self.line_extractor.extract_lines(text_blocks)
        elif self.mode == "table":
            table_data = self.table_detector.detect(text_blocks)
        else:
            table_data = self._auto_detect_structure(text_blocks)

        self.exporter.export(table_data, output_path,
                             first_row_header=first_row_header)
        return table_data

    def _ocr_grid_cells(self, img: np.ndarray,
                         cells: list[CellBBox]) -> list[list[str]]:
        """Crop each cell from image, preprocess, and OCR individually.

        For each cell, first tries tally mark detection (quadradinhos).
        If the cell looks like tally marks, returns the numeric count.
        Otherwise falls back to OCR for text content.
        """
        if not cells:
            return [[]]

        max_row = max(c.row for c in cells) + 1
        max_col = max(c.col for c in cells) + 1
        table = [["" for _ in range(max_col)] for _ in range(max_row)]

        # Shrink margin: crop slightly inside the cell to avoid grid lines
        margin_x = 4
        margin_y = 3

        # First pass: OCR the header row to identify text vs numeric columns
        # (we'll use tally detection on all non-header cells heuristically)

        for cell in cells:
            # Crop cell from original image
            x1 = cell.x + margin_x
            y1 = cell.y + margin_y
            x2 = cell.x + cell.w - margin_x
            y2 = cell.y + cell.h - margin_y

            if x2 <= x1 or y2 <= y1:
                continue

            cell_img = img[y1:y2, x1:x2]

            if cell_img.size == 0:
                continue

            # --- Strategy: try tally marks first on non-header rows ---
            # Header row (row 0) always uses OCR
            if cell.row > 0 and self.tally_detector.looks_like_tally(cell_img):
                tally_count = self.tally_detector.detect(cell_img)
                if tally_count is not None:
                    # 0 means empty cell, show as empty
                    table[cell.row][cell.col] = str(tally_count) if tally_count > 0 else ""
                    logger.debug(f"Cell [{cell.row},{cell.col}]: tally={tally_count}")
                    continue

            # --- Fallback: OCR for text cells ---
            processed = self.preprocessor.preprocess_cell(cell_img)
            text = self.ocr.ocr_cell(processed)

            # Also try on grayscale version if no text found
            if not text.strip():
                if len(cell_img.shape) == 3:
                    gray = cv2.cvtColor(cell_img, cv2.COLOR_BGR2GRAY)
                else:
                    gray = cell_img
                text = self.ocr.ocr_cell(gray)

            table[cell.row][cell.col] = text.strip()

        return table

    def _auto_detect_structure(self, blocks):
        if len(blocks) < 3:
            return self.line_extractor.extract_lines(blocks)
        x_positions = sorted(set(b.x for b in blocks))
        if len(x_positions) < 2:
            return self.line_extractor.extract_lines(blocks)
        gaps = [x_positions[i+1] - x_positions[i]
                for i in range(len(x_positions) - 1)]
        median_gap = sorted(gaps)[len(gaps) // 2] if gaps else 0
        large_gaps = sum(1 for g in gaps if g > median_gap * 2)
        if large_gaps >= 1 and len(x_positions) >= 3:
            return self.table_detector.detect(blocks)
        else:
            return self.line_extractor.extract_lines(blocks)


# ---------------------------------------------------------------------------
# CLI
# ---------------------------------------------------------------------------

def parse_args():
    parser = argparse.ArgumentParser(
        description="Recognize handwriting from JPG images and export to Excel",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog="""
Examples:
  %(prog)s foto.jpg
  %(prog)s foto.jpg -o resultado.xlsx
  %(prog)s foto.jpg --engine easyocr --mode table
  %(prog)s foto.jpg --enhance aggressive --lang pt en
  %(prog)s foto.jpg --mode lines --no-header
        """,
    )
    parser.add_argument("image", help="Path to JPG image with handwriting")
    parser.add_argument("-o", "--output", default=None,
                        help="Output Excel file path (default: <image_name>.xlsx)")
    parser.add_argument("--engine", choices=["tesseract", "easyocr", "both"],
                        default="both", help="OCR engine (default: both)")
    parser.add_argument("--lang", nargs="+", default=["pt", "en"],
                        help="Languages for OCR (default: pt en)")
    parser.add_argument("--enhance", choices=["standard", "aggressive"],
                        default="standard", help="Enhancement level (default: standard)")
    parser.add_argument("--mode", choices=["auto", "table", "lines"],
                        default="auto", help="Detection mode (default: auto)")
    parser.add_argument("--no-header", action="store_true",
                        help="Don't treat first row as header")
    parser.add_argument("--debug", action="store_true",
                        help="Save preprocessed images for debugging")
    parser.add_argument("-v", "--verbose", action="store_true",
                        help="Verbose output")
    return parser.parse_args()


def main():
    args = parse_args()

    if args.verbose:
        logging.getLogger().setLevel(logging.DEBUG)

    if not os.path.exists(args.image):
        logger.error(f"Image file not found: {args.image}")
        sys.exit(1)

    output_path = args.output or f"{Path(args.image).stem}.xlsx"

    logger.info(f"Input:  {args.image}")
    logger.info(f"Output: {output_path}")
    logger.info(f"Engine: {args.engine}")
    logger.info(f"Mode: {args.mode}")

    if args.debug:
        preprocessor = ImagePreprocessor(enhance_level=args.enhance)
        img = preprocessor.load_image(args.image)
        versions = preprocessor.preprocess_for_ocr(img)
        for i, v in enumerate(versions):
            debug_path = f"debug_preprocessed_{i}.png"
            cv2.imwrite(debug_path, v)
            logger.info(f"Debug image saved: {debug_path}")

    pipeline = HandwritingToExcel(
        engine=args.engine, languages=args.lang,
        enhance_level=args.enhance, mode=args.mode,
    )

    table = pipeline.convert(
        image_path=args.image, output_path=output_path,
        first_row_header=not args.no_header,
    )

    print("\n--- Preview ---")
    for i, row in enumerate(table[:10]):
        print(f"  Row {i+1}: {' | '.join(str(cell) for cell in row)}")
    if len(table) > 10:
        print(f"  ... and {len(table) - 10} more rows")
    print(f"\nDone! Excel file saved to: {output_path}")


if __name__ == "__main__":
    main()
