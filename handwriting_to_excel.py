#!/usr/bin/env python3
"""
Handwriting Recognition to Excel Tool
======================================
High-capability tool for recognizing messy/disorganized handwriting
from JPG images and exporting structured data to Excel spreadsheets.

Features:
- Advanced image preprocessing (denoising, deskewing, binarization)
- Dual OCR engine support (Tesseract + EasyOCR) for maximum accuracy
- Automatic table structure detection from handwritten content
- Smart line/column grouping even for messy layouts
- Clean Excel export with formatting

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
from dataclasses import dataclass, field

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
class TableCell:
    """Represents a cell in the reconstructed table."""
    row: int
    col: int
    text: str
    confidence: float = 0.0


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
        """Full preprocessing pipeline for handwriting recognition."""
        logger.info("Starting image preprocessing...")

        # Step 1: Resize if too small or too large
        img = self._normalize_size(img)

        # Step 2: Convert to grayscale
        gray = cv2.cvtColor(img, cv2.COLOR_BGR2GRAY)

        # Step 3: Deskew (fix rotation)
        gray = self._deskew(gray)

        # Step 4: Denoise
        gray = self._denoise(gray)

        # Step 5: Enhance contrast
        gray = self._enhance_contrast(gray)

        # Step 6: Adaptive binarization
        binary = self._binarize(gray)

        # Step 7: Morphological cleanup
        binary = self._morphological_cleanup(binary)

        logger.info("Preprocessing complete.")
        return binary

    def preprocess_for_ocr(self, img: np.ndarray) -> list[np.ndarray]:
        """Generate multiple preprocessed versions for OCR ensemble."""
        versions = []

        # Version 1: Standard preprocessing
        versions.append(self.preprocess(img))

        # Version 2: High contrast grayscale (no binarization)
        gray = cv2.cvtColor(img, cv2.COLOR_BGR2GRAY)
        gray = self._deskew(gray)
        gray = self._denoise(gray)
        clahe = cv2.createCLAHE(clipLimit=3.0, tileGridSize=(8, 8))
        versions.append(clahe.apply(gray))

        # Version 3: Inverted binary (white text on black)
        versions.append(cv2.bitwise_not(versions[0]))

        if self.enhance_level == "aggressive":
            # Version 4: Dilated text for broken characters
            kernel = np.ones((2, 2), np.uint8)
            dilated = cv2.dilate(versions[0], kernel, iterations=1)
            versions.append(dilated)

        return versions

    def _normalize_size(self, img: np.ndarray) -> np.ndarray:
        """Resize image to optimal range for OCR."""
        h, w = img.shape[:2]
        min_dim = 1000
        max_dim = 4000

        if max(h, w) > max_dim:
            scale = max_dim / max(h, w)
            img = cv2.resize(img, None, fx=scale, fy=scale, interpolation=cv2.INTER_AREA)
            logger.info(f"Downscaled to {img.shape[1]}x{img.shape[0]}")
        elif max(h, w) < min_dim:
            scale = min_dim / max(h, w)
            img = cv2.resize(img, None, fx=scale, fy=scale, interpolation=cv2.INTER_CUBIC)
            logger.info(f"Upscaled to {img.shape[1]}x{img.shape[0]}")

        return img

    def _deskew(self, gray: np.ndarray) -> np.ndarray:
        """Correct image rotation/skew."""
        # Use Hough lines to detect dominant angle
        edges = cv2.Canny(gray, 50, 150, apertureSize=3)
        lines = cv2.HoughLinesP(edges, 1, np.pi / 180, threshold=100,
                                minLineLength=gray.shape[1] // 8,
                                maxLineGap=20)

        if lines is None or len(lines) < 3:
            return gray

        angles = []
        for line in lines:
            x1, y1, x2, y2 = line[0]
            angle = np.degrees(np.arctan2(y2 - y1, x2 - x1))
            # Only consider near-horizontal lines
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
        rotated = cv2.warpAffine(gray, matrix, (w, h),
                                 flags=cv2.INTER_CUBIC,
                                 borderMode=cv2.BORDER_REPLICATE)
        return rotated

    def _denoise(self, gray: np.ndarray) -> np.ndarray:
        """Remove noise while preserving text edges."""
        if self.enhance_level == "aggressive":
            denoised = cv2.fastNlMeansDenoising(gray, h=15, templateWindowSize=7,
                                                 searchWindowSize=21)
        else:
            denoised = cv2.fastNlMeansDenoising(gray, h=10, templateWindowSize=7,
                                                 searchWindowSize=21)
        return denoised

    def _enhance_contrast(self, gray: np.ndarray) -> np.ndarray:
        """Enhance contrast using CLAHE."""
        if self.enhance_level == "aggressive":
            clahe = cv2.createCLAHE(clipLimit=4.0, tileGridSize=(4, 4))
        else:
            clahe = cv2.createCLAHE(clipLimit=2.0, tileGridSize=(8, 8))
        return clahe.apply(gray)

    def _binarize(self, gray: np.ndarray) -> np.ndarray:
        """Adaptive thresholding for uneven lighting."""
        # Try both methods and pick the one with better text separation
        adaptive_mean = cv2.adaptiveThreshold(
            gray, 255, cv2.ADAPTIVE_THRESH_MEAN_C,
            cv2.THRESH_BINARY, 21, 10
        )
        adaptive_gauss = cv2.adaptiveThreshold(
            gray, 255, cv2.ADAPTIVE_THRESH_GAUSSIAN_C,
            cv2.THRESH_BINARY, 21, 10
        )

        # Use Gaussian as default (usually better for handwriting)
        return adaptive_gauss

    def _morphological_cleanup(self, binary: np.ndarray) -> np.ndarray:
        """Clean up small noise and reconnect broken strokes."""
        # Remove small noise
        kernel_small = np.ones((2, 2), np.uint8)
        cleaned = cv2.morphologyEx(binary, cv2.MORPH_OPEN, kernel_small)

        # Close small gaps in characters
        kernel_close = np.ones((2, 2), np.uint8)
        cleaned = cv2.morphologyEx(cleaned, cv2.MORPH_CLOSE, kernel_close)

        return cleaned


# ---------------------------------------------------------------------------
# OCR Engine Wrapper
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
        """Initialize available OCR engines."""
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

    def recognize(self, images: list[np.ndarray]) -> list[TextBlock]:
        """Run OCR on preprocessed images and merge results."""
        all_blocks = []

        for i, img in enumerate(images):
            if self._tesseract_available:
                blocks = self._run_tesseract(img)
                logger.info(f"Tesseract (version {i+1}): {len(blocks)} text blocks")
                all_blocks.extend(blocks)

            if self._easyocr_reader is not None and i == 0:
                blocks = self._run_easyocr(img)
                logger.info(f"EasyOCR (version {i+1}): {len(blocks)} text blocks")
                all_blocks.extend(blocks)

        # Merge and deduplicate results
        merged = self._merge_results(all_blocks)
        logger.info(f"Total merged text blocks: {len(merged)}")
        return merged

    def _run_tesseract(self, img: np.ndarray) -> list[TextBlock]:
        """Run Tesseract OCR with detailed output."""
        import pytesseract

        lang_str = "+".join(self._tesseract_languages())
        custom_config = (
            f"--oem 3 --psm 6 "
            f"-l {lang_str} "
            f"-c preserve_interword_spaces=1"
        )

        data = pytesseract.image_to_data(img, config=custom_config,
                                          output_type=pytesseract.Output.DICT)

        blocks = []
        n_items = len(data["text"])
        for i in range(n_items):
            text = data["text"][i].strip()
            conf = int(data["conf"][i])
            if text and conf > 20:
                blocks.append(TextBlock(
                    text=text,
                    x=data["left"][i],
                    y=data["top"][i],
                    width=data["width"][i],
                    height=data["height"][i],
                    confidence=conf / 100.0,
                ))
        return blocks

    def _run_easyocr(self, img: np.ndarray) -> list[TextBlock]:
        """Run EasyOCR."""
        results = self._easyocr_reader.readtext(img, paragraph=False)
        blocks = []
        for bbox, text, conf in results:
            if not text.strip():
                continue
            # bbox is [[x1,y1],[x2,y2],[x3,y3],[x4,y4]]
            xs = [p[0] for p in bbox]
            ys = [p[1] for p in bbox]
            x = int(min(xs))
            y = int(min(ys))
            w = int(max(xs) - x)
            h = int(max(ys) - y)
            blocks.append(TextBlock(
                text=text.strip(),
                x=x, y=y, width=w, height=h,
                confidence=conf,
            ))
        return blocks

    def _tesseract_languages(self) -> list[str]:
        """Map language codes to Tesseract format."""
        mapping = {"pt": "por", "en": "eng", "es": "spa", "fr": "fra", "de": "deu"}
        return [mapping.get(lang, lang) for lang in self.languages]

    def _merge_results(self, blocks: list[TextBlock]) -> list[TextBlock]:
        """Merge overlapping text blocks from multiple engines/versions."""
        if not blocks:
            return []

        # Sort by position (top to bottom, left to right)
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
                    # Keep the one with higher confidence
                    if other.confidence > best.confidence:
                        used.add(i)
                        best = other
                    else:
                        used.add(j)

            merged.append(best)

        return merged

    def _blocks_overlap(self, a: TextBlock, b: TextBlock, threshold=0.5) -> bool:
        """Check if two text blocks significantly overlap."""
        x_overlap = max(0, min(a.right, b.right) - max(a.x, b.x))
        y_overlap = max(0, min(a.bottom, b.bottom) - max(a.y, b.y))
        overlap_area = x_overlap * y_overlap

        area_a = a.width * a.height
        area_b = b.width * b.height
        min_area = min(area_a, area_b) if min(area_a, area_b) > 0 else 1

        return (overlap_area / min_area) > threshold


# ---------------------------------------------------------------------------
# Table Structure Detector
# ---------------------------------------------------------------------------

class TableStructureDetector:
    """Detects table structure from scattered text blocks."""

    def __init__(self, row_tolerance=None, col_tolerance=None):
        self.row_tolerance = row_tolerance
        self.col_tolerance = col_tolerance

    def detect(self, blocks: list[TextBlock]) -> list[list[str]]:
        """Convert text blocks into a 2D table structure."""
        if not blocks:
            return [[]]

        # Auto-calculate tolerances if not specified
        if self.row_tolerance is None or self.col_tolerance is None:
            self._auto_tolerances(blocks)

        # Step 1: Group blocks into rows
        rows = self._group_into_rows(blocks)

        # Step 2: Detect column positions
        col_positions = self._detect_columns(rows)

        # Step 3: Assign blocks to grid cells
        table = self._build_table(rows, col_positions)

        logger.info(f"Detected table: {len(table)} rows x "
                     f"{max(len(r) for r in table) if table else 0} cols")
        return table

    def _auto_tolerances(self, blocks: list[TextBlock]):
        """Calculate row/column tolerances from text block statistics."""
        if len(blocks) < 2:
            self.row_tolerance = 20
            self.col_tolerance = 40
            return

        heights = [b.height for b in blocks]
        median_height = sorted(heights)[len(heights) // 2]

        self.row_tolerance = max(int(median_height * 0.7), 10)
        self.col_tolerance = max(int(median_height * 1.5), 20)

        logger.info(f"Auto tolerances: row={self.row_tolerance}, col={self.col_tolerance}")

    def _group_into_rows(self, blocks: list[TextBlock]) -> list[list[TextBlock]]:
        """Group text blocks into rows based on vertical position."""
        sorted_blocks = sorted(blocks, key=lambda b: b.center_y)

        rows = []
        current_row = [sorted_blocks[0]]
        current_y = sorted_blocks[0].center_y

        for block in sorted_blocks[1:]:
            if abs(block.center_y - current_y) <= self.row_tolerance:
                current_row.append(block)
                # Update running average of Y position
                current_y = sum(b.center_y for b in current_row) / len(current_row)
            else:
                # Sort row left to right
                current_row.sort(key=lambda b: b.x)
                rows.append(current_row)
                current_row = [block]
                current_y = block.center_y

        if current_row:
            current_row.sort(key=lambda b: b.x)
            rows.append(current_row)

        return rows

    def _detect_columns(self, rows: list[list[TextBlock]]) -> list[int]:
        """Detect column boundaries from all rows."""
        # Collect all X positions
        all_x_positions = []
        for row in rows:
            for block in row:
                all_x_positions.append(block.x)

        if not all_x_positions:
            return [0]

        # Cluster X positions to find column starts
        sorted_x = sorted(all_x_positions)
        columns = [sorted_x[0]]

        for x in sorted_x[1:]:
            if x - columns[-1] > self.col_tolerance:
                columns.append(x)
            else:
                # Update cluster center
                pass

        return columns

    def _build_table(self, rows: list[list[TextBlock]],
                     col_positions: list[int]) -> list[list[str]]:
        """Assign text blocks to table cells."""
        table = []

        for row_blocks in rows:
            row = [""] * len(col_positions)

            for block in row_blocks:
                # Find closest column
                col_idx = self._find_closest_column(block.x, col_positions)
                if row[col_idx]:
                    row[col_idx] += " " + block.text
                else:
                    row[col_idx] = block.text

            table.append(row)

        # Normalize: ensure all rows have the same number of columns
        if table:
            max_cols = max(len(row) for row in table)
            for row in table:
                while len(row) < max_cols:
                    row.append("")

        return table

    def _find_closest_column(self, x: int, col_positions: list[int]) -> int:
        """Find the column index closest to the given x position."""
        min_dist = float("inf")
        best_col = 0

        for i, col_x in enumerate(col_positions):
            dist = abs(x - col_x)
            if dist < min_dist:
                min_dist = dist
                best_col = i

        return best_col


# ---------------------------------------------------------------------------
# Excel Exporter
# ---------------------------------------------------------------------------

class ExcelExporter:
    """Export recognized table data to a formatted Excel file."""

    def __init__(self):
        self.header_fill = PatternFill(start_color="4472C4", end_color="4472C4",
                                        fill_type="solid")
        self.header_font = Font(name="Calibri", bold=True, color="FFFFFF", size=11)
        self.cell_font = Font(name="Calibri", size=11)
        self.border = Border(
            left=Side(style="thin"),
            right=Side(style="thin"),
            top=Side(style="thin"),
            bottom=Side(style="thin"),
        )
        self.alignment = Alignment(wrap_text=True, vertical="center")

    def export(self, table_data: list[list[str]], output_path: str,
               first_row_header: bool = True, sheet_name: str = "Dados Reconhecidos"):
        """Export table data to Excel with formatting."""
        wb = openpyxl.Workbook()
        ws = wb.active
        ws.title = sheet_name

        if not table_data:
            wb.save(output_path)
            logger.info(f"Empty spreadsheet saved to {output_path}")
            return

        # Write data
        for row_idx, row in enumerate(table_data, start=1):
            for col_idx, value in enumerate(row, start=1):
                cell = ws.cell(row=row_idx, column=col_idx, value=self._clean_text(value))
                cell.font = self.cell_font
                cell.border = self.border
                cell.alignment = self.alignment

        # Format header row
        if first_row_header and len(table_data) > 1:
            for col_idx in range(1, len(table_data[0]) + 1):
                cell = ws.cell(row=1, column=col_idx)
                cell.fill = self.header_fill
                cell.font = self.header_font

        # Auto-fit column widths
        for col_idx in range(1, (max(len(r) for r in table_data) + 1)):
            max_length = 0
            col_letter = openpyxl.utils.get_column_letter(col_idx)
            for row_idx in range(1, len(table_data) + 1):
                cell = ws.cell(row=row_idx, column=col_idx)
                if cell.value:
                    max_length = max(max_length, len(str(cell.value)))
            ws.column_dimensions[col_letter].width = min(max(max_length + 4, 10), 50)

        # Freeze header row
        if first_row_header:
            ws.freeze_panes = "A2"

        # Add autofilter
        if table_data:
            max_col = max(len(r) for r in table_data)
            ws.auto_filter.ref = f"A1:{openpyxl.utils.get_column_letter(max_col)}{len(table_data)}"

        wb.save(output_path)
        logger.info(f"Excel file saved: {output_path}")
        logger.info(f"  {len(table_data)} rows x {max(len(r) for r in table_data)} columns")

    def _clean_text(self, text: str) -> str:
        """Clean OCR artifacts from text."""
        if not text:
            return ""
        # Remove multiple spaces
        text = re.sub(r"\s+", " ", text).strip()
        # Fix common OCR errors
        text = text.replace("|", "l").replace("}{", "ll")
        # Remove control characters
        text = "".join(c for c in text if c.isprintable() or c in "\n\t")
        return text


# ---------------------------------------------------------------------------
# Line-based mode (for non-tabular content)
# ---------------------------------------------------------------------------

class LineExtractor:
    """Extract text line by line for non-tabular handwriting."""

    def extract_lines(self, blocks: list[TextBlock]) -> list[list[str]]:
        """Group blocks into lines and return as single-column table."""
        if not blocks:
            return [[]]

        sorted_blocks = sorted(blocks, key=lambda b: b.center_y)

        # Calculate median height for row grouping
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
                line_text = " ".join(b.text for b in current_line)
                lines.append([line_text])
                current_line = [block]
                current_y = block.center_y

        if current_line:
            current_line.sort(key=lambda b: b.x)
            line_text = " ".join(b.text for b in current_line)
            lines.append([line_text])

        return lines


# ---------------------------------------------------------------------------
# Main Pipeline
# ---------------------------------------------------------------------------

class HandwritingToExcel:
    """Main pipeline orchestrating the full conversion process."""

    def __init__(self, engine="both", languages=None, enhance_level="standard",
                 mode="auto"):
        self.preprocessor = ImagePreprocessor(enhance_level=enhance_level)
        self.ocr = OCREngine(engine=engine, languages=languages)
        self.table_detector = TableStructureDetector()
        self.line_extractor = LineExtractor()
        self.exporter = ExcelExporter()
        self.mode = mode

    def convert(self, image_path: str, output_path: str,
                first_row_header: bool = True) -> list[list[str]]:
        """Full pipeline: image -> preprocessing -> OCR -> table -> Excel."""

        # Step 1: Load image
        img = self.preprocessor.load_image(image_path)

        # Step 2: Preprocess
        processed_versions = self.preprocessor.preprocess_for_ocr(img)

        # Step 3: OCR
        text_blocks = self.ocr.recognize(processed_versions)

        if not text_blocks:
            logger.warning("No text detected in the image!")
            self.exporter.export([[]], output_path)
            return [[]]

        # Step 4: Structure detection
        if self.mode == "table":
            table_data = self.table_detector.detect(text_blocks)
        elif self.mode == "lines":
            table_data = self.line_extractor.extract_lines(text_blocks)
        else:
            # Auto-detect: if blocks seem tabular, use table mode
            table_data = self._auto_detect_structure(text_blocks)

        # Step 5: Export to Excel
        self.exporter.export(table_data, output_path,
                             first_row_header=first_row_header)

        return table_data

    def _auto_detect_structure(self, blocks: list[TextBlock]) -> list[list[str]]:
        """Automatically detect if content is tabular or line-based."""
        if len(blocks) < 3:
            return self.line_extractor.extract_lines(blocks)

        # Check X-position variance - tabular data has clustered X positions
        x_positions = sorted(set(b.x for b in blocks))
        if len(x_positions) < 2:
            return self.line_extractor.extract_lines(blocks)

        # Calculate gaps between X positions
        gaps = [x_positions[i+1] - x_positions[i]
                for i in range(len(x_positions) - 1)]
        median_gap = sorted(gaps)[len(gaps) // 2] if gaps else 0

        # If there are significant gaps, it's likely tabular
        large_gaps = sum(1 for g in gaps if g > median_gap * 2)

        if large_gaps >= 1 and len(x_positions) >= 3:
            logger.info("Auto-detected: TABLE structure")
            return self.table_detector.detect(blocks)
        else:
            logger.info("Auto-detected: LINE structure")
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
                        default="both",
                        help="OCR engine to use (default: both)")
    parser.add_argument("--lang", nargs="+", default=["pt", "en"],
                        help="Languages for OCR (default: pt en)")
    parser.add_argument("--enhance", choices=["standard", "aggressive"],
                        default="standard",
                        help="Image enhancement level (default: standard)")
    parser.add_argument("--mode", choices=["auto", "table", "lines"],
                        default="auto",
                        help="Detection mode: auto, table, or lines (default: auto)")
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

    # Validate input
    if not os.path.exists(args.image):
        logger.error(f"Image file not found: {args.image}")
        sys.exit(1)

    # Set output path
    if args.output:
        output_path = args.output
    else:
        base = Path(args.image).stem
        output_path = f"{base}.xlsx"

    logger.info(f"Input:  {args.image}")
    logger.info(f"Output: {output_path}")
    logger.info(f"Engine: {args.engine}")
    logger.info(f"Languages: {args.lang}")
    logger.info(f"Mode: {args.mode}")

    # Save debug images if requested
    if args.debug:
        preprocessor = ImagePreprocessor(enhance_level=args.enhance)
        img = preprocessor.load_image(args.image)
        versions = preprocessor.preprocess_for_ocr(img)
        for i, v in enumerate(versions):
            debug_path = f"debug_preprocessed_{i}.png"
            cv2.imwrite(debug_path, v)
            logger.info(f"Debug image saved: {debug_path}")

    # Run pipeline
    pipeline = HandwritingToExcel(
        engine=args.engine,
        languages=args.lang,
        enhance_level=args.enhance,
        mode=args.mode,
    )

    table = pipeline.convert(
        image_path=args.image,
        output_path=output_path,
        first_row_header=not args.no_header,
    )

    # Print preview
    print("\n--- Preview ---")
    for i, row in enumerate(table[:10]):
        print(f"  Row {i+1}: {' | '.join(str(cell) for cell in row)}")
    if len(table) > 10:
        print(f"  ... and {len(table) - 10} more rows")

    print(f"\nDone! Excel file saved to: {output_path}")


if __name__ == "__main__":
    main()
