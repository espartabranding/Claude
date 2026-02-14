#!/usr/bin/env python3
"""
Batch processor for handwriting recognition.
Processes multiple JPG images and creates one Excel file with multiple sheets,
or one Excel file per image.

Usage:
    python batch_process.py images_folder/ -o results/
    python batch_process.py images_folder/ --single-file combined.xlsx
"""

import argparse
import os
import sys
import logging
from pathlib import Path
from concurrent.futures import ProcessPoolExecutor, as_completed

from handwriting_to_excel import HandwritingToExcel

logging.basicConfig(level=logging.INFO, format="%(levelname)s: %(message)s")
logger = logging.getLogger(__name__)

SUPPORTED_EXTENSIONS = {".jpg", ".jpeg", ".png", ".bmp", ".tiff", ".tif", ".webp"}


def find_images(input_path: str) -> list[str]:
    """Find all supported image files in the given path."""
    path = Path(input_path)

    if path.is_file():
        if path.suffix.lower() in SUPPORTED_EXTENSIONS:
            return [str(path)]
        else:
            logger.error(f"Unsupported file format: {path.suffix}")
            return []

    if path.is_dir():
        images = []
        for ext in SUPPORTED_EXTENSIONS:
            images.extend(str(p) for p in path.rglob(f"*{ext}"))
            images.extend(str(p) for p in path.rglob(f"*{ext.upper()}"))
        return sorted(set(images))

    logger.error(f"Path not found: {input_path}")
    return []


def process_single_image(image_path: str, output_path: str,
                         engine: str, languages: list[str],
                         enhance: str, mode: str) -> tuple[str, bool, str]:
    """Process a single image. Returns (image_path, success, message)."""
    try:
        pipeline = HandwritingToExcel(
            engine=engine,
            languages=languages,
            enhance_level=enhance,
            mode=mode,
        )
        pipeline.convert(image_path, output_path)
        return (image_path, True, f"Saved to {output_path}")
    except Exception as e:
        return (image_path, False, str(e))


def main():
    parser = argparse.ArgumentParser(
        description="Batch process handwriting images to Excel"
    )
    parser.add_argument("input", help="Image file or folder with images")
    parser.add_argument("-o", "--output-dir", default="output",
                        help="Output directory (default: output/)")
    parser.add_argument("--single-file", default=None,
                        help="Combine all results into a single Excel file")
    parser.add_argument("--engine", choices=["tesseract", "easyocr", "both"],
                        default="both")
    parser.add_argument("--lang", nargs="+", default=["pt", "en"])
    parser.add_argument("--enhance", choices=["standard", "aggressive"],
                        default="standard")
    parser.add_argument("--mode", choices=["auto", "table", "lines"],
                        default="auto")
    parser.add_argument("--workers", type=int, default=1,
                        help="Number of parallel workers (default: 1)")

    args = parser.parse_args()

    images = find_images(args.input)
    if not images:
        logger.error("No images found!")
        sys.exit(1)

    logger.info(f"Found {len(images)} images to process")

    os.makedirs(args.output_dir, exist_ok=True)

    if args.single_file:
        # Process all into one Excel with multiple sheets
        import openpyxl
        combined_wb = openpyxl.Workbook()
        combined_wb.remove(combined_wb.active)

        pipeline = HandwritingToExcel(
            engine=args.engine,
            languages=args.lang,
            enhance_level=args.enhance,
            mode=args.mode,
        )

        for i, image_path in enumerate(images, 1):
            name = Path(image_path).stem[:31]  # Excel sheet name limit
            logger.info(f"[{i}/{len(images)}] Processing: {image_path}")

            try:
                temp_output = os.path.join(args.output_dir, f"_temp_{i}.xlsx")
                pipeline.convert(image_path, temp_output)

                temp_wb = openpyxl.load_workbook(temp_output)
                temp_ws = temp_wb.active

                new_ws = combined_wb.create_sheet(title=name)
                for row in temp_ws.iter_rows():
                    for cell in row:
                        new_ws.cell(row=cell.row, column=cell.column,
                                    value=cell.value)

                os.remove(temp_output)
                logger.info(f"  OK: {name}")
            except Exception as e:
                logger.error(f"  FAILED: {e}")

        output_path = os.path.join(args.output_dir, args.single_file)
        combined_wb.save(output_path)
        logger.info(f"\nCombined file saved: {output_path}")

    else:
        # Process each image to a separate Excel file
        results = []
        for i, image_path in enumerate(images, 1):
            name = Path(image_path).stem
            output_path = os.path.join(args.output_dir, f"{name}.xlsx")
            logger.info(f"[{i}/{len(images)}] Processing: {image_path}")

            path, success, msg = process_single_image(
                image_path, output_path,
                args.engine, args.lang, args.enhance, args.mode,
            )
            results.append((path, success, msg))
            status = "OK" if success else "FAILED"
            logger.info(f"  {status}: {msg}")

        # Summary
        success_count = sum(1 for _, s, _ in results if s)
        fail_count = len(results) - success_count
        print(f"\n--- Summary ---")
        print(f"  Processed: {len(results)}")
        print(f"  Success:   {success_count}")
        print(f"  Failed:    {fail_count}")

        if fail_count > 0:
            print("\nFailed images:")
            for path, success, msg in results:
                if not success:
                    print(f"  {path}: {msg}")


if __name__ == "__main__":
    main()
