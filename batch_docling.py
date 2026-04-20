from __future__ import annotations

import argparse
from pathlib import Path
from typing import Iterable

from docling.document_converter import DocumentConverter
from docling.datamodel.base_models import InputFormat


SKIP_EXTENSIONS = {".md"}


def build_converter() -> DocumentConverter:
    """Create a Docling converter configured to avoid OCR-heavy PDF rendering."""
    converter = DocumentConverter()
    pdf_options = converter.format_to_options[InputFormat.PDF].pipeline_options
    pdf_options.do_ocr = False
    pdf_options.generate_page_images = False
    pdf_options.generate_picture_images = False
    pdf_options.generate_table_images = False
    pdf_options.generate_parsed_pages = False
    pdf_options.do_picture_classification = False
    pdf_options.do_picture_description = False
    pdf_options.do_chart_extraction = False
    pdf_options.do_code_enrichment = False
    pdf_options.do_formula_enrichment = False
    pdf_options.force_backend_text = True
    return converter


def iter_input_files(data_dir: Path) -> Iterable[Path]:
    """Yield all files under data_dir recursively."""
    for path in data_dir.rglob("*"):
        if path.is_file() and path.suffix.lower() not in SKIP_EXTENSIONS:
            yield path


def convert_file(converter: DocumentConverter, src: Path, dst: Path) -> None:
    """Convert one source file to markdown and write to dst."""
    result = converter.convert(str(src))
    markdown = result.document.export_to_markdown()

    dst.parent.mkdir(parents=True, exist_ok=True)
    dst.write_text(markdown, encoding="utf-8")


def process_all(data_dir: Path, output_dir: Path) -> tuple[int, int]:
    """Process all files and return (success_count, fail_count)."""
    converter = build_converter()
    success_count = 0
    fail_count = 0

    for src in iter_input_files(data_dir):
        relative = src.relative_to(data_dir)
        dst = output_dir / relative.with_suffix(".md")

        try:
            convert_file(converter, src, dst)
            success_count += 1
            print(f"[OK]   {src} -> {dst}")
        except Exception as exc:  # Keep batch running even if one file fails.
            fail_count += 1
            print(f"[FAIL] {src}: {exc}")

    return success_count, fail_count


def main() -> int:
    parser = argparse.ArgumentParser(
        description="Batch convert files in data folder to markdown using Docling."
    )
    parser.add_argument(
        "--data-dir",
        default="data",
        type=Path,
        help="Input folder that contains files to convert (default: data)",
    )
    parser.add_argument(
        "--output-dir",
        default="cleaned data",
        type=Path,
        help="Output folder where markdown files are written (default: cleaned data)",
    )
    args = parser.parse_args()

    data_dir = args.data_dir
    output_dir = args.output_dir

    if not data_dir.exists() or not data_dir.is_dir():
        print(f"Input folder not found: {data_dir}")
        return 1

    output_dir.mkdir(parents=True, exist_ok=True)

    success, failed = process_all(data_dir, output_dir)

    print("\nDone.")
    print(f"Converted: {success}")
    print(f"Failed:    {failed}")

    # Non-zero exit when any conversion failed.
    return 0 if failed == 0 else 2


if __name__ == "__main__":
    raise SystemExit(main())
