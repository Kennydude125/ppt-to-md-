import argparse
import re
from pathlib import Path
from collections.abc import Iterable

# Fallbacks in case libraries are missing (assuming pip install requirements)
try:
    from markitdown import MarkItDown
except ImportError:
    MarkItDown = None

try:
    from docling.document_converter import DocumentConverter, PdfFormatOption
    from docling.datamodel.base_models import InputFormat
    from docling.datamodel.pipeline_options import PdfPipelineOptions
except ImportError:
    DocumentConverter = None

# Import our robust PPTX logic
try:
    from batch_markitdown import slide_to_markdown
    from pptx import Presentation
except ImportError:
    Presentation = None


def clean_markdown(md_text: str) -> str:
    """Universal Beautifier for tables, noise removal, and hierarchies."""
    if not md_text:
        return ""

    lines = md_text.splitlines()
    cleaned_lines = []
    
    for i, line in enumerate(lines):
        stripped = line.strip()
        
        # 1. Table Repair (Ensure correct spacing around |---|)
        # Not a complete markdown parser but enforces basic table pipe rules.
        if re.match(r'^\|[-:| ]+\|$', stripped):
            # Ensure it matches standard markdown table spacer
            cleaned_lines.append(stripped)
            continue
            
        # 2. Noise Removal (Delete recurring page numbers, headers, and footers)
        # Assuming simple integers or "Page X" on a single line are page numbers
        if re.match(r'^(page )?\d+$', stripped, re.IGNORECASE):
            continue
            
        # 3. Hierarchy Optimization (Automated newlines above headings)
        if re.match(r'^#{1,6}\s+', stripped):
            if cleaned_lines and cleaned_lines[-1].strip() != "":
                cleaned_lines.append("")  # Add extra newline above heading
            
        cleaned_lines.append(line)

    # 4. Final join and fixing multiple empty lines down to two
    final_text = "\n".join(cleaned_lines)
    final_text = re.sub(r'\n{3,}', '\n\n', final_text)
    return final_text.strip()


def convert_pptx(file_path: Path) -> str:
    """Convert PPTX using our custom grid-based spatial sorting logic."""
    if Presentation is None:
        raise ImportError("python-pptx and batch_markitdown requirements not fully met.")
    
    presentation = Presentation(str(file_path))
    slide_blocks: list[str] = []

    for slide_number, slide in enumerate(presentation.slides, start=1):
        slide_blocks.append(slide_to_markdown(slide, slide_number))

    return "\n\n".join(block for block in slide_blocks if block).strip()


def convert_pdf(file_path: Path) -> str:
    """Convert PDF using MarkItDown (more stable for large docs) or Docling as fallback."""
    # Prioritize MarkItDown as it safely retrieves embedded text without OOM errors
    if MarkItDown is not None:
        md = MarkItDown()
        return md.convert(str(file_path)).text_content
        
    if DocumentConverter is not None:
        try:
            pipeline_options = PdfPipelineOptions()
            pipeline_options.do_ocr = False
            pipeline_options.do_table_structure = True
            
            converter = DocumentConverter(
                format_options={
                    InputFormat.PDF: PdfFormatOption(pipeline_options=pipeline_options)
                }
            )
            doc = converter.convert(str(file_path))
            return doc.document.export_to_markdown()
        except Exception as e:
            print(f" [Docling failed: {e}]", end="", flush=True)
            return ""
            
    raise ImportError("No PDF conversion library available (need MarkitDown or Docling).")


def convert_docx(file_path: Path) -> str:
    """Convert DOCX, prioritizing heading styles to markdown mappings."""
    if MarkItDown is not None:
        # MarkItDown natively reads docx styles and maps them to Heading 1, 2, 3 nicely.
        md = MarkItDown()
        return md.convert(str(file_path)).text_content
        
    raise ImportError("No DOCX conversion library available (need MarkitDown).")


def convert_generic(file_path: Path) -> str:
    """Fallback converter for other supported files."""
    if MarkItDown is not None:
        md = MarkItDown()
        return md.convert(str(file_path)).text_content
    return ""


def process_file(src: Path, dst: Path):
    """Auto-detect extension and route to specific logic."""
    print(f"Processing: {src.name} -> ", end="", flush=True)
    
    ext = src.suffix.lower()
    raw_md = ""
    
    try:
        if ext == ".pptx":
            raw_md = convert_pptx(src)
        elif ext == ".pdf":
            raw_md = convert_pdf(src)
        elif ext in [".docx", ".doc"]:
            raw_md = convert_docx(src)
        else:
            raw_md = convert_generic(src)
            
        # Apply Universal Beautifier
        cleaned_md = clean_markdown(raw_md)
        
        dst.parent.mkdir(parents=True, exist_ok=True)
        dst.write_text(cleaned_md, encoding="utf-8")
        print("Success")
        return True
        
    except Exception as e:
        print(f"Failed ({e})")
        return False


def main():
    parser = argparse.ArgumentParser(description="Universal Batch Document to Markdown Converter.")
    parser.add_argument("--data-dir", default="data", type=Path)
    parser.add_argument("--output-dir", default="cleaned data", type=Path)
    args = parser.parse_args()

    data_dir = args.data_dir
    output_dir = args.output_dir

    if not data_dir.exists():
        print(f"Directory {data_dir} does not exist.")
        return

    success_count = 0
    fail_count = 0

    for src in data_dir.rglob("*"):
        if src.is_file() and src.suffix.lower() not in [".md", ".log", ".env"]:
            relative = src.relative_to(data_dir)
            dst = output_dir / relative.with_suffix(".md")
            
            if process_file(src, dst):
                success_count += 1
            else:
                fail_count += 1

    print("\n--- Summary ---")
    print(f"Successfully converted: {success_count}")
    print(f"Failed: {fail_count}")


if __name__ == "__main__":
    main()
