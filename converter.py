
import argparse
import re
import sys
from collections import Counter
from pathlib import Path

import pdfplumber
from docx import Document
from pptx import Presentation



def extract_pdf(path):
    """
    Pull text out of a PDF line by line, using relative font size to guess
    which lines are headings. This is heuristic — PDFs carry no semantic
    structure — but it works well for lecture-note style documents where
    headings are visibly larger/bolder than body text.
    """
    raw_lines = []      # (text, size, page_no)
    with pdfplumber.open(path) as pdf:
        for page_no, page in enumerate(pdf.pages):
            # x_tolerance lowered from pdfplumber's default (3): some PDFs —
            # notably LaTeX-generated academic papers — use tighter
            # character spacing than the default assumes, which otherwise
            # merges entire lines into a single run-on "word".
            words = page.extract_words(extra_attrs=["size"], x_tolerance=2)
            if not words:
                continue
            lines = {}
            for w in words:
                key = round(w["top"], 1)
                lines.setdefault(key, []).append(w)
            for top in sorted(lines.keys()):
                line_words = sorted(lines[top], key=lambda w: w["x0"])
                text = " ".join(w["text"] for w in line_words).strip()
                if not text:
                    continue
                avg_size = sum(w["size"] for w in line_words) / len(line_words)
                raw_lines.append((text, avg_size, page_no))

    if not raw_lines:
        return []

    # Body text size = the most common font size in the document.
    body_size = Counter(round(size) for _, size, _ in raw_lines).most_common(1)[0][0]

    # Drop lines that repeat verbatim across many pages (running headers/footers).
    page_count = len({p for _, _, p in raw_lines})
    text_freq = Counter(text for text, _, _ in raw_lines)
    boilerplate = {
        t for t, c in text_freq.items()
        if page_count > 1 and c >= max(2, page_count * 0.4) and len(t) < 80
    }

    blocks = []
    for text, size, _ in raw_lines:
        if text in boilerplate:
            continue
        if re.fullmatch(r"(page\s*)?\d+(\s*(of|/)\s*\d+)?", text, flags=re.I):
            continue  # bare page numbers
        is_heading = size >= body_size + 2  # noticeably larger than body text
        bullet_match = re.match(r"^([•\-\*▪○◦]|\d+[\.\)])\s+(.*)", text)
        is_bullet = bool(bullet_match)
        if bullet_match:
            text = bullet_match.group(2)
        blocks.append({"text": text, "is_heading": is_heading, "is_bullet": is_bullet})
    return blocks


def extract_docx(path):
    """Extract paragraphs from a DOCX, using paragraph style to find headings and lists."""
    doc = Document(path)
    blocks = []
    for para in doc.paragraphs:
        text = para.text.strip()
        if not text:
            continue
        style = (para.style.name or "").lower()
        is_heading = style.startswith("heading") or style == "title"
        is_bullet = "list" in style  # "List Bullet", "List Number", "List Paragraph", etc.
        blocks.append({"text": text, "is_heading": is_heading, "is_bullet": is_bullet})

    # Tables often hold problem sets / data in past questions and handouts.
    for table in doc.tables:
        for row in table.rows:
            cells = [c.text.strip() for c in row.cells if c.text.strip()]
            if cells:
                blocks.append({"text": " | ".join(cells), "is_heading": False, "is_bullet": True})
    return blocks


def extract_pptx(path):
    """Extract slide titles (as headings) and body text (as content)."""
    prs = Presentation(path)
    blocks = []
    for slide in prs.slides:
        title_shape = slide.shapes.title
        title_id = title_shape.shape_id if title_shape else None
        title_text = title_shape.text.strip() if (title_shape and title_shape.text) else None
        if title_text:
            blocks.append({"text": title_text, "is_heading": True, "is_bullet": False})

        for shape in slide.shapes:
            if shape.shape_id == title_id or not shape.has_text_frame:
                continue
            for para in shape.text_frame.paragraphs:
                text = "".join(run.text for run in para.runs).strip()
                if text:
                    blocks.append({"text": text, "is_heading": False, "is_bullet": True})

        if slide.has_notes_slide:
            notes = (slide.notes_slide.notes_text_frame.text or "").strip()
            if notes:
                blocks.append({"text": f"Speaker notes: {notes}", "is_heading": False, "is_bullet": False})
    return blocks


EXTRACTORS = {".pdf": extract_pdf, ".docx": extract_docx, ".pptx": extract_pptx}

# Every converted .md file lands here by default, regardless of where the
# input file lives or where the script is run from.
DEFAULT_OUTPUT_DIR = Path(__file__).resolve().parent / "output"


# ---------------------------------------------------------------------------
# Step 2b: Cleaning
# ---------------------------------------------------------------------------

def clean_text(text):
    """Collapse whitespace and fix common PDF hyphenation breaks."""
    text = re.sub(r"(\w)-\s*$", r"\1", text)      # trailing hyphen from line wrap
    text = re.sub(r"\s+", " ", text).strip()
    return text


def clean_blocks(blocks):
    cleaned = []
    for b in blocks:
        text = clean_text(b["text"])
        if not text:
            continue
        cleaned.append({"text": text, "is_heading": b["is_heading"], "is_bullet": b.get("is_bullet", False)})
    return cleaned


# ---------------------------------------------------------------------------
# Step 3: Structuring into the required Markdown format
# ---------------------------------------------------------------------------

COURSE_CODE_RE = re.compile(r"\b([A-Z]{2,5})\s?-?\s?(\d{3})\b")


def guess_course_code(blocks, filename):
    """Look for a course code (e.g. CSC401) in the filename, then the body."""
    for source in [filename] + [b["text"] for b in blocks[:15]]:
        match = COURSE_CODE_RE.search(source.upper())
        if match:
            return f"{match.group(1)}{match.group(2)}"
    return "UNKNOWN"


def guess_title(blocks, filename):
    """Use the first heading if there is one, else fall back to the filename."""
    for b in blocks:
        if b["is_heading"]:
            # Strip a leading course code from the title if present.
            stripped = COURSE_CODE_RE.sub("", b["text"]).strip(" -–—:")
            return stripped or b["text"]
    return Path(filename).stem.replace("_", " ").replace("-", " ").title()


def structure_markdown(blocks, course_code, title):
    """
    Group blocks into (heading, content) sections and render the required
    Markdown structure:

        # COURSE CODE — Title

        ## Section Heading
        - Sub-topic one
        - Sub-topic two

    List-like content (bullets, slide body text, table rows) is rendered as
    a Markdown list, mirroring how the source document presented it. Plain
    prose content is kept as its own paragraph line instead of being merged
    with everything else.
    """
    sections = []
    current_heading = None
    current_content = []

    def flush():
        if current_heading is not None and current_content:
            sections.append((current_heading, current_content))

    for b in blocks:
        if b["is_heading"]:
            flush()
            current_heading = b["text"]
            current_content = []
        else:
            if current_heading is None:
                # Content appearing before the first detected heading.
                current_heading = "Overview"
            current_content.append((b["text"], b.get("is_bullet", False)))
    flush()

    lines = [f"# {course_code} — {title}", ""]
    for heading, content_items in sections:
        lines.append(f"## {heading}")
        for text, is_bullet in content_items:
            lines.append(f"- {text}" if is_bullet else text)
        lines.append("")
    return "\n".join(lines).rstrip() + "\n"


# ---------------------------------------------------------------------------
# Step 4: Save + CLI entry point
# ---------------------------------------------------------------------------

def convert(input_path, output_dir=None, course=None, title=None):
    input_path = Path(input_path)
    ext = input_path.suffix.lower()
    if ext not in EXTRACTORS:
        raise ValueError(f"Unsupported file type '{ext}'. Supported: {', '.join(EXTRACTORS)}")

    raw_blocks = EXTRACTORS[ext](input_path)
    blocks = clean_blocks(raw_blocks)
    if not blocks:
        raise ValueError(f"No extractable text found in {input_path.name}")

    course_code = course or guess_course_code(blocks, input_path.name)
    doc_title = title or guess_title(blocks, input_path.name)
    markdown = structure_markdown(blocks, course_code, doc_title)

    out_dir = Path(output_dir) if output_dir else DEFAULT_OUTPUT_DIR
    out_dir.mkdir(parents=True, exist_ok=True)
    out_path = out_dir / f"{input_path.stem}.md"
    out_path.write_text(markdown, encoding="utf-8")
    return out_path


def pick_file_dialog():
    """Open the OS's native 'Open File' dialog and return the chosen path, or None if cancelled."""
    try:
        import tkinter as tk
        from tkinter import filedialog
    except ImportError:
        print(
            "Error: the file-picker needs tkinter, which isn't installed.\n"
            "Either install it (e.g. 'sudo apt install python3-tk' on Linux) "
            "or pass a file path directly: python converter.py path/to/file.pdf",
            file=sys.stderr,
        )
        sys.exit(1)

    root = tk.Tk()
    root.withdraw()          # hide the empty main window, we only want the dialog
    root.attributes("-topmost", True)  # bring the dialog to the front
    path = filedialog.askopenfilename(
        title="Select a document to convert",
        filetypes=[
            ("Supported documents", "*.pdf *.docx *.pptx"),
            ("PDF files", "*.pdf"),
            ("Word documents", "*.docx"),
            ("PowerPoint files", "*.pptx"),
            ("All files", "*.*"),
        ],
    )
    root.destroy()
    return path or None


def main():
    parser = argparse.ArgumentParser(description="Convert academic PDF/DOCX/PPTX files to clean Markdown.")
    parser.add_argument("input_file", nargs="?", default=None,
                         help="Path to a .pdf, .docx, or .pptx file. Omit this to open a file picker instead.")
    parser.add_argument("-o", "--output-dir", default=None, help="Folder to save the .md file (default: an 'output' folder next to this script)")
    parser.add_argument("--course", default=None, help="Override the detected course code (e.g. CSC401)")
    parser.add_argument("--title", default=None, help="Override the detected document title")
    args = parser.parse_args()

    input_file = args.input_file
    if input_file is None:
        input_file = pick_file_dialog()
        if not input_file:
            print("No file selected. Exiting.")
            sys.exit(0)

    try:
        out_path = convert(input_file, args.output_dir, args.course, args.title)
    except Exception as e:
        print(f"Error: {e}", file=sys.stderr)
        sys.exit(1)

    print(f"Saved: {out_path}")


if __name__ == "__main__":
    main()
