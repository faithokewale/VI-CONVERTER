# converter.py

Converts academic documents (PDF, DOCX, PPTX) into clean, structured Markdown.

## Objective

Turn lecture notes, past questions, handouts, assignments, and slide decks
into consistent `.md` files — no database, no API, no UI. Just: file in →
clean Markdown out.

## Supported input files

| Format |
|--------|
| `.pdf` | 
| `.docx`|
| `.pptx`|


## How it works

1. Script receives a file (by path, or via a file picker)
2. Extracts raw text from the document (PDF font-size heuristics, DOCX
   paragraph styles, PPTX slide titles/body text)
3. Cleans the text (removes running headers/footers, page numbers, fixes
   line-wrap hyphens) and structures it into headings + bullets/paragraphs
4. Saves the result as a `.md` file

## Installation

```bash
pip install -r requirements.txt
```

Requires `pdfplumber`, `python-docx`, and `python-pptx`. The file-picker
uses `tkinter`, which ships with Python on Windows by default.

## Usage

```bash
python converter.py                          # opens a file picker
python converter.py path/to/file.pdf
python converter.py path/to/file.docx -o somewhere_else
python converter.py path/to/file.pptx --course CSC401 --title "Networks Intro"
```

## Output location

By default, every converted `.md` file is saved into an `output/` folder
next to `converter.py` — created automatically the first time you run it —
regardless of where the source file lives. Pass `-o` to save somewhere else.