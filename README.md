# Publisher to HTML Converter

A Python script that converts Microsoft Publisher (.pub) files to HTML format using COM automation.

Note: While the repository is named "pubtopdf", the tool currently focuses on HTML conversion as the primary output format since this proved to be the most reliable method.

## Requirements

- Windows with Microsoft Publisher installed
- Python 3.6 or later
- Dependencies from requirements.txt

## Installation

```bash
python -m venv .venv
.venv\Scripts\activate
pip install -r requirements.txt
```

## Usage

```bash
python convert.py "input/document.pub"
```

This will create HTML output in the `output` directory with the same base name as the input file.

### Optional Arguments

- `--output-dir`: Directory to save output files in (default: output)
- `--format-constant`: Format constant to use with SaveAs (default: 7 for HTML)

## Format Constants Reference

The following format constants have been tested with Publisher's SaveAs method:

| Constant | Output Format | Notes |
|----------|--------------|-------|
| 7 | HTML | Creates .htm file and supporting _files directory |
| 8 | Text | Creates plain text (.txt) file |
| 9 | Text | Creates plain text (.txt) file |
| 1-6, 10-20 | - | COM Error (-2147024809) |

HTML (constant 7) is the primary supported format for this tool.