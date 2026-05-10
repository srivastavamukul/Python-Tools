# Python Tools

This repository collects small Python utilities designed to automate repetitive tasks and make daily workflows easier.

## Current tool

- `ppt2pdf_batch.py`

This script converts all `.pptx` files from one folder into PDF files in another folder.

### Features

- Batch converts PowerPoint presentations to PDF
- Uses Windows PowerPoint automation via `pywin32`
- Stores output files in a dedicated output folder

### Requirements

- Windows
- Microsoft PowerPoint installed
- Python 3.x
- `pywin32` package

### Install dependencies

```bash
pip install pywin32
```

### Usage

1. Open `ppt2pdf_batch.py`.
2. Update the `input_folder` and `output_folder` variables with your paths.
3. Run the script:

```bash
python ppt2pdf_batch.py
```

### Example

```python
input_folder = r"C:\path\to\your\pptx_files"
output_folder = r"C:\path\to\save\pdfs"
```

## Project structure

For a single utility, the repo root is fine. If this repo grows to include multiple tools, consider organizing them in a dedicated folder such as `tools/` or `src/`.

Example structure for more tools:

- `README.md`
- `requirements.txt`
- `tools/`
  - `ppt2pdf_batch.py`
  - `another_tool.py`

## Future plans

- Add more small, practical utilities
- Keep each tool simple and easy to run
- Document new tools clearly in this README

---

Built to make everyday work easier with simple Python automation.