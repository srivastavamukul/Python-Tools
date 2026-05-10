# Python Tools

This repository is a collection of small Python utilities I build to make my daily work more efficient. The goal is to keep useful tools simple, practical, and easy to run so I can solve real problems quickly.

## Why this repo exists

I create these tools to automate repetitive tasks and improve efficiency in my day-to-day life. The first tool in this repo converts PowerPoint files (`.pptx`) into PDF documents, because the study material from college is often shared as slides, which I find harder to review than PDFs.

## Current tool: `ppt2pdf_batch.py`

This script converts all `.pptx` files from one folder into PDF files in another folder.

### Features

- Converts PowerPoint presentations to PDF in batch
- Uses Windows PowerPoint automation via `pywin32`
- Keeps output files organized in a separate folder

### Requirements

- Windows
- Microsoft PowerPoint installed
- Python
- `pywin32` package

### Installation

```bash
pip install pywin32
```

### Usage

1. Update the `input_folder` and `output_folder` paths in `ppt2pdf_batch.py`.
2. Run the script:

```bash
python ppt2pdf_batch.py
```

### Example

```python
input_folder = r"C:\path\to\your\pptx_files"
output_folder = r"C:\path\to\save\pdfs"
```

## Future tools

This repo is intended to grow with more small utilities that help me with study workflows and daily tasks. Each tool will be added here when it solves a real problem.

---

Created to make everyday work easier by building simple, practical Python tools.