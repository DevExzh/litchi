# litchi-py - Python Bindings for Litchi

High-performance Python bindings for the Litchi Office file format parser. Parse Word documents, PowerPoint presentations, and Excel workbooks with ease.

## Features

- **Fast**: Built on Rust with zero-copy parsing and SIMD optimizations
- **Universal**: Supports legacy and modern Office formats (.doc/.docx, .ppt/.pptx, .xls/.xlsx)
- **Easy to Use**: Pythonic API inspired by python-docx and python-pptx
- **Type Safe**: Complete type stubs for excellent IDE support
- **Cross-Platform**: Works on Linux, macOS, and Windows
- **Python 3.8+**: Compatible with Python 3.8 and above using abi3

## Installation

### From PyPI (when published)

```bash
pip install litchi-py
```

### From Source

```bash
# Install maturin
pip install maturin

# Build and install in development mode
cd pyo3-litchi
maturin develop --release

# Or build a wheel
maturin build --release
pip install target/wheels/*.whl
```

## Quick Start

### Reading Word Documents

```python
from litchi_py import Document

# Open any Word document (.doc or .docx) - format auto-detected
doc = Document.open("document.docx")

# Extract all text
text = doc.text()
print(text)

# Access paragraphs
for para in doc.paragraphs():
    print(f"Paragraph: {para.text()}")
    
    # Access runs with formatting
    for run in para.runs():
        print(f"  Text: {run.text()}")
        if run.bold():
            print("    (bold)")

# Access tables
for table in doc.tables():
    print(f"Table with {table.row_count()} rows")
    for row in table.rows():
        for cell in row.cells():
            print(f"  Cell: {cell.text()}")
```

### Reading PowerPoint Presentations

```python
from litchi_py import Presentation

# Open any PowerPoint presentation (.ppt or .pptx)
pres = Presentation.open("presentation.pptx")

# Extract all text
text = pres.text()
print(text)

# Get slide count
print(f"Total slides: {pres.slide_count()}")

# Access individual slides
for i, slide in enumerate(pres.slides()):
    print(f"Slide {i + 1}: {slide.text()}")
```

### Reading Excel Workbooks

```python
from litchi_py import Workbook

# Open an Excel workbook (.xls, .xlsx, .xlsb)
wb = Workbook.open("workbook.xlsx")

# Get worksheet count
print(f"Worksheets: {wb.worksheet_count()}")

# Access worksheets
for ws in wb.worksheets():
    print(f"Sheet: {ws.name()}")
    print(f"  Rows: {ws.row_count()}")
    print(f"  Cols: {ws.column_count()}")
    
    # Get cell value
    value = ws.cell_value(0, 0)  # Row 0, Column 0
    if value:
        print(f"  A1: {value}")
    
    # Get all rows
    for row in ws.rows():
        print(row)

# Get worksheet by name
sheet = wb.worksheet_by_name("Sheet1")
if sheet:
    print(f"Found sheet: {sheet.name()}")
```

### Format Detection

```python
from litchi_py import detect_file_format, FileFormat

# Detect format from file path
fmt = detect_file_format("document.docx")
print(fmt)  # FileFormat.Docx

# Detect format from bytes
with open("presentation.pptx", "rb") as f:
    data = f.read()
    fmt = detect_file_format_from_bytes(data)
    print(fmt)  # FileFormat.Pptx
```

## API Reference

### Document API

- **`Document.open(path)`**: Open a Word document
- **`Document.text()`**: Extract all text
- **`Document.paragraphs()`**: Get all paragraphs
- **`Document.tables()`**: Get all tables
- **`Paragraph.text()`**: Get paragraph text
- **`Paragraph.runs()`**: Get text runs
- **`Run.text()`**: Get run text
- **`Run.bold()`**: Check if text is bold
- **`Run.italic()`**: Check if text is italic
- **`Run.underline()`**: Check if text is underlined
- **`Table.row_count()`**: Get number of rows
- **`Table.rows()`**: Get all rows
- **`TableRow.cells()`**: Get all cells
- **`TableCell.text()`**: Get cell text

### Presentation API

- **`Presentation.open(path)`**: Open a PowerPoint presentation
- **`Presentation.text()`**: Extract all text
- **`Presentation.slide_count()`**: Get number of slides
- **`Presentation.slides()`**: Get all slides
- **`Slide.text()`**: Get slide text

### Workbook API

- **`Workbook.open(path)`**: Open an Excel workbook
- **`Workbook.worksheet_count()`**: Get number of worksheets
- **`Workbook.worksheets()`**: Get all worksheets
- **`Workbook.worksheet_by_name(name)`**: Get worksheet by name
- **`Worksheet.name()`**: Get worksheet name
- **`Worksheet.row_count()`**: Get number of rows
- **`Worksheet.column_count()`**: Get number of columns
- **`Worksheet.cell_value(row, col)`**: Get cell value
- **`Worksheet.rows()`**: Get all rows

### Utility Functions

- **`detect_file_format(path)`**: Detect file format from path
- **`detect_file_format_from_bytes(data)`**: Detect format from bytes

## Supported Formats

| Format | Extension | Read Support |
|--------|-----------|--------------|
| Microsoft Word 97-2003 | .doc | ✅ |
| Microsoft Word 2007+ | .docx | ✅ |
| Microsoft PowerPoint 97-2003 | .ppt | ✅ |
| Microsoft PowerPoint 2007+ | .pptx | ✅ |
| Microsoft Excel 97-2003 | .xls | ✅ |
| Microsoft Excel 2007+ | .xlsx | ✅ |
| Microsoft Excel Binary | .xlsb | ✅ |
| OpenDocument Text | .odt | ✅ |
| OpenDocument Spreadsheet | .ods | ✅ |
| OpenDocument Presentation | .odp | ✅ |
| Apple Pages | .pages | ✅ |
| Apple Keynote | .key | ✅ |
| Apple Numbers | .numbers | ✅ |
| Rich Text Format | .rtf | ✅ |

## Performance

Litchi is built on Rust and uses:
- Zero-copy parsing where possible
- SIMD instructions for text processing
- Efficient memory management
- Parallel processing for large files

This results in performance that's often **10-100x faster** than pure Python implementations.

## Development

### Building from Source

```bash
# Install dependencies
pip install maturin

# Build in development mode (faster, with debug symbols)
maturin develop

# Build in release mode (optimized)
maturin develop --release

# Run tests
pytest tests/
```

### Project Structure

```
pyo3-litchi/
├── src/
│   ├── lib.rs          # Main module entry point
│   ├── common.rs       # Common types and utilities
│   ├── document.rs     # Document API bindings
│   ├── presentation.rs # Presentation API bindings
│   └── sheet.rs        # Workbook API bindings
├── python/
│   └── litchi_py/
│       ├── __init__.pyi # Type stubs
│       └── py.typed     # PEP 561 marker
├── Cargo.toml          # Rust dependencies
├── pyproject.toml      # Python project config
└── README.md           # This file
```

## Type Checking

Full type stubs are included for excellent IDE support and type checking with mypy:

```bash
pip install mypy
mypy your_script.py
```

## License

This project is licensed under Apache License, Version 2.0 ([LICENSE](../LICENSE))

## Contributing

Contributions are welcome! Please feel free to submit a Pull Request.

## Related Projects

- [Litchi](https://github.com/DevExzh/litchi) - The main Rust library
- [python-docx](https://python-docx.readthedocs.io/) - Pure Python DOCX library
- [python-pptx](https://python-pptx.readthedocs.io/) - Pure Python PPTX library
- [openpyxl](https://openpyxl.readthedocs.io/) - Pure Python XLSX library

