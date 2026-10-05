# litchi-py - Python Bindings for Litchi

High-performance Python bindings for the Litchi Office file format parser. Parse Word documents, PowerPoint presentations, and Excel workbooks with ease.

## Features

- **Native**: The parsing runs in Rust; opening a file and the document-wide reads release the GIL
- **Universal**: Supports legacy and modern Office formats (.doc/.docx, .ppt/.pptx, .xls/.xlsx/.xlsb), OpenDocument and Apple iWork
- **Easy to Use**: Pythonic API inspired by python-docx and python-pptx
- **Thread-Safe**: Handles are immutable and can be used from any thread
- **Typed Errors**: Every error is a `LitchiError` with a stable `code`, a `message` and a `suggestion`
- **Type Safe**: Complete type stubs for excellent IDE support
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

# Open a workbook (.xls, .xlsx, .xlsb, .ods, .numbers)
wb = Workbook.open("workbook.xlsx")

# Get worksheet count and names
print(f"Worksheets: {wb.worksheet_count()}")
for name in wb.worksheet_names():
    print(f"Sheet: {name}")

# Cell values, one row per line with tab-separated cells
# (.numbers lists each sheet, table and cell on its own line)
print(wb.text())
```

There is no per-worksheet or per-cell API yet: `Worksheet` is a placeholder
that no method returns.

### Handling Errors

```python
from litchi_py import Document, LitchiError, UnsupportedError

try:
    doc = Document.open("slides.pptx")
except UnsupportedError as error:
    print(error.code)        # "unsupported_format"
    print(error.message)     # "detected Pptx, which this operation does not handle"
    print(error.suggestion)  # "open it with Presentation.open"
except OSError as error:     # IoError is an OSError
    print(f"cannot read the file: {error}")
except LitchiError as error:
    print(f"{error.code}: {error}")
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

- **`Workbook.open(path)`**: Open a workbook
- **`Workbook.worksheet_count()`**: Get number of worksheets
- **`Workbook.worksheet_names()`**: Get worksheet names
- **`Workbook.text()`**: Extract all cell text

### Utility Functions

- **`detect_file_format(path)`**: Detect file format from path
- **`detect_file_format_from_bytes(data)`**: Detect format from bytes

### Exceptions

Every error is a `LitchiError` with `code`, `message` and `suggestion`
attributes; `str(error)` is the message followed by the suggestion. Each
family also derives from the built-in exception that fits it:

| Class | Also a | Codes |
|-------|--------|-------|
| `LitchiError` | `Exception` | the base class; raised itself for `limit_exceeded`, `internal_error`, `cancelled` and the other codes outside a family |
| `IoError` | `OSError` | `file_not_found`, `io_error` (also a file changed after it was opened), `file_locked`, `file_exists`, `file_required` |
| `UnsupportedError` | `ValueError` | `unsupported_format` (not an Office file, or another opener's format), `unsupported_capability` (for example flat OpenDocument), `unsupported_element`, `unsupported_type`, `unsupported_command` |
| `CorruptFileError` | `ValueError` | `corrupt_file`, `invalid_input` |
| `ProtectedError` | `ValueError` | `document_protected` (encrypted or password-protected) |
| `PropertyError` | `ValueError` | `invalid_value`, `unsupported_property`, `missing_property`, `duplicate_name`, `invalid_argument`, `missing_argument`, `invalid_json`, `corrupt_batch` |
| `PathError` | `LookupError` | `not_found`, `path_not_found`, `zero_matches`, `out_of_range`, `range_target_not_found`, `invalid_path`, `invalid_selector`, `invalid_xpath`, `unsupported_path`, `bare_selector_rejected` |

The read API raises the first five; `PropertyError` and `PathError` are
reserved for the editing API.

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

## Threads and the GIL

Handles are immutable, so a `Document`, `Presentation` or `Workbook` and the
paragraphs, runs, tables and slides read from it can be shared between
threads. Opening a file and the document-wide reads (`text()`, `paragraphs()`,
`tables()`, `slides()`, `slide_count()`, `worksheet_names()`,
`worksheet_count()`) and the detection functions release the GIL while they
run, so other Python threads keep running. Litchi starts no threads of its
own.

## Development

### Building from Source

```bash
# Install dependencies
pip install maturin

# Build in development mode (faster, with debug symbols)
maturin develop

# Build in release mode (optimized)
maturin develop --release

# Run tests (from the repository root; builds with cargo, needs no maturin)
crates/pyo3-litchi/tests/run_tests.sh
crates/pyo3-litchi/tests/run_tests.sh --release
```

### Project Structure

```
pyo3-litchi/
├── src/
│   ├── lib.rs          # Main module entry point
│   ├── common.rs       # Common types and utilities
│   ├── errors.rs       # LitchiError and its subclasses
│   ├── iwork.rs        # Pages, Keynote and Numbers readers
│   ├── document.rs     # Document API bindings
│   ├── presentation.rs # Presentation API bindings
│   └── sheet.rs        # Workbook API bindings
├── python/
│   └── litchi_py/
│       ├── __init__.pyi # Type stubs
│       └── py.typed     # PEP 561 marker
├── tests/              # unittest suite (run_tests.sh)
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

