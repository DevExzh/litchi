# litchi-py - Python Bindings for Litchi

High-performance Python bindings for the Litchi Office file format parser. Parse Word documents, PowerPoint presentations, and Excel workbooks with ease.

## Features

- **Native**: The parsing runs in Rust; opening a file and the document-wide reads release the GIL
- **Universal**: Supports legacy and modern Office formats (.doc/.docx, .ppt/.pptx, .xls/.xlsx/.xlsb), OpenDocument and Apple iWork
- **Unified API**: `litchi_py.open()` returns a `File` that answers OfficeCLI's paths (`/body/p[3]`, `/Sheet1/B2`), selectors (`paragraph[style=Heading1]`) and views, with results as plain dicts, and edits them: `set`, `add`, `remove`, `move`, `swap`, transactions and OfficeCLI batches, each returning a new `File`
- **Easy to Use**: Pythonic API inspired by python-docx and python-pptx
- **Thread-Safe**: Handles are immutable and can be used from any thread
- **Typed Errors**: Every error is a `LitchiError` with a stable `code`, a `message` and a `suggestion`
- **Type Safe**: Complete type stubs for excellent IDE support
- **Python 3.9+**: Compatible with Python 3.9 and above using abi3

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

### The Unified API

`litchi_py.open(path)` (or `open_bytes(data)`) detects the format from the
content and returns a `File`, an immutable handle that is also a context
manager. Its reads use OfficeCLI's addressing and return OfficeCLI's node view
as dicts: `path`, `type`, `text`, `preview`, `style`, `childCount`, `format`
(the element's own properties) and `children`.

```python
import litchi_py

with litchi_py.open("report.docx") as doc:
    print(doc.format)                        # FileFormat.Docx

    paragraph = doc.get("/body/p[1]")        # the node view, one level of children
    print(paragraph["text"], paragraph["format"])
    for run in paragraph["children"]:
        print(run["path"], run["format"].get("bold"))

    doc.find("/body/p[999]")                 # None: nothing there
    headings = doc.query("paragraph[style=Heading1]", limit=10)
    print([heading["path"] for heading in headings])

    print(doc.view("outline"))               # text; view("stats", json=True) is a dict
    print(doc.capabilities()["verbs"]["get"])

with litchi_py.open("data.xlsx") as book:
    cell = book.get("Sheet1!B2", depth=0)    # same as "/Sheet1/B2"
    rows = book.query("row", depth=1)        # each row with its cells
    print(book.view("text"))
```

`depth` is the number of child levels a view lists (default 1 for `get` and
`find`, 0 for `query`; at most 64), and `max_nodes` bounds the nodes one call
returns (default 100,000; more is `LitchiError` with code `output_too_large`).
The unified reader covers .docx, .xlsx, .pptx, .xlsb, .doc, .xls, .ppt and
.rtf; .odt, .ods and .odp open and answer `text()`, `metadata()` and
`capabilities()`, while their `get`, `find`, `query` and `view` raise
`UnsupportedError` (`unsupported_capability`) until their readers land, as
`capabilities()` reports. Apple Pages, Keynote and Numbers packages open as
`File` too (record 0296), and so does an app-authored package directory of
any of the three: the `File` holds its single-file form, and `save()` writes
it back as a directory (records 0299, 0301 and 0306). The family classes
below read iWork inputs through the same `File`.

### Editing

A `File` never changes: every write verb returns a new `File`, and the old one
stays readable. Properties are plain dicts with OfficeCLI's names (dotted keys
are fine) or OfficeCLI's `["key=value"]` list; values keep their Python types
(`None` clears a property, text is coerced through the key's type: `"12pt"`,
`"1F4E79"`).

```python
import litchi_py

doc = litchi_py.open("report.docx")
doc = doc.set("/body/p[1]/r[1]", {"bold": True, "size": 14, "color": "1F4E79"})
doc, path = doc.add("/body", "paragraph", {"text": "Appendix"}, after="/body/p[1]")
doc = doc.move(path, index=0).swap("/body/p[1]", "/body/p[2]").remove("/body/p[3]")

# A transaction: everything or nothing.
with doc.edit() as tx:
    path = tx.add("/body", "paragraph", {"text": "Summary"})
    tx.set(path, {"align": "center"})
doc = tx.commit().file

# OfficeCLI batch items, as dicts or as JSON text; "best_effort" reports
# failed items in commit.results instead of raising.
commit = doc.batch([
    {"command": "set", "path": "/body/p[1]", "props": {"align": "right"}},
    {"command": "add", "parent": "/body", "type": "paragraph", "props": {"text": "End"}},
])
doc = commit.file
doc.save()                          # atomic, back to report.docx
doc.save_as("copy.docx")            # or anywhere; to_bytes() for the bytes

book = litchi_py.create("xlsx").set("/Sheet1/A1", {"value": 42})
```

`set`, `add`, `remove`, `move`, `swap`, `edit`, `batch`, `save` and `create`
work on .docx, .xlsx, .pptx, .xlsb and .rtf, and on the subsets of .doc, .xls
and .ppt that `capabilities()` lists; a verb or element a format does not
support raises `UnsupportedError` with the code the report gives. A signed
file refuses edits (`ProtectedError`) until `unsign()` returns a `File` whose
first write drops the signature; an encrypted file opens locked and refuses
every read and write.

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

### Unified API

- **`open(path)`**: Open a file through the unified API
- **`open_bytes(data)`**: Open in-memory bytes through the unified API
- **`File.get(path, depth=1)`**: The node view of the element at a path (`PathError` when nothing is there)
- **`File.find(path, depth=1)`**: The node view at a path, or `None`
- **`File.query(selector, limit=None, depth=0)`**: The node views of every match, in document order
- **`File.view(mode, json=False)`**: The `text`, `annotated`, `outline` or `stats` view, as text or as a dict
- **`File.capabilities()`**: What the file's format answers (OfficeCLI's help schema), as a dict
- **`File.text()`**: Extract all text
- **`File.metadata()`**: The document properties, as a dict
- **`File.close()`**: Close the file (also on leaving a `with` block)
- `File.format`, `File.closed`, `File.signed`, `File.locked`: The detected format; whether the file is closed, signed or locked
- **`create(format)`**: A new, empty document of a format (`"docx"` or a `FileFormat`)
- **`File.set(target, props)`**: Set properties; returns the new `File`
- **`File.add(parent, type, props=None, index=, after=, before=)`**: Add an element; returns `(File, path)`
- **`File.remove(target)`**, **`File.move(path, to=None, index=, after=, before=)`**, **`File.swap(path, path2)`**: Structural edits; each returns the new `File`
- **`File.edit(fn=None)`**: A transaction: an `Edit` (a context manager), or run `fn(edit)` and return its `Commit`
- **`File.batch(items, mode="atomic")`**, **`File.apply(item)`**: OfficeCLI batch items (dicts or JSON text); return a `Commit`
- **`File.save(path=None, overwrite=True)`**, **`File.save_as(path)`**, **`File.to_bytes()`**: Publish in the file's own format, atomically
- **`File.unsign()`**: A `File` that may edit a signed document
- **`Edit.set(...)`**, **`Edit.add(...)`**, **`Edit.remove(...)`**, **`Edit.move(...)`**, **`Edit.swap(...)`**, **`Edit.get(...)`**, **`Edit.query(...)`**, **`Edit.commit()`**: The transaction's verbs
- `Commit.file`, `Commit.changed`, `Commit.operations`, `Commit.results`, `Commit.value`: A transaction's or batch's result
- `Commit.patch`: The change's reversible `Patch` (None for `add_part` and `raw_set`): `Patch.apply(file)` replays it, `Patch.inverse()` undoes it, `Patch.is_noop()`; the original File and clones use the fast identity path, while a same-format independently opened File is accepted when its published base passes the bounded content proof; changed content, protected state and stale sources raise the corresponding typed error

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

Every error is a `LitchiError` with `code`, `message`, `suggestion` and `item`
attributes (`item` is the zero-based batch item or transaction operation that
failed, else `None`); `str(error)` is the message followed by the suggestion. Each
family also derives from the built-in exception that fits it:

| Class | Also a | Codes |
|-------|--------|-------|
| `LitchiError` | `Exception` | the base class; raised itself for `limit_exceeded`, `internal_error`, `cancelled` and the other codes outside a family |
| `IoError` | `OSError` | `file_not_found`, `io_error` (also a file changed after it was opened), `file_locked`, `file_exists`, `file_required` |
| `UnsupportedError` | `ValueError` | `unsupported_format` (not an Office file, or another opener's format), `unsupported_capability` (for example flat OpenDocument), `unsupported_element`, `unsupported_type`, `unsupported_command` |
| `CorruptFileError` | `ValueError` | `corrupt_file`, `invalid_input` |
| `ProtectedError` | `ValueError` | `document_protected` (encrypted, password-protected or signed; a save that would drop encryption) |
| `PropertyError` | `ValueError` | `invalid_value`, `unsupported_property`, `missing_property`, `duplicate_name`, `invalid_argument`, `missing_argument`, `invalid_json`, `corrupt_batch` |
| `PathError` | `LookupError` | `not_found`, `path_not_found`, `zero_matches`, `out_of_range`, `range_target_not_found`, `invalid_path`, `invalid_selector`, `invalid_xpath`, `unsupported_path`, `bare_selector_rejected` |

`Document`, `Presentation` and `Workbook` raise the first five. `File` also
raises `PathError` for paths and selectors (malformed: `invalid_path`,
`invalid_selector`; nothing there: `path_not_found`; not in the format:
`unsupported_path`), `PropertyError` (`invalid_argument`) for an unknown view
mode or a negative count, `IoError` (`file_required`) after `close()`, and
`LitchiError` itself for `output_too_large` and `limit_exceeded` (a batch of
more than 10,000 operations). The write verbs raise `PropertyError` for an
unknown property (with a did-you-mean suggestion), an invalid value or a
malformed batch item (`corrupt_batch`), `ProtectedError` for a signed or
locked file, and `IoError` for `file_exists`, `file_required` (a file with no
path to save to) or a source that changed. Errors keep the message and the
suggestion Litchi gives, such as `did you mean "paragraph"?` or the elements
that do exist at a path.

## Supported Formats

"Family classes" are `Document`, `Presentation` and `Workbook`. In the
`File` column, "paths" means `get`, `find`, `query` and `view`; every format
`File` opens also answers `text()`, `metadata()` and `capabilities()`.

| Format | Extension | Family classes | `File` |
|--------|-----------|----------------|--------|
| Microsoft Word 97-2003 | .doc | ✅ | ✅ paths; edits: paragraph text, create, save |
| Microsoft Word 2007+ | .docx | ✅ | ✅ paths and edits |
| Microsoft PowerPoint 97-2003 | .ppt | ✅ | ✅ paths; edits: shapes, slide order and removal, create, save |
| Microsoft PowerPoint 2007+ | .pptx | ✅ | ✅ paths and edits |
| Microsoft Excel 97-2003 | .xls | ✅ | ✅ paths; edits: cells, sheet names, create, save |
| Microsoft Excel 2007+ | .xlsx | ✅ | ✅ paths and edits |
| Microsoft Excel Binary | .xlsb | ✅ | ✅ paths and edits |
| OpenDocument Text | .odt | ✅ | opens; paths and edits pending |
| OpenDocument Spreadsheet | .ods | ✅ | opens; paths and edits pending |
| OpenDocument Presentation | .odp | ✅ | opens; paths and edits pending |
| Apple Pages | .pages | ✅ | — |
| Apple Keynote | .key | ✅ | ✅ paths; edits: slides (add from a layout or as a copy, remove, move, swap, hidden), texts, paragraphs, notes, frames, text boxes, shapes, pictures, z-order, create, save |
| Apple Numbers | .numbers | ✅ | — |
| Rich Text Format | .rtf | ✅ | ✅ paths and edits |

Each format's `capabilities()` lists exactly which elements and properties
each verb supports.

## Threads and the GIL

Handles are immutable, so a `File`, `Document`, `Presentation` or `Workbook`
and the paragraphs, runs, tables and slides read from it can be shared between
threads. Opening a file and the document-wide reads (`text()`, `paragraphs()`,
`tables()`, `slides()`, `slide_count()`, `worksheet_names()`,
`worksheet_count()`) and the detection functions release the GIL while they
run, so other Python threads keep running. Litchi starts no threads of its
own.

`File`'s reads release the GIL for all of Litchi's work (`open`, `get`, `find`
and `query` resolve, walk and read every node with the GIL released; `view`,
`text` and `metadata` render with it released) and hold it only to build the
returned Python objects, which are made directly, without JSON text in
between. `capabilities()` reads only static tables, so all its cost is
building the dict, which needs the GIL. `close()` drops the handle's reference
to the document; a read already running on another thread finishes on its
own reference.

The write verbs convert their arguments (property dicts, batch items) to Rust
once with the GIL held, then edit, run the batch, save or publish with the GIL
released. An `Edit` can be shared between threads: its calls run one at a
time.

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
│   ├── file.rs         # File, open(), open_bytes() and create(): the unified API
│   ├── nodes.rs        # Node views as dicts (get, find, query)
│   ├── write.rs        # Commit, Patch, Edit and the write verbs' shared pieces
│   ├── props.rs        # Property dicts as litchi Props
│   ├── pyde.rs         # Python objects read as batch items (serde)
│   ├── pyser.rs        # serde values (views, capabilities) as Python objects
│   ├── iwork.rs        # Pages, Keynote and Numbers projections of a File
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

