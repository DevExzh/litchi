# S8a Python read-API oracle

- Status: test suite only; no binding code changed. S8a has since fixed every
  drift and bug item below; see "Status after S8a" and
  `docs/unified-ops/s8a-python-on-file.md`.
- Date: 2026-10-05
- Branch: `uo/m0a-s8a-pre-pytests`, based on `uo/m0a` at `2d8971feac`. That is
  before S8a, so `litchi::Document`, `Presentation` and `Workbook` still back
  the bindings.
- Scope: the safety net for M0a step S8a (record 0003 §5.1). S8a moves
  `crates/pyo3-litchi` onto the unified `litchi::File`. Design §7 and the goal
  require the Python read API to stay source-compatible. Before this suite,
  the bindings had no Python tests at all.

The suite pins the read API as it behaves today: the module `litchi_py`, its
classes, functions, errors and stub. S8a must finish with every test green,
changing only the tests this record marks as changeable (see "Test tiers").

## How to run

From the repository root:

```sh
export CARGO_TARGET_DIR=$PWD/target CARGO_INCREMENTAL=0 CARGO_PROFILE_DEV_DEBUG=0
crates/pyo3-litchi/tests/run_tests.sh                    # dev build, then the suite
crates/pyo3-litchi/tests/run_tests.sh --release          # release (panic = "abort")
crates/pyo3-litchi/tests/run_tests.sh --no-build -v -k d4  # reuse the library, filter
```

The runner does the following:

1. Runs `cargo build -p litchi-py [--release]`. The crate directory is
   `crates/pyo3-litchi`, but the package is named `litchi-py`.
2. Finds the target directory with `cargo metadata`, so it honours
   `CARGO_TARGET_DIR`.
3. Copies `liblitchi_py.so` to `litchi_py.abi3.so` in a `mktemp -d` directory.
   The module name comes from `#[pymodule] fn litchi_py`; `pyproject.toml` sets
   no `module-name`.
4. Runs `python3 -m unittest discover` on `crates/pyo3-litchi/tests` with
   `PYTHONDONTWRITEBYTECODE=1`, and deletes the staging directory on exit.

Apart from cargo's build output (in `CARGO_TARGET_DIR`, else `<repo>/target`),
it writes nothing into the checkout. Arguments after the runner's own flags go
to unittest. `PYTHON` selects the interpreter. `LITCHI_TEST_DATA` overrides
the fixture root, `test-data/`. A missing fixture fails loudly; it is never
skipped.

To run without the runner:

```sh
mkdir -p /tmp/lp && cp "$CARGO_TARGET_DIR/debug/liblitchi_py.so" /tmp/lp/litchi_py.abi3.so
PYTHONPATH=/tmp/lp python3 -m unittest discover -s crates/pyo3-litchi/tests -t crates/pyo3-litchi/tests
PYTHONPATH=/tmp/lp python -m pytest crates/pyo3-litchi/tests   # pytest from any venv
```

- **Python packages.** The suite needs neither maturin nor pytest. It uses only
  the standard library's `unittest`, and pytest runs it unchanged (results
  below).
- **ABI3 forward compatibility.** `PYO3_USE_ABI3_FORWARD_COMPATIBILITY` was not
  needed: pyo3 0.29.3 built `abi3-py38` against Python 3.14.4 directly.
- **Python versions.** The tests avoid syntax and APIs newer than Python 3.8,
  for example `ast.unparse`. They were run only on 3.14.4.

## Layout

| File | Tests | Contents |
|---|---|---|
| `tests/_support.py` | – | Module import guard, fixture paths, the subprocess runner for panicking code, in-place/atomic file mutation, stub (`.pyi`) parser, README API parser, known-drift constants |
| `tests/test_read_api.py` | 58 | Per-format reads for all three families, signed packages, content-based dispatch, handle semantics, `repr` snapshot |
| `tests/test_formats_and_values.py` | 16 | `FileFormat`, both detection functions, `RGBColor`, `Length` |
| `tests/test_errors.py` | 13 | Path errors, argument types, non-Office input, wrong family, unsupported inputs, source changes, message snapshot |
| `tests/test_stub_sync.py` | 7 | Stub against runtime: names, `__all__`, member kinds, parameter names, bases. Also README API references. |
| `tests/test_known_drift.py` | 27 (14 expected failures) | Today/intended pairs for D1–D5, F10 and B1–B5 |
| `tests/run_tests.sh` | – | Build, stage, run |

In total there are 121 tests. 107 pass and 14 are expected failures. Under
pytest they produce 354 subtests.

## Test tiers

1. **Contract.** These tests must stay green without edits. Exception classes
   are checked with `isinstance`, so S8a may raise subclasses such as a
   `LitchiError` family (design §8). It must keep `except OSError`,
   `except ValueError` and `except Exception` working wherever they work
   today. Messages are checked only for meaningful fragments:
   - the detected format name in a wrong-family refusal (record 0003 D7);
   - "password" for a password-protected `.doc`;
   - "source changed" after an in-place modification.
2. **Snapshot.** These tests pin exact strings and today's routing:
   - `test_errors.MessageSnapshotTests` (Linux only, because I/O messages are
     platform text);
   - `test_read_api.ReprSnapshotTests`;
   - `test_errors.SourceChangeTests.test_eagerly_read_formats_ignore_a_changed_source`.

   S8a may change them, but only on purpose. It updates the table in the same
   change and lists the difference in its record. Design §8 adds `code`,
   `message` and `suggestion`, so error wording is expected to change.
3. **Known drift.** Each item in `test_known_drift.py` has two tests:
   - `test_<id>_…_today` pins the current behaviour;
   - `test_<id>_…_intended` asserts the fix and is marked
     `@unittest.expectedFailure`.

   When a change fixes an item, the `_today` test fails and the `_intended`
   test passes unexpectedly. Both unittest and pytest fail the run on an
   unexpected success (checked). The person making the fix then deletes the
   `_today` test, removes the decorator and updates the table below.

   Before this record was written, every `_intended` test was run with its
   decorator removed, to confirm it fails on its intended assertion and not on
   an incidental error.

## Coverage

What the bindings accept today:

| Opener | Formats |
|---|---|
| `Document.open` | docx, doc, rtf, odt, pages |
| `Presentation.open` | pptx, odp, key |
| `Workbook.open` | xlsx, xlsb, ods, numbers |

`.ppt` and `.xls` are not accepted (B1). The public API has no metadata
accessor, no open-from-bytes and no cell API. Bytes are accepted only by
`detect_file_format_from_bytes`, and cell values appear only through
`Workbook.text()`. `Worksheet` is an empty placeholder that cannot be
constructed and that nothing returns.

Legend:
- `Y`: asserted.
- `0`: asserted to be empty today.
- `R`: the read is refused after the source changes.
- `E`: the format was read at open, so a later change is not seen.
- `–`: not applicable.
- `·`: not asserted for this format.
- IDs (B1, B4, …) point to the drift table below. The O-numbers in parentheses
  point to the "Other behaviour pinned today" list.

| API | docx | doc | rtf | odt | pages | pptx | odp | key | xlsx | xlsb | ods | numbers | ppt | xls |
|---|---|---|---|---|---|---|---|---|---|---|---|---|---|---|
| `detect_file_format(path)` / `(Path)` | Y | Y | Y | Y | Y | Y | Y | Y | Y | Y | Y | Y | B1 | B1 |
| `detect_file_format_from_bytes` | Y | Y | Y | Y | Y | Y | Y | Y | Y | Y | Y | Y | B1 | B1 |
| `open(str)` | Y | Y | Y | Y | Y | Y | Y | Y | Y | Y | Y | Y | B1 | B1 |
| `open(Path)`, `open(path=…)` | Y | · | · | · | · | Y | · | · | Y | · | · | · | – | – |
| content-based dispatch (misleading extension) | Y | Y | · | · | Y | · | · | · | Y | · | Y | · | – | – |
| `text()` | Y (order; B3) | Y exact | Y exact | Y exact | Y exact | Y exact | Y exact | Y exact | Y rows | Y rows | Y rows | Y exact | – | – |
| `paragraphs()`, `Paragraph.text()` | Y | Y | Y | Y | Y | – | – | – | – | – | – | – | – | – |
| `Paragraph.runs()`, `Run.text()` | Y | Y | Y | 0 (O1) | Y | – | – | – | – | – | – | – | – | – |
| `Run.bold()`, `Run.italic()` | Y | Y | Y | – | Y | – | – | – | – | – | – | – | – | – |
| `Run.underline()` | · | · | B4 | – | Y | – | – | – | – | – | – | – | – | – |
| `tables()`, `Table.row_count()`, `rows()` | Y | Y | Y | Y | 0 (O2) | – | – | – | – | – | – | – | – | – |
| `TableRow.cells()`, `TableCell.text()` | Y | Y | Y | Y | – | – | – | – | – | – | – | – | – | – |
| `slide_count()`, `slides()`, `Slide.text()` (shape text) | – | – | – | – | – | Y | Y | Y (O2) | – | – | – | – | – | – |
| `worksheet_count()`, `worksheet_names()` | – | – | – | – | – | – | – | – | Y | Y | Y | Y | – | – |
| cell values through `Workbook.text()` | – | – | – | – | – | – | – | – | Y | Y | Y | Y (O6) | – | – |
| `repr()` (ASCII previews) | Y | Y | · | · | · | Y | Y | · | Y | · | · | · | – | – |
| `repr()` with multibyte text | B2 | · | · | · | · | · | · | · | – | – | – | – | – | – |
| wrong-family refusal names the format | Y | Y | Y | Y | Y | Y | Y | Y | Y | Y | Y | Y | – | – |
| read after an in-place change | R | E | E | R | E | R (F10) | R | E | R | R | R | E | – | – |
| read after replace or unlink | Y | Y | Y | Y | Y | Y | Y | · | Y | Y | Y | Y | – | – |
| signed package reads | Y | – | – | – | – | Y | – | – | Y | – | – | – | – | – |

The following are not tied to one format:

- **`FileFormat`.** The 20 runtime variants. `str()` is the variant name and
  `repr()` is `FileFormat.<name>`. Equality with detection results works,
  values are hashable (dictionary keys, set members), and `int()` gives
  distinct values. The exact integers are not pinned.
- **Detection of the other ODF formats.**
  - `odg`, `odc`, `odf` and `odm` are named by detection, but every opener
    refuses them with a non-`OSError` exception.
  - Flat ODF (`fodt`, `fods`, `fodp`, `fodg`) is detected as its packaged
    format. `Document` and `Presentation` refuse it with `ValueError`.
  - `odb` is detected as `Doc` (D1).
- **Detection edge cases.** Detection returns `None` for empty, truncated,
  plain-text, random and unrelated-ZIP bytes, and for a missing path or a
  directory.
- **`RGBColor`.** Its components, positional and keyword construction, `str`
  and `repr`, `OverflowError` for values outside 0–255, and `TypeError` for
  non-integers.
- **`Length`.** EMU, point and inch conversions, keyword arguments, `str` and
  `repr`, `TypeError` for a float EMU value, and `OverflowError` above `i64`.
- **Errors.**
  - A missing file or a directory raises `OSError` from `Document` and
    `Presentation`, and `Exception` from `Workbook` (D4a).
  - A path argument of `None`, `42` or a list raises `TypeError`.
  - Non-Office input raises a non-`OSError` exception. The inputs tested are
    plain text, random bytes, an unrelated ZIP, an empty `.docx`, plain text
    named `.docx`, and a truncated `.docx`.
  - A password-protected `.doc` raises `ValueError` whose message contains
    "password".
  - A Word 6 `.doc` raises `ValueError`.
- **Handles.**
  - Collections are `list`s of the documented classes.
  - Counts agree with the collections.
  - Children outlive their parent handle.
  - Reads are repeatable.
- **Stub and README.** Names, kinds and parameters must match exactly, apart
  from the known drift (D2, D3, D5). New API therefore needs a stub entry in
  the same change.

## Known drift and bugs

All fourteen pairs were fixed in S8a: each `_today` test is deleted and each
`_intended` test is an ordinary test. The table keeps the state before S8a.

| ID | Today (before S8a) | Intended | Root cause | Fix expected in |
|---|---|---|---|---|
| D1 | `FileFormat` has 20 variants and no `Odb`. `.odb` is detected as `FileFormat.Doc` (path and bytes), yet `Document.open` refuses it. | `FileFormat.Odb` exists and detection returns it; all 21 core variants. | `From<litchi::FileFormat>` ends with `_ => FileFormat::Doc` (`src/common.rs`). | S8a (record 0003 §5.1, F14) |
| D2 | The stub's `FileFormat` declares 14 of 21 formats. `Odg`, `Odc`, `Odf`, `Odi`, `Odm` and `Oth` exist at runtime but are missing from the stub. | The stub declares every runtime variant. | `python/litchi_py/__init__.pyi` | S8a (record 0003 §5.1) |
| D3 | The stub says `class FileFormat(Enum)`, but the runtime is a pyo3 class, not an `enum.Enum`. | Either the runtime becomes an `Enum` (keeping `str(v) == name`) or the stub drops the base. | `__init__.pyi` | open |
| D4a | `Workbook.open` raises a bare `Exception` for a missing file or a directory. | `OSError`, as the stub's `Raises: IOError` says. | `boxed_err_to_py_err` maps every workbook error to `Exception`. | open; S8a deletes `boxed_err_to_py_err` |
| D4b | Non-Office, wrong-family and corrupt-ZIP inputs raise a bare `Exception` from all three openers. | `ValueError`, as the stub's `Raises: ValueError` says, unless the stub changes instead. | `to_py_err` maps only `Io`, `InvalidFormat` and `Unsupported`. | open; S8a remaps errors to codes |
| D4c | Flat ODF through `Workbook.open` raises a bare `Exception`, while `Document.open` raises `ValueError` for the same input. | `ValueError` | as D4a | open |
| D5 | The README's "API Reference" lists `Workbook.worksheets()`, `worksheet_by_name()` and `Worksheet.name/row_count/column_count/cell_value/rows`. None of them exist. | The README lists only existing API, or the API is implemented. | `crates/pyo3-litchi/README.md` | open |
| F10 | After an in-place change, PPTX reads raise `ValueError("PPTX OPC error: OPC source changed …")`. DOCX, XLSX, XLSB and ODF raise a bare `Exception("positional source changed …")`. | PPTX raises the same class as DOCX. | PPTX `check_source()` errors go through `map_ooxml_error`. | S8a. S5b's `Pptx` handle already keeps `SourceChanged` typed; Python sees the fix only when it switches over (record 0003 F10, R7). |
| B1 | `.ppt` and `.xls` are not recognised. Detection returns `None`, and every opener says "Not a valid Office file". | `.ppt` opens as a `Presentation` and `.xls` as a `Workbook`, as the README, the module docstring and the `open` docstrings claim. | `crates/pyo3-litchi/Cargo.toml` enables the facade's `doc` feature but not `ppt` or `xls` (or `legacy`), and OLE2 classification is gated per feature (`crates/litchi-ops/src/detect/ole2.rs`). | open (enabling the features changes the binding's build, so it was not done here) |
| B2 | `repr()` of `Paragraph`, `Run` and `TableCell` panics when the preview cut (50 or 30 bytes) falls inside a UTF-8 character. A dev build raises `pyo3_runtime.PanicException`, a `BaseException` that `except Exception` does not catch. A release build aborts the interpreter (exit 134). The trigger is a real fixture, `drawing.docx` (Russian text). `Slide` has the same code but no fixture reaches it. | The preview is cut at a character boundary. | `&text[..50]` / `&text[..30]` in `src/document.rs` and `src/presentation.rs` | open |
| B3 | DOCX `Document.text()` concatenates paragraphs with no separator, for example `"Following is a hyperlink.EricWhite.com"`. | Paragraph boundaries are kept, as in every other `Document` format. | The source-backed route calls `extract_word_text`, which joins `<w:t>` text. The eager route uses a different method. | open (leaf or facade) |
| B4 | `Run.underline()` always returns `None`, including for `\ul` text in `tdf108947.rtf`. | It reports `True` for underlined text. | A placeholder in `src/document.rs`; the facade `Run` has no underline accessor. | open |
| B5 | Every handle is `#[pyclass(unsendable)]`, so using one from another thread panics. A dev build raises `PanicException`; a release build aborts. | A read from another thread works. Raising a catchable exception would also be acceptable; if that route is taken, adjust the `_intended` test. | pyo3 `unsendable` thread check | open (S8a's `Send + Sync` `File` makes this fixable) |

B2 and B5 run in a subprocess, because the release profile has
`panic = "abort"` and the panic would otherwise kill the test run.

## Other behaviour pinned today

These behaviours have no `_intended` test. The read tests pin them so that a
change in S8a is visible. Parity of the iWork formats (R6) depends most on O2
and O6.

- **O1.** ODT paragraphs return no runs, even for styled text.
- **O2.** iWork structure:
  - Each Pages text storage is one `Paragraph`, which can contain `\n`, with
    one run.
  - Pages `tables()` is empty even when the document has native tables. Its
    `text()` carries one U+FFFC per table.
  - Keynote slide text leaves out native table cells.
- **O3.** `paragraphs()` includes table-cell paragraphs for DOCX, DOC and ODT,
  but not for RTF.
- **O4.** Each format joins text differently:
  - DOCX uses no separator (B3).
  - DOC ends each paragraph with `\r` and keeps field codes.
  - RTF, ODT and Pages separate paragraphs with `\n`.
  - PPTX and ODP join non-empty slides with `\n\n`.
  - XLSX, XLSB and ODS write rows as lines of tab-separated cells.
- **O5.** Run formatting has three states:
  - DOCX and DOC report `None` when unspecified.
  - RTF reports an explicit `False`.
- **O6.** Numbers `text()` names each sheet and table, then lists one cell per
  line, showing formula source (`=(B2+1)`) rather than a value.
- **O7.** Detection and opening ignore the file extension.
- **O8.** The ODF variants that are detected but have no opener (`odg`, `odc`,
  `odf`, `odm`, and `odb`) are refused with "Not a valid Office file" rather
  than with a message that names the detected format.
- **O9.** An open handle keeps reading the file it opened, even after the path
  is replaced atomically or unlinked. The source pins its descriptor.

Seen but not tested:

- `RGBColor` and `Length` compare by identity, because they define no
  `__eq__`.
- Every class reports `__module__ == "builtins"`, because `#[pyclass]` sets no
  `module`.
- `pyproject.toml` has no `python-source`. It is therefore unclear whether
  maturin installs `python/litchi_py/__init__.pyi` beside the module. This was
  not verified, because maturin is not installed.

## Notes for S8a

- **Exception classes.** Keep the contract classes:
  - `Document` and `Presentation` raise `OSError` for I/O failures.
  - They raise `ValueError` for flat ODF, a password-protected `.doc` and
    Word 6.
  - Everything else is at least an `Exception`.

  A `LitchiError` hierarchy fits if its I/O and unsupported/invalid families
  also derive from `OSError` and `ValueError`. A password-protected `.doc`
  becomes `File::Locked` (0003 §2.6); through `Document.open` it must still
  raise a `ValueError` that mentions "password".
- **Wrong-family messages** must keep naming the detected format.
  Case-insensitive matching is enough, so "docx" in lower case passes.
- **R6 (iWork moves to `litchi::iwork::Document`).** The Pages, Keynote and
  Numbers tests pin paragraph boundaries, runs, missing tables, text layout and
  sheet names exactly.
- **R7 (signed input becomes `File::Signed`).** `SignedPackageReadTests` must
  stay green, which means reads go to the inner file.
- **Source changes.**
  - The seven source-backed formats must keep refusing reads after an in-place
    change.
  - The eager list (DOC, RTF, Pages, Keynote, Numbers) is a snapshot. If S8a
    makes one of them source-backed, move its row deliberately.
  - Fixing F10 flips its pair.
- **Stub.** `test_stub_sync` fails on any API that is added to only one of the
  runtime and the stub. Fixing D1 and D2 flips their pairs. After both fixes,
  the runtime and the stub list the same 21 variants.

## Results (2026-10-05, Linux x86_64)

Toolchain: rustc and cargo 1.95.0, pyo3 0.29.3, Python 3.14.4. All runs used
`CARGO_TARGET_DIR=<worktree>/target CARGO_INCREMENTAL=0 CARGO_PROFILE_DEV_DEBUG=0`.

| Command | Result |
|---|---|
| `cargo build -p litchi-py` (dev, cold target) | finished in 55.8 s. `liblitchi_py.so` is 56.6 MB, sha256 `c937116611b69302db2c7f06f7a7f5846c4a7d93ac111e2a7c366c25bdc4ea0e`. |
| `cargo build -p litchi-py --release` | 2 min 29 s wall time (fat LTO). The artifact is 11.9 MB, sha256 `1d86ddb935b65f4254b099d847df21f381bf931566d7114c9d3724c47cabb6dd`. |
| `crates/pyo3-litchi/tests/run_tests.sh` | `Ran 121 tests in 1.272s` / `OK (expected failures=14)` |
| `crates/pyo3-litchi/tests/run_tests.sh --release` | `Ran 121 tests in 0.154s` / `OK (expected failures=14)` |
| `python -m pytest -p no:cacheprovider crates/pyo3-litchi/tests` (pytest 9.1.1 in a scratch venv, dev library) | `107 passed, 14 xfailed, 354 subtests passed in 1.25s` |
| the same with the release library | `107 passed, 14 xfailed, 354 subtests passed in 0.29s` |
| each `_intended` test with its decorator removed | all 14 fail on their intended assertion |

The target directory took 3.7 GB after both profiles. The libraries are build
outputs and are not committed (AGENTS.md); the hashes above are only for
tracing. B2 and B5 were exercised in both profiles:

| Profile | B2 subprocess exit | B5 subprocess exit | Message seen |
|---|---|---|---|
| dev | 1 (`PanicException`) | 3 | the panic text on stderr |
| release | −6 (SIGABRT) | −6 (SIGABRT) | the panic text on stderr |

## Status after S8a

S8a (`docs/unified-ops/s8a-python-on-file.md`) moved the bindings onto
`litchi::File` and fixed D1–D5, F10 and B1–B5. The suite now has 132 tests,
all ordinary (no expected failures): 13 `_today` tests were deleted and 24
tests were added. The snapshot tier changed on purpose (messages, the
source-change prefixes); the S8a record lists every difference.

## Not covered, not claimed

- **API the bindings do not expose** cannot be tested: metadata, open from
  bytes, a cell API, and shapes as objects. Shape text is covered only through
  `Slide.text()`.
- **Other platforms and versions.** The tests ran only on Linux with Python
  3.14.4. The runner names the macOS library `.dylib`, but that path was not
  tried. Windows is not supported by the runner.
- **Performance.** No claims.
