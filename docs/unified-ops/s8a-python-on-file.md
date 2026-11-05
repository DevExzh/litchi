# S8a — Python bindings on `litchi::File`

- Status: implemented on `uo/m0a-s8a-python` (base `uo/m0a` at `20692d6444`).
- Date: 2026-10-05
- Scope: M0a step S8a (record 0003 §5.1, risk R6): move `crates/pyo3-litchi`
  (package `litchi-py`, module `litchi_py`) onto the unified `litchi::File`,
  keep the Python read API source-compatible, add the `LitchiError`
  exception hierarchy, and fix the bugs the oracle
  (`docs/unified-ops/s8a-python-oracle.md`) recorded: D1–D5, F10 and B1–B5.

## 1. What changed

### Bindings (`crates/pyo3-litchi`)

- **Opening.** `Document.open`, `Presentation.open` and `Workbook.open` call
  `litchi::open` with the GIL released, then check the family view
  (`File::document`, `presentation`, `workbook`), so a file of another family
  is refused at open, as before. Each handle holds a binding-internal
  `Source`: `File(litchi::File)`, or the iWork document of its family (below).
  Signed packages read through `File::Signed`; encrypted ones are
  `File::Locked` and are refused at open.
- **Reads** go through the family views and the `litchi::compat` wrappers
  (paragraphs, runs, tables, rows, cells, slides) and `File::text()`.
- **iWork (deviation from record 0003 §5.1).** The plan named
  `litchi::iwork::Document` for the iWork arm. Its format-neutral snapshot
  cannot reproduce three outputs the oracle pins (contract tier, risk R6): it
  has no Numbers sheets (`worksheet_names()` needs them), it keeps only
  textual Numbers cells (`text()` lists numbers and booleans such as `42`,
  `true`), and it has no Keynote show title (part of `Presentation.text()`).
  The bindings therefore read Pages, Keynote and Numbers through the
  format's archive-free semantic document, re-exported by the facade as
  `litchi::{pages, keynote, numbers}::semantic::Document`, which is what the
  old facade types projected. A probe over every iWork fixture in `test-data`
  found identical text, paragraphs, slides, slide counts, sheet names and
  sheet counts on all 123 regular files (31 Pages, 54 Keynote, 38 Numbers).
  `litchi::open` refuses iWork with `unsupported_format` and the detected
  format; the opener of that family then opens the semantic document.
- **Thread safety (B5).** Every class is `#[pyclass(frozen)]` and no longer
  `unsendable`; `litchi::File`, the compat wrappers and the semantic iWork
  documents are `Send + Sync` (pyo3 checks this at compile time).
- **GIL.** Released around open, `Document.text/paragraphs/tables`,
  `Presentation.text/slide_count/slides`, `Workbook.text/worksheet_count/
  worksheet_names`, both `detect_file_format*` functions, and the `repr` of
  `Presentation` and `Workbook` (which count slides or sheets). Element reads (`Paragraph`, `Run`, `Table`, ... on data
  already in memory) keep the GIL. The work runs on the calling thread; no
  thread is started.
- **Errors.** `to_py_err` maps the unified `Code` (section 2);
  `boxed_err_to_py_err` is deleted. Suggestions that name Rust API are
  replaced by Python ones (section 2).
- **`FileFormat` (D1).** All 21 core formats. `Odb` is added last, so the
  `int()` values of the other 20 do not change (`Odb` is 20). A format a later
  Litchi adds before the bindings learn it raises `LitchiError`
  (`internal_error`) instead of being reported as `Doc`.
- **`repr` previews (B2)** are cut at 50 or 30 characters, as Python slices a
  `str`, never inside a character. ASCII previews are unchanged.
- **`Run.underline()` (B4)** returns the run's value (`litchi::compat::word::
  Run::underline`, below).
- **Cargo features (B1).** `litchi` is built with `legacy` (doc, ppt, xls)
  instead of `doc`; the other features are unchanged. The dependency is still
  `litchi`, not `litchi-ops` (plan finding F5: the bindings need the facade's
  iWork re-exports).
- **Classes report module `litchi_py`** (was `builtins`).

### `litchi-docx` and `litchi-ops` (for B3 and B4)

- **B3, DOCX text separators.** `File::text()` for DOCX joined the `w:t` text
  of the whole story with no separator, so Python's `Document.text()` read
  `"Following is a hyperlink.EricWhite.com"`. `Docx::text()` (litchi-ops)
  now returns each paragraph's text separated by `\n`; Python's
  `Document.text()` is `File::text()`. The separators come from the leaf's
  validating single-pass extractor, which gained a paragraph mode:
  - `litchi_docx::source_backed::Document::extract_text_with_paragraph_separator`,
    `litchi_docx::Document::text_with_paragraph_separator` and
    `DocumentPart::extract_text_with_paragraph_separator` (new public API);
  - the scanner writes the separator before each `w:p` that is not nested in
    another `w:p` (a text box's paragraphs stay in their anchor, as the
    paragraph range scanner counts them); an empty `<w:p/>` adds an empty
    segment;
  - for well-formed XML the result equals `paragraphs()` joined by `\n`
    (tested over every DOCX main part in `test-data`), and malformed XML fails
    exactly as `extract_text()` fails;
  - the existing `extract_word_text` path is a const-generic instance of the
    same scanner with the paragraph mode compiled out.
- **Rejected first attempt.** Joining `paragraphs()` in litchi-ops was
  rejected: the paragraph range scanner tolerates unterminated XML, so
  `text()` of a malformed main part returned `""` instead of `corrupt_file`
  (three ops tests caught it), and it cost +4.6% median, up to +31%, on
  open+text over the DOCX corpus.
- **B4.** `litchi::compat::word::Run::underline()`: DOCX reports a direct
  `w:u` (`False` for `w:val="none"`), DOC `True` for any underline and `None`
  otherwise, RTF always `True`/`False`, ODT `None`.

## 2. Exception hierarchy

Defined in Python at module initialization (a pyo3 class can extend only one
native base, and the I/O family must extend `LitchiError` and `OSError`), once
per process, with `__module__ == "litchi_py"`.

| Class | Also a | Codes | Default `code` |
|---|---|---|---|
| `LitchiError` | `Exception` | base; raised itself for `limit_exceeded`, `internal_error`, `cancelled`, `validation_error`, `output_too_large`, `max_depth_exceeded`, `merge_not_exact`, `refresh_failed`, and codes added later | `internal_error` |
| `PathError` | `LookupError` | `not_found`, `path_not_found`, `zero_matches`, `out_of_range`, `range_target_not_found`, `invalid_path`, `invalid_selector`, `invalid_xpath`, `unsupported_path`, `bare_selector_rejected` | `path_not_found` |
| `PropertyError` | `ValueError` | `invalid_value`, `unsupported_property`, `missing_property`, `duplicate_name`, `invalid_argument`, `missing_argument`, `invalid_json`, `corrupt_batch` | `invalid_value` |
| `UnsupportedError` | `ValueError` | `unsupported_format`, `unsupported_capability`, `unsupported_element`, `unsupported_type`, `unsupported_command` | `unsupported_capability` |
| `IoError` | `OSError` | `file_not_found`, `file_locked`, `file_exists`, `file_required`, `io_error` | `io_error` |
| `CorruptFileError` | `ValueError` | `corrupt_file`, `invalid_input` | `corrupt_file` |
| `ProtectedError` | `ValueError` | `document_protected` | `document_protected` |

- **Attributes:** read-only properties `code`, `message` and `suggestion`.
  `args == (message,)`. `str(error)` is the message, followed by
  ` (suggestion: ...)` when there is a suggestion. Instances pickle and copy
  with all three values. `LitchiError(message, code=None, suggestion=None)`;
  a missing `code` is the class's default.
- **Contract kept:** I/O failures are `OSError` (`IoError`); flat ODF and
  Word 6 are `ValueError` (`UnsupportedError`); a password-protected `.doc`,
  now `File::Locked`, is `ValueError` (`ProtectedError`) and its text says
  "password"; wrong-opener messages name the detected format.
- **Python suggestions:** a wrong-family refusal names the opener that reads
  the format (`open it with Document.open`); formats no opener reads (ODG,
  ODC, ODF, ODI, ODM, OTH, ODB) get none; `document_protected` says to remove
  the password, since the bindings cannot unlock; `limit_exceeded` gets none
  (Python has no open options). Other suggestions pass through.
- **For M3 (provisional):** the read API raises only `IoError`,
  `UnsupportedError`, `CorruptFileError`, `ProtectedError` and the base.
  `PathError` and `PropertyError` are reserved for the editing API, and their
  code sets and built-in bases are proposals to confirm there.

## 3. Behaviour and message changes

| Input or call | Before S8a | After S8a |
|---|---|---|
| Missing file, all openers | `OSError` (`Workbook`: bare `Exception`) | `IoError`, `file_not_found`; message unchanged |
| Directory, all openers | `OSError` (`Workbook`: `Exception`) | `IoError`, `io_error`; message unchanged |
| Plain text, random bytes, unrelated ZIP | `Exception("Not a valid Office file")` | `UnsupportedError`, `unsupported_format`: "not a recognized Office file" |
| Empty, plain-text or truncated `.docx` | `Exception`, ZIP message | `CorruptFileError`, `corrupt_file`; message unchanged |
| ODG, ODC, ODF, ODM, ODB | `Exception("Not a valid Office file")` (O8) | `UnsupportedError`: "detected Odg, which this operation does not handle" |
| Wrong opener | `Exception("Detected Docx, which this opener does not handle")` | `UnsupportedError`: "detected Docx, which this operation does not handle (suggestion: open it with Document.open)" |
| Password-protected `.doc` | `ValueError("DOC password required")` | `ProtectedError`: "the Doc document is encrypted (suggestion: remove the password in the application that created the file; litchi_py cannot unlock encrypted documents yet)" |
| Word 6 `.doc` | `ValueError("Word 6.0 documents (nFib 0x0065) are not supported")` | `UnsupportedError`, `unsupported_format`: "... predate the Word 97 binary format and are not supported" |
| Flat ODF | `ValueError` (`Workbook`: `Exception`), "flat OpenDocument Ods is detected but the dedicated family facade exposes packaged parsing only" | `UnsupportedError`, `unsupported_capability`: "flat OpenDocument Ods is detected, but only packaged OpenDocument opens" |
| Read after an in-place change (F10) | PPTX `ValueError`, others `Exception` | `IoError`, `io_error`, all formats. DOCX "OPC error: OPC source changed from ...", XLSX "OPC package error: ...", XLSB "OPC source error: ...", PPTX unchanged, ODF unchanged ("positional source changed ...") |
| DOCX `Document.text()` (B3) | concatenated paragraphs | one line per paragraph (`\n`); Rust `File::text()` likewise |
| `.ppt`, `.xls` (B1) | not recognized | open; detection returns `Ppt`/`Xls`; encrypted ones raise `ProtectedError` |
| `.odb` detection (D1) | `FileFormat.Doc` | `FileFormat.Odb` |
| iWork package directories | refused (`OSError`, not a regular file) | open in their opener; `detect_file_format` still returns `None` for them |
| `repr()` with multibyte text (B2) | panic (abort in release) | cut at 50/30 characters |
| `Run.underline()` (B4) | always `None` | the run's value |
| Use from another thread (B5) | panic (abort in release) | works |
| `Class.__module__` | `builtins` | `litchi_py` |

Observed, not changed: the CryptoAPI-encrypted
`poi/test-data/slideshow/Password_Protected-hello.ppt` raises
`UnsupportedError` (`unsupported_capability`, "unsupported PPT encryption:
CryptoApi") rather than `ProtectedError`, because the PPT reader reports it
that way to `litchi::open`.

Stub and README: the stub declares all 21 formats as `ClassVar[FileFormat]`
without the false `Enum` base (D2, D3), the seven exception classes, and
`Raises:` notes that name the classes (D4). The README documents only existing
API (D5), the exceptions and threading; it drops the unsupported performance
claims ("10-100x faster", "parallel processing") and points tests at
`run_tests.sh`.

## 4. Oracle results

All runs: Linux x86_64, rustc 1.95.0, pyo3 0.29.3, Python 3.14.4,
`CARGO_TARGET_DIR=<worktree>/target CARGO_INCREMENTAL=0
CARGO_PROFILE_DEV_DEBUG=0`.

| Run | Before (base `20692d6444`, original bindings) | After |
|---|---|---|
| `crates/pyo3-litchi/tests/run_tests.sh` | `Ran 121 tests in 1.340s` / `OK (expected failures=14)` | `Ran 132 tests in 1.814s` / `OK` |
| `crates/pyo3-litchi/tests/run_tests.sh --release` | `Ran 121 tests in 0.152s` / `OK (expected failures=14)` | `Ran 132 tests in 0.226s` / `OK` |

The libraries are build outputs and are not committed. After: dev
`liblitchi_py.so` 65,248,600 bytes, sha256 `e9f401de…4a51`; release
13,109,872 bytes, sha256 `fe440f50…d535`.

Test changes:

- **Fixed pairs.** All 14 `_intended` tests are ordinary tests; the 13
  `_today` tests are deleted (`test_known_drift.py`, 27 → 14 tests).
- **Snapshot tier, on purpose.** `MessageSnapshotTests.test_messages` (every
  row of section 3) and `test_source_change_messages` (the DOCX and XLSX
  prefixes).
- **Helpers and constants (`_support.py`).** `RUNTIME_FORMAT_NAMES_TODAY` is
  renamed `RUNTIME_FORMAT_NAMES` and holds all 21 formats (D1);
  `STUB_FORMAT_NAMES_TODAY` is removed; `KNOWN_RUNTIME_ONLY_MEMBERS`,
  `KNOWN_STUB_BASE_DRIFT` and `KNOWN_README_ONLY_API` are empty.
  `public_members` counts the names a class defines itself (the exception
  classes inherit `args`, `errno`, ... from built-ins, which stubs do not
  redeclare; for the pyo3 classes the set is the same as before), and
  `resolve_stub_base` resolves `litchi_py` classes and built-ins.
- **Contract tests**: assertions unchanged. Edited only where they encoded a
  fixed drift: `test_runtime_variant_names` uses the renamed constant;
  `DETECTION_CASES` gains `.ppt`, `.xls` and `.odb` rows; stale comments in
  `test_errors.py` and `test_read_api.py` are updated.
- **New tests (24).** `test_read_api.py` +8: `.ppt` (2), `.xls` (2), iWork
  package directories (2), DOCX one line per paragraph, multibyte `repr`.
  `test_exceptions.py` 14: hierarchy, defaults, pickling, read-only
  attributes, and the class/code/suggestion of each failure kind.
  `test_threads.py` 2, in subprocesses: eight threads reading nine shared
  handles (DOCX, DOC, Pages, PPTX, PPT, Keynote, XLSX, XLS, Numbers) and
  their paragraphs, tables and slides agree with the opening thread; handles and an exception made on a worker are used on the
  main thread.
- **Rust tests.** litchi-docx: paragraph-mode unit tests (separators, nested
  and empty paragraphs, malformed XML parity, corpus parity with the
  paragraph scanner) and a public-API test on both owners. litchi-ops: B3 on
  both DOCX owners; `underline()` for DOCX and RTF; the DOCX oracles are the
  paragraph texts joined by `\n`. litchi-py: the `preview` unit test.

## 5. Performance (indicative; one shared host, pinned CPU)

- **DOCX extraction.** Instruction counts (cachegrind) over 20 extra passes
  of the 60 DOCX main parts in `test-data`: paragraph mode 478.1M vs plain
  469.7M, +1.8%. The plain `extract_text()` path is +0.04% against HEAD
  (938.71M vs 938.31M for 40 passes), i.e. unchanged; a +3.8% wall-clock
  shift seen between two separately built probes was layout noise.
- **Python, old vs new release library** (median of per-process medians,
  three processes each, ABBA order, 300 iterations):

  | Case | Before µs | After µs | After/before |
  |---|---|---|---|
  | DOCX open+text `drawing.docx` | 2926.1 | 2966.3 | 1.014 |
  | DOCX open+text `table-alignment.docx` | 382.6 | 386.5 | 1.010 |
  | DOCX open+paragraphs / +tables | 463.5 / 432.7 | 470.9 / 432.3 | 1.016 / 0.999 |
  | DOC, RTF, ODT open+text | 37.2, 11.1, 85.0 | 36.4, 10.7, 85.6 | 0.979, 0.966, 1.006 |
  | PPTX open+text / +slides | 530.9 / 536.5 | 537.4 / 542.3 | 1.012 / 1.011 |
  | ODP open+text | 191.9 | 194.2 | 1.012 |
  | XLSX open+text / +names | 184.9 / 63.3 | 184.6 / 64.1 | 0.999 / 1.012 |
  | XLSB, ODS open+text | 155.1, 103.9 | 155.7, 104.0 | 1.004, 1.001 |
  | Pages, Keynote, Numbers open+text | 1047.1, 1791.2, 1232.3 | 1112.1, 1860.7, 1305.3 | 1.062, 1.039, 1.059 |
  | Wrong-family error | 39.3 | 41.5 | 1.056 |

- **iWork cost.** The iWork openers pay one extra container detection
  (+65–73 µs): `litchi::open` indexes the ZIP and classifies it before it
  refuses iWork, and the semantic reader detects again. It ends when iWork
  joins `File` (M6), or earlier if the facade can hand the refused prepared
  source to the reader. Other open paths are within about ±2%.
- The error path costs about 2 µs more (a Python-level exception class).

## 6. Findings and follow-ups

1. **`litchi::iwork` snapshot gaps** (for the iWork owners and M6): no
   Numbers sheets, no non-text Numbers cell rendering, no Keynote show
   title. The bindings use the per-format semantic documents meanwhile.
2. **iWork double detection** in the Python openers (section 5).
3. **For S8b/S10:** DOCX `File::text()` output changed (paragraph
   separators, +1.8% instructions), so DOCX full-text routes measured before
   this change are not like-for-like.
4. The CryptoAPI-encrypted PPT is `unsupported_capability`, not
   `document_protected` (section 3); a classification question for
   litchi-ops or litchi-ppt.
5. `detect_file_format` returns `None` for an iWork package directory that
   the openers now open.
6. `PathError` and `PropertyError` code sets and bases are provisional (M3).

## 7. Gates

- `cargo fmt --check -p litchi-py -p litchi-ops -p litchi-docx`: clean.
- `cargo clippy -p litchi-py --all-targets -- -D warnings`, and with
  `--all-features --all-targets` on litchi-docx and litchi-ops: clean.
- `cargo test -p litchi-ops --all-features`: 397 unit and 55 doc tests pass.
- `cargo test -p litchi-docx --no-fail-fast`: 1,999 pass; the 7 failures
  need the untracked `3rdparty` corpus (`drawing_svg_lifecycle`,
  `drawing_svg_source`) and pass (30 of 30) with `3rdparty` linked
  temporarily from the main checkout.
- `cargo test -p litchi-py --lib`: 1 test passes.
- `cargo check -p litchi --all-features`: passes.
- `python3 tools/check_crate_boundaries.py`: valid (66 packages, 267
  declarations; dependencies unchanged). `python3 tools/non_iwork_gate.py
  verify`: passes (47 bulk roots).

## 8. Not covered, not claimed

- pytest was not run (not installed); the suite is plain `unittest`.
- Only Linux and Python 3.14.4 were run; the `abi3-py38` floor, macOS and
  Windows were not.
- GIL release is not tested directly; it is not observable without timing.
- Encrypted OOXML through Python was not run: the test data has no such file
  (litchi-ops builds one in its own tests).
- maturin packaging of the stub was not checked (maturin is not installed).
- Performance numbers are indicative (section 5): no perf-baseline route was
  added or run, and no claim is made beyond the measured rows.
