# 0255 — Python M3 surface (plan 0248 step 14)

- Status: implemented on branch `uo2/p14-python` (cut from trunk
  `feat/unified-ops-api` at `908f9cccdd`).
- Date: 2026-10-13.
- Scope: design §8 and goal M3 for `crates/pyo3-litchi` (package
  `litchi-py`, module `litchi_py`): the R9 read surface ported from the
  stale reference `668d66883c` (`uo/m1`, record `r9-python-read.md` there),
  plus the write surface. Only `crates/pyo3-litchi` changed, apart from
  this record and its evidence file. The `litchi-ops` public API is consumed
  unchanged; what it lacks is listed as follow-ups (§6).
- `performance_claim: indicative` (§5): one host, one pinned CPU, ABBA
  process order; no perf-baseline route.

## 1. What landed

**Read surface (R9 port).** The R9 sources (`src/file.rs`, `src/nodes.rs`,
`src/pyser.rs`, the `.pyi` additions and `tests/test_file_api.py`) applied
cleanly to trunk, because the crate had not changed since R9's base. The
adaptation to trunk's ops API:

- `litchi::Node` is now `litchi::node::Node`.
- `view` takes `mode` and `json` only. Trunk keeps `view::mode::Options`
  (`start`, `end`, `max_lines`, `columns`, `range`) crate-private, so those
  keywords are gone (follow-up F2). `Mode::parse` replaces `FromStr`. The
  bindings give the did-you-mean suggestion for an unknown mode themselves,
  on the error path only.
- Test oracles follow trunk's behaviour: canonical XLSX paths are positional
  (`/sheet[1]/A1`); path-error messages are content-free; PPTX, XLSB and the
  legacy formats now read, so `PptxPendingTests` became `PptxReadTests`; DOCX
  headers, notes and comments are listed gaps (L-DOCX-R5), and the test
  checks that they refuse with `unsupported_capability`. Trunk has no Rust
  goldens for those three elements any more.

**Write surface (new).**

| Python | Rust | Returns | GIL released for |
|---|---|---|---|
| `create(format)` (`FileFormat` or a name such as `"docx"`) | `litchi::create` | `File` | building the document |
| `File.set(target, props)` | `File::set` | new `File` | the edit |
| `File.add(parent, type, props=None, *, index, after, before)` | `File::add` with `NewNode` + `Placement` | `(File, path)` | the edit |
| `File.remove(target)` | `File::remove` | new `File` | the edit |
| `File.move(path, to=None, *, index, after, before)` | `File::move_to` | new `File` | the edit |
| `File.swap(path, path2)` | `File::swap` | new `File` | the edit |
| `File.edit(fn=None)` | `File::edit` (staged runs, §2) | `Edit`, or `Commit` with `fn` | every run |
| `File.batch(items, *, mode="atomic")` | `File::batch` | `Commit` (`results`) | JSON decoding and the batch |
| `File.apply(item)` | `File::apply` | `Commit` | the command |
| `File.save(path=None, *, overwrite=True)`, `save_as(path, *, overwrite=True)` | `File::save_with` (`SaveTo`, `Overwrite`) | bytes written | the save |
| `File.to_bytes()` | `File::to_bytes` | `bytes` (one copy) | publication |
| `File.unsign()` | `Signed::unsign` | new `File` | — |
| `File.signed`, `File.locked` | `File::is_signed`, `File::Locked` | `bool` | — |

- **Props** are a dict (dotted keys) or OfficeCLI's `["key=value"]` list.
  `None`, `bool`, `int` (up to 128 bits, as `PropValue` holds them),
  `float` and `str` keep their types. Anything else is `PropertyError`
  (`invalid_value`). Coercion, unknown keys and did-you-mean suggestions are
  the Rust registry's. A test checks that dict and list forms publish the
  same bytes.
- **Batch items** are dicts in OfficeCLI's `BatchItem` shape, or JSON text
  (`str` or `bytes`). Dicts go through `BatchItem`'s own `Deserialize`
  implementation, via a small serde deserializer over Python objects
  (`src/pyde.rs`). That gives exactly the wire rules of the JSON decoder:
  case-insensitive keys, `op` as an alias, unknown or repeated keys refused,
  `None` as an absent field, i32 fields, both props forms. Every item is
  converted once, with no JSON text in between. JSON text is decoded by
  `litchi::batch::from_bytes` with the GIL released, borrowing Python's
  buffer. Tests check that dicts, JSON `str` and JSON `bytes` publish the
  same bytes, and that one batch equals the same sets made one at a time.
- **`Commit`** has `file`, `changed`, `operations`, `results` and `value`.
  `results` holds one dict per item in OfficeCLI's `BatchResult` members:
  `index` and `success`, then `changed` and `path` for an add, or `code`,
  `error` and `suggestion` for a failure. `value` is what `edit(fn)`
  returned.
- **Errors.** The S8a family classes (one per group of `Code`s) gain `item`,
  the zero-based batch item or transaction operation that failed. `item` is
  pickled too. For `document_protected`:
  - A signed file gets the Python suggestion "File.unsign()".
  - A locked file keeps the existing "cannot unlock" text, because the
    bindings build without the `encryption` feature (follow-up F4).
  - Any other protection suggestion is passed through.
- **The `.pyi` stub** declares `create`, the write verbs, `signed` and
  `locked`, `Commit`, `Edit`, the typing-only `_ItemResult`, `_Props` and
  `_PropValue`, and `LitchiError.item`. The stub-sync test passes.

## 2. Transactions: why `Edit` stages

A Rust `transaction::Edit` exists only inside the closure that
`File::edit` calls. A Python object cannot keep it between calls without
`unsafe` (`forbid(unsafe_code)` stays) or a hidden thread (goal
constraint 5). The Python `Edit` therefore works like this:

- It stages `set`, `remove`, `move` and `swap`.
- It runs the staged operations in one `File::edit` call, with the GIL
  released, at these points:
  - before `add` (so `add` returns the new element's path);
  - before a read (`get`, `find`, `query`);
  - at `commit`.
- Each run starts from the previous run's immutable snapshot, so nothing is
  visible before `commit` and the source `File` never changes.
- Consecutive sets in one run share the Rust edit's native coalescing.
  Rust's own `Edit` also closes runs at reads and structural commands.
- A test checks that `edit(fn)` with 500 cell sets publishes the same bytes
  as a batch of the same 500 sets.
- The first failure poisons the transaction. Its `item` is the global
  operation index: the Rust item index plus the operations of earlier runs.
- A staged operation's error surfaces at the next run (`add`, a read or
  `commit`), as Rust's native-commit errors surface at the next barrier.
  Argument errors (placement, props types) are raised at the call itself.
- The first run happens even when nothing is staged, so a locked or signed
  file is refused as Rust's `File::edit` refuses it.
- `with` commits on a normal exit and discards on an exception.
- `commit()` is idempotent: it returns the same `Commit`.
- Thread safety:
  - The state sits behind a `Mutex` that is only locked with the GIL
    released, so a thread waiting for the lock never holds the GIL.
  - A commit closes the transaction (`Committing`) inside the same critical
    section as its last run, so no operation can slip in after it.

Difference from Rust: with more than one run, `Commit.changed` is "some run
changed the document". Rust compares the final endpoints, so an `Edit` whose
runs cancel out reports `True` where Rust's single closure would report
`False` (finding 3).

## 3. Format coverage of the tests

Every verb is tested on every format whose `capabilities()` reports it
supported:

| Format | `set` | `add` | `remove` | `move` | `swap` | `batch`/`edit` | `save` | `create` |
|---|---|---|---|---|---|---|---|---|
| docx | run and paragraph props, scoped selector | paragraph/run, index/after/before | ✅ | index/after/before/to | ✅ | ✅ | ✅ | ✅ |
| xlsx | value, formula, text | sheet | sheet | sheet | sheet | ✅ | ✅ | ✅ |
| pptx | shape text/geometry, slide hidden | slide (layout), shape | slide, shape | slide | slide | ✅ (edit) | ✅ | ✅ |
| xlsb | cell | sheet | sheet, cell | sheet | sheet | ✅ | ✅ | ✅ |
| doc | paragraph text | refused | refused | refused | refused | ✅ | ✅ | ✅ |
| xls | cell, sheet name | refused | cell | refused | refused | ✅ | ✅ | ✅ |
| ppt | shape text and geometry (`datetime.ppt`) | refused | slide | slide | slide | ✅ | ✅ | ✅ |
| rtf | paragraph align, run bold | paragraph, run | paragraph | paragraph | paragraph | ✅ | ✅ | ✅ |
| odt/ods/odp | every write verb refuses with `UnsupportedError` |||||||

"Refused" is `CapabilityAgreementTests`: every verb a format's report marks
unsupported raises `UnsupportedError` with the report's `code`.

PPT shape writes work only on records the native editor models. Of the
`.ppt` fixtures in `test-data`, only `datetime.ppt` accepts them, and only
for text of the same length. PowerPoint-authored decks such as
`SampleShow.ppt` refuse with `unsupported_capability`, and the test asserts
that refusal. RTF structural edits need a body of one contiguous span, so the
test opens the literal input of `crates/litchi-ops/tests/rtf_unified_move.rs`.

Protection policies are tested as typed exceptions:

- A signed DOCX or XLSX:
  - `set` and `batch` raise `ProtectedError` with the `unsign()` suggestion.
  - An unchanged signed file publishes with its signature.
  - `unsign()` then `set` publishes an unsigned file.
- A locked DOC or XLS refuses `set`, `add`, `remove`, `batch`, `edit`,
  `to_bytes` and `save_as` with `ProtectedError`.

Other save cases:

- `file_required` for a created file.
- `file_exists` with `overwrite=False`.
- `invalid_argument` for another format's extension.

Thread tests (subprocesses, as S8a's):

- Eight threads write to one shared `File` with `set`, `batch` and `edit`.
  Every result holds only its own change, and the shared File publishes the
  same bytes before and after.
- Eight threads share one `Edit` and stage 200 sets; the commit holds all
  200.
- A ticker observes that `batch` (9,000 sets), `save_as` and `to_bytes` of
  an edited 75k-cell workbook release the GIL. `open` is accepted as
  "released" or as too short to observe, but never as "held".

## 4. Results

`crates/pyo3-litchi/tests/run_tests.sh` (Python 3.14.4, Linux x86_64) ran
**242 tests, OK**, in both dev and release builds:

- 132 pre-R9 tests;
- 58 R9 tests, adapted (one bounds test replaced by a view-option test, one
  DOCX story-gap test added);
- 49 write tests (`test_write_api.py`);
- 3 write-thread tests.

`cargo test -p litchi-py --lib` passes 2 tests. The gate commands and their
exit codes are in the commit message and the final report.

## 5. Python overhead vs the Rust call (indicative)

Method:

- Release builds (`lto`, `codegen-units = 1` for `litchi-py`), pinned with
  `taskset -c 2`.
- Process order R P P R R P P R.
- Each process takes the median of 21 batch means per scenario. Each cell
  below is the median over the four processes of each leg.

The Rust leg is a scratch example, not committed (scenarios below). It was
built as `cargo build --release -p litchi-ops --features
docx,xlsx,pptx,xlsb,doc,xls,ppt,rtf,odt,ods,odp --example p14_timing`
(binary sha256 `b254c1f1…a0210`). The Python leg is a scratch script on the
release `liblitchi_py.so`: 33,888,552 bytes, sha256
`628d1fb3fa1e996bf2ea7b7de1d039300308e92e3f15c748e50064f20b912129`.

Each Rust read leg is the call plus `serde_json` of its node views into a
reused buffer: the Rust counterpart of producing the whole result. The
Python legs return dicts. Raw runs:
`docs/unified-ops/results/0255-python-overhead.json`.

| Scenario | Rust | Python | Python ÷ Rust | Δ per call |
|---|---|---|---|---|
| DOCX `get("/body/p[10]", 0)` (`drawing.docx`) | 60.71 µs | 62.66 µs | 1.032 | 1.96 µs |
| DOCX `get("/body/p[10]")` (depth 1) | 154.84 µs | 158.50 µs | 1.024 | 3.66 µs |
| XLSX `get("/ExcelTable/C2", 0)` | 6.01 µs | 6.62 µs | 1.102 | 0.62 µs |
| DOCX `query("paragraph")` (227 matches) | 7.117 ms | 7.406 ms | 1.041 | 290 µs |
| XLSX `query("cell", 100)` (75k-cell sheet) | 476.1 µs | 482.7 µs | 1.014 | 6.6 µs |
| DOCX `set` run bold (`Hyperlink.docx`) | 106.52 µs | 109.31 µs | 1.026 | 2.8 µs |
| XLSX `set` one cell (created workbook) | 21.70 µs | 22.14 µs | 1.020 | 0.44 µs |
| XLSX `batch` of 1k cell sets (dicts) | 2.070 ms | 2.419 ms | 1.169 | 349 µs |
| XLSX `batch` of 1k, JSON text | 2.070 ms (dict leg) | 2.341 ms | 1.131 | 271 µs |
| XLSX `edit(fn)` with 1k `tx.set` | 2.012 ms (`File::edit` closure) | 2.279 ms | 1.133 | 267 µs |
| DOCX `batch` of 1k run-bold sets | 22.32 ms | 23.20 ms | 1.040 | 882 µs |

Spread (largest minus smallest of four processes, over the median) is at
most 2.4% for every cell except these:

- Rust `query("cell", 100)`: 17.5%, from one slow process.
- Python DOCX `set`: 12.9%, from one slow process.

Neither changes the reading.

Reading:

- **Single calls** (`get`, `query`, `set`) cost 1.4–10% over the Rust call
  plus its JSON. The fixed cost is about 0.4–0.6 µs per call (argument
  parsing, GIL release and reacquire, wrapping the new `File`). Only the
  6-µs XLSX cell read shows it as 10%.
- **Batch of 1k:**
  - XLSX costs 17% more. That is about 0.35 µs per item to read a 3-key dict
    with a nested props dict through serde.
  - JSON text is cheaper (0.27 µs per item, decoded with the GIL released).
  - `edit(fn)` costs 0.27 µs per `tx.set`.
  - DOCX's heavier per-item work brings it to 4%.
- There is no budget for Python overhead in the goal. Against the goal's
  1k-batch budget (within 2× of a hand-written leaf edit), the Python leg
  adds 0.13–0.17× of the Rust unified leg.

## 6. Gaps and follow-ups

- **F1 `dump`.** `litchi-ops` has no `dump` (OfficeCLI batch export), so
  Python has none. Add `File::dump` in ops, then a one-line binding.
- **F2 view options.** `view::mode::Options` is crate-private, so Python
  cannot pass OfficeCLI's `--start/--end/--max-lines/--cols/--range` (R9 had
  them). Needs a public typed options API (`File::view_with`) in ops.
- **F3 `raw()`.** There is no part-level raw API in ops (design §12). A
  zero-copy `bytes` would need the buffer protocol, which needs `unsafe` in
  PyO3. So `to_bytes()` makes one copy (publication into a `Vec`, then
  `PyBytes`); the record says so. Revisit when ops gains `raw(part)`. A
  `memoryview` over `Arc<[u8]>` would need a PyO3 safe buffer API.
- **F4 unlock and re-encrypt.** The bindings build without the facade's
  `encryption` feature, so neither `Locked::unlock` nor
  `SaveOptions::reencrypt` is exposed. Locked files stay read- and
  write-refused with the existing suggestion.
- **F5 `query` streaming.** `query` still returns a list: R9's reason,
  because `Matches` borrows the `File`.
- **F6 Rust syntax in suggestions.** Some ops suggestions name Rust APIs,
  for example "save it with `File::save_as(path)`". The bindings rewrite
  only the protection suggestions. A per-surface suggestion channel would
  fix the rest (R9 finding 2).
- **F7 exception classes.** Exceptions are one class per error family (the
  goal's "subclasses for the major families"), not one class per `Code`;
  `code` names the exact code.
- **F8 property name in errors.** Rust's `Error::property()` is not
  surfaced as an attribute. Messages are content-free, as in Rust.

## 7. Findings

1. `capabilities()` still keeps the GIL (R9 finding 1, unchanged).
2. Python `add` returns `(File, path)`, not a bare `File`, because the new
   element's path is part of the result (Rust: `Added { file, path }`).
3. `Edit.changed` across several runs is "some run changed the document"
   (§2).
4. A batch is bounded by the default `SemanticLimits` of 10,000 operations.
   Python exposes no limits; the 10,001-item test raises `limit_exceeded`.

## 8. Not claimed

- Only Linux x86_64 and Python 3.14.4 ran. The `abi3-py38` floor, macOS
  and Windows were not tested.
- The stub was checked by the stub-sync test and against real results, not
  by mypy or pyright.
- The timings are indicative (§5).
- No ODF or iWork write path exists. Those are other streams.
