# 0619: the standing XLS lifecycle failure is change 0565's, and it is an assertion the record already superseded

Status: retained, harness-only correction and attribution. `performance_claim:
none` — this record carries a bisect, a per-read attribution trace and a test
correction, not a claim-registry entry. **No file under `crates/` was
modified.**

OLE2 and OOXML remain the active priority. ODF optimization stays deferred until
that goal completes; iWork is excluded.

Change [0601](0601-perf-harness-real-producer-shape.md) reported a standing
harness failure it did not introduce and did not fix:
`tests::xls_source_backed_lifecycle_selectors_are_matched_and_local` asserts
`open_reads_zero_worksheet_payload == [true]` and gets `[false]`, and its panic
poisons the shared allocation-metrics mutex so six further tests fail as
cascades. It named change 0595's area as the place to look. This record
establishes that **0595 is not responsible, 0605 is not responsible, and no
library behaviour regressed**: the first bad commit is `c1d2caf85`, change
[0565](0565-xls-globals-single-pass.md), which bounded the read-locality *gate*
and left the same pre-change contract asserted at three more sites in the
harness's own unit test — sites that change 0565's record has already described
in prose and measured to the byte.

## What was changed

One test in `tools/perf-baseline/src/lib.rs`
(`tests::xls_source_backed_lifecycle_selectors_are_matched_and_local`). Three
assertion groups now state the contract change 0565 established instead of the
one it replaced:

| Case | Was asserted | Now asserted |
| --- | --- | --- |
| `XlsSourceBackedOpen` | `open_reads_zero_worksheet_payload == [true]`, `unselected_worksheet_read_bytes == [0]` | the published boolean equals its own definition (`unselected_worksheet_read_bytes[0] == 0`), `selected_worksheet_read_bytes == [0]` (unchanged), and `unselected_worksheet_read_bytes[0] <= XLS_GLOBALS_MAX_WINDOW_BYTES` |
| `XlsSourceBackedOpenListWorksheets` | `open_reads_zero_worksheet_payload == [true]` | the same three |
| `XlsSourceBackedOpenOneCell` | `selected_query_reads_only_selected_worksheet == [true]`, `unselected_worksheet_read_bytes == [0]` | the published boolean equals its own definition, `selected_worksheet_read_bytes[0] > 0` (unchanged), and `unselected_worksheet_read_bytes[0] <= XLS_GLOBALS_MAX_WINDOW_BYTES` |

**No published field changed.** `open_reads_zero_worksheet_payload` and
`selected_query_reads_only_selected_worksheet` keep their strict definitions and
their current values, because their names state the strict property and `false`
is the honest value for this corpus. `docs/performance/results/change-0412/verify-capture.py`
reads `xls.open_reads_zero_worksheet_payload` by name and retained packets
carry it (`results/change-0533/candidate/analysis.json`, among others);
redefining it would silently reinterpret every one of them. Nothing else moved:
no production crate, no corpus, no selector, no JSON schema, no gate function.

## Why it is sound

### The failure is an over-read change 0565 designed, documented and measured

`GlobalsBuffer::fill_cap`
(`crates/litchi-xls/src/workbook/source.rs`) clamps each fill by the smallest
`BoundSheet8` stream position **once one has been framed**. A fill issued before
any `BoundSheet8` exists is bounded by the stream length and by
`max_global_bytes` instead, so on a workbook whose globals are shorter than the
fill schedule reaches, the last fill lands past the globals end. The bytes are
dropped by `bytes.truncate(global_len)` before framing and are never
interpreted, hashed, logged, placed in an error message or published.

Change 0565's own record says so under *A third site held the old contract*:
"that corpus has only 1,483 bytes of globals, so the last fill is issued before
any `BoundSheet8` exists and reads **93 bytes** of an unselected worksheet
body", and its *Limitations* repeat the figure: "zero bytes on the flagship
fixture, **93 on the harness locality corpus**, at most 3,949 across the
surveyed corpus". The measurement below reproduces 93 bytes exactly, 54 records
later.

### What 0565 changed, and the one thing it did not

0565 relaxed `validate_xls_source_locality` — the gate every measured sample
runs — from "an open reads zero worksheet bytes" to "an open reads **no byte of
the selected worksheet** and at most one globals window of any worksheet body".
That is a strictly weaker statement in one respect and a strictly stronger one
in another, and it still fails an open that materializes a worksheet, which is
the regression the gate was written to catch: this corpus's unselected
worksheet body is 78,849 bytes against a 65,536-byte window.

What it did not do is update the two published booleans, which are computed
with the old strict `== 0`, or the unit test that asserts them. `tools/perf-baseline`
is a separate Cargo project; 0565's correctness evidence covers `litchi-xls`
(1,351 tests) and one facade test in `crates/litchi/src/sheet/workbook.rs`, and
its gate list does not include `cargo test --manifest-path tools/perf-baseline/Cargo.toml`.
The harness test therefore went red at `c1d2caf85` and stayed red.

### Error identity, refusals, limits

Untouched. This change edits assertions inside one `#[cfg(test)]` function. No
typed error, no limit, no refusal point, no validation order, no output byte and
no public API is reachable from it.

## Measured

Host: AMD EPYC 9R45, 32 cores, 123 GiB, Linux 7.0.0-1012-aws, rustc 1.95.0.
Every leg is `cargo test --release --locked` over the single test, each in its
own detached checkout with its own external `CARGO_TARGET_DIR`. The test is
deterministic — one generated corpus, no clock, no PRNG, no ambient I/O — so one
run per leg is the whole evidence, and no timing was taken.

### Bisect (measured)

| Commit | Change | `xls_source_backed_lifecycle_selectors_are_matched_and_local` |
| --- | --- | --- |
| `6b13261e5` | parent of 0565 | **ok** |
| `c1d2caf85` | **0565**, globals single windowed pass | **FAILED**, `left: [false] right: [true]` |
| `2391a3462` | 0568/0570/0571 | FAILED |
| `93a610ded` | 0577-0583 wave | FAILED |
| `08d968f8e` | 0587 survey base | FAILED |
| `c1503db2b` | **0595**, frame loop and SST walk | FAILED |
| `f8cf7d2a1` | 0600/0601 base | FAILED |
| `1e4198321` | current head (0606) | FAILED |

First bad commit: **`c1d2caf856fdbfa3b60b6de6cb2e85dc9ee9abab`**. The failure
predates change 0595 by four production changes and change 0605 by ten, so
neither record's "reads, bytes and observations unchanged per scenario" is
contradicted by it and neither needs a correction.

### Per-read attribution at `1e4198321` (measured)

Every `InstrumentedSource::read_at` of one `XlsSourceBackedOpen`, with each
read's overlap against the five classification range sets. The probe is
retained as `results/change-0619/trace-probe.patch`; its output is
`results/change-0619/trace/source-backed-open-read-trace.txt`.

Corpus `xls-comments-opaque-heavy`: archive 16,995,840 B, Workbook stream
80,946 B. Classification: `cfb_structural` 267 runs / 136,704 B;
`workbook_global` 3 runs / 1,483 B, physical `512..1995`; `selected_worksheet`
("Untouched") 3 runs / 614 B; `unselected_worksheets` ("Comments") 155 runs /
78,849 B, beginning at physical 1995.

Fifteen reads. Eight are structural (header, directory, FAT runs, 136,704 B).
The other seven are the globals scan:

| # | Offset | Length | Globals bytes | Unselected bytes | Phase |
| ---: | ---: | ---: | ---: | ---: | --- |
| 9 | 512 | 4 | 4 | 0 | exact prologue, record 1 header |
| 10 | 516 | 20 | 20 | 0 | exact prologue |
| 11 | 536 | 6 | 6 | 0 | exact prologue |
| 12 | 542 | 6 | 6 | 0 | exact prologue |
| 13 | 548 | 4 | 4 | 0 | exact prologue, 5th header |
| 14 | 552 | 512 | 512 | 0 | window fill 1 (`GLOBALS_FIRST_WINDOW_BYTES`) |
| 15 | 1064 | 1024 | 931 | **93** | window fill 2 (doubled) |

The prologue covers logical `[0, 40)`; fill 1 covers `[40, 552)`; fill 2 is
issued at logical 552 with no `BoundSheet8` framed yet and runs to logical
1,576, which is **93 bytes past the 1,483-byte globals end**. Totals for the
open: `cfb_structural` 136,704 B, `workbook_global` 1,483 B,
`selected_worksheet` **0 B**, `unselected_worksheets` **93 B**,
`opaque_payload` **0 B**; 15 reads, 138,280 bytes.

93 ≤ 65,536, so `validate_xls_source_locality` passes on every measured sample
and always has — the gate is not what fails. `open_zero`, computed as
`selected == 0 && unselected == 0`, is `false`, and the test asserted `true`.

### After the correction (measured)

`cargo test --release --locked ... xls_source_backed_lifecycle_selectors_are_matched_and_local`
at `1e4198321` with the corrected assertions: **ok, 1 passed, 0 failed**.

Evidence tiers: **measured** for all eight bisect legs, the fifteen-read trace
and every byte total above. **Modelled**: nothing. **Unknown**: whether any
fixture outside this corpus and change 0565's 104-fixture survey drives the
over-read above 3,949 bytes.

## Correctness evidence

| Gate (`tools/perf-baseline`, a separate Cargo project) | Result |
| --- | --- |
| `cargo fmt --all --check --manifest-path tools/perf-baseline/Cargo.toml` | exit 0 |
| `cargo clippy --locked --manifest-path tools/perf-baseline/Cargo.toml --all-targets` | exit 0, no warning |
| `cargo doc --locked --manifest-path tools/perf-baseline/Cargo.toml --no-deps` | exit 0, no rustdoc warning |
| `cargo test --locked --release --manifest-path tools/perf-baseline/Cargo.toml` | **508 passed, 0 failed, 1 ignored**, 82.27 s; `grep -c FAILED` over the whole run is 0 |

No crate under `crates/` was touched, so no production crate's `clippy`, `test`
or `doc` gate is in scope. The seven Rust failures change 0601 reported —
the XLS lifecycle assertion and its six `PoisonError` cascades — are all
gone: change 0601 recorded 501 passed / 7 failed / 1 ignored at `f8cf7d2a1`,
and this branch records 508 passed / 0 failed / 1 ignored over 509 tests. The
allocation-metrics mutex is no longer poisoned, so nothing cascades.

The four Python failures change 0601 reported (`test_check_crate_boundaries`,
`test_native_odf_resave`, `test_perf_claims`, `test_perf_compare`) are
untouched by this change and remain pre-existing; nothing here reaches them.

## Validation preserved

Nothing about validation changed, because no production code changed. The
read-locality gate `validate_xls_source_locality` is byte-identical: it still
refuses a source-backed XLS sample whose open reads any byte of the selected
worksheet, more than one globals window of any worksheet body, any opaque
payload byte, or zero structural or zero globals bytes. Every measured sample
in every retained packet ran through it unchanged.

## Limitations

- **Nothing here makes anything faster**, and nothing here is a claim. No
  speedup, regression, allocation, peak-RSS, cold-cache, physical-I/O,
  range-source or cross-platform result is stated.
- **The published booleans keep the strict pre-0565 meaning.** After this change
  the harness publishes `open_reads_zero_worksheet_payload: false` for a
  source-backed XLS open on this corpus and
  `selected_query_reads_only_selected_worksheet: false` for its one-cell query,
  which is honest under the names but is not the invariant the library
  guarantees. A future change that wants the fields to state the bounded
  invariant must rename them and migrate
  `results/change-0412/verify-capture.py` with the rename; that is a schema
  change and is deliberately not made here.
- **The over-read is corpus-shaped.** 93 bytes is this corpus's figure, driven
  by 1,483 bytes of globals and the 512-byte-doubling fill schedule. Change
  0565 surveyed 104 container fixtures and found at most 3,949 bytes; no
  measurement here widens that.
- **Only one test was run per bisect leg.** The legs establish where the
  assertion started failing, not that nothing else changed at those commits.
- **No production fix exists for the over-read** and none is proposed. The
  globals end is not knowable until a `BoundSheet8` has been framed, which is
  the reason 0565 bounded the contract rather than preserving it; clamping
  earlier is the "exact until the clamp is known" schedule 0565 measured at
  7,872 reads against 434.

## Retained evidence

[`results/change-0619/`](results/change-0619/README.md) — the eight bisect
transcripts, the fifteen-read attribution trace, the probe patch, the gate
tails, `decision.json` and `log-sections.md`.
