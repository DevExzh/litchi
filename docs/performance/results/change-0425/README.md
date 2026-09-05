# 0425: non-iWork strict verification maintenance

Rust 1.98 strict Clippy rejected constant-size chunk iterators throughout the
non-iWork codecs. This batch uses borrowed fixed-array chunk views, preserving
existing bounds, full-chunk iteration, mutable buffers and explicit remainder
checks. Runtime-sized chunk widths remain dynamic. No benchmark speedup,
allocation, peak memory or I/O claim is introduced.

The initial four PPTX findings were fixed first. Its scoped strict OPC/PPTX
Clippy gate and all 844 PPTX tests/doctests passed before the broader migration.
The wider 45-package audit then exposed shared and downstream findings. Each
failed stage and subsequent check is retained under `checks/`, with commands,
source hashes, raw output and terminal status. Test counts from earlier stages
must not be added to later repeated full suites as unique tests.

The package selection is derived from all 64 workspace manifests in
`package-selection.json`. Seventeen iWork packages are excluded. `litchi-py`
unconditionally enables iWork and remains outside this gate. The 45 other
packages use all features; root `litchi` requires a separate explicit non-iWork
feature selection. Shared ZIP, XML and Office substrate owners stay included.
Standalone performance and native-resave tools have their own manifests.

The ODF `ArchiveReaderKind` large-variant warning is an unchanged layout
question: its borrowed reader occupies at least 312 bytes, while the prepared
handle is eight. Adding a box would introduce allocation; that is not a
mechanical compatibility fix. `checks/layout-debt.json` binds the unchanged
source and compiler finding. Diagnostic runs that exempt this one lint
category do not establish a strict full-workspace pass. No production lint
allowance was added.

Read-only reviews cover shared substrates, DOC/PPT, XLS/XLSB and crypto/font/
image/OLE codecs. They check MD4 padding, AES alignment, UTF-16 terminators,
remainder rejection, iterator ordering and array typing. Test setup and
assertion lints are fixed without changing their intended cases. The XML
attribute decoder's late-assigned result becomes an expression while retaining
its checked numeric conversion, scratch buffer and fallible output reservation.

`check.py --tag NAME -- COMMAND...` runs one source-bound command and refuses
existing receipts/logs. All Cargo/test/profiling workloads stay serialized;
Rust 1.98.1, four build jobs and one test thread are explicit. No native Office,
fuzz, sanitizer, hardware-counter or performance recapture is implied by these
checks. The full non-iWork goal remains active and incomplete.

## Test corrections and remaining gates

The 45-package full test run contains 15,775 passes, three failures and 98
ignored tests. All three failures reproduce exactly at the clean pre-batch
revision with identical locked dependencies. The ODP fixture now records its
validated content range before corrupting central size metadata; its payload
read-limit assertion is unchanged. ODT now refuses CRC and both size mutations
and asserts byte-exact atomicity. XLSB retains six parsed corpus anchors,
including a chart-sheet anchor, and separately tests all five supported
worksheet transfers with exact one-anchor reopen checks. Their three complete
integration targets pass 27 tests. Only those three test files differ from the
full run, giving composite coverage of 15,778 passing tests, with the same 98
ignored tests. This is a full run plus target reruns, not a second full run.

The explicitly configured facade has 455 passes, six failures and 11 ignored
tests. The same six failures reproduce at the clean baseline in its two failing
targets (360 passes, six failures). Detection/limit arbitration, deferred XLSX
access, XLSB wrong-format errors and legacy XLS formula extraction remain open.
The test assertions are retained; this batch makes no facade correctness-pass
claim. Source-only diagnoses and the precise test names are under `checks/`.

The strict 44-package gate passes, excluding ODF's unchanged layout finding;
the corrected ODP/ODT/XLSB integration targets also pass strict Clippy. The
45-package run with a command-local layout-lint exception is diagnostic only.
The facade retains 18 findings, the standalone performance harness retains 29,
and native-resave's locked command stops before compilation because its
unchanged tracked lockfile needs updating. These are explicit open gates.

`workspace-Cargo.lock` retains the workspace's ignored dependency lockfile.
Baseline checkouts require copying it to their root before using `--locked`.
The first baseline attempts failed before compilation without that file; those
attempts remain recorded. Earlier receipts hash tracked Rust/manifests/locks;
subsequent receipts additionally bind the ignored workspace lock. The full
suite driver's historical `passed_tests` value counts only successful target
summaries and omits passes within failed targets. `verify.py` independently
counts all summaries and checks the composite result and source transition.

## Final passing checks and replay

The affected standalone XLS/CFB helpers pass eight harness tests. Warning-denied
rustdoc passes for all 45 selected packages and the explicit facade feature set.
Crate boundaries pass with 17 declared iWork-only debt edges; the CRUD index
validates 15 categories and 32 selectors; all nine registered claims replay in
strict mode. Rustfmt passes for all 194 changed Rust files. These checks are
separate from the open test/lint/lockfile gates above.

From the repository root, replay retained output and composite coverage with:

```sh
python3 -B docs/performance/results/change-0425/verify.py
```

The replay verifies all 34 command receipts, original log hashes, exact failure
censuses, the test-only source transition, strict package selection, warning-
denied rustdoc flags, and existing claim replay. Logs are stored losslessly as
`.log.gz`; `log-storage.json` records both original and stored hashes. Failed
attempts are retained. `SHA256SUMS` inventories every other bundle file;
`checks/replay.json` records the final replay after temporary checkout removal.
