# Reference and worksheet metadata implementation

Status: implementation, independent reviews, isolated gates, and performance
review complete. The broader specification audit remains open.

This batch implements `AREAS`, `COLUMN`, `COLUMNS`, `ISREF`, `ROW`, `ROWS`,
`SHEET`, and `SHEETS` in the existing ODS formula evaluator, against baseline
`049c09cdde3978593149079c4257df047a3fa419`.
[contract.md](contract.md) defines the normative implementation profile.

## Resulting behavior

The resolver evaluator consumes reference descriptors before intersection or
cell materialization. It counts logical reference records, reports dimensions
and physical sheet order including hidden sheets, and produces bounded matrix
row/column vectors. Direct metadata operations make zero cell reads.
`AREAS` and `ISREF` accept reference lists; the six other functions reject list
identity even when a list retains only one record. No-argument coordinate and
sheet functions use the fixed explicit formula position.

Computed arguments retain scalar or matrix evaluation context. `SHEET` maps
computed Text arrays elementwise, including arrays produced by IF, IFERROR,
and IFNA inside projected branches. ROW/COLUMN reject computed value arrays
at their ordinary pseudotype boundary. Descriptor probing distinguishes
unknown geometry, an internal reference-operand refusal, and provider failures.
Only metadata discovery may defer the internal refusal to value evaluation;
ordinary reference operators preserve their typed refusal. Provider failures
from reads, sheet lookup, range geometry, and reference/MUNIT shape discovery
propagate immediately without retrying or becoming formula errors.

Source-qualified references use a distinct private marker. Scalar selection
and error-handler fallbacks preserve it until ISREF classifies it or
SHEET/SHEETS apply their formula-error constraint. Arithmetic and selected
matrix cell values still require external dereferencing and retain typed
capability failures. Source policies restore across scalar handlers and reset
for each selected matrix cell. The context-free scalar API supports only the
subset available without workbook context; other cases return typed refusals.
No public API, ambient external fetch, named-expression resolver, workbook
recalculation engine, or formula-cache publication was added.

## Correctness and review evidence

The corrected isolated run passes all seven checks and 1,655 tests, with no
failures or ignored tests: package tests, strict all-target Clippy, strict
rustdoc, package formatting, selected-source formatting, crate boundaries, and
diff checks. [gates/freeze.json](gates/freeze.json) and adjacent receipts bind
the exact source, commands, logs, and isolated lock.

Focused coverage comprises 19 semantic tests, 12 resource tests, and an
independent executable oracle with 91 observations. The native fixture contains
32 formula rows and an actually hidden sheet: 28 exact comparisons plus four
retained LibreOffice Err:504 divergences for locally rejected reference lists.
Native observations do not override the ODF contract; see
[native/README.md](native/README.md).

Independent [semantic](semantic-review.md) and
[resource/cache](resource-review.md) source reviews pass and are bound by
[review-receipt.json](review-receipt.json). Resource tests cover zero cell reads,
checked storage/output limits, cancellation, source fences, borrowed text,
cache contexts, and typed provider failures during nested shape discovery.

## Performance evidence and limitations

The superseded source produced one complete 4,650-sample capture: 840 baseline
and 3,810 candidate measurements, with three warmups and 15 fresh processes
per case and phase. Its 56 matched groups had identical allocation, work,
read, and output accounting. Median latency changes ranged from −3.854% to
+5.865%; RSS changes ranged from −4.989% to +4.311%.

The evaluate-only SUMIFS control was the sole latency review flag: 83,295 to
88,180 ns/repeat (+5.865%). The profile's descriptive 100,000-resample interval
was +0.528% to +6.831%; the independent root 5,000-resample audit gave +0.528%
to +6.818%. Parse-evaluate was −1.179%. Root accepts and retains this observed
regression with unchanged work (1,839), reads (1,792), allocator calls (42),
and output accounting. The cause is unestablished; unchanged counters do not
prove timing noise. These are evaluator measurements on a non-isolated host,
not end-to-end workbook speedup evidence.

The final corrected-source capture contains 4,920 samples (840 baseline and
4,080 candidate), with 136 candidate preflight cases. Together the two complete
captures retain 9,570 timed samples. Final matched latency changes range from
−4.571% to +6.284%, and RSS from −5.436% to +4.722%. Accounting sets remain
identical. SUMIFS evaluate is 83,060.5 to 88,280 ns/repeat (+6.284%); the root
5,000-resample interval is −0.937% to +12.014%. Parse-evaluate is −1.852%.

Root accepts this measured limitation for the added functionality. Independent
static review found no added per-cell work in the SUMIFS scan; new dispatch
and source-marker checks occur outside it. Code layout and host variability
are possible explanations, but the cause remains unestablished. No timing
no-regression claim is made and no further capture was used to select a result.
See [root-performance-audit.json](root-performance-audit.json).

The captured standalone performance verifier had a stale zero-read assumption
for six computed ROW/COLUMN cases. Their frozen case matrix, oracle, and Rust
preflight correctly require two reads. The independent root audit validates
exact reads against that contract and retains the original verifier failure;
see [performance-verification-note.md](performance-verification-note.md).
One setup-only failed invocation produced zero timed samples.

## Diagnostics, custody, and cleanup

Three superseded snapshots are retained as diagnostic evidence:

- `diagnostics/source-policy-gates/`: earlier seven-gate pass before the
  scalar error-handler source-policy correction.
- `diagnostics/computed-array-preflight/`: 1,647-test pass and original complete
  capture before computed-array and typed-provider probe fixes.
- `diagnostics/reference-operator-refusal/`: the full integration test caught
  generic reference operators incorrectly accepting array operands; six other
  gates passed. No timing capture used that intermediate source.

The isolated lock is the retained `58b4be6c…` copy; the ambient root lock
`aa945c79…` remains unchanged. Unrelated tracked edits recorded in
[baseline.json](baseline.json) remain outside this batch. Owned checkouts, targets, registry, and bytecode were removed;
[gates/cleanup.json](gates/cleanup.json) records cleanup. Retained-only
verification is recorded in [root-verification.json](root-verification.json). Verifier infrastructure landed separately in
`08e09647e2`; it does not represent completion of the implementation batch.
