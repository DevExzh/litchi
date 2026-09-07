# 0457 integration evidence

Both candidate source epochs have completed output capture and formal
measurement derivation. Performance review, initial-source profiling, and both
ASAN fuzz lanes are retained. Final precleanup and portable-copy verification
pass; owned staging cleanup is recorded in `cleanup.json`. The full non-iWork
performance goal remains open.

The initial compiler attempts are retained as `checks/integration-check-1`
through `-5`; `integration-check-6` passed. The corrections included callback
error conversions, explicit lifetimes, quick-xml API use, and preserving the
primary source-member error in the ODP layer.

The first ZIP release test run exposed compression that depended on callback
write boundaries. Deflate now receives fixed 64 KiB input windows plus one
final tail window. Both chunk-boundary regressions pass in `zip-r1`; the
decoded and compressed SHA/CRC/length replay checks remain enforced. This adds
one fixed input buffer, for 192 KiB of replay/copy scratch plus compressor
state, independent of the member length.

The initial XML guard test incorrectly assumed its first read bypassed the
three-byte BOM probe. Its correction checks accumulated bytes, the absence of
an exposed extra marker, and the typed token-ceiling failure. The ODP empty
presentation test initially failed while constructing a noncompact fixture;
the compact empty fixture now reaches the actual append refusal.

Generated-page semantic readback uses the established ODP parser on the
bounded authored fragment. Its memory reservation includes actual submitted
string capacities, expanded text, parser strings and tokens, fragment growth,
the wrapper, and fixed parser state. Space runs are split at the parser's
per-control ceiling; whitespace-only body text is checked in its retained
object shape. Inputs whose carriage returns or Unicode whitespace would be
normalized are explicitly refused. Final review additionally released the
submitted strings before the readback reservation ends, ahead of common
planning; this small lifetime correction follows the full `*-r1` test runs.

The parent sink now invalidates flush state before nonempty writes and before
another flush. Error-only ZIP replay measurements and common transport replay
diagnostics are boxed to keep success result values small. Original replay
diagnostics remain available with the mapped public transport cause.

`strict-r4` and `harness-strict-r3` passed with warnings denied. The focused
`insertion-tests-r2` run passed all 20 tests. Full release retries passed 481
ZIP tests (2 ignored), 497 ODF-common tests (1 ignored), and 368 ODP tests.
Receipts bind the exact source snapshot for each result; later source changes
must not be represented as covered by an earlier receipt.

The first full harness run passed 344 tests and failed one new assertion that
compared a runtime provider ID with an independently created fixture provider
ID. Runtime proof ID/revision vectors now bind each publication report to its
own plan and provider outside timing. Fixture identity remains separate. The
full `harness-r1` retry passed 383 tests (1 ignored). `opc` passed 497 tests
(1 ignored), and `odp-append-final` passed all 10 append tests after the final
input-lifetime correction. The native probe now checks exact body equality.

The full workspace format attempt retains its Keynote failure. The user
excluded iWork. `format-non-iwork` and `harness-format-r2` passed after formatting
the affected files and pre-existing differences in the DOCX glossary test,
PPTX boundary harness, and XLSX filesystem test. Final formatting receipts are
`format-non-iwork-r1` and `harness-format-r3`. Rustdoc with warnings denied,
the workspace feature check, and crate-boundary checks also passed.
The existing inline borrowed
archive-reader enum has a narrow documented Clippy allowance to preserve its
inline storage without adding a heap allocation; its representation was not changed.

The synthetic fixture export and independent ZIP/XML output binding passed for
all three sizes. Native publication and the independent oracle passed for all
ten retained LibreOffice fixtures. These are fixture and library-reopen checks;
no LibreOffice application was launched. The first aggregate native preflight
mistook a retained gzip path for the live producer output path. Its verifier now
checks the exact runner path, `/tmp/litchi-goal-0457/native/{index:02d}.odp`,
separately from the retained archive path. `native-record-preflight-r1` passed
all ten record checks and oracle replays; the original failure is retained.

Both candidate performance phases completed all six lanes, retaining 360
samples in total. The first candidate precleanup check compared build custody
against an absent duplicate field in the candidate protocol. The verifier now
authenticates the actual role, binary binding, source manifest, and before/after
build records, and checks the duplicate if present. Frozen capture inputs and
receipts were not rewritten. The failed `candidate-precleanup` and `-r1`
receipts remain: a second direct reference to the same absent duplicate in the
phase verifier also needed correction. `candidate-precleanup-r2` passed all
12 lanes and 360 samples. `candidate-derive` and `comparison` passed, retaining
the source vectors, recomputed quantiles, explicit flags, and deterministic
10,000-resample median-delta bootstrap intervals.

The final `strict-final` check passed after the input lifetime correction.
Each ZIP/XML fuzz lane passed offline lock generation, ASAN build, and 1,000
iterations. `fuzz-binary-retention` passed after lossless gzip storage and
streamed verification of the two captured executables; original inventories
and shared target-cache binaries remain intact.

Final code review found a separate plan-lifetime budget defect: the authored
fragment retained its caller-provided Vec capacity after the local preparation
reservation ended. The corrected plan charges actual capacity before source
reads and holds that lease through fragment drop; publication independently
charges capacity under its own options. The two new regressions are described
in `retained-plan-budget-amendment.md`. `retained-plan-tests` passed 13 tests,
`odf-common-final` passed 499 with one ignored, and `odp-retained-plan` passed
368. A deliberate drop-order move triggered Clippy's redundant-local lint;
renaming the input parameter preserved the move without an allowance.
`strict-retained-plan-r1`, `harness-strict-retained-plan`, `doc-retained-plan`,
and `format-retained-plan` passed. The failed lint receipt remains.

`candidate-final-build` binds the corrected source epoch. Its fresh fixture
export and independent output binding passed for all three sizes, with exactly
the initial source/output archive and XML identities. `native-final` and its
preflight passed all ten fixtures and oracle replays. The final R1/R2 capture
retains another 360 samples; `candidate-final-precleanup`,
`candidate-final-derive`, and `comparison-final` passed. This final matrix is
the current performance evidence; the original candidate matrix remains
historical and is not relabeled.

The ASAN ZIP and XML targets ran on the initial source epoch. Their target and
scanner/replay code is unchanged by the retained-fragment lease correction;
the new managed-plan behavior is covered by the integration/common/ODP tests
above. Initial-source profiling retains the failed source-epoch preflight and
sample-count oracle attempt. Candidate R2 profiling passed. The control R2
workload, sampling and postprocessing succeeded, then passed the already
accepted amended control oracle in `profile-control-oracle-r1`; the recorder's
original oracle failure is preserved and explicitly bound by
`profiling/control-oracle-amendment.json`.

Final bundle integration corrected the fuzz input-seal coverage rule: the
immutable pre-run seal excludes only the two explicitly named post-run binary
retention files, which remain authenticated by the root seal and the original
artifact inventory. Its original input hashes were not rewritten.

The first altered-copy checks exposed a portable synthetic-oracle failure:
loading the native oracle beside its static inventory assumed six repository
parent levels. Both candidate verifiers now replay the authenticated original
oracle bytes from a private temporary directory without that unrelated native
inventory. Synthetic archive/member bindings and exact output replay remain
mandatory. The original and R1 negative-check failures are retained; the first
also used an incorrect expected diagnostic for summary mutation. R2 checks the
actual summary diagnostic and both fuzz input-integrity cases. No production
source, protocol, oracle, or captured report changed in these corrections.

R2 rejects all three intended mutations with their specific inner-check
diagnostics. The complete unmodified bundle passes precleanup and separate
portable-copy verification. Cleanup removed 284 inventoried files totaling
375,674,717 bytes from `/tmp/litchi-goal-0457`; shared target caches remain.
