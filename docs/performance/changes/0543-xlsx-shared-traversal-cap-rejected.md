# 0543: XLSX shared traversal rejected at the valid event boundary

The shared worksheet traversal candidate is rejected and production restored.
The retained change is a public regression test for post-EOF typed raw errors.
This campaign makes no retained runtime speedup claim.

## Mechanism and admission

The candidate shares a borrowed namespace-aware XML traversal between the
independent validator and ordinary raw parser. Successful validated EOF forwards
the completed store or typed raw materialization error, preserving the historical
x14ac retry. Early failure discards provisional parser and validator state before
complete authoritative validation/raw parsing. Existing source, cancellation,
resource, namespace, publication and MCE/x14ac boundaries remain in force.

The original frozen pilot used medium and dense-sparse workflow fixtures plus
96×96 and 128×128 valid/late-error standalone grids. Every workflow p50/mean,
planning p50, planning allocated-byte/peak and refusal-envelope gate passed.
Conditional planning profiles, whole-child hardware counters and eager controls
were captured only after that pilot passed. All 475 review flags remain individually
recorded: 470 main/eager/cap rows plus five whole-child hardware diagnostics.

| Main shape | Repeat | Planning p50 reduction | Workflow p50 reduction |
| --- | ---: | ---: | ---: |
| dense-sparse | 1 | 28.50% | 5.89% |
| medium | 1 | 28.60% | 9.85% |
| dense-sparse | 2 | 31.60% | 8.56% |
| medium | 2 | 30.64% | 11.91% |

Planning Ir decreases 24.917–25.353%. Ordinary planning allocated bytes decrease
0.078–0.104%; incremental peak increases 0.0025–0.0062%, within the frozen 1%
envelope. No ordinary-allocation adverse increase exceeds 5%. Hardware data are
whole-child diagnostics with full counter coverage, not isolated planning cost.
Eager p50/mean show no adverse rows; one medium open p99 rises 11.21%, and two
same-build p99 drift rows remain individual review triggers.

The late-raw p50 improves 17.05–20.86% against the same invalid baseline. The
late-validator p50 instead increases 174.54–184.69%; its allocated bytes rise
616–961%. All remain inside the separately frozen baseline-valid latency/peak
envelopes. Passing an envelope does not erase those invalid-input regressions.
Allocator Region peak is neither process RSS nor a general memory/OOM guarantee.

## Supplemental valid-input boundary

Source review found an omitted performance boundary: valid input over 131,072
provisional events discards partial shared parsing and then repeats the original
validation/raw passes. A separate plan, source freeze, release build pair and
ABBA capture matrix test this before retention. It tightens admission to candidate
planning p50 and mean at most 1.05 times matched baseline for every size/repeat;
the original gates are unchanged.

All fixtures are stored ZIP XLSX, UTF-8, marker-free, below 8 MiB XML and below
the ordinary million-event limit. Exact event counts include EOF. Each of the
12 captures has ten warmups and 100 native planning observations. Setup, source
opening, selector construction, correctness checks and no-op commits are outside
the clock. Fixture and binary SHA-256 are bound by parent receipts.

| Grid | Events | Repeat 1 p50 change | Repeat 2 p50 change |
| --- | ---: | ---: | ---: |
| 160×160 | 128,326 | −33.01% | −32.56% |
| 164×164 | 134,814 | +58.97% | +61.39% |
| 256×256 | 328,198 | +25.84% | +23.86% |

The cap lane rejects the candidate. Candidate final warning-denied Clippy also
fails `large_enum_variant` on `SourceParseAttempt::Complete(Result<Store>)`.
The measured candidate is preserved exactly; no lint fix or admission precheck
is silently substituted into its performance evidence.

## Validation, preservation and limits

The final public test asserts the exact post-EOF boolean error twice, unchanged
source version and bytes, recovery on an unaffected second sheet, and byte exact
no-op publication. Its initial misuse of the single-sheet snapshot API is
preserved in the failed baseline attempt and corrected before successful capture.
The cap harness also has a retained pre-capture compile correction. All failures,
source snapshots, raw timing/allocation vectors, analyzer corrections and reviews
remain in the [sealed bundle](../results/change-0543/README.md).

Final test counts and all eight required checks are recorded in
[quality-summary.json](../results/change-0543/quality-summary.json). The candidate
passes 1,302 tests; restored baseline plus the retained public test passes 1,299.
No new fuzz, native Office, concurrency-scaling, cold-cache, remote-source or
cross-platform performance claim is made. No production API, dependency, unsafe
policy or preservation contract changes. Owned temporary targets and executable
copies are removed after hash-bound verification.

## Next work

A conservative lexical event upper bound could decline above-cap input before
building provisional parser state. It is an unmeasured follow-up proposal whose
proof and fresh matched measurements are required. Keep authoritative validation
and runtime caps; do not add hot-path allocation solely to appease the enum lint.
After admission succeeds, fresh commit-path profiling is the next overall workflow
priority identified by native phase shares. OLE2 and OOXML remain first; ODF work
is deferred until that optimization goal completes.
