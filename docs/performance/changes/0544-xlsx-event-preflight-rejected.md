# 0544: XLSX event preflight rejected after matched measurement

The candidate is rejected and production restored. The retained change is a
reproducible cap-boundary benchmark example. No runtime speedup is retained.
OLE2/OOXML optimization remains active; ODF work is deferred.

## Mechanism and correctness

The candidate computes a conservative lexical event bound before constructing
provisional reader/parser state. It counts markup/reference starts and possible
text starts with two delimiter scans. Inputs above the bound use complete
authoritative validation/raw parsing directly, avoiding 0543's discarded parser
prefix. Eligible inputs retain shared traversal, validation-first failure
handling, post-EOF materialization results and historical x14ac retry.

A direct pinned NsReader oracle checks every event kind, BOM, text, declarations,
embedded delimiters, named/numeric/malformed references and malformed markup.
Exact 131,072 and 131,073 event streams pin the boundary; a deliberately inflated
bound checks safe fallback and public no-op publication. The candidate passes
1,305 tests and warning-denied Clippy before release capture. A narrow expectation
keeps the private owned result on the stack rather than adding a heap allocation
only to silence the prior large-enum lint.

## Matched results

All primary and guard captures use fresh hash-bound baseline/candidate binaries,
CPU 2 and A1/B1/B2/A2 order. The original primary workflow and planning gates are
unchanged; the valid cap lane is required before any conditional diagnostics.

| Main shape | Repeat | Planning p50 reduction | Workflow p50 reduction | Workflow mean reduction |
| --- | ---: | ---: | ---: | ---: |
| dense-sparse | 1 | 18.226% | 5.194% | 4.851% |
| medium | 1 | 20.147% | 6.568% | 6.720% |
| dense-sparse | 2 | 19.641% | 5.822% | 5.767% |
| medium | 2 | 20.444% | 2.857% | 2.917% |

The last workflow p50/mean pair fails the required 3% improvement. No rerun or
relaxed threshold replaces that result. Planning allocated bytes fall
0.078–0.104%; incremental Region peak rises 0.0025–0.0062%, within the 1% gate.
Calls/reallocations remain diagnostic. All refusal envelopes pass, but the same
late-validator invalid input costs 182.7–188.2% more native p50 time and
616–961% more allocated bytes. Region peak is not RSS or an OOM guarantee.

| Cap grid | Events including EOF | Repeat | p50 change | Mean change |
| --- | ---: | ---: | ---: | ---: |
| 160×160 | 128,326 | 1 | −25.740% | −26.873% |
| 160×160 | 128,326 | 2 | −25.584% | −26.547% |
| 164×164 | 134,814 | 1 | +5.449% | +5.358% |
| 164×164 | 134,814 | 2 | +4.521% | +2.619% |
| 256×256 | 328,198 | 1 | +3.281% | +3.041% |
| 256×256 | 328,198 | 2 | +6.940% | +6.887% |

The 164 repeat1 and 256 repeat2 rows fail the 5% valid planning envelope. The
preflight removes the discarded shared prefix but still adds scan work before
the original passes. These fresh comparisons establish the remaining cost;
older 0543 results are motivation, not a matched estimate of this revision's gain.
No cap RSS adverse or same-build drift flag exceeds 5% in this matrix.

All 450 adverse/drift flags are individually preserved and interpreted in the
[bundle](../results/change-0544/README.md). Native phase/tail flags are kept
separate from primary gates; allocator timing is not native latency evidence.
Conditional profiles, hardware counters and eager controls were not executed
after the failed pilot. No cold-cache, remote, scaling, native Office or fuzz
performance claim is made.

## Retained benchmark and custody

`cargo run --release --locked -p litchi-xlsx --all-features --example
perf_cap_boundary -- --size 164 --warmup 10 --samples 100 --json REPORT
--fixture-out FIXTURE` reproduces the cap scenario. Use a new explicit output
location; the campaign driver also pins CPU and binds all output hashes.
The benchmark times only public `edit_sheets`; setup, inspection, no-op commit,
source checks and byte-exact no-op publication are outside the clock. Deterministic
stored ZIP fixtures expose 160,164,256 square grids with source/event metadata.

The initial benchmark Clippy failure is preserved before correction; the
interrupted baseline build retains partial logs and no fabricated terminal status.
The resumed build uses identical frozen inputs and a fresh successful receipt.
Nine final checks on the restored source, cleanup and seal are recorded in the
bundle. The candidate-only tests and production patch remain evidence artifacts;
the example is the only retained Rust change.

## Next work

A coarser single-scan bound could reduce preflight overhead, but static review
finds it also declines the current 128 grid (bound131,593 exceeds131,072),
losing a primary shared path. It is not ready for the next freeze. Its design
review and proof requirements remain, without implementation or performance admission.
A fresh campaign must show useful whole-workflow gains and acceptable valid
fallback cost. After a traversal passes, obtain fresh commit-path attribution
before pursuing smaller planning micro-optimizations.
