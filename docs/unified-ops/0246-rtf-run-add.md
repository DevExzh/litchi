# 0246 — Native-backed RTF run insertion

## Contract

RTF paragraph parents now accept `NewNode::new("r")` / `"run"` with required,
nonempty `text` and the ten existing local formatting properties. Default
placement appends; start, end, zero-based position and same-paragraph relative
placements return the canonical path of the complete inserted run.

Native `RunTextInsertion`, `Edit::insert_run` and `Edit::insert_body_runs` own
staging, source splicing, candidate validation and exact source-bound inverse
history. Batch positions refer to the original source and must be nondecreasing;
equal boundaries retain input order. The full right-neighbor format is inherited,
or the last left neighbor at append. Empty paragraphs inherit their retained
native delimiter formatting. Explicit properties override only their facets.
Scoped Unicode encoding keeps each inserted string in one semantic run, including
mixed ASCII, Unicode and supplementary characters.

Body opaque bytes and order remain exact. Anchor projection uses semantic UTF-8
positions and raw source offsets to resolve boundary ties, with one prefix table
and binary searches. A facade regression caught and now covers the zero-prefix
case before the first insertion. Existing source-closure refusals remain typed.
Native mixed-operation families and durable replay remain unsupported. LF is
refused in run insertion. Direct insertion preserves literal CR, VT, NUL and tabs;
existing command coercion normalizes CR/VT to LF and therefore refuses those
command inputs. This record does not claim identical control-character admission
through direct and command APIs.

Capabilities and property add flags now reflect this admitted surface. Admission
charges the native insertion payload and checks operation, staging and snapshot
limits before publication; failed transactions leave the source unchanged.

## Verification

At measured implementation `d2bd303ccd`, native tests/doctests pass **1,525**,
full-facade tests/doctests **1,746**, and RTF-only facade tests/doctests **611**.
Warning-denied full/minimal all-target Clippy, rustdoc, no-feature checking,
native/facade glossary JSON gates, and all three parser smoke targets pass.
The tooling suite passes **1,166** tests; crate boundaries retain **66** packages,
**268** dependencies and **11** tracked debts. The generated OfficeCLI capability
delta check passes. Repository lockfile hashes remain unchanged.

The focused regressions cover placement and returned paths, inherited and
compound formatting, Unicode, opaque syntax/inverse history, atomic/best-effort
and sequential editing, typed refusals, resource bounds, protection/origin and
capability flags. Final lint-driven native storage/indexing changes are covered
by a repeated native suite and focused facade insertion suite.

The frozen performance harness passes all **nine** RTF tests. Scoped formatting
and diff checks pass. Paired release measurements are recorded below.

## Paired measurements

The new standalone `unified-rtf-run-add-baseline` harness compares native and
unified start, middle and end insertion. Thirty-six cases cover 0–10,000 groups
of three runs, placed in separate paragraphs or one dense paragraph. Owners,
selectors, node properties and native formatting are prepared before timing;
staging and publication are timed. Exact native/unified bytes, created path,
reopened text/format/counts and source immutability are checked outside timing.
Two warm-up rounds and nine ABBA rounds provide 18 samples per route.

Three CPU-0-pinned trials at `d2bd303ccd` pass all **108** case validations.
Receipts are [trial 1](results/0246-rtf-run-add/trial-1.json),
[trial 2](results/0246-rtf-run-add/trial-2.json), and
[trial 3](results/0246-rtf-run-add/trial-3.json); each contains raw timings,
compiler/platform metadata, source tree and binary SHA-256.

| Shape, 10,000 groups | Native median range | Unified/native range |
|---|---:|---:|
| Separate paragraphs | 45.94–47.17 ms | 0.9917–1.0074 |
| One dense paragraph | 29.07–29.93 ms | 1.0015–1.0149 |

Ranges cover all placements and three trials. All cases with at least ten groups
are within the 10% overhead budget. **27 of 108** individual case results exceed
it, all on empty or one-group fixtures; the maximum ratio is **1.1449**.
These misses remain open: profile facade admission/path/property costs before
claiming small-document acceptance. Native publication also grows faster than
linearly over the largest input-size step; this matrix does not qualify native
scaling or the repeated-operation budget. No allocation or complete acceptance
claim is made.

Reproduce with:

```sh
cargo build --manifest-path tools/perf-baseline/unified-rtf-run-add/Cargo.toml --release
taskset -c 0 tools/perf-baseline/unified-rtf-run-add/target/release/unified-rtf-run-add-baseline
```

The recorded runs used external copied locks/targets with cached `cc 1.6.0`,
offline resolution and incremental compilation disabled, as receipts specify.
All three implementation worktrees/branches and owned temporary build artifacts
were removed after integration and verification.

## Remaining work

Run move/swap, font/color authoring, broader mixed-operation composition,
repeated-command coalescing, durable replay for new run operations, allocation
and scaling qualification, wider RTF elements, other format coverage, complete
glossary migration and Python bindings remain required by the full goal.
