# Change 0479: existing DOCX plain-paragraph tail append baseline

`performance_claim: none; baseline and phase attribution`

`claim_authorized: false`

This batch measures the existing public source-backed DOCX paragraph-copy
transaction. It copies source paragraph zero to the final insertion slot in
documents containing 64, 8,192, or 131,072 plain paragraphs. The ordinary
transaction retains materialized snapshots and reversible patch bytes. The
baseline does not establish the explicit-window append bound required by the
full performance goal.

The [contract audit](contract-audit.md) corrects two assumptions in the prior
handoff: this edit permits exactly one operation, and it refuses section
properties. Repeated 64/256 appends require separate reopen/edit/publication
lifecycles; a document containing `sectPr` needs a different capability. Those
scenarios remain open rather than being counted as covered by this baseline.

## Frozen measurement scope

The 24-process protocol uses two reverse-order repeats, two instrumentation
builds, three source sizes, and two observation modes. Each process performs
three warmups and 30 measured operations, pinned to CPU 2. Source archive and
untimed oracle storage already exist when each operation begins. Constructing
the source adapter, opening the package, capturing the snapshot, staging the
edit, committing, sequential publication, and dropping the operation's owners
belong to the lifecycle.

The `total` mode measures one complete allocator region. The `phases` mode
uses separate, non-overlapping regions for open, snapshot, stage, commit,
publication, and destruction. These are separate executions of the same
semantic operation. Summing phase peaks cannot recover the whole-operation
peak, and phase observation overhead is not an optimization result.

Logical `ReadAt` counts and sink writes describe the supplied in-memory
source and sequential sink. They are not physical disk traffic. Full-process
RSS includes corpus construction, oracles and teardown. Separate perf and
strace runs include this wider process scope and are excluded from the formal
720 operations. No latency, constant-memory, durability or scaling speedup
claim is authorized by this baseline.

## Evidence workflow

`capture.py --freeze` records the exact case order, commands and environment
before the final builds. `build.py` binds source manifests, eight DOCX template
inputs and separately copied normal/allocator executables. `pilots.py` checks
each binary and observation mode before formal captures. `gate.py` retains
command outputs and before/after source manifests without overwriting earlier
attempts. `run-gates.py` records the applicable harness and DOCX contract gates.

The final matrix contains 24 successful reports and 720 samples. Normal total
means are 0.169641/0.169123 ms, 17.011582/16.251897 ms and
271.346097/271.230824 ms at the three sizes. Incremental allocator peaks are
509,974, 3,071,310 and 41,793,870 bytes, with zero lifecycle live-byte exit
deltas. All 38 repeat review flags remain in `summary.json`; the small normal
process RSS grows 5.277% between repeats. This is a baseline with no
candidate/control or speedup claim.

The [change record](../../changes/0479-docx-tail-append-baseline.md) contains
individual percentiles, mean intervals, allocator and I/O tables. Independent
[measurement review](measurement-review.md), [allocation source audit](allocation-source-audit.md)
and [CPU review](cpu-review.md) explain the observed owners and scope. The
[harness review](harness-review.md) and [ADR matrix](adr-matrix.md) document
preservation and contract checks. [Next work](next-work.md) keeps the full
non-iWork objective open.

The diagnostic source is commit `3e8022336`; both binary builds bind the same
7,048-file source manifest. `rust-validation.json` records 16 required final
gates and all 30 attempts (26 successes and four retained development failures).
One analyzer attempt correctly refused to overwrite the agent's existing
development summary; that artifact is retained as `development-summary.json`,
and `analyze-final` generated the final summary. The other development failures
were two harness compilation errors and module-order formatting, all corrected
before final source validation. No capture is excluded from the formal matrix.

## Reproduction and custody

From the repository root with the recorded Rust 1.98.1 environment, the
recorded command sequence is:

```sh
python3 -B docs/performance/results/change-0479/build.py
python3 -B docs/performance/results/change-0479/pilots.py
python3 -B docs/performance/results/change-0479/freeze-corpus.py
python3 -B docs/performance/results/change-0479/capture.py
python3 -B docs/performance/results/change-0479/profiles.py
python3 -B docs/performance/results/change-0479/analyze.py
```

Run heavy commands serially under the protocol's `cpu.lock`. These capture
helpers intentionally refuse existing output, so reproduce in a fresh result
location with a freshly frozen protocol rather than overwriting this bundle.
Exact original commands, timestamps, environment, outputs, binary identities,
source manifests and template copies are retained in the receipts.

Portable verification needs only this bundle and Python's standard library:

```sh
python3 -B verify.py
```

`live-state.json` checks the final source, templates, binaries and protected
user files before removal. `cleanup.json` records authenticated removal of
both copied executables and the lock, preserving shared Cargo caches.
`evidence-validation.json` binds helper bytes and the final evidence test
receipt. The full verifier recomputes the main and profile summaries, validates
profile gzip identity, and checks the complete validation ledger and seal.
`portable.json` retains a copied-bundle pass and eight independently resealed
corruption refusals after the original runtime binaries are gone.

