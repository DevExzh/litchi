# 0356c — PPTX manifest fidelity with compact patch state

- Date: 2026-10-21.
- Status: coordinator review follow-up to [0356b](0356b-pptx-retire-source-backed-editor.md).
- Scope: the owned PPTX part-edit engine, private package graph ownership,
  and OPC retained-manifest comparison. Public family signatures are unchanged.
- `performance_claim`: all 16 routes are within 1.05 times baseline; retained
  patch-history peak heap is below baseline.

## Problem and implementation

A raw OPC edit that changed only the lexical representation of content types
or relationships could commit as an empty patch. The semantic graph was equal,
but producer XML bytes had changed. Structural inverse replay could also
regenerate the content-types manifest instead of restoring its original bytes.

Part patches now retain exact content-type and relationship images as well as
semantic part state. Capture recognizes lexical-only edits and exact no-ops;
replay checks the recorded source images and restores the changed images.
Payload-only replay updates the payload in place, preserving relationship
spelling and explicit empty-member presence. All capture and restoration
reads use the package's retained source limits.

An eager validated-token implementation passed the fidelity tests but regressed
commit costs to 1.87–4.04 times baseline. A lazy graph-sharing implementation
fixed those timings, but retaining 128 payload patches while dropping their
published packages raised peak heap from 2.98 MB to 14.75 MB. Both were rejected.

The final design shares the immutable OPC graph inside the private PPTX owner,
starts staging without cloning it, and clones only on a graph update. Patches
retain compact manifest images, never the package graph. Retained producer XML
uses shared byte allocations. Ordinary structural patches represent canonical
content types with a marker and regenerate bytes only if replay needs them.
Explicit lexical replacements bind a canonical base with SHA-256, keeping the
exact guard in both directions without charging ordinary structural commits.
The OPC layer exposes read-only retained-byte projections and bounded canonical
serializers separately from validated source tokens. Replacement still checks
current bytes, required limits, XML validity and publication policy.

An explicit raw content-type replacement carries the same exact-image guard in
both directions. Pure structural changes whose original source token merely
became incompatible still compose with unrelated part additions, including an
inverse that restores producer markup. Dropping an edit discards it; a failed
update leaves the previous staged graph intact. The public const read-only OPC
projection remains const.

OPC relationship provenance comparison checks its retained semantic binding
without allocating an intermediate sorted binding (commit `db1c77b012`).
Content-type comparison recognizes a shared manifest with unchanged coverage
without rebuilding two sorted part-name indexes. It checks part names, content
types, relationship presence and retained relationship-member presence; the
general path remains for other graphs. A focused test compares the fast path
with the full retained-manifest computation in both directions.

## Correctness and API validation

`tests/opc_manifest_fidelity.rs` checks exact commit, independent-copy replay
and reopened inverse bytes for lexical content types, root relationships and
part relationships; structural inverse producer bytes; structural composition
with an unrelated added part; stale lexical conflicts; lexical change-and-restore
no-ops; payload changes retaining relationship spelling; mixed lexical and
structural edits in both directions; and replay onto independently reopened
content after the retained source manifest loses coverage. The first three
tests fail on the original implementation because diagnostics report no change.

Validation on this change:

- PPTX all-feature library: 1,280 passed, one ignored.
- Manifest fidelity: all 11 tests passed on the final code. Signature-policy,
  comments and hyperlinks integration targets: 14 passed.
- OPC retained-coverage equivalence and bounded compact-view tests: passed.
- OPC and PPTX all-target/all-feature clippy with warnings denied: passed.
- PPTX full and minimal glossary profiles: passed.

Logs remain outside the checkout under
`(deleted-coord-artifact)`.

## Performance evidence

Baseline is `86d708878f`; candidate is this record's implementation on
`f93b9e1c32`. Both use the identical committed
`crates/litchi-pptx/examples/opc_manifest_bench.rs` source and tracked
`test-data/poi/test-data/slideshow/45545_Comment.pptx` fixture. The baseline was
exported with `git archive`, with the same Cargo.lock copied into its scratch
tree. The builds use **distinct target directories** to avoid Cargo reusing an
artifact from another tree with different source contents.

```sh
CARGO_INCREMENTAL=0 CARGO_BUILD_JOBS=4 cargo build --locked --release \
  -p litchi-pptx --all-features --example opc_manifest_bench
taskset -c 7 <binary>
```

Three A-B-B-A trials, three warmups and 31 samples per scenario. Values below
are medians of six per-run medians. Workers were explicitly held; continuous
process monitoring rejected any run containing cargo, rustc, rustdoc or clippy.
Earlier contaminated trials and the rejected eager implementation are not part
of the accepted evidence.

| Route | Before µs | After µs | After/before |
|---|---:|---:|---:|
| no-op update and commit | 80.856 | 39.531 | 0.489 |
| payload commit | 58.140 | 43.200 | 0.743 |
| payload commit and save | 12,219.624 | 12,133.414 | 0.993 |
| payload commit and two saves | 24,320.569 | 24,188.317 | 0.995 |
| part-add commit | 54.056 | 52.221 | 0.966 |
| part-add commit and save | 12,239.964 | 12,142.884 | 0.992 |
| root-relationship commit | 57.666 | 49.060 | 0.851 |
| payload replay (warm) | 82.310 | 73.930 | 0.898 |
| payload inverse (warm) | 84.266 | 75.900 | 0.901 |
| payload commit and first replay | 143.146 | 122.526 | 0.856 |
| payload commit and first inverse | 148.206 | 127.051 | 0.857 |
| 128 retained payload patches | 7,576.723 | 6,281.051 | 0.829 |
| part-add replay (warm) | 73.370 | 72.631 | 0.990 |
| part-add inverse (warm) | 74.245 | 72.760 | 0.980 |
| 128 retained structural patches | 7,638.673 | 7,251.091 | 0.949 |
| 1,000 updates and commit | 28,169.380 | 28,442.277 | 1.010 |

Heaptrack on the same complete 16-route probe reports 2.98 MB peak heap
before and 2.93 MB after (its displayed precision), with allocation calls
falling from 58,831,913 to 49,341,461. Both 128-patch scenarios retain patches
while dropping their published packages, so this includes history lifetime
cost. Patches retain changed payloads and compact manifest images, with no
reference to either whole package graph. These instrumented runs are separate
from the timing trials; their times are not performance evidence.

`results/0356c/` contains twelve small timing receipts, the summary, process
monitor receipt, probe hash, binary hashes and heap summaries.
`heap-artifacts.json` records external capture filenames, sizes, SHA-256 values
and regeneration commands. No binaries or raw profiler captures are committed.
`p0356c` adds all 16 routes to the generated benchmark matrix; its unit tests and
`--check` pass. Workspace integration with W4, W4b and W8 remains a separate gate.
