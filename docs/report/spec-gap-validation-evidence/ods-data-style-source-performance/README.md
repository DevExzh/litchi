# ODS source-qualified data-style performance

This directory records a reproducible characterization of the final
source-qualified ODS data-style candidate. It does not establish a speedup:
there is no equivalent pre-change baseline with the same source, fixtures,
operations, allocator instrumentation, and build. The candidate is pinned to
commit `8e3ad310426a534c0bb17a789eb2c42a96e61310`, with base
`2e023cbd253a9b46ec2582e5319cbcbcddac5433`.

The earlier
[`ods-data-style-performance`](../ods-data-style-performance/README.md)
capture is historical pre-correction smoke evidence. It used different
synthetic inputs and cumulative allocation totals, so it is retained for
history and excluded from this characterization and from any comparison.

The source was copied into a private isolated checkout and verified against
the 26-file freeze before building. The source checkout was reset to the
candidate commit before capture; the harness and build target were separate
from the repository's shared target. The before/after hashes in
[`provenance-before-after.json`](results/final-v9-8e3ad3104/provenance-before-after.json)
show that the source and harness did not change during either measured pass.

## Fixture and operation matrix

The harness creates deterministic ZIP packages at scales 8, 128, and 512.
ZIP member timestamps are `2020-01-02 03:04:05`, permissions are fixed, XML
uses CRLF, `mimetype` and the manifest are stored, XML and the unrelated
`Extras/foreign.bin` member are deflated, and all fixture bytes are generated
from fixed strings. The fixture contains an effective cell style, automatic
number styles, common styles, metadata, settings, and a binary payload that
is checked for scalar edits. Graph lanes check selected and unselected style content; they do not assert the full member inventory or unrelated binary payload.

Each scale runs these lanes with one correctness warm-up followed by seven
measured repetitions:

* `source_query` resolves the effective cell data style by its source-qualified
  owner and family.
* `snapshot_clone` clones one retained snapshot 10,000 times. It reports zero
  allocator calls and zero observed copy bytes; the reported time is for the
  whole 10,000-clone loop.
* `metadata_noop` applies an empty metadata patch and requires exact package
  bytes.
* `scalar_patch` changes only the title, checks the opaque child and unrelated binary payload. The later inverse check also checks the comment.
* `graph_put` adds scientific and percentage nodes and checks the preserved
  opaque node.
* `graph_replace` replaces a typed decimal node and checks an unselected node.
* `graph_replace_same_value` supplies the typed value `1` for source text
  `number:decimal-places="01"`; the semantic replacement preserves the
  noncanonical lexical source and is an exact byte no-op at every scale.
* `graph_remove` removes one unreferenced node and checks an unselected node.

The scalar lane also performs a reverse patch. It compares every package
payload except the generated manifest, compares member inventory, and checks
the binary payload. The manifest is excluded because the package writer may
regenerate harmless lexical formatting there.

## Counter semantics

The harness installs a process-local counting allocator. For every measured
operation it records allocation and deallocation calls, cumulative requested
and released bytes, live bytes before and after, and `peak_live_bytes_delta`.
The peak value is allocator-observed logical live bytes after resetting the
operation baseline, rather than cumulative requested bytes. It excludes
allocator overhead and transient old-plus-new storage internal to
`System::realloc`; process RSS is measured separately. Deallocation call counts
cover explicit `dealloc` calls; reallocations contribute released bytes but
not deallocation calls. Every measured lane has zero net live-byte change after its
temporary values are dropped.

`copy_bytes_observed` counts only the explicit source `Vec` clone used to feed
the API. It is a lower-bound harness ingress count and does not claim to count
internal library copies. `requested_bytes` is cumulative allocator traffic
within an operation and must not be read as a peak. Wall time is reported in
nanoseconds, with the warm-up excluded from the receipts. Timed closures include
source parsing, staging, member extraction and their correctness assertions;
no durable commit or write is timed. The warm-up exercises code, allocator and
OS paths; no process parser-cache behavior is established. `/usr/bin/time -v`
records process RSS for each full run; one additional `perf stat` receipt is
retained for context.

## Characterization

The table gives median wall time, median allocation calls, and median allocator-observed
peak live-byte delta from pass 1. Full per-iteration JSONL receipts and both
pass summaries are in the result directory.

| operation | scale 8 | scale 128 | scale 512 |
| --- | ---: | ---: | ---: |
| `source_query` | 0.249 ms / 2,463 alloc / 145,858 B | 1.254 ms / 10,411 / 263,851 B | 4.649 ms / 35,822 / 1,006,739 B |
| `snapshot_clone` (10,000 clones) | 0.195 ms / 0 / 0 B | 0.196 ms / 0 / 0 B | 0.195 ms / 0 / 0 B |
| `metadata_noop` | 0.384 ms / 3,898 / 158,580 B | 1.391 ms / 10,882 / 272,154 B | 4.693 ms / 33,219 / 956,706 B |
| `scalar_patch` | 0.747 ms / 6,209 / 454,652 B | 2.744 ms / 19,940 / 813,284 B | 9.249 ms / 63,848 / 1,964,364 B |
| `graph_put` | 1.053 ms / 9,218 / 454,045 B | 3.471 ms / 24,997 / 760,143 B | 11.355 ms / 75,437 / 1,743,415 B |
| `graph_replace` | 0.674 ms / 6,202 / 457,846 B | 1.238 ms / 10,908 / 704,940 B | 2.987 ms / 25,950 / 1,500,534 B |
| `graph_replace_same_value` | 0.348 ms / 3,790 / 155,054 B | 0.789 ms / 8,017 / 233,702 B | 2.198 ms / 21,524 / 885,462 B |
| `graph_remove` | 0.795 ms / 7,709 / 445,094 B | 1.417 ms / 13,835 / 584,228 B | 3.319 ms / 33,422 / 1,035,776 B |

At scale 512, `graph_put` has the highest median wall time and cumulative
allocation calls among graph operations, while `scalar_patch` has the highest
allocator-observed peak live-byte delta. The typed same-value lane remains source-exact and
has the lowest peak among the graph edit lanes. These observations locate
where later optimization work should start; they do not compare the candidate
with another implementation. Graph put uses the opaque fixture (8,610 bytes at
scale 512), while replace/remove use the typed fixture (about 3,899 bytes), so
this ranking is specific to these workloads. Graph lanes do not verify an
inverse; their preservation assertions have the narrower scope listed above.

The two seven-repeat passes completed successfully. Process maximum RSS was
7,564 KiB and 7,216 KiB, respectively. Deterministic fixture hashes match
between passes. The raw receipts are:

* [`raw-pass1.jsonl`](results/final-v9-8e3ad3104/raw-pass1.jsonl)
* [`raw-pass2.jsonl`](results/final-v9-8e3ad3104/raw-pass2.jsonl)
* [`summary.json`](results/final-v9-8e3ad3104/summary.json)
* [`summary-pass2.json`](results/final-v9-8e3ad3104/summary-pass2.json)
* [`time-pass1.stderr`](results/final-v9-8e3ad3104/time-pass1.stderr)
* [`time-pass2.stderr`](results/final-v9-8e3ad3104/time-pass2.stderr)
* [`perf-repeats1.stderr`](results/final-v9-8e3ad3104/perf-repeats1.stderr)

## Replay

[`replay.sh`](replay.sh) checks that the source files still match the checked-
in candidate commit, uses `--locked --offline --release`, allocates a fresh
Cargo target directory, and allocates a fresh output directory for every
invocation. It does not use the historical source checkout or a shared build
target. Run it from the candidate checkout with:

```sh
ODS_PROFILE_REPEATS=7 ./docs/report/spec-gap-validation-evidence/ods-data-style-source-performance/replay.sh
```

The replay smoke was executed with one repetition and its receipt is retained
in [`replay-test.log`](results/final-v9-8e3ad3104/replay-test.log). The
checked-in harness lockfile, source freeze, build receipt, and before/after
hashes are retained in the result directory.

Independent review approved this candidate-only characterization with these
measurement boundaries. Root replay receipts in [results/root-replay](results/root-replay/verification.json)
confirm 168 measurements, matching fixture hashes, unchanged source/harness
hashes and exact no-op assertions. Root independently recomputed both captured
summaries. The original capture retains one binary hash, not a before/after
binary pair; root replay records a separate rebuilt binary hash.
