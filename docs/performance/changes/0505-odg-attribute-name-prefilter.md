# 0505: filter ODG attributes by local name before namespace lookup

Two private helpers resolved the namespace of every checked XML attribute,
even when its local name could not match the requested field. A six-line
change in `attribute` and `attribute_source_span` skips that lookup for
nonmatching names. It retains full checked iteration, namespace matching,
selected-value decoding, duplicate detection, source-span ownership and errors.
The quick-xml 0.41 local-name and resolver paths use the same first-colon split;
resolution has no mutation or validation side effect that this filter bypasses.

## Primary comparison

[Evidence](../results/change-0505/README.md) identifies the **final/** capture
as the shipped filter-only comparison against clean `df2e73fa3`.
The unchanged, locked 0502 release probe runs four deterministic corpora,
serial A1/B1/B2/A2 blocks pinned to CPU 2, 25 warmups and 200 samples per child:
16 children and 3,200 measured samples. Rust/Cargo 1.95.0, build commands,
source hashes, binary hashes, raw samples and process RSS receipts are retained.
The timer includes owned-byte open, traversal, checksum assertion and drop;
fixture generation and hashing are outside it. The traversal checksum does not
cover every metadata field; focused tests and source review cover semantics.

| Corpus | Before p50 ms, r1 / r2 | After p50 ms, r1 / r2 | p50 change, r1 / r2 |
| --- | --- | --- | --- |
| plain-small | 2.121 / 2.141 | 1.940 / 1.936 | -8.51% / -9.58% |
| plain-large | 65.559 / 65.027 | 59.582 / 59.547 | -9.12% / -8.43% |
| metadata-small | 1.766 / 1.759 | 1.597 / 1.595 | -9.59% / -9.35% |
| metadata-large | 25.006 / 24.862 | 22.537 / 22.591 | -9.88% / -9.13% |

No paired measured latency, throughput or RSS adverse change exceeds 5%.
RSS changes range from -3.20% to +0.11%. The summary retains p95/p99,
throughput, and 10,000-resample median ratio intervals using the 0502 seed.
Intervals describe within-run uncertainty, not host-to-host variability.
These two reversed repeats do not enter the ten-entry strict claim registry.

## Attribution and rejected experiment

Fresh metadata-large whole-child Callgrind instruction references fall
1,305,166,661 to 1,143,560,704 (12.38%). Namespace resolution exclusive
references fall 149,857,714 to 36,768,716. Attribute iterator and duplicate
checker references remain exactly 231,525,188 and 84,841,334 respectively.
These are instruction references, not hardware cycles or cache measurements;
`perf_event_paranoid=4` prevents hardware counters. The child includes fixture
setup, hashing and two opens (preflight plus one sample). The remaining checked
attribute walks are a substantial hotspot, but removing validation is not an
acceptable optimization.

Heaptrack reports the same 507,836 allocation calls, 344,771 temporary
allocations and rounded 6.98M peak heap on both sides. These are whole-child
measurements with profiler overhead, not isolated operation allocation counts.
No allocation reduction is claimed for the shipped change.

An initial combined candidate also borrowed the selected QName instead of
copying it. Its root-level capture improved p50 but flagged plain-large RSS
increases of 9.97% and 10.09%. Supplementary reversed short RSS runs varied,
and heap peaks were unchanged; this does not establish borrowing as the cause.
That experiment has no accepted incremental benefit evidence and was removed.
Its patch, 3,200 samples, and profiles remain separately identified. The final
filter-only candidate was rebuilt and measured afresh. The 400 pilot samples
and 180 supplementary RSS-review samples are not pooled with either capture.

## Correctness, architecture and limits

Four new integration tests cover namespace aliases, foreign/default/unbound
attributes with the same local name, exact preservation during shape rename,
semantic alias duplicates, and malformed or duplicate unrelated trailing
attributes. Invalid-input fixtures bypass the validating package writer;
valid-control assertions ensure rejection is due to the intended malformed
attribute. Initial fixture-authoring failures and a corrected formatting check
remain in the evidence, alongside the successful final checks.

The ODG all-target suite passes 96 unique tests. A final focused rerun passes
all four new tests; it is not counted again. Scoped formatting, warning-denying
Clippy, doctests, warning-denying rustdoc, the ODF umbrella's ODG-only build,
and crate boundaries are recorded in `gates.json`.

The architecture contract is unchanged: helpers remain private to ODG;
no public API, dependency, thread, unsafe code, ownership or source-byte policy
changes. Full attribute validation and existing resource budgets remain active.
Tests include existing fixture coverage, but no live Office application,
full-workspace build, sanitizer or fuzz campaign was run for this change.
The result applies to these synthetic owned-byte open/traversal scenarios.
Full CRUD, provider/scaling intersections and the older 0502 parser comparison
remain open; this batch does not complete the broader goal. Batch-owned scratch
is removed after checks and retained evidence is hashed before commit.
