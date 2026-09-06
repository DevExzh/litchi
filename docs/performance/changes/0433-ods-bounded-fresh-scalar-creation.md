# 0433: bounded ODS fresh scalar creation

Change 0433 records a bounded ODS scalar-row creation path beside the existing
buffered builder. It measures fresh one-sheet creation only. Logical append to
an existing document, adding a package Part, arbitrary editing/repackaging,
native-application compatibility, and broad memory or scaling claims remain
separate requirements. The derived report keeps `claims: []`.

The complete result bundle is [`change-0433`](../results/change-0433/README.md).
Its frozen [protocol](../results/change-0433/protocol.json),
[summary](../results/change-0433/summary.json), build/capture receipts, and
[profile index](../results/change-0433/profile-index.json) bind the source and
artifact identity; the retained [checks](../results/change-0433/checks/)
directory records validation and custody. The [ADR record](../results/change-0433/adr-compliance.md)
records the common XML/package and ODS ownership boundaries. Sealed copied-bundle
replay and mutation checks pass both [before cleanup](../results/change-0433/checks/precleanup-portable.json)
and [after cleanup](../results/change-0433/checks/aftercleanup-portable.json).
The five task temporary directories were removed; repository build targets
and the user goal file were preserved.

## Protocol and scope

The matrix contains 36 formal reports and 1,080 retained samples: three roles
(`before-buffered`, `after-buffered`, and `after-streaming`), normal and
allocator modes, 64/8,192/32,768 rows, four scalar cells per row, two reversed
repeats, three warmups, and thirty retained samples per report on CPU 2. The
streaming role uses a fixed 4,096-byte row-authoring window. The timed operation
creates deterministic rows, authors and validates the package, finalizes its
compression, and writes to a hashing discard sink. Artifact construction,
reopen/oracle work, procfs probes, sink construction, and digest extraction
remain outside the timer.

The before and after buffered roles use the existing builder at their respective
captured revisions. The after-streaming role uses the new sequential writer.
The cross-role oracle compares semantic rows, cells, scalar values, and
`Sheet1`; it does not require lexical ZIP/XML equality. Each role retains its
own deterministic archive identity.

## Retained observations

Normal after-streaming p50 latency is 59.017–62.137% below the before-buffered
role across the three shapes and two repeats. The after-buffered control is
close to its before revision except for the 64-row R1 normal p99 comparison,
which is a retained +7.037% regression flag. These are descriptive matched
rows from this harness and do not establish a production speedup or a stable
latency claim.

Allocator-mode operation observations show the following p50 values; the
regional value is the serialized operation-region peak above entry:

| Rows | Requested bytes, before → streaming | Requested calls, before → streaming | Regional peak−entry, before → streaming |
| ---: | ---: | ---: | ---: |
| 64 | 1,366,338 → 882,097 | 3,583 → 1,059 | 461,977 → 419,347 |
| 8,192 | 66,578,958 → 6,230,321 | 426,260 → 122,979 | 17,758,827 → 419,347 |
| 32,768 | 263,900,025 → 22,401,329 | 1,704,218 → 491,619 | 71,050,076 → 419,347 |

The large regional observation is approximately 71.05 MB → 0.419 MB in the
instrumented requested-allocation region. These counters do not measure
physical copy volume, allocator-internal overhead, total process RSS, or all
retained objects. The fixed row window is not a total-memory bound.

## Validation and remaining limits

Focused generated-XML and ODS receipts record 19 and 11 passing tests; the
retained common and ODS suites record 437 and 444 passing tests. Documentation
generation passes. The unscoped production strict-lint receipt remains failed
on the pre-existing `ArchiveReaderKind` `large_enum_variant` finding; the
bundle preserves that receipt and does not relabel it as a 0433 regression.
Profile artifacts and source/build custody are retained under the result
bundle. Exact profile interpretation remains separate from the timing and
allocator rows.

This batch establishes measured fresh ODS scalar creation with explicit
semantic, output, cancellation, and resource boundaries. It does not close
logical append, package-Part addition, arbitrary repackaging, native or
cold/remote I/O, total-RSS attribution, or broader CRUD coverage. No broad
optimization claim is registered.
