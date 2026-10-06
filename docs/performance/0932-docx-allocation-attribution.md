# 0932 — DOCX construction allocation attribution

This batch investigates the large cumulative allocation requests observed in
[0931](0931-docx-retention-memory.md). It changes no production code. The
retained allocator executable uses the adopted 0930 implementation plus the
0931 harness; its build receipt binds that exact source independently of the
current documentation commit.

## Counter meaning and workload decomposition

The counting allocator records the full requested new size of every successful
reallocation as allocated bytes and the old size as deallocated bytes. A
reallocation also increments both allocation-call and reallocation-call
counters. These totals are not bytes copied, allocator-internal writes or
unique live heap. In-place reallocations can contribute to the requested-byte
total without copying their full payload.

The 8,192-section source is 1,170,499 uncompressed XML bytes, packaged into a
25,498-byte generated DOCX. The 16-section counterpart is 2,316 XML bytes and
1,052 package bytes. Construction counters end at `after_snapshot_drop`,
before measured-commit publication and inverse checks. Later cumulative
counters include that correctness work and cannot be substituted for this
boundary.

The three 0931 repetitions agree exactly on construction counters. Source
inspection shows one package/query path followed by one commit in the
one-commit case; shared eight-commit construction uses one package/query path
and eight independently created commits; independent construction repeats
the complete path eight times. The observed totals decompose as follows:

| Sections | Derived component | Requested bytes | Allocation calls | Reallocation calls |
| ---: | --- | ---: | ---: | ---: |
| 16 | One package/query path | 595,620 | 3,887 | 270 |
| 16 | One changed commit | 296,881 | 4,298 | 335 |
| 8,192 | One package/query path | 139,348,382 | 1,892,561 | 122,928 |
| 8,192 | One changed commit | 139,000,751 | 1,892,972 | 122,993 |

These are topology-derived components, not directly instrumented subphases.
The package/query component includes opening and source ownership; the commit
component includes target construction and validation. Neither identifies
parser-only or copy-only traffic. The
[source/counter audit](results/change-0932/source-counter-audit-v1.json)
binds the raw counter evidence, source hashes and derivation.

## Profiling protocol

Two fresh processes run the retained frame-pointer allocator binary under
heaptrack 1.5.0: 8,192 sections, eight commits, exact retention, shared snapshots
followed by independent sources. Both use CPU 12, the explicit 0931 child
environment and heaptrack-managed injection. Each has a 600-second deadline;
neither may be retried or selectively replaced. Child terminal records and the
original outer result precede report/profile analysis. Raw captures and large
exports remain outside the checkout.

The profile covers the entire process: initialization, fixture generation,
prevalidation, cap admission, construction, publication/inverse checks and
release. It is not a direct profile of the construction counter interval.
The native stack identifies `run_capture`, while the construction routine is
inlined; source-level phase names alone cannot recover a missing native frame.
No timing or speedup claim follows from profiler execution time.

## Measured attribution

Both children exit successfully, pass their correctness checks and return to
baseline after the final commit drop. No child is replaced. The corrected raw
decoder and heaptrack's independent allocation-size histogram agree exactly:

| Ownership | Allocation events | Cumulative requested bytes |
| --- | ---: | ---: |
| Shared snapshots | 125,004,757 | 9,178,781,390 |
| Independent sources | 178,052,276 | 13,077,355,214 |

These whole-process totals include much more work than the construction
interval above. They must not be compared as if they measured the same boundary.

The [nearest-Litchi attribution](results/change-0932/heaptrack-nearest-litchi-v1.json)
assigns each trace once, to its first resolved physical `litchi_` function while scanning
from leaf toward root, with the final Rust symbol hash removed. Inline
locations are not separate frames in this classifier. The largest
buckets are:

| Nearest Litchi function | Shared requested bytes | Independent requested bytes |
| --- | ---: | ---: |
| `Inventory::parse_with_limits` | 1,925,267,601 | 2,742,047,691 |
| `ensure_source_section_inventory_xml` | 1,616,580,372 | 2,302,402,348 |
| Footnote codec `read_with_context` | 1,091,634,180 | 1,554,744,898 |
| Section codec `attributes` | 1,023,042,831 | 1,457,022,545 |
| `Section::from_xml_bytes` | 747,758,580 | 1,064,986,736 |
| Section codec `decode` | 669,962,016 | 954,159,454 |

The first two scanners together account for 38.5873% of shared and 38.5739%
of independent requested bytes. This
supports investigating their event, namespace-resolver and element-name copies.
It does not establish that every byte in those buckets comes from these copies,
or predict a construction-phase fraction, CPU saving or speedup. Each profile
retains an explicit unknown bucket of one allocation requesting 73,728 bytes;
nine other events requesting 2,620 bytes have no Litchi frame.
The separate [direct-caller summary](results/change-0932/heaptrack-attribution-v2.json)
uses a different grouping rule; its categories are not added to this table.

The file-format decoder uses implicit sequential string, instruction-pointer
and trace IDs, and resolves each `+` record through its allocation-info index.
The `s` record's first field is a hexadecimal byte length. An initial decoder
draft incorrectly used the current allocation-info record; it was superseded
before execution. Only corrected v2 produced accepted raw attribution. The
[draft receipt](results/change-0932/heaptrack-raw-decoder-v1-draft-not-used.json)
preserves that distinction. No capture was repeated to correct analysis.

## Independent verification

An independently implemented raw decoder reproduces all 1,375 physical-stack
symbol groups per child, with exact event counts and requested bytes, after
normalizing `<no-frame>` to `<unknown>` for the one unresolved allocation.
Inline locations are excluded by this definition. This verifies stack-level
assignment as well as the independent allocation-size histogram totals.

The canonical folded allocation exports also conserve the full event totals.
They expand inline locations, so exact equality of their text with the
physical-stack representation is not claimed. The independent raw
physical-stack comparison supplies the attribution check. The
[independent audit](results/change-0932/independent-audit-v1.json) records the
schema, verifier, output identities and comparison limits.

## Source candidates and semantic boundaries

Every changed commit builds a complete target XML buffer and parses a fresh
section inventory. Reusing the old inventory wholesale would be incorrect:
the edited descriptor changes, and later physical source spans can shift.

Two smaller allocation candidates preserve that validation work. Full-document
scanners clone namespace resolvers, and the inventory scanner owns events and
element-stack names even though its reader borrows immutable XML. Borrowing
these values for the current event or input lifetime may remove copies without
changing parsing order. Separately, each section fragment undergoes raw codec
validation, footnote-column validation and a second raw codec parse for
semantic decoding.

The fragment path has stricter constraints: construction currently validates
codec structure before footnote extensions, while semantic failures remain
lazy for public `Section` callers. Inventory parsing also inherits footnote
refusals even though inventory descriptors do not expose that field. Any
consolidation must preserve error ordering, namespace context, tolerance
offsets, unknown markup, limits and exact bytes. The
[candidate review](results/change-0932/candidate-review-v1.json) records these
unimplemented candidates. No bypass, new cache or production optimization is
adopted by this batch.

## Evidence and reproduction

The [capture plan](results/change-0932/heaptrack-plan-v1.json),
[capture index](results/change-0932/capture-index-v1.json), original terminal
receipts and frozen analysis scripts are committed as compact evidence. The
[artifact receipt](results/change-0932/canonical-evidence-receipt-v1.json)
records the binary, immutable source, tool and script identities, every large
member's size and SHA-256, and capture/export/decoder regeneration commands.

Canonical raw profiles and full exports are retained outside the repository in
`/home/zhuhe/litchi-performance-0932/canonical-evidence-v1.tar.zst`:
2,660,533 bytes, SHA-256
`b56d0734752f6bff1cf26bdc837459803d6fbcf3278d0c50e0da2bf30927d6e1`.
The archive passes `zstd -t`. The retained 0931 executable and immutable source
remain available at their original paths. No executable or large evidence file
is committed. Two original terminal records contain a trailing literal `\n`
after their JSON object; byte-identical originals remain, alongside normalized
derivatives that remove those two bytes and append one newline. The artifact
receipt binds both forms and the exact transformation.

[Cleanup](results/change-0932/cleanup-v1.json) removes disposable exploratory print outputs and audit scratch files
totaling 19,768,825 logical bytes. Canonical evidence
and the unused decoder draft are retained.

No production or harness code changes, build or Rust test run occur in this
batch. The [validation plan](results/change-0932/validation-plan-v1.json) defines
checks for a future implementation; it is not a claim that those tests ran.
The next bounded candidate is borrowed scanner state while preserving fresh
validation; fragment-parser consolidation requires separate semantic review.
