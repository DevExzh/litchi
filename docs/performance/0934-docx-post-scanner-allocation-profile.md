# 0934 — DOCX allocation profile after borrowed scanners

## Purpose and scope

[0933](0933-docx-borrowed-inventory-scanners.md) adopted borrowed scanner state
with 21.46–36.02% fewer construction-requested bytes in its four measured cells.
This batch profiles that adopted implementation to rank the remaining work.
It changes no production or harness code and builds no new executable.

The retained counting-allocator executable is the 0933 candidate, built from
review HEAD `3ae856d76424f2f9b58d26dcd20fd074b8e97541` plus the recorded scanner
patch. Its runtime and harness inputs match adopted commit
`5f05944e9994a4cc1b387823c503fbab8650aca3`. The
[independent source binding](results/change-0934/independent-source-binding-v1.json)
records the full source-list verification and explicitly distinguishes the
isolated canonical lock from the checkout's ignored root lock. The new commit
identity does not rewrite the historical build receipt.

## Frozen protocol

The [prospective plan](results/change-0934/heaptrack-plan-v2.json) fixes exactly
two heaptrack 1.5.0 processes: 8,192 sections, eight live commits, exact retention,
shared snapshots followed by independent sources. Both use CPU 12 and the
0932 explicit child environment with heaptrack-managed injection. Each has a
600-second deadline. There are no retries or replacement children.

The profile includes initialization, fixture creation, prevalidation, cap
admission, construction, correctness/publication/inverse checks and release.
It is not the `after_snapshot_drop` construction interval measured in 0933.
Cumulative requested bytes are neither live memory nor bytes copied. Profiler
runtime is not a normal-release timing comparison.

Raw format decoding and allocation-size histograms must conserve both event
counts and requested bytes. Nearest-Litchi classification uses the first resolved
physical `litchi_` frame from leaf to root, stripping the final Rust symbol hash;
inline locations are not separate frames. Compiler inlining can move costs
between symbol buckets, so a bucket difference is not proof of allocation
ownership at an individual source statement.

## Results

Both planned children exited successfully without a timeout or retry. Their
creation-interval counters reproduce the adopted 0933 candidate exactly:
800,608,673 requested bytes / 11,875,584 calls for shared snapshots and
1,425,467,600 bytes / 21,109,592 calls for independent sources. Both return to
the measured allocator baseline after release.

Whole-process allocation traffic is:

| Ownership | Allocation events | Requested bytes | Bytes below 0932 |
| --- | ---: | ---: | ---: |
| Shared snapshots | 87,159,235 | 5,873,313,617 | 36.0121% |
| Independent sources | 124,151,078 | 8,369,567,879 | 35.9995% |

Event counts are respectively 30.2753% and 30.2727% below 0932. These are
single diagnostic profiles of each build under the same workload, not repeated
normal-release measurements. They support allocation attribution, not a latency
or construction-phase claim.

The largest mutually exclusive nearest-physical-Litchi buckets are:

| Symbol within `litchi_docx::section` | Shared requested bytes | Independent requested bytes |
| --- | ---: | ---: |
| `footnote_columns::codec::read_with_context` | 1,091,634,180 | 1,554,744,898 |
| `codec::attributes` | 1,023,042,831 | 1,457,022,545 |
| `model::Section::from_xml_bytes` | 747,758,580 | 1,064,986,736 |
| `codec::decode` | 669,962,016 | 954,159,454 |
| `inventory::append_section` | 413,527,488 | 588,963,392 |
| `codec::root_namespace_bindings` | 342,786,048 | 488,210,432 |
| `codec::validate_attribute_qnames` | 276,840,192 | 394,280,704 |
| `inventory::Inventory::parse_with_limits` | 236,342,238 | 336,608,642 |

The first seven listed buckets retain their 0932 byte totals. Removing scanner
copies therefore makes fragment codec and extension validation the next area to
investigate. Those buckets contain necessary work as well as potential repeated
work; their sum is not an estimate of removable allocations.

## Remaining source candidate

Inventory fragment construction still calls `Section::from_xml_bytes`, which
validates raw codec structure and footnote-column extensions. Inventory then
requests local semantic state, triggering a second raw codec parse. A private
inventory path may be able to reuse the first raw parse while preserving that
validation order. Public `Section` construction must retain its existing lazy
semantic errors; namespace context, limits, lexical preservation and tolerance
offsets remain constraints. No such change is implemented in this batch.

The [source review](results/change-0934/candidate-review-v1.json) proposes a
private helper returning semantic state and tolerance data from one raw codec
parse, with the independent footnote-extension validation retained before
semantic decoding. Holding that raw parse across footnote validation could
increase temporary peak memory and must be measured.

The [validation review](results/change-0934/validation-review.json) specifies
eight differential cases for lazy errors, validation precedence, namespace
injection/rebinding, limits, original-coordinate tolerance offsets, footnote
context and absent/zero semantics, and publication/inverse parity. Inherited
`mc:Ignorable` is not injected by the existing fragment materializer; a candidate
must preserve the current refusal or explicitly carry the appropriate context.
No tests or builds are claimed by these read-only design reviews.

## Evidence and custody

The [comparison evidence receipt](results/change-0934/comparison-evidence-v1.json)
binds the capture reports, original terminal records, offline tools, exact
conservation checks and nearest-frame classifications. Raw decoding, histogram
integration, trace aggregation and heaptrack footer event counts agree. Each
profile has one unknown allocation of 73,728 bytes, retained in the totals.
The [independent audit](results/change-0934/independent-profile-audit-v1.json)
reconciles a separate Perl raw-schema verifier, exact allocation-size maps,
nearest-Litchi categories, physical-stack totals and folded event counts. It
also verifies the capture terminal status, report correctness flags, retention
boundaries and construction-counter identity against 0933.

The initial plan and source-equivalence preflight required metadata corrections
before capture; their original records remain available. The first two offline
`heaptrack_print` commands succeeded, but their terminal writer appended a literal
`\n`, producing malformed JSON. Those outputs and original terminal bytes were
preserved, the writer was corrected, and only the two offline export commands
were repeated. No workload capture was repeated. Heaptrack produced actual raw
files named `heaptrack-data.gz.zst`; terminal records preserve the originally
requested `.gz` names and the receipt binds the actual files separately.

Large evidence stays outside the checkout. The canonical compressed archive is
`/home/zhuhe/litchi-performance-0934/canonical-evidence-v1.tar.zst`,
1,983,540 bytes, SHA-256
`653f292cf9f729e4f04c8cb89d5a91326e23e2bc776a4f2828ca3400bb9fee4c`.
The comparison receipt records its contents, uncompressed identity and exact
regeneration commands, plus individual artifact paths, sizes and hashes.
Measured executables and source snapshots retain their original 0933 locations.
Disposable Python caches were removed; canonical evidence and failure-history
receipts remain available. This batch runs no Rust tests because it changes only
performance documentation and diagnostic evidence.
