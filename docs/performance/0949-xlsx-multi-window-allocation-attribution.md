# 0949 — multi-window allocation costs traced to storage and parser state

Four fresh heaptrack captures reproduce the allocation regression of the
[rejected 0948 candidate](0948-xlsx-multi-window-audit-rejected.md).
Retained `Vec<Window>` growth contributes **176,640 / 397,440 requested bytes**
on the medium / dense-sparse one-percent edits. Fresh parser buffers contribute
**279 / 531 allocation requests**. These are measured candidate contributions;
subtracting the baseline and other path changes gives the net regressions below.
Production remains unchanged and 0948 remains rejected.

## Exact publication accounting

| Workload | Baseline calls / bytes | Candidate calls / bytes | Net increase |
| --- | ---: | ---: | ---: |
| Medium | 406 / 2,495,613 | 681 / 2,677,821 | +275 / +182,208 |
| Dense-sparse | 406 / 2,765,840 | 931 / 3,172,336 | +525 / +406,496 |

`allocation_calls` includes reallocations. The separate reallocation counter
is a subset: baseline/candidate 50/54 for medium and 49/50 for dense-sparse.
Adding it to allocation calls would double-count requests. Bytes mean requested
allocation bytes, including the new size of each reallocation, before frees;
they are neither peak live memory nor physical RSS.

Raw trace selection requires both the resolved
`SourceBackedEditor::publish_multi_commit_to_stream` frame and
`run_xlsx_cell_values_edit_save` at `tools/perf-baseline/src/lib.rs:44840`.
The four selected sums match the application publication counters exactly.
This is a checked stack/call-site selection, not a temporal heaptrack marker.
Selecting only the publisher frame includes other calls and is too broad.

## Disjoint candidate source contributions

| Source class | Medium calls / bytes | Dense-sparse calls / bytes |
| --- | ---: | ---: |
| Three parser buffers per replacement window | 279 / 4,464 | 531 / 8,496 |
| `MultiWindowSearch::consider` window vector growth | 16 / 176,640 | 11 / 397,440 |
| `MultiWindowSearch::open` vector growth | 8 / 1,536 | 6 / 1,152 |
| Candidate subtotal | 303 / 182,640 | 548 / 407,088 |
| Net other/common-path change after baseline subtraction | −28 / −432 | −23 / −592 |
| Net publication increase | 275 / 182,208 | 525 / 406,496 |

The verifier checks that the classes are disjoint and entirely contained in
the exact publication call-site selection. It computes the residual from the
pair delta, rather than attributing overlapping inclusive stacks twice.

There are 93 / 177 multi-window parser states, each allocating quick-xml's
`opened_starts` and `opened_buffer`, and the auditor's `State::spaces`.
The dense case also retains one existing single-window route. The baseline
uses four complete replacement audits for medium and three complete audits
plus one single-window audit for dense-sparse.

The window vector makes four initial allocations and twelve reallocations for
medium, and three initial allocations and eight reallocations for dense-sparse.
The request sizes imply 736 bytes per inline window element; this is inferred
from growth sizes, not measured with a Rust `size_of` probe. Its candidate byte
contribution is 96.94% / 97.77% of the respective **net** publication byte delta.
These percentages compare a gross candidate contribution with the net delta.

No resolved allocation in either candidate capture enters the namespace clone
helpers. This fixture-specific zero does not establish that nonempty namespace
clones are free. Inline namespace state still occupies window storage.

## Capture and validation

The frozen v3 plan runs exactly four serial children: baseline/candidate 0948
allocator binaries, each on medium and dense-sparse one-percent edits, CPU 12,
C/UTC environment, zero warmups, one sample, and a 600-second child deadline.
All four finish successfully without retries. These profiled runs support
allocation attribution only; their timing does not establish native latency.
No production or harness source was edited and no binaries were rebuilt.

Before capture, the retained 0948 strict verifier rechecked both 10,789-file
source censuses and the source/build/binary chain. Current committed runtime
inputs are unchanged from that evidence. The runner validates bindings before
launch; start and terminal receipts precede report consumption. All captured
files were hashed before offline decoding. A postflight checks the bindings
again. The retained 0948 archive supplies the source and executable dependency.

The raw whole-process totals are deliberately kept separate:

| Profile | Raw allocation requests | Requested bytes |
| --- | ---: | ---: |
| Baseline medium | 979,147 | 475,804,603 |
| Candidate medium | 979,434 | 476,010,899 |
| Baseline dense-sparse | 2,596,848 | 667,496,174 |
| Candidate dense-sparse | 2,597,386 | 667,935,103 |

Independent validation matches every size bucket, heaptrack footer count,
folded-stack call total, and decoder aggregate. Each profile retains one
unknown initial allocation of 73,728 bytes and an allocation/deallocation
count difference of one; there are no missing allocation metadata or adjacency
mismatches. Unknown records are retained, not silently removed.

An exploratory attribution run used different configuration versions between
profiles. It is retained as excluded evidence. The authoritative
`attribution-final-v1` uses one root-sealed configuration for all four profiles
and passes independent admission and conservation checks. An initial export
helper failed before launching a subprocess because of a receipt field name;
the corrected helper exported the existing captures, without rerunning them.
The exact source-site reconciliation uses the stable raw decoder output.

## Next candidate boundary

The next bounded experiment should reduce retained window representation and
reuse parser buffer capacity across windows. It must preserve complete original
auditing, inherited namespace/depth state, exact outside gaps, aggregate limits,
fallible allocation, and canonical complete-audit errors. The ADR obligations
in 0948 still apply; this diagnostic changes none of them. Namespace-rich and
malformed-input coverage must accompany any implementation. Fresh prospective
correctness, native latency, allocation, and memory gates remain required.

This batch establishes the source of the rejected candidate's allocation costs;
it does not adopt an optimization or extrapolate to other Office formats,
cold-cache workloads, concurrency, or general namespace distributions.
The [reduced evidence packet](results/change-0949/README.md) records immutable
inputs, replay commands, independent reviews, and the external archive.
