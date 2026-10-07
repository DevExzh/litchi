# XML multi-window source/profile reconciliation (0949, v2)

Status: read-only analysis of the four sealed 0949 heaptrack captures. No
workload, build, decoder, raw capture, or repository source was changed.

This report uses only the stable raw decode at
`/home/zhuhe/litchi-performance-0949/raw-decoder-v1/`. The exploratory
`attribution-v1` output is excluded. The decoder is
`tools/decode_heaptrack_0949_v1.py`, SHA-256
`c87e7089774f0eac24ffde262b7fe27b1f1ff4586798683bcc44fa68049bc7bf`.
The candidate source reviewed is the preserved 0948 v3 tree under
`/home/zhuhe/litchi-performance-0948/candidate-build-v3`; its relevant
`audit.rs` SHA-256 is
`29efd9aa99c9491a1e86cfc25f1bbb31b3da7b6cbfc4438a1638c15b18b95cee`.

The standalone create-only verifier is
`xml-multiwindow-source-profile-reconciliation-v2.py`, SHA-256
`e039a99aa4d505692cd12d0a2c4ef17de37b0ebf1546d295f7c362765a7a1cd4`,
and its generated JSON is
`xml-multiwindow-source-profile-reconciliation-v2.json`, SHA-256
`18f27ae921871a9321069f6fe4f72919b140d873aac35100f1f88dc7d5813d84`.
The verifier records the source-file hashes and target lines in that JSON,
refuses an existing output, and reads only the stable decoded aggregates.
A replay with a new output path is:

```text
python3 /home/zhuhe/litchi-performance-0949/review/xml-multiwindow-source-profile-reconciliation-v2.py \
  --packet /home/zhuhe/litchi-performance-0949 \
  --source-root /home/zhuhe/litchi-performance-0948/candidate-build-v3 \
  --output /home/zhuhe/litchi-performance-0949/review/replay.json
```

## Accounting scope

The decoder resolves every raw `+` record to a requested allocation size and
trace. Its byte total is requested bytes before frees over the whole process
profile. A trace whose first allocator frame is `std/src/sys/alloc/unix.rs:54`
is a reallocation request; it remains part of the raw requested-byte/event
total. These are heap requests, not RSS or allocator-rounded physical size.

The publication counters below are a different interval and metric: the
counting allocator's `allocation_calls`/`allocated_bytes` counters for the
publication region in each child `report.json`. They are included only to
show the boundary between the stable raw profile and the published phase
comparison.

| profile | raw `+` events | raw requested bytes |
| --- | ---: | ---: |
| baseline medium | 979,147 | 475,804,603 |
| candidate medium | 979,434 | 476,010,899 |
| candidate - baseline medium | **+287** | **+206,296** |
| baseline dense | 2,596,848 | 667,496,174 |
| candidate dense | 2,597,386 | 667,935,103 |
| candidate - baseline dense | **+538** | **+438,929** |

All four decodes have zero unresolved trace definitions and zero `+` records
without allocation metadata. Each has one identical unknown initial-stack
allocation (73,728 B), so that bucket does not explain either delta.

## Resolved allocations for each successful replacement window

The candidate's `MultiWindowSearch` records a window during the complete
original scan, then `prove_windows` calls `audit_window` for the corresponding
replacement slice. The successful window stacks end at
`audit.rs:2651` (`prove_windows`) and `audit.rs:1173`
(`verify_source_replacement_with_spans`). They are therefore separable from
ordinary full audits and from the one-span path.

The three rows below are physical heap requests made by each such parser
state. `quick-xml` is version 0.41.0 in the resolved stack.

| resolved source frame | medium | dense | allocation made |
| --- | ---: | ---: | --- |
| `quick-xml/src/reader/state.rs:346`, via `audit.rs:2061`, trace `8c4d` / `8cef` | 93 calls, 2,976 B | 177 calls, 5,664 B | `opened_starts` growth, 32 B/request |
| `quick-xml/src/reader/state.rs:347`, via `audit.rs:2061`, trace `8c50` / `8cf2` | 93 calls, 744 B | 177 calls, 1,416 B | `opened_buffer` growth, 8 B/request |
| `audit.rs:2106`, via `audit_window` `2710`, trace `8c52` / `8cf4` | 93 calls, 744 B | 177 calls, 1,416 B | `State::spaces.try_reserve(1)`, 8 B/request |
| **successful-window parser subtotal** | **279 calls, 4,464 B** | **531 calls, 8,496 B** | **3 heap requests/window** |

The 93 and 177 counts are measured trace call counts. They correspond to the
multi-window portions of the medium and dense candidates. The dense input has
one single-span part that delegates the existing one-window path, leaving 177
multi-window windows.

These are separate heap blocks owned by each `audit_window` parser state. The
`Window` value that describes the state is not allocated once per row; it is
stored inline in the `Vec<Window>` block described next.

## `Vec<Window>` growth at `MultiWindowSearch::consider` (`audit.rs:2509`)

The request sizes are exact powers of two in element capacity. They infer a
736-byte `Window` element in this binary (`2,944 / 4 = 736`, and every larger
request is the corresponding capacity times 736). This is an inference from
allocator requests, not a direct `size_of` probe.

### Medium

Trace `8c45` is four initial allocations of 2,944 B. Trace `8c47` is twelve
reallocation requests with this size histogram:

| request size | calls | requested bytes | inferred capacity |
| ---: | ---: | ---: | ---: |
| 2,944 B | 4 alloc | 11,776 B | 4 elements |
| 5,888 B | 4 realloc | 23,552 B | 8 elements |
| 11,776 B | 4 realloc | 47,104 B | 16 elements |
| 23,552 B | 4 realloc | 94,208 B | 32 elements |
| **total** | **16 requests** | **176,640 B** | — |

This is four affected parts starting with capacity four; each grows through
capacity 8, 16, and 32. Every `Window` (including its inline
`Namespaces` field, ranges, counters, and metadata) resides inside these
blocks. There is no separate per-window heap block in this trace class.

### Dense

Trace `8ce7` is three initial allocations of 2,944 B. Trace `8ce9` has eight
reallocation requests:

| request size | calls | requested bytes | inferred capacity |
| ---: | ---: | ---: | ---: |
| 2,944 B | 3 alloc | 8,832 B | 3 initial blocks, capacity 4 |
| 5,888 B | 2 realloc | 11,776 B | capacity 8 |
| 11,776 B | 2 realloc | 23,552 B | capacity 16 |
| 23,552 B | 1 realloc | 23,552 B | capacity 32 |
| 47,104 B | 1 realloc | 47,104 B | capacity 64 |
| 94,208 B | 1 realloc | 94,208 B | capacity 128 |
| 188,416 B | 1 realloc | 188,416 B | capacity 256 |
| **total** | **11 requests** | **397,440 B** | — |

The three multi-window parts have approximately 164, 10, and 3 spans: the
first reaches capacity 256, the second capacity 16, and the third stays at
capacity 4. The capacities are inferred from the request histogram; the span
counts are the source/workload shape used for the capture.

## `MultiWindowSearch::open` growth (`audit.rs:2554`)

The open-element stack stores `(start, Counters)` inline. Its request sizes
infer a 32-byte element: 128 B is capacity four and 256 B is capacity eight.

| profile | initial allocations | reallocations | total requests / bytes |
| --- | --- | --- | ---: |
| medium, traces `8c3c` / `8c42` | 4 x 128 B = 512 B | 4 x 256 B = 1,024 B | **8 / 1,536 B** |
| dense, traces `8cde` / `8ce4` | 3 x 128 B = 384 B | 3 x 256 B = 768 B | **6 / 1,152 B** |

## Namespace clone result

`MultiWindowSearch::consider` checks `Namespaces::clone_memory_charge` before
calling `Namespaces::try_clone`; the charge is a conservative logical
admission value, not an RSS or physical allocator measurement. The retained
`Namespaces` value is part of the 736-byte inline `Window` element when its
vectors/indexes are empty.

No candidate-medium or candidate-dense raw trace resolves to
`audit/namespaces.rs:try_clone_vec` (571--577), `try_clone_index` (579--584),
or `Namespaces::try_clone` (140--150). Thus these captures show no dynamic
namespace-vector/index heap blocks. This is consistent with the captured
worksheet roots having only the default namespace. It does not mean that a
nonempty namespace clone is free; it means this workload did not produce a
resolved allocation at those helper frames.

## Cross-accounting reconciliation

The measured candidate-only multi-window classes are:

| profile | parser-window requests | `Vec<Window>` requests | `open` requests | candidate class subtotal |
| --- | ---: | ---: | ---: | ---: |
| medium | 279 / 4,464 B | 16 / 176,640 B | 8 / 1,536 B | **303 / 182,640 B** |
| dense | 531 / 8,496 B | 11 / 397,440 B | 6 / 1,152 B | **548 / 407,088 B** |

These subtotals are resolved contributions in the candidate whole-process
profile. They must not be equated directly with candidate-minus-baseline raw
profile deltas: the baseline has corresponding ordinary audit paths, and the
candidate also changes the common path/source line layout. The measured net
residual after subtracting the resolved candidate classes from the raw pair
delta is:

* medium: raw pair delta minus class subtotal = **-16 requests, +23,656 B**;
* dense: raw pair delta minus class subtotal = **-10 requests, +31,841 B**.

The residual is whole-process activity outside the successful multi-window
classes, not an unresolved namespace-clone claim. In particular, raw trace
reallocation requests are included in the class subtotal, while the pair
comparison also includes baseline/common reallocations.

The separate publication-region counters from the sealed child reports are:

| profile | baseline calls / bytes | candidate calls / bytes | delta calls / bytes | reallocation delta |
| --- | ---: | ---: | ---: | ---: |
| medium | 406 / 2,495,613 B | 681 / 2,677,821 B | **+275 / +182,208 B** | +4 (50 -> 54) |
| dense | 406 / 2,765,840 B | 931 / 3,172,336 B | **+525 / +406,496 B** | +1 (49 -> 50) |

The `Vec<Window>` request class alone accounts for 176,640 / 182,208 =
**96.94%** of the medium publication byte delta and 397,440 / 406,496 =
**97.77%** of the dense delta. The parser and `open` classes are close to the
remaining phase bytes but do not exactly equal them because the phase metric
subtracts the baseline publication path and is not the whole-profile raw
trace total.

For the source-state model used in the 0947 review, medium has four complete
replacement parser states in the baseline and 93 candidate multi-window
states, so the 89 extra states imply 89 x 3 = 267 extra parser allocation
calls. Adding four initial `Vec<Window>` allocations and four initial `open`
allocations gives 275, exactly matching the medium publication
`allocation_calls` delta. Dense has three complete baseline states plus one
old one-window state versus 177 multi-window states plus that old path: 174 x
3 = 522 extra parser calls. The publication call delta is 525; the remaining
three calls are phase/path accounting outside this simple state-instance
model. These are inferred state-instance comparisons; the per-frame counts
and bytes above are the measured stable raw evidence.

## Limits of this attribution

The stable raw decode identifies the successful multi-window parser and both
new vector growth sites by source frame and request-size histogram. It does
not provide publication markers, so it cannot assign its whole-profile
residuals to the publication interval. It also cannot turn the inline
`Namespaces` field into a physical RSS measurement. A future namespace-rich
capture would be needed to measure dynamic clone blocks; the current capture
only establishes that their resolved count is zero here.

## Exact publisher symbol/call-site slice

As a separate check, I selected raw traces only when the resolved stack
contained both the exact function
`SourceBackedEditor::publish_multi_commit_to_stream` and the benchmark caller
`litchi_perf_baseline::run_xlsx_cell_values_edit_save` at `src/lib.rs:44840`.
This is a symbol/call-site scope in the whole-process trace; heaptrack does
not mark the allocator publication interval, so this selection is not a
claim about temporal profiling boundaries.

A publisher-frame-only selection is broader and is not usable as a phase
proxy: it totals BM 1,087 calls / 38,082,187 B, CM 1,362 / 38,264,875 B,
BD 1,089 / 56,181,853 B, and CD 1,614 / 56,588,829 B. The additional
`44840` caller frame isolates the intended sample:

| role | exact symbol/call-site raw requests | `alloc` requests | `realloc` requests | requested bytes | sealed phase `allocation_calls` / `allocated_bytes` |
| --- | ---: | ---: | ---: | ---: | ---: |
| baseline medium | 406 | 356 | 50 | 2,495,613 B | 406 / 2,495,613 B |
| candidate medium | 681 | 627 | 54 | 2,677,821 B | 681 / 2,677,821 B |
| baseline dense | 406 | 357 | 49 | 2,765,840 B | 406 / 2,765,840 B |
| candidate dense | 931 | 881 | 50 | 3,172,336 B | 931 / 3,172,336 B |

The exact symbol/call-site raw totals therefore match the phase
`allocation_calls` and `allocated_bytes` fields byte-for-byte and count-for-count.
The separate `reallocation_calls` field is a subset of `allocation_calls`,
because `allocation_metrics::reallocation_locked` increments both counters
while adding only the new size to `allocated_bytes`. Adding the two counters
would double-count reallocations; the raw `+` totals and publication
`allocation_calls` already include each successful reallocation once.

The candidate multi-window class is entirely inside this exact symbol/call-site
slice. For medium it is 303 requests / 182,640 B; for dense it is 548
requests / 407,088 B. Subtracting those disjoint candidate classes from the
exact dual-filter pair deltas gives the corrected common-path residuals:
medium **-28 requests / -432 B** and dense **-23 requests / -592 B**. The
verifier computes and asserts these values. Reallocation counts are reported
as a subset diagnostic and are never added to `allocation_calls`.

The exact symbol/call-site selection covered 125, 136, 129, and 141 aggregate
trace rows for BM, CM, BD, and CD respectively. Each row's `calls` and
`requested_bytes` was included, regardless of whether its allocator leaf was
`alloc` or `realloc`; no physical or inline request class was filtered out.
