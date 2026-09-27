# 0793 independent raw-results review

## Disposition

**The frozen allocation qualification fails in all ten traces.** The failure
is reproducible from the retained stack files and operation reports. The
owner count is not an authorized allocation fraction, and this review does
not promote an allocation source, speedup, memory reduction, or production
change.

This is an offline review of the terminal packet. I read
`0793-pptx-capture-allocation-attribution.md`, the control and heaptrack
reports, decoded stacks, histograms and logs, `root_audit.py`,
`root-audit.json`, and the retained allocator sources. I also ran the
packet-local `root_audit.py --check` and independently recomputed the stack
and counter rows. No build, native child, profiler, or production source was
run or changed.

## Corpus and reconciliation

The retained receipts contain 40 control reports with 120 samples and ten
heaptrack reports with ten samples. Every report receipt exits successfully;
the source digest, output digest, and semantic readback checks remain true.
The independent report walk therefore reproduces 50 reports and 130 samples.

For each repetition and shape, I parsed each `stack count` line in the whole
and owner files. The owner file is an exact subset of the whole file under
the packet's capture-wrapper frame predicate. The whole stack sum equals the
whole print-summary count and the size-histogram count. The owner print and
histogram files retain that same whole-process total; they are not owner-only
counter summaries. The allocation and profile-allocation reports also agree
on the operation counters used by the qualification rule.

The ten rows are:

| Repeat | Shape | Whole stack calls | Owner stack calls | `allocation_calls` | `reallocation_calls` | Frozen expected | Frozen result |
|---:|---|---:|---:|---:|---:|---:|---|
| 0 | tiny | 13,634 | 1,338 | 1,338 | 150 | 1,488 | fail |
| 0 | medium | 24,071 | 2,869 | 2,869 | 385 | 3,254 | fail |
| 0 | large | 608,290 | 72,106 | 72,106 | 2,603 | 74,709 | fail |
| 0 | vendor | 33,002 | 3,157 | 3,157 | 577 | 3,734 | fail |
| 0 | unicode-vendor | 33,038 | 3,157 | 3,157 | 577 | 3,734 | fail |
| 1 | tiny | 13,634 | 1,338 | 1,338 | 150 | 1,488 | fail |
| 1 | medium | 24,071 | 2,869 | 2,869 | 385 | 3,254 | fail |
| 1 | large | 608,290 | 72,106 | 72,106 | 2,603 | 74,709 | fail |
| 1 | vendor | 33,002 | 3,157 | 3,157 | 577 | 3,734 | fail |
| 1 | unicode-vendor | 33,038 | 3,157 | 3,157 | 577 | 3,734 | fail |

The owner stack sum equals `allocation_calls` in every row, and
`frozen_expected - owner_calls` equals `reallocation_calls` in every row.
The repeated values are evidence of stable retained observations; they do
not turn the failed rule into a pass.

## Why the frozen rule fails

The retained probe source makes the counter relationship explicit.
`counting_allocator.rs:52-58` calls `record_reallocation` for each successful
system `realloc`. `allocation_metrics.rs:713-723` then increments both
`allocation_calls` and `reallocation_calls`, while recording the new bytes
and old bytes. Thus every successful realloc is already included in
`allocation_calls`; adding `reallocation_calls` counts that same operation a
second time. Failed allocation calls are tracked separately and are zero in
all ten operation rows.

The post-capture comparison against `allocation_calls` alone explains why
the owner and operation rows agree, but it is supplementary evidence. It does
not amend the frozen plan, repair the qualification gate, or authorize owner
fractions after collection. The retained result remains `fail` for all ten
traces.

## Nested stack diagnostics and limits

The owner stacks contain these overlapping diagnostic counts, identical in
both repetitions:

| Shape | `check_for_duplicates` | `notes::codec::inspect_element` |
|---|---:|---:|
| tiny | 444 | 455 |
| medium | 1,038 | 1,046 |
| large | 61,342 | 61,274 |
| vendor | 1,254 | 1,262 |
| unicode-vendor | 1,254 | 1,262 |

These are stack costs nested inside the owner total. They overlap with one
another and with other frames, so they cannot be added, divided by the whole
count, or presented as allocation shares. They support a follow-up source
hypothesis only; they do not identify which allocation site caused any byte
or latency effect.

The whole-process totals include fixture construction, capture, output
verification, and teardown outside the wrapper. The owner stack is therefore
a bounded attribution subset, not a process-wide allocation total. Heaptrack
histogram byte totals and intercepted process high-water observations remain
process-scoped; this packet provides no operation RSS or native latency
result. The retained semantic/readback checks establish output validity, not
qualification of the failed allocation formula.

No production conclusion follows from these counts. A later experiment must
declare the existing counter semantics before capture, preserve the same
exact owner-stack and whole-process reconciliation, and qualify any source
claim independently.
