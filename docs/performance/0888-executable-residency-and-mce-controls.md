# 0888 Executable residency and MCE control audit

The large RSS flags in [0887](0887-mce-inline-attribute-storage.md) do not
establish a library heap regression. Live mapping samples place the difference
mainly in executable-file residency, and byte-identical copies of the binaries
produce substantially different process peaks. The precise kernel mechanism
remains unresolved. The candidate stays **rejected** because its original
native latency comparison failed the frozen adoption gate.

This batch changes no production or Rust benchmark code and claims no speedup.
It improves interpretation of the RSS evidence and records the controls needed
before another MCE candidate.

## Diagnostic protocol and evidence

At capture, HEAD was `b1f142e1f9`; production and harness source still matched
`335a40350b`. The before executable is retained from 0885; the after executable
is the rejected 0887 candidate. No binaries were rebuilt. Hashes bind both
executables to their existing source/build records.

The [first plan](results/change-0888/plan.json) runs four counterbalanced
before/after pairs of the producer XLSX full-cell scan. Each child has three
warmups and 200 measured iterations. A separate observer samples the live
benchmark child's `/proc/PID/status` and `smaps`, while GNU `time %M` records
process high-water RSS. The benchmark is pinned to CPU 12 and the observer to
CPU 13. Sampling requests a 10 ms interval; reading and compressing the data
adds work, so this is not an exact sample frequency.

The [copy-control plan](results/change-0888/copy-probe/plan.json) adds sixteen
processes: the original and a newly copied executable for each version, in
four rotated blocks. Each variant occupies each launch position once. Copies
are ordinary new files with different device/inode identities from their
sources and identical SHA-256, size, and permission mode. Complete hash reads
warm each file before capture; no cache was evicted, and file/cache histories
are not guaranteed equivalent. Each child again has three warmups and 200
measured iterations.

All **24 processes, 4,800 measured iterations, and 72 warmups** completed
without retry or replacement. The archive, 4,096-cell count, ordered-value
digest, producer census, source-version, and route checks all pass. Raw status
and mapping text is retained losslessly in compressed JSONL streams, with
command, report, snapshot-index, PID, executable-path, and artifact hashes.
There are 726 snapshots in the initial capture and 1,455 in the copy control.

These monitored populations are distinct from the uninstrumented 0887 latency
population. Their elapsed values are not pooled, substituted into its
confidence intervals, or used to reconsider adoption.

## What the mapping samples show

For the first four pairs:

| Executable | GNU peak RSS median [range], KiB | Sampled executable-file RSS median, KiB | Sampled total Anonymous median, KiB | Sampled heap RSS median, KiB | Sampled stack RSS median, KiB |
| --- | ---: | ---: | ---: | ---: | ---: |
| Before original | 21,644 [20,368–22,204] | 17,496 | 2,684 | 1,706 | 632 |
| Rejected candidate original | 24,310 [23,680–24,852] | 20,242 | 2,542 | 1,566 | 632 |

The sampled columns are medians of each child's independently computed
maximum for that metric. They need not occur at the same instant and must not
be added. `Anonymous` is the sum of that smaps field over **all** VMAs,
including heap, stack, unnamed mappings, and copy-on-write portions of named
file mappings. It is not just the separately classified unnamed-VMA bucket.

The executable-file category includes small anonymous/COW regions as well as
clean file pages. The read/execute mapping's extra residency is predominantly
`Private_Clean`; the name of a mapping alone does not prove that all its
resident pages remain file-backed. No sampled mapping reports nonzero
`FilePmdMapped` or `AnonHugePages`. That does not establish the physical
page-cache folio size or rule out fault-around/readahead effects.

The [static audit](results/change-0888/rss-static-audit.json) finds matching
entry-point stack sizes and a fixed 64 KiB executable-hash buffer outside the
timer. The candidate MCE start frame grows by 144 bytes and `.text` by 9,584
bytes. These observations narrow the investigation but do not by themselves
explain the process RSS difference.

GNU high-water values and sampled smaps totals have different accounting and
sampling behavior. Some smaps totals exceed the reported GNU maximum. They
are retained separately; neither is silently substituted for the other.
The captures are non-atomic, can miss transients or exit activity, and can
perturb execution. CPU affinity does not provide an exclusive host.

## Byte-identical executable control

The second, separately counterbalanced population gives:

| Variant | GNU peak RSS median [range], KiB | Sampled executable-file RSS median, KiB | Sampled total Anonymous median, KiB |
| --- | ---: | ---: | ---: |
| Before original | 21,812 [21,644–21,908] | 17,560 | 2,644 |
| Before copy | 19,748 [19,684–19,948] | 15,588 | 2,664 |
| Rejected candidate original | 24,568 [24,132–24,992] | 20,382 | 2,666 |
| Rejected candidate copy | 19,400 [19,176–19,624] | 15,180 | 2,588 |

Every copy has identical program bytes to its corresponding original, yet
its process peak is lower in all four blocks. The executable-file category
accounts for most of that difference. With the copies, the candidate's peak
is slightly below the before copy in each block; this is a diagnostic memory
observation, not a native performance or adoption result.

Copying changes pathname, inode, backing-file/cache history, and subsequent
loading conditions together. This experiment does not isolate which of those
causes the change, nor does it prove a particular kernel mechanism. It does
show that the earlier RSS gap cannot be attributed solely to a change in
program bytes or to the inline attribute buffer's heap cost.

The 0887 RSS observations remain valid observations of those processes. Its
rejection remains valid independently of RSS: native paired p50 did not meet
the required practical or statistical improvement. Historical capture and
review artifacts have not been rewritten.

For future comparisons, prepare both execution files through the same
recorded procedure and retain their content/build identities. Keep the whole-
process RSS review threshold, but investigate flags with byte-identical
execution-file controls and separate file/anonymous mapping evidence before
claiming a library heap regression. This is an additional diagnostic, not an
automatic exemption from a resource regression.

## MCE controls before another candidate

The [control audit](results/change-0888/mce-control-audit.json) records exact
fixture/member hashes, current limits, refusal boundaries, and a proposed
panel. It is a static audit; these controls were not built or run in this batch.

| Proposed real control member | Uncompressed bytes |
| --- | ---: |
| `StructuredRefs-lots-with-lookups.xlsx`, `xl/worksheets/sheet3.xml` | 1,558,391 |
| `drawing.docx`, `word/document.xml` | 288,070 |
| Optional `NumberedList.docx`, `word/styles.xml` | 173,620 |

The next control qualification should cover codec and stream paths; small MCE
branch/preservation and prefix-alias cases; inline/spill-sized tags; and
exact-limit/+1 attribute, namespace, directive, depth, input/output, name, and
attribute-byte cases. Count namespace declarations among raw tag attributes.
Distinguish introducing a new in-scope prefix from rebinding an existing one.
Large floods and long-URI cases remain separate safety diagnostics.

The existing `xml_attribute_bounds` timing wrappers do not provide a complete
oracle by themselves: some retain only output SHA/error text, without all MCE
report counters. Its differential traversal can also skip unreadable packages
or members. Qualification must explicitly account for every intended input,
fail or report skips, compare complete output/report/event evidence, and
check typed refusal behavior. No passing runtime qualification is inferred
from this source review.

## Replay, review, and cleanup

[Full analyses](results/change-0888/analysis.json.gz) and
[copy-control analyses](results/change-0888/copy-probe/analysis.json.gz) retain
every per-snapshot aggregation. A second reader independently recomputes
custody, semantics, and mapping maxima. The
[combined check](results/change-0888/analysis-verification.json) matches all
24 children and all 2,181 raw snapshots against the full analysis replay.

The first derived analysis export remains compressed as `analysis-v1.json.gz`.
The final reader adds stricter capture/schedule/executable metadata checks;
statistics are unchanged. Export receipts preserve both hashes and verify
lossless compression before removing duplicate uncompressed exports. No
workload was repeated to produce the final analyses.

```sh
python3 -B docs/performance/results/change-0888/check_analysis.py \
  --output /tmp/litchi-0888-replay-new.json
```

Replay expects the recorded Linux workspace/tool paths and the locally
retained 0885/0887 original executables, their frozen metadata, and captured
text. Temporary copied executables are reconstructable from
the originals and are not required for replay. There are no new Git worktrees
or build directories.

[Independent raw review](results/change-0888/rss-review.json) passes.
[Cleanup](results/change-0888/cleanup.json) removed only the marked
`/tmp/litchi-0888-copies` directory after verifying both copies against their
retained originals. Original binaries, source archives, and unrelated
worktrees are unchanged.
[Post-cleanup replay](results/change-0888/post-cleanup-verification.json)
passes without the copies, including complete analysis recomputation and
independent mapping totals. The broader performance goal remains active.
