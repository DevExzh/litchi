# 0820 protocol review — explicit OOXML save-durability attribution

## Review status

This review covers the frozen 0820 plan and the current production save
implementation at base `096b810f23cc66fe01ec8c96c74f36888f954eff`. The
final 0820 driver handoff carries the frozen schemas, policy matrix, and
report/sample counts, and its static source checks pass. Post-capture reader
validation remains a separate unfrozen handoff in `reader-review.md`; no
capture may be interpreted until those reader checks are complete.

The scope is configuration attribution for three real OOXML files and four
save policies. It does not authorize a production change or a recommendation
to weaken the ordinary-save default.

## Frozen scope and custody

The plan is `litchi.performance.0820.plan.v1`, SHA-256
`782d4931ff396f8fb6e463cd2eb054515b941a3f5b18a14542de4813a8c74597`. Its
base is the same source revision recorded in `origin.json`; production and
the performance harness are declared unchanged. The three timed inputs are:

| Format | Input | Bytes | SHA-256 |
| --- | --- | ---: | --- |
| DOCX | `test-data/ooxml/docx/documentProperties.docx` | 23,503 | `1cff7a0a94dfce307a70032d21070d26ae34b9fdf742cf70fa66d4a2078ec9d5` |
| XLSX | `test-data/libreoffice-core/sc/qa/unit/data/xlsx/dateAutofilter.xlsx` | 8,435 | `d7ab3dbb59388d245ee779bf8547748dc6bac70f3c7216e673e0d97dbbbd6bc4` |
| PPTX | `test-data/ooxml/pptx/shapes.pptx` | 68,822 | `19fde9b87e33dd1a95fdbba0cf6abc2278bf03874f4665c7f8b88b6afe4a2571` |

The workload edits the first admitted semantic target for each real file:
DOCX appends one paragraph, XLSX edits the first worksheet's `A1`, and PPTX
edits the first admitted slide/shape text. The three generated corpora are
artifact qualification controls only. They must not be added to the timed
policy matrix.

The matrix is 3 formats × 2 phases × 4 policies = 24 selectors. The phases
are lifecycle (`open + edit + save`) and atomic publication (`save-to-path`).
The planned acquisition is 24 qualification reports/24 samples, 144 native
reports/4,320 samples (six 30-sample blocks, three warmups), and 48 observer
reports/144 samples (two three-sample blocks), for 216 reports and 4,488
samples. Native and observer elapsed values remain separate populations.

The six real/generated artifact cases must first pass the independent XML/
relationship audit and the ZIP preservation audit for all five outputs:
`default`, `full`, `file-only`, `no-sync`, and `stream`. Artifact equality and
reopen checks establish the output contract; they do not turn the generated
controls or untimed artifacts into timing samples.

## Policy meaning and attribution

The policy labels have these exact meanings:

| Label | Entry point | Synchronization requested |
| --- | --- | --- |
| `default` | documented `save` | `Full`: temporary file sync, rename, parent-directory sync where supported |
| `full` | explicit `save_with_durability(Full)` | same `Full` sequence |
| `file-only` | explicit `save_with_durability(FileOnly)` | temporary file sync and rename; no parent-directory sync |
| `no-sync` | explicit `save_with_durability(NoSync)` | rename only; no temporary-file or parent-directory sync |

All four policies still validate the destination, create the complete sibling
temporary artifact, finish the writer, preserve existing permissions, and
replace the destination with one same-directory rename. The policy controls
only the synchronization calls in that publication sequence. It does not
change the semantic edit, serialization bytes, source authorization, cleanup
contract, or typed failure behavior before replacement.

`default` and `full` are a deliberate route control. `save` is defined as
`save_with_durability(Full)` in each of the three format owners, so a measured
default/full difference is dispatch or run-to-run noise, not a measured
synchronization saving. It must not be described as a benefit of either
policy.

`file-only` and `no-sync` compare configurations that give up crash-persistence
guarantees. A lower elapsed value, if observed, is a configuration difference
on this warm host and filesystem. It is not a general filesystem-sync cost,
power-loss result, or reason to adopt a weaker default. The policy comparison
should report the full save-route medians and within-block policy/default
ratios, with the atomic phase as the primary attribution view. Lifecycle
values show end-to-end impact but include open and edit work, so they dilute
the publication difference.

The policy matrix is not additive. Do not subtract lifecycle, edit, counting,
or stream medians to derive a sync cost. Edit and counting are intentionally
not policy selectors in this experiment. The counting route is a separate
serialization control, and PPTX `to_bytes()` materializes the complete buffer
before it is written to the bounded sink; it is not streaming evidence.

## Timing brackets and order

For lifecycle, the timer starts immediately before opening the source path and
stops immediately after the selected save returns. Destination readback,
digest, owner destruction, and cleanup are outside the interval. For atomic
publication, opening and editing happen before the timer; the timer covers
only the selected save-to-path call, including temporary creation, output
serialization, conditional file sync, rename, and conditional parent sync.
Readback, digest, owner destruction, and cleanup remain outside it.

Each sample starts from the same staged source and removes the destination
after verification. Consequently the timed publication normally targets an
absent destination; it still performs the destination validation/permission
probe, but this matrix does not measure replacement over an existing file or
permission preservation on that existing file. Policy order rotates inside
each process block as frozen by the plan, while format/phase block order is
forward/reverse/forward/reverse/reverse/forward for native and forward/reverse
for observer. This supports within-block configuration comparison while
retaining order effects as diagnostic evidence. It does not establish
exclusive-host or device-level control.

Use nearest-rank p95/p99 and the plan's median-of-six process p50 statistic.
Preserve the raw harness p50 midpoint definition separately from planned
quantiles. Bootstrap only the specified median of six matched process-block
policy/default p50 ratios with seed `820820`, 10,000 resamples, and ranks
250/9749. Keep spread and tail flags; do not remove or replace samples.

The native binary has no allocator or procfs observer feature. The observer
binary's allocator and process counters are diagnostic only. Retain 32 empty
procfs controls per child without subtracting them from operation deltas;
report whole-child RSS and quantized CPU counters with their stated limits.

## Interpretation limits

The source and artifacts are warm-cache, small checked-in files on one shared
Linux host and caller-selected CPU affinity. The result cannot establish
cold-cache, physical-I/O, device-bandwidth, power-loss, operating-system
crash durability, replacement-over-existing-file behavior, Windows/macOS,
external Office, or general workload performance behavior. Equal output
bytes and successful reopen are publication correctness checks; they are not
crash or power-loss evidence.
Logical source/output bytes per second are workload descriptors, not physical
I/O or memory-bandwidth measurements.

The generic `timing_scope` field emitted by `ordinary_save.rs` describes the
phase and therefore mentions the full save sequence for every atomic policy.
For `file-only` and `no-sync`, readers must use the report's
`save_durability` and `atomic_publication_steps` fields for the actual policy
steps. The default deliberately has an absent `save_durability`; normalize
that absence to documented `save`/`Full` in prose while retaining the raw
field. A report that presents the generic field as proof that a skipped sync
was executed is invalid.

## Required final-driver checks

Before capture is authorized, the refreshed drivers must prove all of the
following:

1. Every report carries one of the four policy labels and the expected phase;
   default has no explicit durability field, while the other three labels
   carry the matching explicit level.
2. Artifact admission, qualification admission, and every capture receipt
   bind the 0820 plan hash, source hashes, binary hashes, quality/build
   evidence, and the policy selector. No 0819 schema, count, target, scratch
   marker, or recovery path remains.
3. The default/full route is retained as a control, and weaker-policy output
   hashes and semantic/reopen checks remain tied to the same admitted output.
4. Native and observer builds are separate, observer controls are retained
   without subtraction, and all child processes terminate before offline
   readers run.
5. The production source and unrelated workspace hashes remain unchanged.

Subject to those driver checks and the independent artifact gates, this is a
valid durability-attribution baseline. Any policy result must remain a
configuration observation under the existing Full-by-default contract.
