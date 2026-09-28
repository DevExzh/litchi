# 0821 protocol review — real-file ordinary-save durability attribution

## Review status

This is an independent static review of the 0821 protocol and execution
boundary. It covers the real-file ordinary-save matrix, its artifact controls,
quality reuse, counterbalancing, and interpretation limits. It does not run
Cargo, a release build, an exporter, a workload, or an offline reader. The
reviewed source revision is the current repaired base
`8312aaa29b59f73d2a7409cb501828989320a8c9` (`test: respect process-wide
allocator live-byte accounting`); the abbreviated revision is retained here
only as a human-readable identity and the packet must retain the full hash.

The protocol is admissible after the 0821 plan, source/quality descriptors,
and root-owned driver inputs have been frozen with 0821 schema names. The
offline readers are intentionally unfrozen until every capture child is
terminal; they are a later handoff and must then replay the retained packet
fail-closed. Explicit paths to the sealed 0820 repair quality receipts are
legitimate quality-reuse inputs. A copied 0820 schema, target, count, or
source-revision identity in an 0821 driver or reader is a packet error; the
0820 quality-reuse references themselves must remain explicit. No production
or runtime harness edit is authorized by this review.

## Scope and matrix

The timed workload uses three checked-in caller-named OOXML files:

| Format | Input | Bytes | SHA-256 |
| --- | --- | ---: | --- |
| DOCX | `test-data/ooxml/docx/documentProperties.docx` | 23,503 | `1cff7a0a94dfce307a70032d21070d26ae34b9fdf742cf70fa66d4a2078ec9d5` |
| XLSX | `test-data/libreoffice-core/sc/qa/unit/data/xlsx/dateAutofilter.xlsx` | 8,435 | `d7ab3dbb59388d245ee779bf8547748dc6bac70f3c7216e673e0d97dbbbd6bc4` |
| PPTX | `test-data/ooxml/pptx/shapes.pptx` | 68,822 | `19fde9b87e33dd1a95fdbba0cf6abc2278bf03874f4665c7f8b88b6afe4a2571` |

The real-file selectors are 3 formats × 2 phases × 4 policies = 24 rows.
The phases are `lifecycle` (`open + edit + save`) and `atomic_publish`
(`save-to-path` with open and edit outside the clock). The policy labels are
`default`, `full`, `file-only`, and `no-sync`. The `default` route is the
documented Full-durability route; `full` is the explicit Full control;
`file-only` and `no-sync` are weaker caller-selected configurations. A policy
comparison is therefore configuration attribution. It is not a production
optimization, a historical speedup, or evidence for changing the
Full-by-default contract.

Untimed artifact admission covers six cases: each of the three real files and
each deterministic generated corpus. Every case must export five outputs:
`default`, `full`, `file-only`, `no-sync`, and the sequential `stream` control.
The stream output is an artifact correctness control and contributes no timing
row. It must not be silently pooled with the four filesystem policies.

The semantic operation is fixed per format: DOCX appends the marker paragraph,
XLSX edits the first admitted worksheet `A1`, and PPTX edits the first
admitted slide/shape text. The generated controls remain qualification
fixtures, not additional timed real-file cases. A typed edit refusal is
retained as an outcome and cannot be converted into a guessed edit or a
replacement sample.

## Acquisition and statistical controls

The native lane should retain six process blocks with 30 samples and three
warmups per selector. The observer lane should retain two blocks with three
samples and no warmups per selector; its elapsed values and operation metrics
are diagnostic and remain separate from native timing. Qualification is a
separate one-sample-per-selector lane. With the 24-row matrix this yields 24
qualification reports/24 samples, 144 native reports/4,320 samples, and 48
observer reports/144 samples: 216 reports and 4,488 samples. These are report
and sample counts, not counts of saved files.

The six native format/phase groups run in forward/reverse/forward/reverse/
reverse/forward order. Within the groups, policy order is counterbalanced as:

```text
default, full, file-only, no-sync
no-sync, file-only, full, default
full, file-only, no-sync, default
default, no-sync, file-only, full
file-only, no-sync, default, full
full, default, no-sync, file-only
```

Observer groups use forward/reverse order. The policy order controls position
effects while preserving the paired policy/default comparison inside each
native block. The required summary is nearest-rank p50/p95/p99 within each
process, then the median of six process p50 values. Preserve the harness p50
integer midpoint separately. Bootstrap only the specified median of six
matched process-block policy/default p50 ratios, with 10,000 resamples and
seed `821821`. The plan and reader must agree exactly. Sorted interval ranks
are 250 and 9749 at 95% confidence.

Do not remove, replace, or pool samples based on spread or tail flags. Keep
raw vectors, block order, policy order, warmup counts, process identity, and
the paired ratios. Lifecycle and atomic values are independently prepared and
are not additive; subtracting them does not measure synchronization cost.

## Clock and durability boundary

For `lifecycle`, the clock begins before the documented path open and ends
after the selected save returns. For `atomic_publish`, open and semantic edit
occur before the clock; the clock contains only the selected save-to-path
publication. Readback, digesting, owner destruction, destination removal, and
offline verification are outside both clocks.

Every filesystem save still validates the destination, creates a sibling
temporary file in the destination directory, writes and finishes the complete
artifact, preserves existing regular-file permissions where promised, and
publishes with one same-directory rename. Full synchronizes the temporary and
then the parent directory where supported. FileOnly synchronizes the
temporary but skips the parent directory. NoSync skips both synchronization
calls. A skipped operation must not be described as having run or as having
failed.

The report's generic `timing_scope` is a phase-level description and may name
the complete atomic sequence for every policy. Readers must use the exact
`save_durability` and `atomic_publication_steps` fields to attribute policy
steps: default omits `save_durability` because it calls documented `save`,
explicit Full carries `full`, FileOnly carries `file-only`, and NoSync carries
`no-sync`. A generic timing string cannot prove that a skipped synchronization
was executed.

## Quality reuse and custody

The 0820 repair quality receipt is reusable only as exact evidence, not as a
new quality run. The packet must bind the committed 0820 repair seal and
receipts, including six passing gates and the full all-feature test result of
641 passed, zero failed, and one ignored. The 0821 packet must independently
verify that:

* the 9,197 production-file source descriptor matches the current tree;
* `tools/perf-baseline/src/ordinary_save.rs` and all runtime production files
  match the repaired quality source descriptor;
* the only source difference introduced by 0820 is the committed
  `#[cfg(test)]` allocator assertion repair, with no production or ordinary-save
  runtime change;
* both root and tool `Cargo.lock` files, rustfmt configuration, all 35
  normative documents, and corpus inputs match the sealed 0820 witnesses;
* quality evidence is referenced by digest and is not relabeled as a fresh
  0821 test execution.

The current 0821 build, artifact, qualification, and capture receipts must be
fresh and must bind the current full source identity, both locks, the repaired
quality evidence, the 0821 plan, source hashes, and binary hashes. The root
driver inputs are frozen before execution; all child processes must be
terminal before `analyze.py` or `validate.py` runs. The owned target and
scratch paths must be 0821-specific, and cleanup must not touch the 0820
target, prior packets, or unrelated workspace files.

## Host and interpretation limits

Inputs and staged sources are warm-cache observations on one shared Linux
host with caller-selected CPU affinity. A fresh destination is removed after
each sample, so the timed publication normally targets an absent destination.
This does not measure replacement over an existing file or permission-copy
latency. The requested `cold` state, if retained as metadata, is not evidence
of a physical cold filesystem cache. The packet must make no claim about
physical I/O, device bandwidth, power-loss durability, operating-system crash
behavior, Windows/macOS, remote storage, external Office compatibility, or
exclusive-host stability.

Observer procfs counters and allocator metrics describe the whole child and
retain their fixed empty controls without subtraction. Logical byte rates are
workload descriptors only. Equal output bytes, successful reopen, XML/relationship
closure checks, and untouched-member preservation establish publication
correctness; they do not establish crash or power-loss durability.

## Static decision

The matrix, policy controls, clock boundary, preservation requirements, and
interpretation limits are suitable for execution. Root may execute once the
0821 plan, source/quality-reuse witnesses, and root-owned driver inputs prove
the exact 24-row matrix and fresh-build/admission/capture custody described
above. After all children terminate, the unfrozen readers must be checked
against the same packet before analysis and sealing. Any remaining 0820
schema, path, count, or source-revision identity in an 0821 driver or reader
is a release blocker; explicit 0820 quality-reuse paths are expected and are
not a blocker.
