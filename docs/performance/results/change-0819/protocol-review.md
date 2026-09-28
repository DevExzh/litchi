# 0819 protocol review

This packet is a baseline-only measurement of the ordinary documented OOXML
save route at revision `0c784f7aecda2a38bbe6cc98a684ce68da1f0dc8`. It makes no
optimization, adoption, speedup, regression, or historical timing claim. The
three real-file cases are the same checked-in inputs queued by 0817; the DOCX
relationship-preservation repair in 0818 is part of this packet's base.

## Corpus and custody

The caller-named inputs are fixed by path, size, and SHA-256:

| Format | Input | Bytes | SHA-256 |
| --- | --- | ---: | --- |
| DOCX | `test-data/ooxml/docx/documentProperties.docx` | 23,503 | `1cff7a0a94dfce307a70032d21070d26ae34b9fdf742cf70fa66d4a2078ec9d5` |
| XLSX | `test-data/libreoffice-core/sc/qa/unit/data/xlsx/dateAutofilter.xlsx` | 8,435 | `d7ab3dbb59388d245ee779bf8547748dc6bac70f3c7216e673e0d97dbbbd6bc4` |
| PPTX | `test-data/ooxml/pptx/shapes.pptx` | 68,822 | `19fde9b87e33dd1a95fdbba0cf6abc2278bf03874f4665c7f8b88b6afe4a2571` |

Each input is classified from its package main part: `word/document.xml`,
`xl/workbook.xml`, or `ppt/presentation.xml`. The real-file path is bounded at
32 MiB and the artifact auditor applies finite ZIP/XML limits. Missing,
duplicate, ambiguous, malformed, encrypted, unsafe, or traversal-bearing
entries are refused. The packet retains source member inventories,
compression methods, sizes, relationship/content-type graphs, edit targets,
and source identities.

`test-data/office-interop/PROVENANCE.md` is a historical Litchi edit /
LibreOffice resave / Litchi readback record. It is retained as provenance for
the checked-in corpus only. Current outputs are not externally resaved in this
batch, and no Microsoft producer is inferred for the DOCX or PPTX files. Old
timings and another dependency graph are not pooled with this baseline.

The existing tracked `tools/perf-baseline/Cargo.lock` graph is used exactly as
checked in. Its differences from the workspace lock are recorded in
`lock-parity.json`; neither lock is updated. Production and tracked harness
sources are frozen at the base revision for every driver child.

## Admission and preservation

Before any qualification or timing, the exporter creates six untimed cases:
three generated controls and the three caller-named real files. Each case
retains the source archive plus five policy outputs: `default`, explicit
`full`, `file-only`, `no-sync`, and a separate sequential `stream` output.
The generated controls are exporter controls and are not native timing cases.
The default public save and explicit `full` policy have the same full
durability semantics; weaker policies and the stream output do not substitute
for the native default save.

Admission requires both independent checks. `artifact_audit.py` verifies finite
ZIP/XML structure, source identity, the format-specific semantic edit and
closure, relationships and content types, untouched decoded XML or binary
members, output reopening, and policy identity. `preservation.py` separately
checks the default output's ZIP member order and archive comment, every
untouched member's ZIP metadata and compressed payload, and the expected DOCX
edit closure. All five policy outputs are required to be byte-identical by the
artifact audit. The DOCX relationship repair in the base is specifically
covered by exact untouched `word/_rels/document.xml.rels` bytes; an equal edge
set with changed relationship XML is insufficient.

The audit also binds each real staged archive byte-for-byte to its caller-named
path. A typed edit refusal remains a refusal with its reason and output
preserved; deterministic output and a successful process exit do not turn a
refusal into an admitted edit. No qualification or timing denominator is
replaced when an independent preservation gate fails.

## Phase boundaries

The twelve selectors are three real inputs × four independent phases. They are
not an additive decomposition, and no phase is inferred by subtracting other
phase medians.

- `lifecycle` opens the path, makes the one semantic edit, and saves to a
  path inside one interval. Destination preparation, readback, digest, and
  cleanup are outside the interval.
- `edit` opens the path before timing and measures only the semantic edit and
  its commit. It publishes no per-sample file; its retained
  `published_sha256` vector is intentionally empty.
- `atomic_publish` opens and edits before timing, then measures only the
  documented save-to-path. With no `--save-durability` override, public
  `save` is full durability: sibling temporary creation, publication write,
  temporary-file `sync_all`, rename, and parent-directory synchronization are
  inside the interval. Readback and cleanup are outside it.
- `counting_publish` opens and edits before timing and measures sequential
  serialization into a bounded counting sink. DOCX uses `to_stream` and XLSX
  uses `write_to`. PPTX has no sequential serializer here: `to_bytes()`
  materializes the complete package and the result is then accepted by the
  sink inside the interval. PPTX counting is therefore materialization-plus-
  sink evidence, not a streaming serializer or bounded-memory streaming
  claim. Byte accounting runs after timing.

Semantic reopening, hashes, relationship/XML checks, source preservation, and
the edit-outcome checks occur outside the measured interval. The owner is
destroyed after each timed region; retained live bytes therefore do not mean
that the operation leaked memory.

## Capture matrix

The plan retains the following lanes:

| Lane | Binary/features | Blocks × order | Warmup | Samples/selector | Reports | Samples |
| --- | --- | --- | ---: | ---: | ---: | ---: |
| Qualification | observer; both diagnostic features | 1 × forward | 0 | 1 | 12 | 12 |
| Native | `litchi-perf-baseline`; no features | 6 × `F,R,F,R,R,F` | 3 | 30 | 72 | 2,160 |
| Observer | `litchi-perf-baseline-alloc`; `allocator-metrics,ordinary-save-process-metrics` | 2 × `F,R` | 0 | 3 | 24 | 72 |
| **Total** |  |  |  |  | **108** | **2,244** |

All child processes are pinned to CPUs 12–19 on the recorded shared EPYC host.
This is warm filesystem/provider-cache evidence with no exclusive-CPU,
physical-cold-cache, network, or physical-device claim. Native elapsed values
alone support the baseline latency table. Observer elapsed, allocator, and
procfs values remain diagnostic and are never pooled with native latency.

The observer captures 32 adjacent empty procfs snapshot pairs before its
warmups. It retains their deltas, includes probe activity in the operation
window, and subtracts none of them. CPU ticks may quantize short operations to
zero; process read/write counters are same-process observations, not physical
I/O attribution. The external `/usr/bin/time` RSS value is whole-child RSS,
including setup and corpus qualification, rather than an operation-local
allocation peak.

## Statistics and claim boundary

Each raw report retains all samples and uses the harness's integer midpoint for
`p50`; its `p95` and `p99` are nearest-rank values. The frozen analysis plan
recomputes nearest-rank `p50`, `p95`, and `p99` within each process block, then
uses the median of the six process-block nearest-rank `p50` values as its
primary cross-process summary. Packet analysis must retain both definitions:
the raw report midpoint is a replay field and must not be silently substituted
for the planned nearest-rank process `p50`. Block values remain retained.
Spread is flagged when max/min block metric exceeds 1.05, and the tail flag is
median process p99 divided by median process p50 above 1.05. If bootstrap
intervals are reported, they use 10,000 resamples, seed `819819`, and sorted
endpoints 250 and 9749.

Report source and published bytes/sec only when the corresponding workload
bytes are retained; do not call them physical I/O throughput or memory
bandwidth. Keep operation, sink, allocation, procfs, and RSS vectors separate
and label unavailable metrics explicitly. This packet makes no cold-cache,
network, physical-device, scaling, optimization, or before/after claim.

## Driver and completion gates

The quality driver runs six checks: formatting; offline locked all-feature,
all-target checking; the full offline locked all-feature `cargo test` command
without `--all-targets` (so Cargo's normal test selection includes doctests);
warning-denied all-target Clippy; warning-denied rustdoc; and the
crate-boundary checker against the frozen harness lock. The separate
all-target check and Clippy commands cover the broader target set, while the
test command retains doctest execution. The release driver
builds feature-off native, the untimed exporter, and the combined-feature
observer serially with the explicit release profile in `plan.json`. It retains
failed attempts, commands, environments, logs, binary hashes, and source
witnesses.

Every driver rechecks production and harness hashes, both lock identities,
the architecture inputs, corpus/provenance, host descriptors, packet/driver
hashes, and unrelated workspace files around each child. Capture cannot open
qualification or timing handles before fresh artifact and qualification
admission. Owned target and scratch directories are removed only after
validation and independent review, with the cleanup witness retained for the
final seal.
