# 0819 source and scope review

The frozen base is `0c784f7aecda2a38bbe6cc98a684ce68da1f0dc8`. Production and
the tracked `tools/perf-baseline` harness are unchanged for this batch; the
DOCX relationship-preservation repair and its regression tests are already in
the base. No dependency, Cargo lock, runtime-harness, or benchmark-source
change is introduced by the 0819 packet.

The 0818 repair in `crates/litchi-docx/src/package/codec.rs` replaces an
existing main document part in place and keeps its package-level relationship
provenance. The missing-main-part path still uses the ordinary add path. This
allows an ordinary paragraph edit whose relationship binding is unchanged to
reuse the producer's exact relationship member, including order and lexical
bytes. The committed regression covers both exact untouched relationship
bytes and a changed external hyperlink relationship that must be serialized.
That repair is a preservation prerequisite for this batch, not a performance
change being measured against an earlier implementation.

## Reviewed ordinary-save path

`tools/perf-baseline/src/ordinary_save.rs` classifies caller-named OOXML files
from their main package member, bounds the input, stages it once in an owned
workspace, and derives the format-specific edit target from the opened file:

- DOCX appends one ordinary paragraph containing the fixed harness marker;
- XLSX edits `A1` on the first worksheet and commits the resulting edit; and
- PPTX probes the first admitted slide/shape position and replaces that shape's
  text, retaining a typed refusal if the opened deck admits no position.

The corpus builder probes the edit before any samples, freezes the exact
outcome, performs two fresh open/edit/save cycles, and saves one edited owner
twice. It refuses nondeterministic output or a changing edit outcome. These
probes are untimed. They prove a stable workload identity; they do not replace
the independent ZIP/XML preservation admission.

`run_case_with_durability` has four distinct brackets. It opens and edits
outside the `atomic_publish` and `counting_publish` clocks; `edit` has no
publication call in its loop. For path saves, an absent durability option
calls the documented public `save` method. The three format implementations
bind that method to `Durability::Full`, so the ordinary baseline includes the
temporary-file data sync and parent-directory sync. The explicit reduced
durability policies are exporter controls only. Owner destruction, file reads,
hashing, byte-split calculation, semantic checks, and cleanup occur outside
the timed path.

The counting sink reserves a four-times-output budget before the clock and
retains bytes for post-clock accounting. DOCX and XLSX write sequentially to
the sink. The PPTX branch first calls `Package::to_bytes()` and then writes
that buffer to the sink, so its result is not evidence of streaming or a
bounded-memory serializer. The reported byte split is derived from source and
published ZIP members and is explicitly an upper-bound comparison, not a
production copy-through counter.

## Instrumentation and report identity

The native release binary is built with no Cargo features. It therefore has no
allocator or ordinary-save procfs probe and its operation allocation metric is
unavailable or absent as defined by the report schema. The observer release
binary enables exactly `allocator-metrics` and
`ordinary-save-process-metrics`; its allocation region brackets the same
operation and its procfs snapshots are additive diagnostic observations. The
observer's 32 empty controls are retained without subtraction. The
feature-dependent instrumentation identity is an existing harness contract;
the packet does not alter it.

The exporter uses the same corpus construction, public editor, save, and
sequential serialization paths as the timed selectors. It writes the source
plus `default`, `full`, `file-only`, `no-sync`, and `stream` outputs for three
generated controls and three real files. The independent artifact auditor and
the separate ZIP preservation replay must pass before any qualification or
timing output is admissible. Generated controls do not enter the twelve-case
real-file matrix.

## Scope limits

The source and protocol preserve the 0817 lock graph, three-input corpus, and
real-file route while carrying forward the 0818 DOCX relationship repair. The
historical provenance record is not a current Office resave and cannot support
a new producer identity or an external round-trip claim. Native and observer
latencies are separate; lifecycle, edit, atomic-save, and counting values are
separate observations and must not be subtracted or pooled. No result may be
described as a physical-I/O, memory-bandwidth, cold-cache, network, scaling,
or optimization claim without additional evidence.
