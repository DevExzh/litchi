# 0798 source-only review

Review basis: the frozen 0798 plan and README, the before/after helper
archive, the instrumentation design and manifest, the isolated canonical and
census tests, the direct probe, `root_audit.py`, `analyze.py`, `validate.py`,
and the recorded source route. This review was source-only; it did not build,
test, capture, profile, or replay anything.

## Disposition

No source-level blocker was found for the root-owned diagnostic packet. The
after helper preserves the checked iterator's parser paths and adds a
calling-thread census with explicit lifecycle and overflow failure states.
The probe boundaries and aggregation are suitable for the stated diagnostic
scope. This review does not authorize production adoption, a latency or
allocation claim, or coverage beyond the generated PPTX operations named in
the protocol.

## Instrumentation semantics

The before module is the exact OPC owner at the frozen base commit. In the
after module, `CheckedAttributes::next` still calls the underlying
`Attributes::next` once, performs the same phase transitions, and returns the
same item. The census observes the returned `Option<Result<...>>` after that
operation; it does not inspect, reorder, replace, or recover an item. The
`Phase::QuickXml`, `Own`, and `Done` paths, the first-error stop, and repeated
terminal `None` behavior remain in the parser code. `FusedIterator` is still
implemented.

The manual `Clone` implementation copies the original `tag`, quick-xml
iterator, and parser phase exactly as the former derived implementation did.
Its only additional action is to allocate a diagnostic row when the source
token is valid. `Drop` performs bookkeeping and a separate unchecked scan;
the scan cannot change the checked iterator's state or returned values. With
the census disabled, the hook checks only a thread-local enabled bit and does
not copy the tag, scan attributes, or allocate a registry row. The diagnostic
build therefore has intentional hook and drop costs, and its elapsed fields
must remain opaque.

The 32-name handoff is not altered by the hook. The first 32 calls stay on
quick-xml's checked iterator; the later calls enter the existing ordered-map
path. The census call is present on both returns from `next`, so a terminal
`None` or an error is recorded without changing which parser branch produced
it.

## Clone, partial, and error accounting

Each constructor starts a row with a fresh instance and lineage ID. A clone
gets a new row in the same lineage, resets local event counters, and carries
the source's successful prefix in `starting_successful_yields`. The prefix is
the source prefix plus its local successful yields, so a clone made after
partial progress is not misreported as having consumed the entire tag. The
probe validates unique IDs, one origin per lineage, matching source and
element bytes, and prefix bounds before rows are aggregated.

Drop-time lexical counts use `unchecked_attributes()` from the beginning of
the copied tag and stop at the first lexical error. Repeated names therefore
count as lexical attributes, while checked iteration still stops at its first
duplicate or other error. `partial_consumption` compares successful checked
yields plus a clone prefix with the complete lexical attribute prefix;
`early_drop` is set only when no error or end result was observed. An iterator
that returns an error and is then dropped remains `Error` with `early_drop`
false, while a valid tail found by the unchecked scan still makes partial
consumption explicit. Empty quoted values and trailing XML whitespace are
handled by quick-xml's unchecked iterator and by the retained parser tests.

The session tracks starts, clones, drops, live rows, events, IDs, lexical
counts, and aggregation frequencies with checked increments that saturate at
`u64::MAX` and retain a diagnostic flag. `finish` disables the registry first,
marks undropped rows as `live_at_finish`, and never infers a drop from the
absence of a later hook. Stale tokens from a replaced session are rejected by
generation; a token dropped on another thread cannot mutate the source
thread's registry. These rules make the conservation equation

```text
iterator_starts + iterator_clones
    = iterator_drops + live_instances_at_finish
```

explicit and fail closed when it does not hold. A generation counter reaching
its saturation point is also diagnostic failure rather than silent ID reuse.

## Caller-thread and lifetime scope

The instrumented owner is `litchi-opc::xml_attributes::CheckedAttributes`.
The OOXML-common re-export therefore routes its PPTX callers to this owner,
including notes validation. That caller skips iterator construction when the
raw attribute tail is empty; those skipped tails cannot be reconstructed from
zero-count rows. The census includes every owner invocation on the probe's
calling thread during the enabled region, regardless of Rust call site. It
does not include OLE, signing, XLDM, or XML-minifier copies, unchecked or
lenient iterators, or events on other threads. Element-name and raw-tag bytes
help group observations but do not identify a unique caller.

The probe starts the census immediately before each public operation and
finishes it after stopping the clock and finishing allocation bookkeeping.
Capture measures `opened_presentation`; commit measures `commit` after one
staged edit; lifecycle measures capture, edit, commit, publication, and
serialization. Ingress, fixture creation, output serialization for capture
and commit, reopening, and verification stay outside their regions. The
scope guard finishes and discards an incomplete report on error, preventing a
stale enabled session from contaminating a later sample.

The operation locals remain in scope while allocation and census reports are
finished. This is a deliberate lifetime check: if an iterator escaped the
public call, it is retained as `live_at_finish` and qualification fails rather
than being counted as dropped. The protocol therefore measures only owners
that complete inside the public operation; it does not claim global lifetime
accounting for an iterator moved away and dropped on a foreign thread.

## Probe aggregation and independent audit

The probe first validates raw instance and lineage identities, then groups by
all observable row fields except instance and lineage IDs. Frequencies and
exact source/name bytes are retained, so grouping cannot merge different
consumption states. It separately checks row-frequency conservation and
start/clone/drop/live conservation before serializing the aggregate.

`analyze.py` independently parses each generated tag, checks lexical counts,
next-result conservation, terminal classification, drop and live flags, and
repeat equality. `root_audit.py` independently reconstructs the valid quoted
attribute counts, checks every event and lifecycle total, compares the two
census blocks, and compares control semantic/publication identities with the
sealed 0794 records. Its lexical oracle is intentionally limited to generated
well-formed tag bodies and fails on a row outside that grammar; it is not a
general malformed-XML oracle. The audit excludes only `elapsed_ns`, as the
frozen protocol requires, while retaining the other semantic metrics.

## Tests and protocol limits

The isolated helper suite covers differential parser behavior around the
32-name transition, duplicate and malformed values, terminal behavior, clone
prefixes, no-consumption and partial drops, stale sessions, live-at-finish,
foreign-thread isolation, and saturation. The canonical copy-parity test is
deliberately skipped in the OPC-only mirror because the other helper copies
are not instrumented; that is appropriate for this diagnostic archive but
does not constitute a full-workspace regression run.

The frozen protocol keeps 15 plain controls and two 15-case census repeats,
with one sample, no warmup, serial CPU-12 execution, exact source/binary
receipts, and no production adoption. Production restoration and the
source-allowlist witness remain required custody gates. The resulting census
can describe checked-iterator counts and consumption for these generated
PPTX regions; it cannot estimate workflow speedup, allocator calls, native
latency, cold/range behavior, concurrent behavior, other producers, or the
skipped empty-tail population.
