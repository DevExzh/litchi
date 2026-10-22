# 0362c — ODraw required-limit entry points

- Date: 2026-10-21.
- Parent: 0362b and the 0351 API contract.
- Status: constructor slice validated. This is one required constructor slice;
  the remaining ODraw owner, raw framing, error and builder work stays open.

The image record API now has one entry point for each operation:
`Blip::parse(data, Limits)`, `Blip::from_record(record, Limits)`,
`Entry::parse(record, Limits)`, `Store::parse(data, Limits)`,
`Store::from_record(record, Limits)` and `Delay::new(data, Limits)`.
The default-only twins and the configured suffix names are removed.
`TryFrom<Record>` for BLIP is removed because it silently selected defaults.

Shape traversal is `shape::parse(data, Limits)` and
`Shape::from_record(record, Limits)`. The latter replaces the implicit-policy
`TryFrom<Record>` and enforces the same hard recursion ceiling as `parse`.
`Container::find_recursive(kind, Limits)` is the sole recursive collector;
`Parser::shapes(Limits)` and `textboxes(Limits)` forward the selected policy.
All DOC, PPT, XLS and image-conversion consumers migrate in the same commit,
including tests and examples. Previously default-policy callers pass that
same policy explicitly. No compatibility aliases or deprecations remain.

Delayed-image resolution now honors both the entry's retained image bound and
the delay context's bound. Previously a broadly configured context could
resolve a BLIP larger than its entry allowed. The image-file helper also
constructs its delay context from the entry's policy. Regression coverage
uses a small generated record: a 17-byte payload ceiling rejects a 20-byte
BLIP body, including through a broader context, and the exact 20-byte ceiling
returns the original borrowed payload. The record size excludes its eight-byte
wire header, consistently with direct BLIP admission.
`max_store_entries` bounds store admission and delay iteration. Direct delay
offset lookup parses one block, applying the BLIP byte ceiling without
enumerating preceding blocks; its constant-time lookup is preserved.

Six configured constructor bodies are unchanged after normalizing renamed
calls and whitespace. Delayed resolution intersects entry and context bounds
before decoding, avoiding a second size check. Recursive collection keeps its
active cursor in a local variable and allocates a stack only for suspended
parents; this removes the initial allocation and its first growth. No image
or shape representation gains a field. BLIP selection reuses the instance
range check performed by the sole immutable entry constructor, instead of
repeating a fallible conversion. Boundary tests retain rejection at 256 and
4095 while admitting the full-byte value 255. Depth-first wire order, record counts
and depth ceilings are covered at exact boundaries. Shape-from-record coverage
checks record-budget rejection and rejects an unsafe configured depth before
recursion. Existing tests continue to cover
malformed framing, unknown records, host interpretation and byte preservation.

The new `odraw` glossary profile bans the four removed names and checks seven
read-only projections. The full glossary gate discovers this profile from the
policy; no separate manual gate invocation is needed after integration.

The initial complete five-crate consumer suite passed 5,141 tests across 232
targets (six existing ignored tests). After the traversal and byte-policy repair,
282 ODraw/image-conversion tests and 245 host shape tests passed; the final
admission reuse passed all 82 ODraw tests. All-target/all-feature Clippy for
the five crates, the locked performance harness, glossary, doctests (eight
passed, one existing ignored), formatting, boundaries, 36 tooling tests and
the benchmark matrix pass. The receipt distinguishes each validation stage.

Final native ABBA results use one-second samples on one pinned CPU with all
workers held and a clear process audit. Three blocks cover every route; nine
additional blocks resolve the initially variable delayed/shape-record results.
The table includes every final block, including the initial three for those
two routes; no sample is dropped.

| Route | Blocks | Median after/before | Block range |
| --- | ---: | ---: | ---: |
| Small BLIP | 3 | 1.004 | 0.936–1.005 |
| 2 MiB BLIP | 3 | 0.999 | 0.999–1.021 |
| Delayed BLIP | 12 | 1.009 | 0.955–1.061 |
| Shape parse | 3 | 0.990 | 0.990–1.006 |
| Shape from record | 12 | 0.995 | 0.877–1.046 |
| Recursive traversal | 3 | 0.965 | 0.941–0.966 |

The focused nine-block medians are 0.999 for delayed resolution and 0.992
for shape-from-record, so the earlier small increases do not repeat. Delayed
resolution's Callgrind count differs by only 20 instructions over 100,000
operations; shape-from-record drops from 101,917,258 to 100,717,148 instructions.
Recursive traversal improves in every final block. Borrowed payload
pointers, unknown shape records, checksums and representation sizes agree.
These generated-record probes cover this constructor slice; the final workspace
gate remains separate.

`results/0362c/odraw-validation.json` records commands, validation stages,
source and binary digests, every measurement and external artifact provenance.
Initial candidates and the cursor-repair intermediate remain recorded alongside
the final runs. The probe source and runner are kept there; generated crates,
binaries, source archives and raw logs remain outside the checkout.

Required next work remains explicit: raw framing/iteration admission; the
image-file scanning caps and owned-file reparse policy; typed error fields;
`prop::picture::Commit` still publishes a Snapshot and must migrate to the
shared triad with its owner and consumers (three remaining inventory rows);
write/draft namespace review; retained host OpenOptions/EditOptions policy;
and the rest of the 0362b constructor sweep. This slice does not flip their
inventory or glossary status and creates no accepted exception. W4, W4b and W8
integration and the final workspace gate remain required.
