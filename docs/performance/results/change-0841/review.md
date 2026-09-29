# 0841 diagnostic review

## Disposition and frozen scope

This packet is an attribution diagnostic for the rejected 0782
`Option<&str>` PPT candidate. It is based on current HEAD
`307434b454f622023818f86f724b56ae5c99ee15` and does not authorize production
adoption. The archived candidate patch is applied only to produce a matched
diagnostic binary and is restored after capture, regardless of the result.

The scope correction is resolved in the frozen plan:

- the native lane retains all ten 0782 write/lifecycle cases, six paired
  blocks, and 30 measured samples after three warmups (`120` reports and
  `3,600` samples);
- the Callgrind lane has `16` profiles: `many`, `unicode`, `payload`, and
  `rich`, write mode only, two independent one-operation repeats per leg
  (`before` and `after`);
- there is no three-operation-per-process claim. Each profile process has one
  `profile_write` owner invocation and one expected Callgrind publication;
- qualification remains the ten before/after public cases. The profile lane
  is diagnostic and is not a CI adoption gate.

`profile_write` is a shared `#[inline(never)]` probe helper around
`Writer::write_to(&mut Cursor<Vec<u8>>)`. Its result and destination are
barriered, while writer/sink destruction, input generation, readback, and
verification stay outside the owner. Callgrind Ir is guest instruction
accounting; it is not a native-cycle, wall-time, page-fault, or physical-copy
measurement. A profile is useful only when it has one positive owner
publication, a non-empty nested child, and no setup/readback rows in the
collected region.

## Exact 0782 source parity

The candidate patch is byte-identical to
`docs/performance/results/change-0782/candidate/applied-model.patch`:

```text
bytes:   22105
sha256:  b9ea9b31f2daff2513aeb9e735324d6aab81671f1a7385a9a011abb929136d59
```

`candidate-source-parity.json` reports `all_equal: true` for every affected
file. These are the exact current and historical before-source SHA-256 values
used for the twelve affected files:

| file | current = historical SHA-256 |
| --- | --- |
| `crates/litchi-ppt/src/writer/core/codec.rs` | `2e4845b5c386de6df236494f0bc4f29756b7da1b69b0cffda1acaaa2d8d09a98` |
| `crates/litchi-ppt/src/writer/core/package.rs` | `39cc8a3d63e4b9d7eba53fb2b0150f2c05fa8f0eeda5e7125e73ac5bc91a57a9` |
| `crates/litchi-ppt/src/writer/core/tests.rs` | `b680781d140706b05136d018b527c4d834f323f25dcd47266802c4d2a2e5b62b` |
| `crates/litchi-ppt/src/writer/escher/codec/drawing.rs` | `43c6674baf53e6dc69947fc6cbd1f4d91c08b7d7d34adfc08bf4e993a84e7ff2` |
| `crates/litchi-ppt/src/writer/escher/codec/group.rs` | `68b64bf79d7a05a6b1e172b52829d3c70227e8d642cfda010b86f7ea452964f0` |
| `crates/litchi-ppt/src/writer/escher/codec/properties.rs` | `d2142b311e3e2d35af64a28ef300c9b077e74347d8cdd98b058bfc4b59b12a39` |
| `crates/litchi-ppt/src/writer/escher/codec/shapes.rs` | `35c64874c970df3b26738741ef464aa6a68015fc90c21b427684e684f8ee48cf` |
| `crates/litchi-ppt/src/writer/escher/codec/validation.rs` | `49e4dd10845a56491897e87e6fbf6732cea3a537dbda77a805b0877107ce7c4c` |
| `crates/litchi-ppt/src/writer/escher/model.rs` | `ad9f6a906d96a424ec671998b08d8f43c0e46cc097af71254785a6c9e14b5263` |
| `crates/litchi-ppt/src/writer/escher/semantic.rs` | `3690925f3ce466ecf32787c7a51076f6d95d539964e0905e65b16ec32333b018` |
| `crates/litchi-ppt/src/writer/escher/tests/groups.rs` | `9c24b134f085ca9ebbfdd0828d98a56c4811bd4a5ca5076676ab3e3afaa5412d` |
| `crates/litchi-ppt/src/writer/escher/tests/legacy.rs` | `123e5a026180db41085f70d0b337a53e3348502a5d9d5b350327e8f207089346` |

## Historical reason for attribution

The 0782 borrowed candidate failed the frozen latency guard in four cases:

| case | after/before p50 change | bootstrap ratio interval |
| --- | ---: | ---: |
| `many/write` | `+12.897%` | `[1.110164, 1.135918]` |
| `many/lifecycle` | `+11.089%` | `[1.108498, 1.123324]` |
| `unicode/write` | `+20.650%` | `[1.202215, 1.209476]` |
| `unicode/lifecycle` | `+19.069%` | `[1.189756, 1.192544]` |

The payload cases improved by roughly 22–24%, while tiny and rich were near
controls. Heaptrack removed the expected plain conversion clone ancestry, but
whole-process allocation and `perf stat` observers did not establish an
operation-local cause: instructions were about `-0.046%` while cycles rose by
roughly 11–14%, and live/peak allocation guards did not improve. This pattern
requires attribution before another representation change is considered.

## Operations to attribute

The candidate target is
`crates/litchi-ppt/src/writer/core/codec.rs:190-307`, especially the
`props.text.clone()` conversion at line 306. The relevant write chain is:

1. `core/package.rs:759`, `Writer::write_to`;
2. `escher/codec/drawing.rs:146-212`, drawing/container traversal;
3. `escher/codec/shapes.rs:84-297`, shape validation, property assembly,
   client-data parsing, and plain/rich text dispatch;
4. `escher/codec/properties.rs:8`, `build_shape_properties`, including its
   per-shape vector and sort;
5. `escher/codec/text.rs:39-106`, plain text encoding, and
   `text.rs:116-171`, the unchanged rich-text control;
6. `escher/codec/validation.rs:17`, if it remains as a separately emitted
   function.

The symbol inventory should query `append_dg_container_children`,
`build_shape_properties`, and
`build_client_textbox_formatted_with_interactions` in addition to the current
conversion, plain-text, shape-child, and validation queries. A missing symbol
means inlining or LTO and must be reported as unknown rather than zero work.
Caller-separated Callgrind annotation and DWARF inline stacks are preferable
when the compiler folds these helpers into their callers. Inclusive descendants
overlap and must not be summed as independent phases.

No standalone struct-layout instrumentation is planned: adding it to the
production diagnostic probe would change the measured executable. If the
owner Ir split identifies a boundary that needs explanation, capture assembly
from the same admitted binaries for the selected top functions. Compare
prologues, stack spills, loads/stores, branches, and clone/allocation/drop or
memcpy calls under the same compiler, target, LTO, and codegen-unit settings.
Any layout or prologue explanation remains a hypothesis until a subsequent
operation-scoped native measurement confirms it.

## Follow-on OLE2/OOXML opportunity

DOC currently has no obvious intermediate-copy elimination: `TextStream` writes
into its final buffer, DOC stream vectors are moved into CFB, and CFB emits
large stream slices directly. The remaining first-touch/page-fault path should
be profiled before proposing an unsafe or spare-capacity rewrite.

The higher-impact production discovery lane is opened-package save. In
`crates/litchi-opc/src/source_backed.rs:11721-11935`,
`write_changed_overlays_with_appended_inner` allocates scratch, rebuilds a
`PreservationIndex`, scans entries for replacement/omission decisions, scans
again to build a `PreservationPlan`, and then emits raw copies and regenerated
members. A one-edit source-backed profile should separate those phases before
considering bounded reuse of immutable central metadata. Exact no-op raw-source
passthrough is already implemented and is not a new opportunity.

That lane remains constrained by ADR0005 budgets and explicit I/O, ADR0006
exact preservation/refusal behavior, ADR0030 source freshness and lazy decode,
and ADR0032's restriction against an unbounded parsed-index memo. No such
candidate is implied by this PPT diagnostic.

## Review status

This review records source and packet scope only. It does not pool historical
0782 timings with current 0841 measurements, does not claim a causal result
from Callgrind alone, and does not authorize adoption. The production source
restore receipt remains mandatory after all diagnostic captures.
