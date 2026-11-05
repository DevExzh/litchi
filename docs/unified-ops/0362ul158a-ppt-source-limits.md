# 0362ul158a — PPT source-limit propagation repair

Date: 2026-10-25. Parent: 0362ul158. Status: source changes are ready for
coordinator verification; no Cargo, rustc, native, archive, or gate command
was run in this follow-up.

## Source binding

This leaf is `uo2/ppt-hyperlink-crud-repair`, based on frozen C
`f8cd398a770f7363f628be40edb3ce4219566c16`. The preceding production repair
is `72baeeb13d22358e578e262d11b81bda3b5ccbae`; the honest test/record
follow-up is `d7e03720a7`. This record is the source-only continuation of that
repair and must be reviewed and run through the coordinator's approved
`run_safe_gate.run_step` before acceptance.

The return audit is
`(deleted-coord-artifact)/README.md`,
SHA-256
`d02439f327440abd75a3f55fa81dfac4012fc14bd098e660b86ec19cf2bad1b4`.
It identified a unit mismatch in the OfficeArt adapter, dropped payload and
event ceilings, and neutral-default reparses in presentation, notes,
picture-bullet, and text projections.

## Repair

`odraw::limits_for_source` now maps the PPT encoded record ceiling to an
OfficeArt body ceiling using the stricter of both source dimensions:

```text
min(max_record_bytes.saturating_sub(8), max_record_payload_bytes)
```

The result is clamped to OfficeArt's integer width. Source depth remains
bounded by the existing public OfficeArt safety policy, and source events map
to OfficeArt record events. `RawRecords` and every `visit_host_record` caller
now receive that caller-selected event ceiling instead of an unconditional
one-million fallback.

The retained source profile now reaches:

- presentation drawing text boxes, shape/tolerance projections, master and
  slide header/footer projections, and optimized/full text extraction;
- lazy speaker-notes shape, placeholder-ID, and body-text projections;
- picture-bullet OfficeArt validation and source reference collection;
- slide and notes shape conversion, including detached `ClientTextbox` and
  outline-body reparses;
- incremental text-edit readback of a changed shape.

The public raw `odraw::text_from_drawing` entry point now requires a
`SourceLimits` argument. All repository consumers and tests were migrated in
this change; no defaulted raw-byte source path remains. Existing public
wrapper APIs retain their established behavior, while source-backed callers
use an internal source-bound constructor.

## Regression coverage

The OfficeArt limit test now parses a real `DgContainer`/`SpContainer` shape
tree rather than empty input. It checks the default hard depth intersection,
an allowed caller depth of 64, an explicit caller depth failure on nested
groups, caller event rejection, independent payload rejection, and the
encoded-header/body-unit boundary. The embedded PPT text-record test checks
that a caller event limit rejects a second record.

The existing 0294 hyperlink fidelity regression remains unchanged and still
covers hyperlink retention, chart and text fidelity, and exact inverse bytes.
The source-only changes add no capability waiver, public limit alias, or
`3rdparty/` fixture. They perform scalar limit propagation and do not add a
new traversal to a hot path.

## Verification boundary

The following checks are source-only observations from this leaf:

```text
rustfmt --edition 2024 --check <changed Rust files>
git diff --check
```

Both completed successfully. No compiled test result is claimed here. Root
must bind this exact committed source and run, at minimum, the focused
`litchi-ppt` OfficeArt limit test, the text/notes/picture-bullet restrictive
tests, the 0294 hyperlink fidelity test, and the relevant PPT CRUD/gate
checks via the approved capped runner. Any resulting logs, process receipts,
target bounds, and source hashes remain external; this record
does not add binaries or large evidence.
