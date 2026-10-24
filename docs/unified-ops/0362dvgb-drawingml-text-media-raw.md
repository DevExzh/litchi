# 0362dvgb — DrawingML raw adapter review repairs

- Date: 2026-10-24.
- Parent: 0362dvga, 0362dvg, 0362dv.
- Status: source repairs and focused validation complete; coordinator review pending.

The second independent review found two adapter-boundary issues:

- PPTX `model3d::codec::write` no longer deep-clones the scene wire before
  asking the shared writer to validate structure and output size. The shared
  writer already emits missing model and relationship namespace declarations,
  so the PPTX adapter now passes the borrowed wire directly. This preserves
  the previous emitted XML while ensuring retained opaque children are not
  copied before an immediate output refusal and avoids a redundant candidate
  buffer.
- PPTX `prepare_fragment` now maps source and completed-preparation byte
  refusals to the shared typed `DrawingML::Error::Limit` boundary, retaining
  `LimitKind::XmlBytes`, the observed source/prepared length, and the active
  adapter maximum. The public error remains content-free at display time.

Regression coverage includes typed source and prepared-size refusals and a
large retained opaque scene whose narrow output policy is refused before the
writer can publish output. The source-only packet was checked with rustfmt and
`git diff --check`; no generated artifacts, binaries, or third-party paths are
tracked.

The follow-up review found that the no-clone path suppressed the shared
writer's hard-coded `m3d:` and `r:` declarations whenever an alternate prefix
bound the same URI. The writer now checks the prefix and URI that it actually
emits, retains all existing declarations, and injects only missing bindings.
The same writer path is used by output-size preflight, so the injected bytes
are counted before an output refusal. PPTX regressions cover successful
alternate-prefix output and reparsing, plus a one-byte-narrower policy that
confirms the complete emitted size is refused.

Validation receipt for the corrected source tree (parent `8ce87b6143`, before
the follow-up commit), using `Cargo.lock` SHA-256
`52ed0cacac15b5e440349bfc0f4aca2665bcf50d59d4321624e596bb410783c6`,
`--locked --offline`, `CARGO_BUILD_JOBS=1`, the coordinator cargo shim, and a
20 GiB memory cap:

- `cargo test --lib -p litchi-pptx 'model3d::tests::codec_'`: 6 passed.
- `cargo test --lib -p litchi-drawingml`: 241 passed.
- `cargo test -p litchi-drawingml --test svg_blip`: 24 passed.
- `cargo check -p litchi-spreadsheet-drawing -p litchi-docx -p litchi-pptx -p litchi-xlsx`: passed.

The external generated target was cleaned after validation; no generated
target files are part of the commit.

The final review then found a valid collision case where a retained foreign
`m3d:` or `r:` declaration prevented the canonical spelling from being
redeclared safely. The writer now builds one bounded prefix plan before size
preflight: it reuses a safe existing nonempty model/relationship prefix, or
selects a fresh prefix that is absent from root, raster, and blip declarations.
The plan is propagated to model3d roots, raster/blip names, relationship
attributes, closing tags, and both counting and streaming writers. Retained
opaque bytes and their foreign bindings remain unchanged.

Final collision coverage includes PPTX round-trip fixtures with foreign
`m3d:`/`r:` bindings and opaque children, plus a shared-codec default-namespace
fixture that selects fresh prefixes and verifies the narrow output admission
boundary. With the same lock and capped offline setup:

- `cargo test --lib -p litchi-pptx 'model3d::tests::codec_'`: 7 passed.
- `cargo test --lib -p litchi-drawingml`: 242 passed.
- `cargo test -p litchi-drawingml --test svg_blip`: 24 passed.
- `cargo check -p litchi-spreadsheet-drawing -p litchi-docx -p litchi-pptx -p litchi-xlsx`: passed.
