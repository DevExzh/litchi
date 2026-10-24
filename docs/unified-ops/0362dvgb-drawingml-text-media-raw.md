# 0362dvgb — DrawingML raw adapter review repairs

- Date: 2026-10-24.
- Parent: 0362dvga, 0362dvg, 0362dv.
- Status: source repairs complete; coordinator validation is pending rescheduling.

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
writer can publish output. No Cargo command was run while this repair was
source-only; rustfmt and `git diff --check` are the only checks performed so
far. No generated artifacts, binaries, or third-party paths are tracked.
