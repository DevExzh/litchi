# 0362ul87aa — OOXML properties replay and core-only publication repair

Status: source-only follow-up to 0362ul87a, based on its committed source.
No Cargo, rustdoc, native qualification, performance measurement or final gate
was run in this follow-up.

## Repairs

The shared properties `Patch::apply` now publishes a
`litchi_core::edit::Commit`, so every eager OOXML owner consumes the returned
package rather than mutating and then discarding an intermediate clone:

* XLSX `metadata::Package::apply_metadata` adopts
  `patch.apply(self.opc())?.into_package()`;
* PPTX `metadata::Package::apply_metadata` wraps
  `patch.apply(self.opc.as_ref())?.into_package()`;
* XLSB `metadata::Edit::commit` and `metadata::Package::apply_metadata` both
  adopt the returned package before rebuilding their owner projection.

The existing public metadata tests now exercise changed replay, inverse replay
and exact no-op replay for these three adapters; XLSB has an explicit no-op
byte-copy regression alongside its changed replay test.

## DOCX core-only scope

DOCX's `metadata::Package` historically published only core fields through
the private core writer while application/custom changes were replayed by a
separate properties patch. Calling the unified `set_metadata` with that
core-only value would make absent application fields look like deletions and
could conflict with the later application/custom patch. The shared edit now
provides `set_core_metadata`, which validates that the input contains only
core fields and merges those fields onto the staged/source application
projection. DOCX uses this method for its core route; the clear route remains
an explicit core-family removal. Regressions cover an existing app part and a
combined DOCX core/app/custom edit with forward and inverse replay.

## Source-only checks

`git diff --check` and scoped `rustfmt --edition 2024 --check` pass for the
follow-up Rust sources. No build or runtime result is claimed here; the
coordinator should run the affected common/DOCX/PPTX/XLSX/XLSB tests after
review.
