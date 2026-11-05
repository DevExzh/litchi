# 0362ul67a — PPT VBA owner-policy correction

This return packet is based on `b72afa8303` from the VBA raw-policy packet
`0362ul67`. It closes the remaining legacy PPT VBA boundary found during
source review.

## Owner classification

`litchi_ppt::Presentation` is an immutable read-only projection. Its canonical
`Presentation::vba()` operation therefore inherits the `VbaProjectLimits`
selected on `OpenOptions` when the projection is constructed and retains that
policy in the projection. The public `vba_with` twin is removed; callers that
need a different profile construct a new projection through
`Reader::presentation_with_options` or the corresponding source/package
options path. No public mutable owner API is introduced.

The eager `raw::Reader` and positional `SourceBackedPackage` retain the
selected VBA profile so their no-argument `presentation()` projections do not
silently revert to defaults. The projection exposes the retained profile for
policy diagnostics, while VBA parsing remains inert and never executes macro
code.

## Nested CFB boundary

PPT's outer `SourceLimits` continues to govern the PowerPoint CFB. Only the
decompressed `VbaProjectStg` payload receives the derived inner
`OleFileLimits`: its input, directory, and allocation-table ceilings are
bounded by the selected VBA CFB budget (the stricter outer/project CFB ceiling)
before nested CFB parsing. The outer storage checks the selected stored and
decompressed budgets before allocation, and the inner project parser keeps its
typed MS-OVBA limits. A regression test proves the derived directory and
allocation ceilings follow the selected profile.

## Consumer and glossary closure

PPT VBA writer tests now create projections with `OpenOptions::vba_project_limits`
and call the canonical `vba()`. They verify retained owner policy and both
stored-payload and decompressed-CFB refusals. The DOC and XLS raw reader
`vba(&mut self, &litchi_vba::Limits)` entries are registered in their minimal
and full glossary profiles, while PPT full registers the options setter and
bans `Presentation::vba_with`.

No Cargo, native, rustdoc, benchmark, or macro-execution command was run in
this source-only correction. The coordinator should run the PPT feature
matrix, the DOCX encrypted gate, the glossary/rustdoc profiles, and the full
consumer gate after merging.
