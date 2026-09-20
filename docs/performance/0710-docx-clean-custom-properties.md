# 0710: preserve untouched DOCX custom properties

`performance_claim: none`. This is a preservation correction prompted by the
[0709 baseline](0709-docx-ordinary-save-baseline.md), not a measured speedup.

A normal save previously called `custom_props.write_for` even when the custom
properties had not been edited. For the checked-in `alt-chunk-header.docx`,
that converted a valid 632-byte custom XML member into a 602-byte canonical
serialization. The same rewrite occurred after a refused document edit and
when no edit was attempted. The expanded XML structure was equivalent, but
namespace spelling and other lexical details were lost.

The DOCX writer now serializes custom properties only when
`custom_props_dirty` is set. Opening a package and committing a raw OPC edit
already validate and load the custom-property model. The mutable accessor
marks it dirty; explicit edits still take the existing host validation,
package-graph validation, creation, and removal path. The flag is cleared only
after complete stream publication succeeds, preserving failed-write retries.
The generic OPC XML publication audits remain in place. As before, calling
`custom_props_mut()` records edit intent even if the caller makes no value
change; this patch does not introduce semantic no-op tracking for that accessor.

Five focused regression tests cover clean and refused-edit saves through both
public routes, dirty-property retry after a partial sink failure, a raw OPC
replacement, and an empty custom-properties part preserved until explicit
clear. On the unmodified writer, four preservation tests fail and the retry
control passes; all three pre-existing tests pass. The candidate passes all
1,453 DOCX tests with all features and targets, formatting, and warning-denied
Clippy. The unchanged strict oracle from 0709 now passes. All four clean/refused-edit
outputs across `save` and `to_stream` are identical to the 77,621-byte source
archive (SHA-256 `ee33b932a6c31f430bc8a6e450592970419b0bcc155e5c9884065602b9c733b8`).
An independent Python ZIP check confirms every decoded member is unchanged.
The admitted fixture still passes its paragraph-and-marker checks, with the
same main-document and relationship changes recorded in 0709.

No native Office interoperability, full-package edit preservation, latency,
allocation, RSS, or throughput claim is made. In particular, the admitted
paragraph-edit oracle still does not certify unrelated relationship rewriting.

[Evidence packet](results/change-0710/README.md).
