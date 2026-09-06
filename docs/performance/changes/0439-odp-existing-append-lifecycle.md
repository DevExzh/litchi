# 0439: Existing ODP append lifecycle baseline

The previous `odp_semantic_one_edit_save` timer opened its source before the
clock. The new opt-in `odp_existing_append_lifecycle` includes owned snapshot
opening, transaction creation, one title/body append, commit and sequential
output of the committed bytes. It establishes an existing-document append
baseline without changing production code.

The deterministic source has 64, 4,096, or 8,192 mixed ASCII/Unicode/XML-text
slides plus a 65,536-byte opaque member. The append uses the next indexed
fixture title/body. Untimed checks verify every source slide, exactly one tail
append, all six package members, manifest bindings, untouched decoded bytes
including the manifest, opaque compressed data-span equality, reversible
source-checked patches, and exact no-op shared bytes. These are owned writer
observations; compressed equality does not establish passthrough.

Input cloning, append argument creation and sink setup precede the clock.
Source and commit remain live at the allocator endpoint; oracles, digest
finalization and destruction follow it. The materialized committed archive is
then discarded by the sink. Neither bounded commit memory nor streaming save
is claimed. The report's generic sink `input_bytes` and `authored_part_bytes`
identify full output semantic text and content.xml, respectively.

See the [data path](../results/change-0439/data-path.md),
[design](../results/change-0439/design.md), and
[validation history](../results/change-0439/validation-notes.md).
The selector registry grows to 436; the default 36-case contract is unchanged.
The representative CRUD index has 15 categories and 33 mappings. Its new
correctness-only mapping does not promote an opt-in baseline into the index's
default-matrix measured status.

The [retained baseline](../results/change-0439/measurements.md) contains 12
reports and 360 samples over two repeats. Normal p50 is 2.075–2.081 ms for 64
source slides, 85.178–85.979 ms for 4,096, and 170.742–175.416 ms for 8,192.
The large case records 1,462,779 allocation calls, 236,704,188 cumulative
requested bytes, a 39,890,548-byte peak above entry, and an 8,074,341-byte
retained live delta. Both source and commit remain live at that endpoint.
No normal latency repeat crosses the 5% review trigger; one allocator p99
repeat does. This is a baseline, with no before/after speedup claim.

The final release harness suite passes 368 tests with one ignored; all 35
coverage-index tests, formatting, harness boundaries, and documentation pass.
Strict Clippy retains 29 existing diagnostics in 17 distinct groups, with no
new group or multiplicity. Two successful whole-process profiles retain raw
data and diagnostics, including 13 symbolization warnings in each text
conversion. The [evidence bundle](../results/change-0439/README.md) documents
portable verification and the retained failed setup attempts.

This covers one owned logical-append slice. Package-Part addition, arbitrary
repackaging, richer and native producer fixtures, source-backed/cold/range I/O,
scaling, and the broader non-iWork goal remain open. The shared ODF fragment
validation hypothesis remains future work.
