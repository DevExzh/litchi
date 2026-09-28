# 0804 bounded linear equality control

Before is the exact 0803 after-control. After changes only the bounded linear
comparison call and adds an inline byte-slice equality helper. The helper
increments the existing test-only comparison counter once per visited name.
Name equality/ordering traits and the ordered backend remain unchanged.
All five copies have the same normalized implementation change. Shared tests
are byte-identical. Empty-check placement, raw key/error preflight, offsets,
clone/fusion and the bounded 32-name handoff remain fixed.

Neither leg is production. No workflow advancement/adoption follows from this
diagnostic, and no historical timing is pooled. Compiler/code-layout effects
remain associated with this source intervention. The patch is archive-relative
and applies only to candidate/before. Metadata was completed after source-only
quality capture and before release build; see execution-note.txt.
