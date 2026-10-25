# 0362ul62g — PPTX replay map live-state correction

This return tightens the persistent replay index from 0362ul62f.

The immutable AVL map now supports path-copied deletion with successor
extraction and rebalancing. Notes entries, notes-theme users, media reverse
edges, and missing-theme witnesses are stored as live keys only: removing a
notes part, clearing its last theme user, removing the last media slide, or
repairing a missing theme physically removes that key instead of retaining a
tombstone or empty value. The missing-theme map is therefore a live-only
ordered witness; its minimum key is found in AVL height time and no historical
key scan is needed during replay updates.

Private map tests cover ascending, descending, and zig-zag insertions;
replacement; leaf, one-child, two-child, root, and absent deletion; stored
height and balance invariants after each operation; old-root persistence; and
the distinction between an absent key and a present `Option::None` value.
Domain regressions cover notes-master removal and clearing a malformed
missing-theme relationship. The duplicate media `HashSet` import was removed
so the shared module import is the sole binding.

Validation for this source-only return is limited to changed-file
`rustfmt --edition 2024` and `git diff --check`; Cargo builds/tests remain for
the coordinator's composed validation slot.
