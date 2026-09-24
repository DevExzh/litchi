# DOC OLE2 stream adapters

DOC managed embedded objects expose typed OLE2 presentation and native streams through storage IDs or DOC references, using the bounded common stream owner. Edits and patches validate selected streams before cloning, validate the candidate's DOC references, and publish atomically. Supplied reference names and numeric IDs must agree with a unique managed ObjectPool target; ambiguous numeric spellings are refused.

The isolated Rust 1.95.0 gate passed 1,175 tests and 14 doctests, strict all-target/all-feature Clippy, changed-file formatting, and whitespace checks. Two existing external-fixture tests and 12 doctests remain ignored. Synthetic regressions cover tight/stale patches, failure atomicity, forged and ambiguous references, and changing then restoring both streams within one DOC transaction: the final patch is a no-op, whole-document bytes are exact, and the source allocation is shared.

The receipt binds the five source files and compressed validation logs. These results do not establish native-producer interoperability, rendering, activation, timing, or total-memory improvement. Existing eager whole-DOC retention and whole-storage replacement admission remain separate follow-up work; this batch claims stream-specific admission only.
