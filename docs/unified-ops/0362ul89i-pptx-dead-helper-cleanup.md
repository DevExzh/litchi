# 0362ul89i — PPTX compiler-dead helper cleanup

Status: source-ready for independent review; runtime and Cargo validation were
not run under the assigned source-only lane.

Base: `237d6801a641b85b4477c432a26ac53e531987ed`.

## Scope

This lane owns only the opened transaction helpers, PPTX open options and
package model, and the source-backed presentation snapshot in the sparse
worktree. The supplied all-features compiler attempt reported 113
`dead-code` errors; its log is
`(deleted-coord-artifact)/01-pptx-all-features-check-attempt-actual-f5-01.log`
with SHA-256
`e3004cf8f79ec32a2dd2b082f4204dc1efe5a673dab7df97dcb63f9203aa06ea`.

Zero-consumer default twins were removed. The selected-limit production paths
remain the replacements:

| Removed helper family | Retained route |
| --- | --- |
| `OpenOptions::semantic_limits_value`, `Package::derive` | retained semantic profiles and `derive_with_part_bytes` |
| chunk, enclosed-removal, property, removal-proof and remove-plan default wrappers | corresponding `*_with_limits` or `*_with_semantic_limits` route |
| text-format and XML rewrite/skeleton/staging/removal default wrappers | corresponding selected-limit route |
| `opened::Patch::{apply_to_new,apply_exact_revision}` | semantic-limit variants |
| `model::{capture_with_parent_digests}` | `capture_with_parent_digests_and_semantic_limits` |

Helpers whose only consumers are unit tests retain their capability behind
`cfg(test)`: text-chunk and geometry convenience calls, model capture
conveniences, the raw patch apply convenience, compaction semantics, source
text-box stripping, and lazy laser-usage inspection. Definition-only property
chunk and compaction-scene twins are removed even from the test configuration.
No production path regains a default policy;
the existing selected-limit reexports and callers are unchanged.

## Checks

`git diff --check` passed. Symbol searches against the base tree found no
callers for the removed private/default helpers outside their definitions.
`rustfmt --check --config skip_children=true` parsed the edited files; it
reported pre-existing formatting differences in `package/model.rs` before the
source-only check was stopped. Cargo, rustdoc, native builds, and runtime tests
were intentionally not run in this lane.

No performance or capability waiver is claimed. The compiler log, this record,
and the commit are the complete source evidence for the cleanup.
