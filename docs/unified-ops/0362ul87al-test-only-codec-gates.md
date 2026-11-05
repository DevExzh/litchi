# 0362ul87al — gate test-only OOXML codec helpers

This source-only follow-up starts from `b6328f489577e8a36331847b7d5f4bc58fda4101`
(tree `b7ac524492aa937ff3116f4da28636ee3c2afa43`). It closes the remaining
87ai release dead-code surface without retaining `dead_code` allowances or
removing the low-level fidelity tests.

The core-properties test-only publisher and its physical-graph helpers now
use `#[cfg(test)]`, including `write_core`, `clear`, `sync`, the convenience
`encode`, candidate-name and relationship-ID helpers, and their constants and
imports. `Core::xml`, the `Slot` staging guard, and `Slot::flush`/`stage` are
gated with the test-only path they serve. The custom-properties encoder and
relationship target, extended-properties `splice`, and core graph deletion
checks receive the same treatment. The existing test-only re-exports remain
test-only.

Production publication keeps the preflighted codecs and bounded transaction
path: core `encoded_len`/`encode_preflighted`, custom `encoded_len`/
`encode_preflighted`, and application `plan_splice`/`splice_planned`/
`encode_new_preflighted` are unchanged. The direct XML, graph, and exact-byte
tests still use their private helpers under `cfg(test)`. Current custom-
properties module documentation now points semantic publication at
`properties::Edit::set_custom_props` and `remove_custom_props`.

`rustfmt --edition 2024 --config skip_children=true` and `git diff --check`
were run on the changed source paths, together with static symbol and
allowance audits. No Cargo, rustdoc, glossary, test, native, performance, or
F5 command was run for this tip. The preliminary unapproved glossary receipt
from the earlier 87aj turn remains preserved outside the checkout and is
described in `0362ul87aj-stale-ooxml-publishers.md`.
