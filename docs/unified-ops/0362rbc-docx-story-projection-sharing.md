# 0362rbc — DOCX raw story projection sharing

The first raw selected-story projection still paid for a complete copy of the
already admitted main-document transaction XML. `story_text::Payload` now has
an internal validated `OwnedRange` representation. A main-story target keeps
the transaction's `Arc<Vec<u8>>` and uses the full `0..len` range, so target
construction borrows the projected state instead of allocating a second full
XML buffer. The selected-story scan, fingerprint, output limit, and execution
checks still run on the shared bytes. Collection stories retain their existing
full-part splice because their prefix/entry/suffix artifact must remain one
contiguous exact source part; glossary batch replacements retain the same
bounded full-artifact path.

The public managed path remains explicit: `Snapshot::edit_engine` checks the
execution context and refuses a managed payload before creating the raw edit
engine. `managed_story_reads_are_bounded_and_edits_are_typed_refusals` covers
that reachable path. No synthetic managed raw tree is constructible or used.

## Matched evidence

The matched route was
`raw_story_tree_caches_projection_and_invalidates_after_mutation`. It covers
the first projection, the repeated cached projection, invalidation after a
mutation, the changed projection, and patch target access. The baseline is the
protected 35b candidate binary
`(deleted-coord-artifact)`,
SHA-256
`5d0d5d1bf65e700bab70cb41fabb03446abf6fd1a70069d85826a4080fa24af9`. Its
`story_text.rs` source SHA-256 is
`e046b0aecc6a89e088b567e3e334de77b3726b449486054ef4265e724f151eb9`.
The candidate binary is
`(deleted-coord-artifact)`,
SHA-256
`6f75def75b7372d7aa5056cbbc3a244e6b1365d28b81240d108ceea7258da9c8`,
with `story_text.rs` SHA-256
`c306e66c249a2cbdf667cf296d74b4718b37314fea70bd3f11e2f7d8aebd899a`.
Both binaries used the same locked dependency graph, fixture, test source,
and debug profile; the only relevant execution-path difference in this
comparison is the range-sharing change.

Heaptrack reported 949 allocations, 148 temporary allocations, and 503.58 KiB
peak heap for the baseline. The candidate reported 943 allocations, 148
temporary allocations, and 503.56 KiB peak heap: six fewer allocations and no
peak increase. Callgrind recorded 18,108,160 Ir for the baseline and
18,103,411 Ir for the candidate (4,749 fewer instructions, 0.026%). Raw
captures and analyses remain outside the repository at
`(deleted-coord-artifact)`:

- baseline heaptrack Zstandard capture SHA-256
  `35e324f610a0c3875ca75b08b6f6c1d26e44f4c1b7f3a3f6246e1dce5ec967dc`;
- candidate heaptrack Zstandard capture SHA-256
  `fe26a7f957cd7e6b125f6376ec4556f362e50b2971fb69cfa7eec0f1a0d9d0f4`;
- baseline callgrind capture SHA-256
  `2a744bec2192d125faa50cda634b6bf5c8fecb595c2c1a4a5ee3f0e792ac91df`;
- candidate callgrind capture SHA-256
  `c8c300b2f4bb9fd411be3c5b6db51fd8098eebebd8d37d020c366097eb432413`.

This is a small fixture and a combined first/repeated route; the numbers are
allocation and instruction evidence for that route, not a claim about every
document size.

## Validation

The affected story and glossary suites passed 14 and 19 tests. The external
log is `(deleted-coord-artifact)/range-test.log`,
3,137 bytes, SHA-256
`415190998248ca73b1016805b0ab8fd57d7234c1bbb90ef4e9abbc43bd907f8e`.
All-target/all-feature `litchi-docx` Clippy passed with `-D warnings`; the log
is `(deleted-coord-artifact)/range-clippy.log`,
171 bytes, SHA-256
`c3eefbff404c831928bb90cf349fa91e32634f9409f693f72e1cb98786f0a5ea`.
`cargo fmt --all -- --check` and `git diff --check` passed. No binaries or
large evidence files are committed.
