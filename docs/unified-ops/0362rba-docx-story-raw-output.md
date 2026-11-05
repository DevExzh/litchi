# 0362rba — DOCX story raw output clone repair

This follow-up closes the raw selected-story output regression identified in
the 0362rb review. `source_backed::story_text::Edit::target` now reuses the
already validated projected `DocumentSnapshot` for both main and collection
stories. It no longer calls `DocumentEdit::try_clone().commit()` merely to
produce the first raw projection. The source-backed owner path remains the
consuming `Edit::finish(self)` path, so package publication keeps its existing
Commit, inverse, source-freshness, and full-artifact checks.

The public raw model still refuses managed payloads before creating an edit
engine. The existing `managed_story_reads_are_bounded_and_edits_are_typed_refusals`
test exercises that reachable public path; no synthetic cache-execution test
was added for a managed raw tree that cannot be constructed. Raw owner replay
is covered by `selected_story_replays_on_independent_matching_copies_and_chains_edits`:
an independently opened equal-content source accepts the patch, a second edit
chains from the first publication, and the second inverse restores the exact
first publication bytes. The existing owner test continues to verify exact
full-artifact forward and inverse restoration.

## Allocation evidence

The matched filtered test was
`raw_story_tree_caches_projection_and_invalidates_after_mutation`, switched to
the main story so that the pre-repair non-envelope branch was exercised. The
baseline is committed `8dc298fbb7004bf17d354859f4c684c5df616796`, whose
`story_text.rs` SHA-256 is
`f034166f63bc874fa0a73d271692ef9832f82695748c6e0bcc0bc4cd44515522`.
The candidate `story_text.rs` SHA-256 is
`e046b0aecc6a89e088b567e3e334de77b3726b449486054ef4265e724f151eb9` and the
test source SHA-256 is
`5cd739779794daf9626c5ca3e9430b8505d7c72dd390ce2d22e7489104d0dbeb`.
Both binaries used the same locked dependency graph and fixture source:

- baseline binary SHA-256:
  `6e4f251dd321411c55869f9783665474af15fa7bf3f06c2a943cb89fc1444eb2`;
- candidate binary SHA-256:
  `5d0d5d1bf65e700bab70cb41fabb03446abf6fd1a70069d85826a4080fa24af9`.

Heaptrack reported 994 allocations, 150 temporary allocations, and 5.55 KiB
peak allocation consumption for the baseline. The candidate reported 949
allocations, 148 temporary allocations, and the same 5.55 KiB peak on this
small fixture: 45 fewer allocations (about 4.5%) without a measured peak
increase. The heaptrack peak is the allocation-trace peak for this filtered
test, not a process RSS claim. Raw captures and analysis are outside the
repository under
`(deleted-coord-artifact)`.

The binaries were built with:

```text
source (deleted-coord-artifact)
export CARGO_BUILD_JOBS=1 RUST_TEST_THREADS=1 LITCHI_MEM_MAX=20G
export TMPDIR=(deleted-coord-artifact)
cargo test --offline --locked -p litchi-docx --test source_backed_story_text --no-run
```

Each binary was run with:

```text
heaptrack --output <external-output> <protected-binary> --exact raw_story_tree_caches_projection_and_invalidates_after_mutation --nocapture
heaptrack --analyze <external-output>.zst
```

## Validation

The final focused command passed 58 tests (14 story, 25 secondary-story, and
19 glossary-story tests). Its external log is
`(deleted-coord-artifact)/story-focused-final.log`,
5,248 bytes, SHA-256
`eccfff9bd1af551776b6a3d7f80ac5a189c1aa04ddaacdc627023f5fb0df5096`.
The command was:

```text
cargo test --offline --locked -p litchi-docx \
  --test source_backed_story_text \
  --test source_backed_secondary_story_text \
  --test source_backed_glossary_story_text --no-fail-fast
```

Strict all-target/all-feature Clippy passed with `-D warnings`; its final
external log is
`(deleted-coord-artifact)/story-clippy-final.log`
(171 bytes), SHA-256
`baf608fec5bce8dff5e8910f207ba20ba4374af6b03ae7a9fd6b2cffdc1145ba`.
Targeted rustfmt and `git diff --check` also passed. No binaries or large
evidence files are committed.
