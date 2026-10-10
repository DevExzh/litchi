# 0128 — PPTX empty slide-list prerequisite

## Change

The internal range-preserving slide-binding insertion codec now supports an
existing empty `sldIdLst`, including paired and self-closing forms. It preserves
surrounding bytes, BOM, comments, processing instructions, attributes and the
original list QName. The new entry binds its own prefixes and uses the list's
strict or transitional dialect; empty lists whose dialect differs from the
presentation root are refused. Missing lists are still refused.

Both insertion paths bound new relationship IDs with the existing native
4 MiB string limit and validate the native authoring ID grammar before placing
them in XML attributes. Populated insertion's allocation preflight now includes
the ID length. Empty insertion reserves the entry before formatting and reserves
the output before copying. The shared grammar scan continues to enforce depth,
node and exact current-binding checks. This does not add transaction-specific
source limits to the codec or broaden populated-list dialect validation.

Four regressions cover lexical preservation across six list/dialect variants,
second insertion and order, malformed/missing/duplicate lists, stale bindings,
position bounds, mixed dialects, hostile IDs, oversized IDs and XML depth.
No public slide-creation method is introduced in this checkpoint.

## Required follow-up: durable metadata preservation

An insertion prototype was withdrawn after inspection showed that native opened
patches encode semantic part types and relationships without the retained XML
for `[Content_Types].xml` or relationship members. A diagnostic using OPC's
source-token insertion API confirms that direct insertion preserves those XML
members while serialized/deserialized native patch replay canonicalizes them.
The diagnostic's inverse restores the original members for this fixture; it
does **not** demonstrate an inverse failure.

The small source artifact
[0128-metadata-replay-probe.rs](results/0128-metadata-replay-probe.rs) reproduces
this finding. It is intentionally outside normal test targets: its assertions
confirm the current defect rather than enshrine desired behavior. Temporarily
include it as a `#[cfg(test)]` child module at the end of `opened/patch.rs`:

```rust
#[cfg(test)]
mod metadata_replay_probe_0128 {
    include!(concat!(env!("CARGO_MANIFEST_DIR"),
        "/../../docs/unified-ops/results/0128-metadata-replay-probe.rs"));
}
```

Run the command below, then remove that temporary module. No generated binary
or large capture is retained.

```sh
CARGO_NET_OFFLINE=true cargo test -p litchi-pptx --features encryption --lib metadata_replay_probe -- --nocapture
```

Before adding public slide creation, extend the durable patch representation
with bounded metadata edits and source provenance. Direct commit, durable replay
and inverse must agree on preserved XML while unrelated destination metadata
edits survive. A whole-member overwrite alone would not satisfy transport across
an independently edited destination. Cover wire limits/versioning, composition,
conflicts and rollback along with lexical preservation. The existing OPC owned
metadata edit plans provide the native lowering boundary; do not introduce a
second presentation model in the facade.

## Verification and status

The native encryption unit run passed 967 regular tests and the temporary
metadata replay diagnostic (one existing ignore). The four new empty-list tests
passed. All 99 selected integration tests passed: 58 source cross-copy, seven
adversarial cross-copy, 30 source-edit and four signature-policy tests. All-target native Clippy passed with warnings denied. The
probe's temporary child module was removed after its test process terminated.
Read-only review checked the lexical preservation, allocation and refusal paths;
its broader populated-list dialect/source-limit observations remain outside this
empty-list prerequisite and are not claimed as solved.

```sh
CARGO_NET_OFFLINE=true cargo test -p litchi-pptx --features encryption --lib
CARGO_NET_OFFLINE=true cargo test -p litchi-pptx --features encryption --test source_backed_cross_copy --test source_backed_cross_copy_adversarial --test source_backed_edit --test opened_signature_policy
CARGO_NET_OFFLINE=true cargo clippy -p litchi-pptx --features encryption --all-targets -- -D warnings
```
The full goal, broader PPTX edits and all-format/performance acceptance remain
open. No performance acceptance is claimed.


## Commit and cleanup

The rewritten base is already an ancestor of HEAD; no rebase replay is needed.
Staging failed with exit 128 because Git cannot create
`/home/zhuhe/code/litchi/.git/worktrees/unified-ops-api/index.lock` on the read-only
metadata filesystem. No commit was created. Both required staged artifact guards
printed nothing; the changed/untracked evidence and executable audit found no
forbidden artifacts.

After all build/test processes terminated, the temporary probe registration,
`/tmp/litchi-unified-0128-*` scratch, generated root `Cargo.lock`, `target` and
Python cache were removed. The small diagnostic source under `docs/unified-ops`
is an intentional review artifact, not a temporary build input. No fixture links
or temporary worktrees were created in this checkpoint. Inherited modifications,
tracked tool lockfiles and existing fixture directories were preserved. Delegated
review was read-only. The full goal remains active.
