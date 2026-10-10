# 0142 — Primary source-backed PPTX slide ordering

## Implementation

The primary native source package now composes slide moves, swaps and shape-text
edits in one immutable revision. Its logical catalog uses stable physical part
URI, slide ID and presentation relationship ID keys. A retained hash index maps
logical positions back to the source catalog; text snapshots and changes remain
bound to physical slide identities. Moves and swaps accept checked `Position`
values and resolve against the currently staged order.

Root XML and slide overlays share one bounded OPC overlay publication. Reverting
to the source order restores the original root payload, preserving exact no-op
ZIP output. Root reservations follow immutable packages, patches and inverses;
failed later operations leave the preceding candidate intact. Content comparison,
patch no-op detection, inverse and retained-byte checks include the root/order
state. No eager materialization or serialize/reopen editing fallback was added.

Durable source patches now use `LPSX0002`. The wire includes exact before/after
root payloads, ordered stable keys and selected slide preimages. Its source digest
binds topology and a canonical root permutation, allowing replay and inverse
replay after reopening a reordered file without accepting unrelated root drift.
Decoding bounds catalog counts by available framing, validates complete input and
pure root permutations before retaining payloads, and rejects redundant slide
deltas. Replay checks every root/slide preimage before creating a candidate.
A text-only patch authored after ordering requires that order context. Unchanged
roots do not inherit order-only metadata restrictions or edit-memory charges.
The current wire retains both full root preimages; compact optional root deltas
remain an optimization opportunity, not a performance acceptance claim.

The unified adapter now routes retained-source slide Move/Swap through native
source commits. Catalog IDs, text, scenes, facts, notes, legacy comments and modern
comments all follow the revised logical order. Source adoption checks owner
lineage and candidate freshness. Native stale-source errors classify as `IoError`.
Capabilities and the OfficeCLI delta remove the blanket source-order gap while
retaining explicit positional-metadata gaps, signing, encryption and retention
policy. Source-backed Add/Remove still require topology integration.

## Verification

The full native PPTX/encryption suite passed 1,572 tests/doctests with three
existing ignores across 100 result groups. The PPTX/encryption facade suite
passed 561 across 29 groups. The new primary source-order integration suite
passed nine tests; primary unit coverage includes laziness, large catalogs,
forged root/slide preimages, decoder framing and positional-owner compatibility
for text-only replay. Managed tests distinguish retained metadata from root edit
reservations and verify exact release on inverse staging, cancellation and final
patch/inverse drop.

The capability golden was regenerated, the OfficeCLI delta check passed, and all
14 delta-tool tests passed. The combined DOCX/XLSX/PPTX/encryption facade suite
passed 1,077 tests/doctests across 29 groups. Native and facade all-target Clippy
passed with warnings denied, as did rustdoc for both crates. The no-format capability test also passed. Both glossary
profiles passed (79 minimal and 70 full banned names). Scoped rustfmt and
`git diff --check` passed. Read-only review's concrete decoder/proof findings
were corrected and covered by regressions.

Checks were serialized and offline:

```sh
CARGO_NET_OFFLINE=true cargo test -p litchi-pptx --features encryption
CARGO_NET_OFFLINE=true cargo test -p litchi-ops --no-default-features --features pptx,encryption
CARGO_NET_OFFLINE=true cargo test -p litchi-ops --no-default-features --features docx,xlsx,pptx,encryption
CARGO_NET_OFFLINE=true cargo clippy -p litchi-pptx --features encryption --all-targets -- -D warnings
CARGO_NET_OFFLINE=true cargo clippy -p litchi-ops --no-default-features --features pptx,encryption --all-targets -- -D warnings
CARGO_NET_OFFLINE=true RUSTDOCFLAGS='-D warnings' cargo doc -p litchi-pptx -p litchi-ops --no-default-features --features litchi-pptx/encryption,litchi-ops/pptx,litchi-ops/encryption --no-deps
CARGO_NET_OFFLINE=true python3 tools/check_api_glossary.py --profile pptx-minimal --target-dir target
CARGO_NET_OFFLINE=true python3 tools/check_api_glossary.py --profile pptx-full --target-dir target
```

The synthetic positive fixtures use minimal presentation properties. Builder
stock properties contain extension content that the existing native source-order
policy refuses; a separate regression preserves that typed refusal. Reopening
establishes a new physical source catalog, so `source_position` assertions use
that newly opened catalog as their basis.

As in record 0141, three missing SVG corpus fixtures were temporarily copied from
the base checkout for the full native suite. Their paths, sizes and SHA-256 hashes
are recorded there; they are verification inputs only.

## Remaining work and repository checkpoint

Source Add/Remove and richer positional metadata preservation, broader format and
element coverage, Python bindings and paired performance acceptance remain open.
The full goal remains active. No performance acceptance claim is made.

The rewritten base remains an ancestor, so no rebase is necessary. No temporary
worktree was created. The inherited working changes and evidence are preserved.
Staging failed because Git could not create
`/home/zhuhe/code/litchi/.git/worktrees/unified-ops-api/index.lock` on the read-only
filesystem. No commit was created and no permission workaround was attempted.
Both required staged guards printed nothing; the changed/untracked artifact audit
found no prohibited evidence or compiled binaries.

All verification handles finished before the three temporary corpus copies,
checkpoint scratch logs, generated root lockfile, build output and two generated
Python cache files were removed. Inherited fixtures, tracked lockfiles and
intentional evidence were preserved.
