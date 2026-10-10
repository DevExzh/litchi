# Guarded source topology metadata restoration

## Contract

Reversible source-backed PPTX removal needs to restore relationship and
content-type XML after reopening a changed package. Typed OPC add/remove
operations alone can regenerate or append metadata, losing the original
lexical image. This record tracks the source-native prerequisite identified in
record 0144; it does not claim that primary PPTX removal is complete.

The topology plan remains the authority for semantic changes. Raw restoration
supplies an expected physical source image and a replacement image. The
expected image must match before any publication, and the parsed replacement
must agree with the effective typed graph and content-type coverage. The
restoration operation cannot introduce relationships that are absent from the
typed plan. Absent and present-empty relationship members are distinct states.

Exact no-op requests still check their source preconditions. Changed signed
packages remain refused, and malformed or semantically opaque metadata must not
be admitted merely because a tolerant parser omitted unknown content. Parsing,
retained metadata and publication remain subject to finite source and managed
resource limits. Preparing the candidate must not materialize untouched Part
payloads; eventual raw-copy publication necessarily reads their archive bytes.

This primitive restores decoded metadata bytes and physical member presence.
It does not promise original ZIP member ordering or complete archive-byte
reconstruction after topology changes. Format-owned durable patches still need
to retain and authenticate the metadata images, removed resources and revision
bindings, and compose them with root/order/text edits.

## Implementation and verification

`SourceTopologyPlan` now accepts guarded relationship XML images per owner and
one guarded content-types image. Preparation authenticates physical source
presence and bytes, validates strict metadata grammar, and compares replacement
relationships with the effective typed graph. Publication retains the supplied
metadata image rather than regenerating it.

Content-type admission compares coverage of effective physical members. This
allows an inverse to recover an original explicit relationship-member override
when typed insertion would otherwise use the existing `rels` default. Orphan
overrides and mismatched effective types are refused. Removing an empty physical
relationship member also removes its obsolete explicit override.

The regression fixture removes a part and its incoming edge, publishes and
reopens the result, then restores the typed resources and exact metadata images.
Additional cases exercise stale guards, hidden XML markup, empty-member
presence, finite request bounds, source limits, cancellation, signatures and
lazy preparation.

Review tightened admission of metadata parser allocations and coverage indexes,
replaced a raw-only graph clone with a borrow, and routed source content-type
reads through direct payload accounting. Generic entry/part bounds now apply
before and after that read. Regression fixtures were also corrected to reach
the intended strict XML check rather than an earlier dangling-edge refusal.

The final OPC suite passed 1,135 tests/doctests, including nine new raw-metadata
integration tests. The native PPTX suite passed 1,583 tests/doctests and
the combined DOCX/XLSX/PPTX/encryption facade passed 1,078. Native OPC/PPTX and
PPTX facade all-target Clippy passed with warnings denied. Checks were serialized
and offline:

```sh
CARGO_NET_OFFLINE=true cargo test -p litchi-opc
CARGO_NET_OFFLINE=true cargo test -p litchi-pptx --features encryption
CARGO_NET_OFFLINE=true cargo test -p litchi-ops --no-default-features --features docx,xlsx,pptx,encryption
CARGO_NET_OFFLINE=true cargo clippy -p litchi-opc -p litchi-pptx --features litchi-pptx/encryption --all-targets -- -D warnings
CARGO_NET_OFFLINE=true cargo clippy -p litchi-ops --no-default-features --features pptx,encryption --all-targets -- -D warnings
CARGO_NET_OFFLINE=true RUSTDOCFLAGS='-D warnings' cargo doc -p litchi-opc -p litchi-pptx -p litchi-ops --no-default-features --features litchi-pptx/encryption,litchi-ops/pptx,litchi-ops/encryption --no-deps
CARGO_NET_OFFLINE=true python3 tools/check_api_glossary.py --profile pptx-minimal --target-dir target
CARGO_NET_OFFLINE=true python3 tools/check_api_glossary.py --profile pptx-full --target-dir target
python3 tools/check_crate_boundaries.py
```

The boundary check passed for 66 workspace packages, 268 internal dependency
declarations and 11 explicit migration-debt items. The OfficeCLI delta check and
its 14 generator tests passed. The three native SVG corpus fixtures absent from
this worktree were temporarily copied from the base checkout; their paths,
sizes and SHA-256 hashes are recorded in 0141. They are verification inputs only.

Rustdoc passed with warnings denied. Both glossary profiles passed (79 minimal
and 70 full banned names), as did scoped rustfmt and `git diff --check`.
Review confirmed the parser and candidate reservations; its final per-part
content-types ceiling finding was fixed and covered by the final OPC suite.

The rewritten base remains an ancestor, so no rebase was necessary. Primary
source PPTX removal, durable topology integration, broader format CRUD, Python
bindings and paired performance acceptance remain open. The full goal is active.

Staging failed because Git could not create
`/home/zhuhe/code/litchi/.git/worktrees/unified-ops-api/index.lock` on the read-only
filesystem. No commit was created and no permission workaround was attempted.
Both required staged guards printed nothing; the changed/untracked artifact
audit found no prohibited large evidence or compiled binaries.

All verification handles completed before cleanup. Temporary corpus copies,
scratch logs, the generated root lockfile, build output and this turn's Python
caches were removed. No temporary worktree was created. Inherited working
changes, existing fixtures and tracked lockfiles were preserved.
