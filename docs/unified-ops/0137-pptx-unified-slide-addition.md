# 0137 — Unified PPTX slide addition

## Implementation

`File::add`, transaction `add`, and batch add commands now dispatch PPTX slide
insertion to the native immutable transaction. The parent must resolve to the
presentation root. An exact, nonempty `layout` property is required; no layout is
guessed. The new slide is empty and shares the selected native layout. The
returned path identifies its position in the resulting snapshot.

Created presentations and explicitly eager-opened presentations support this
route. Placement accepts start, end, checked position, before/after a resolved
slide, and the presentation root with an optional position. The facade contains
no XML mutation or serialize/reopen fallback. Native insertion owns authoring,
relationship/content-type updates, dialect handling, dependency admission and
source preservation. Semantic limits admit the selector bytes and target before
native editing; the transaction kernel also accounts for the retained native
patch proof. Failed operations and batches retain the caller's original snapshot.

The add registry exposes only the layout assignment. Contextual capability
reports advertise slide addition for eligible eager owners with retained source,
while default/source-backed reports expose the explicit source-owner gap.
Signature and encryption policy remain ahead of ordinary editing. The PPTX
golden now carries the source-owner refusal and follow-up, and the generated
OfficeCLI delta table includes the same gap.

Four public regression tests cover all placement forms, new empty presentations,
text edits before and after insertion, canonical result paths, native proof replay
and inverse through facade transactions, reopened results, invalid inputs,
semantic limits, poisoned transactions, and atomic/best-effort batches. Capability
tests cover the default source route, eager route and released-source refusal.
Read-only review found no correctness or ownership blocker in the adapter.

## Verification

The PPTX/encryption facade suite passed 538 tests/doctests across 26 result
groups. The combined DOCX/XLSX/PPTX/encryption suite passed 1,054 across 26 groups.
Both runs include all 16 PPTX write integration tests and the contextual
capability regressions. Clippy passed for every facade target with warnings
denied; rustdoc also passed with warnings denied. After removing the redundant
iterator field, combined capability tests, Clippy and rustdoc passed again. Both PPTX glossary profiles
passed (79 minimal and 70 full banned names). The generated OfficeCLI delta
passed its deterministic check and all 14 harness tests.

The full run caught an outdated generic capability expectation; it now asserts
the specific source-owner refusal. The combined build also caught a test import
qualification that was only redundant with PPTX enabled alongside DOCX; the
import is now local to its test. A DOCX-only build exposed redundant PPTX
state in the property iterator; that field was removed in favor of the existing
element state, and the DOCX-only check passed. The capability golden was reviewed before
regenerating the delta: only the default add refusal and explicit source-owner
gap changed.

Commands were offline and builds were serialized:

```sh
CARGO_NET_OFFLINE=true cargo test -p litchi-ops --no-default-features --features pptx,encryption
CARGO_NET_OFFLINE=true cargo test -p litchi-ops --no-default-features --features docx,xlsx,pptx,encryption
CARGO_NET_OFFLINE=true cargo check -p litchi-ops --no-default-features --features docx
CARGO_NET_OFFLINE=true cargo clippy -p litchi-ops --no-default-features --features pptx,encryption --all-targets -- -D warnings
CARGO_NET_OFFLINE=true RUSTDOCFLAGS='-D warnings' cargo doc -p litchi-ops --no-default-features --features pptx,encryption --no-deps
CARGO_NET_OFFLINE=true python3 tools/check_api_glossary.py --profile pptx-minimal --target-dir target
CARGO_NET_OFFLINE=true python3 tools/check_api_glossary.py --profile pptx-full --target-dir target
python3 tools/officecli_capability_delta.py --officecli /home/zhuhe/code/litchi/3rdparty/OfficeCLI --check
python3 -m unittest tools.test_officecli_capability_delta
```

## Source-owner prerequisite and remaining scope

Opened source-backed presentations still refuse this element-level operation.
They already support read/write of shape text, so this is not a whole-format
read-only fallback. Source-native slide insertion requires the following work:

- Build and retain one OPC `SourceTopologyPlan`/`PreparedTopology` containing the
  presentation splice, new slide and both relationship bindings. Use effective
  candidate validation and native source-preserving publication.
- Retain candidate payload reservations with owned effective payloads or an
  `Arc<PreparedTopology>`; borrowed effective bytes cannot outlive their owner.
- Construct a candidate catalog from `EffectiveTopology`. The original source
  catalog cannot see a changed presentation root or an added slide.
- Compose structural changes and shape overlays in one topology plan. Publishing
  a prepared topology and then the old overlay writer is not a valid composition.
- Extend local/durable source patches with structural proofs. Reindex position
  overlays or key them by stable slide URI; candidate baseline and closure root
  bindings must describe the new revision.
- Route all facade position-sensitive reads through that revision, including
  slide count/catalog, facts, notes and both comment families. Reading the
  original owner after insertion would select the wrong slide.

The source cross-copy planner contains reusable topology assembly and validation,
but invoking cross-copy itself is unsuitable: it refuses same-lineage sources.
The eager insertion allocator also contains the required dangling-target checks.
These are native leaf prerequisites, not workarounds to reproduce in the facade.

Unified shape creation, slide removal/order verbs, source-native insertion,
remaining formats and performance acceptance remain unfinished. No performance
acceptance claim is made. The full goal remains active.

## Repository checkpoint

The rewritten base remains an ancestor of the current branch; no rebase is
needed. Staging failed because Git cannot create
`/home/zhuhe/code/litchi/.git/worktrees/unified-ops-api/index.lock` on the
read-only filesystem. No commit was created or permission workaround attempted.
Both required staged evidence guards printed nothing, and the changed/untracked
artifact audit found no prohibited evidence or compiled binaries.

All verification handles were terminal before checkpoint scratch logs, the
generated root lockfile and build output were removed. No temporary worktree
was created. Inherited changes and intentional evidence were preserved.
