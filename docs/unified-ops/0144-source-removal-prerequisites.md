# Native prerequisites for source-backed PPTX removal

## Implemented

The eager slide-removal path now delegates its lexical root change to a shared
binding helper. The helper accepts the complete ordered `(slide ID,
relationship ID)` list and an optional execution context. It proves the complete
list before removing one exact source span, preserves all surrounding bytes,
checks cancellation during scanning and copying, and uses fallible allocation.
Tests cover lexical trivia, malformed and stale bindings, absent selection and
cancellation. This prepares the native root operation for a source catalog
without requiring eager `Slide` values. Physical XML admission, root/MCE policy
and the final-slide rule remain planner responsibilities; the helper itself can
produce an empty slide list.

Source cross-copy now uses a shared registered-layout validator suitable for
source removal. It reuses the existing master/layout/theme graph traversal and
checks layout membership by OPC URI equivalence. The shared traversal now
enforces caller-selected physical XML limits before parsing master, layout and
theme payloads. Regressions cover excessive attribute bytes and event counts in
each of these three owners, requiring the native typed limit error.

No source-removal capability is advertised by this change. The primary source
package and durable patch still require the topology work below; the full goal
is unchanged and remains active.

## Remaining removal implementation

The current source package retains root/slide overlays, a physical source catalog
and a logical order. Its durable v2 proof requires before/after catalogs to be
permutations of one topology. Removal must extend that representation rather
than silently convert the package to an eager owner.

The required forward operation is a single OPC `SourceTopologyPlan`: remove the
selected root binding, its presentation relationship, and its slide Part. OPC
already handles the removed Part's owned relationship member and content-type
Override, and checks retained inbound edges. The PPTX planner must additionally
prove final-slide policy, signatures/macros/protection, modeled slide content,
dependency-free layout ownership, and positional metadata. The shared source
layout graph policy is conservative and is not claimed to match every eager
graph admission case.

Before mutation, retain the selected slide payload and relationships, exact root
and catalog states, and the relevant relationship/content-type metadata. Compose
these with prior text/order edits and bind replay to the effective revision.
Forward publication alone is insufficient for the primary transaction: the
patch must also restore removed resources and compose its inverse correctly.

`SourceTopologyPlan` currently has semantic add/remove operations but no guarded
installation of exact raw metadata preimages after reopening. Re-adding a Part
and relationship does not by itself restore original lexical metadata or member
placement. A bounded authenticated metadata/topology delta is the next OPC
prerequisite. Exact physical archive restoration, where required, also needs a
journal of removed raw records and archive metadata or an explicitly retained
original artifact; an original-artifact hash alone cannot reconstruct bytes.
The existing `SourceArtifactRestoreProof` supports the retained-artifact case,
but that provider cannot be serialized into the ordinary durable patch.

This is an implementation dependency, not a permanent capability exemption or
a reason to weaken the patch/inverse contract. Once it is satisfied, integrate
`source::Edit::remove_slide`, local and durable replay/inverse, logical reads,
and the facade's existing path/query removal in that order. Acceptance includes
mixed text/order/remove edits, untouched payload laziness, failure atomicity,
reopen/inverse, source freshness and managed reservations.

## Verification and checkpoint

The full native PPTX/encryption suite passed 1,583 tests/doctests across 100
groups, with three existing ignores. The focused lexical suite passed 16 tests;
the two physical-limit tests cover all three graph owners. Initial new-test
fixture/configuration mistakes were corrected before the final successful run.
Read-only review found no remaining concrete blocker in the implemented helpers.
The OfficeCLI delta check and all 14 generator tests passed.

Checks were serialized and offline:

```sh
CARGO_NET_OFFLINE=true cargo test -p litchi-pptx --features encryption
CARGO_NET_OFFLINE=true cargo test -p litchi-ops --no-default-features --features docx,xlsx,pptx,encryption
CARGO_NET_OFFLINE=true cargo clippy -p litchi-pptx --features encryption --all-targets -- -D warnings
CARGO_NET_OFFLINE=true cargo clippy -p litchi-ops --no-default-features --features pptx,encryption --all-targets -- -D warnings
CARGO_NET_OFFLINE=true RUSTDOCFLAGS='-D warnings' cargo doc -p litchi-pptx -p litchi-ops --no-default-features --features litchi-pptx/encryption,litchi-ops/pptx,litchi-ops/encryption --no-deps
CARGO_NET_OFFLINE=true python3 tools/check_api_glossary.py --profile pptx-minimal --target-dir target
CARGO_NET_OFFLINE=true python3 tools/check_api_glossary.py --profile pptx-full --target-dir target
python3 tools/check_crate_boundaries.py
```

The three missing SVG corpus fixtures were temporarily copied from the base
checkout for the native suite, as in record 0141, which records their paths,
sizes and SHA-256 hashes. They are verification inputs only.

The combined DOCX/XLSX/PPTX/encryption facade suite passed 1,078 tests/doctests
across 29 groups. Native and facade all-target Clippy and rustdoc passed with
warnings denied. Both glossary profiles passed (79 minimal and 70 full banned
names), along with scoped rustfmt and `git diff --check`.

The rewritten base remains an ancestor, so no rebase was necessary. Staging
failed because Git could not create
`/home/zhuhe/code/litchi/.git/worktrees/unified-ops-api/index.lock` on the read-only
filesystem. No commit was created and no permission workaround was attempted.
The required staged guards printed nothing, and the changed/untracked artifact
audit found no prohibited evidence or compiled binaries.

The crate-boundary checker passed for 66 workspace packages and 268 internal
dependency declarations, retaining 11 explicit migration-debt items.

All verification handles finished before cleanup. This turn's temporary corpus
copies, scratch logs, generated root lockfile, build output and Python caches
were removed. No temporary worktree was created. Inherited working changes,
fixtures and tracked lockfiles were preserved. Broader CRUD, Python and paired
performance acceptance remain unfinished; this record makes no performance
acceptance or full-goal completion claim.
