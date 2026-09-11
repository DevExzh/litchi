# ODG transition and drawing metadata batch

This batch completes the pending ODG metadata implementation. It does not
certify all OpenDocument CRUD scenarios or complete the performance program in
[`docs/GOAL.md`](../../../docs/GOAL.md).

The supported surface is packaged ODG transition editing and inert inventory
of 3D shapes, custom-shape enhanced geometry, image maps, contours, and glue
points. Playback, rendering, geometry evaluation, and activation of links or
sounds are outside this surface. The
[feature matrix](FEATURE_MATRIX.md) distinguishes reading from editing.

## Contract and evidence map

| Requirement | Evidence |
|---|---|
| Namespace-aware owner and lexical validation | [`page_transitions.rs`](../tests/page_transitions.rs) checks alternate prefixes, inherited styles, malformed 3D vectors, misplaced owners, and enhanced-geometry child grammar. |
| Atomic transition edits and exact inverse | [`transition_review.rs`](../tests/transition_review.rs) checks automatic and named owners, exact durable replay/inversion, namespace collisions, and disjoint joins. Lexically irreversible edits retain bounded package projections for exact inversion. |
| Shared style ownership | [`transition_ownership.rs`](../tests/transition_ownership.rs) checks direct and transitive inheritance, atomic refusal, semantic no-ops, and independent child overrides. |
| Unknown source preservation | Transition sound comments and adjacent unknown XML remain source-backed; paired contour/glue-point tests retain comments. |
| Bounded auxiliary inventory | Image-map retained-area and owner XML accounting has a focused rejection test. |
| Optional-inventory dispatch | [`inventory_dispatch.rs`](../tests/inventory_dispatch.rs) checks local aliases, default namespaces, foreign rebinding, and malformed-owner rejection when optional passes are selected during the mandatory scan. |
| Optional metadata storage | Metadata-free shapes retain one optional pointer; enhanced geometry, image maps, contours, and glue points allocate their shared private metadata owner only when present. See the measured layout and memory receipts in change 0502. |
| Existing CRUD behavior | [`package_snapshot.rs`](../tests/package_snapshot.rs) exercises the existing package, edit, patch, resource, and preservation surface. |
| Dependency ownership | `litchi-odg` continues to depend on `litchi-core` and `litchi-odf-common`; ZIP remains a development-only direct dependency. |
| Independent consumer | [LibreOffice receipt](NATIVE_METADATA.md): headless Draw opens/resaves the generated artifact but drops transition metadata. No native transition-persistence claim follows. |

## ADR compliance

| ADR | Application to this batch |
|---|---|
| 0001, 0004 | Semantic selectors and inert bounded values cross the public boundary; retained source fragments are read-only and cannot mutate attached snapshots. |
| 0002, 0010, 0023, 0024 | Drawing grammar remains in its concrete family owner; package I/O remains below it in ODF common. No production dependency is added. |
| 0003 | Edits stage against immutable snapshots, publish only after reopen/readback, and retain exact-source reversible patches. Shared ownership must be proven before changing a style. |
| 0005 | Existing bounded parsing and output limits remain in effect. Performance must be measured separately; additional semantic inventory is not itself a speedup. |
| 0006 | Source bytes remain authoritative. Unsupported edits fail rather than normalize unknown markup; links and sounds remain inert. |
| 0008 | Focused Rust, dependency, and independent consumer evidence are separate gates. A passing round-trip does not establish full specification conformance. |

The repository's [CRUD checklist](../../../docs/CRUD_Scenario_Checklist.md)
remains the broader certification taxonomy. Evidence here is limited to the
named metadata subset.

## Reproduction commands

The final pre-rebase run passed all 87 ODG tests, ODG doctests and rustdoc,
the ODF umbrella feature check, workspace formatting, and both 47-phase
non-iWork check/Clippy sweeps. The retained
[validation receipts](../../../docs/performance/results/change-0502/validation/)
are correctness/build evidence; their elapsed times are not performance claims.
The dependency-boundary checker passed for 64 workspace packages and 239
internal declarations, retaining the 11 explicitly recorded iWork debt items.
The repository-wide example-name checker still reports existing duplicate
iWork example names outside this batch.

Run from the repository root with the pinned Rust 1.95.0 toolchain:

```sh
cargo test -p litchi-odg --all-targets
cargo test -p litchi-odg --doc
cargo test -p litchi-odf --no-default-features --features odg
cargo clippy -p litchi-odg --all-features --lib --no-deps -- -D warnings
RUSTDOCFLAGS=-Dwarnings cargo doc -p litchi-odg --no-deps
cargo fmt --all -- --check
python3 tools/check_crate_boundaries.py
python3 tools/non_iwork_gate.py check
python3 tools/non_iwork_gate.py clippy
```

The [0502 performance record](../../../docs/performance/changes/0502-odg-metadata-open.md)
retains separate before/after measurements and limitations. The existing ODF
`parse_odt` fuzz target does not exercise the ODG-specific metadata parser;
hostile ODG owner coverage comes from the focused tests above.
