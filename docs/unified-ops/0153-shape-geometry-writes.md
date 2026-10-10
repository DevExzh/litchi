# Atomic PPTX text and geometry writes

This increment implements L-PPTX-W3's shape-bound setters through the native
eager and source Edit owners and the unified property registry. It does not
complete the unified operations goal or its performance/acceptance gates.

`ShapePropertiesReplacement` carries a checked shape `Position`, optional text
and optional complete `(x, y, width, height)` bounds. `Edit::set_shape_properties`
applies the whole same-slide batch atomically. The facade merges partial
geometry assignments with the current local bounds before calling the native
operation. Missing local transforms are not synthesized. Requested extents
must be positive; coordinates and extents obey the native DrawingML domains.
Ordinary `p:sp` shapes, including placeholders with explicit local transforms,
are supported. Other transform-owner kinds and inherited/missing geometry
remain separate capability gaps.

The codec locates the selected ordinary shape's direct PresentationML shape
properties and DrawingML transform by namespace and ownership. It splices
numeric attribute values while retaining surrounding bytes. Text and geometry
share one staged candidate and one publication; a later failure cannot publish
an earlier text update. The native batch has a finite 256-shape limit. Work on
the geometry spans is shared across the batch, avoiding one full-slide rewrite
per target. Equal coordinate values keep their original lexical spelling even
when another coordinate changes. This is an implementation property, not a
performance result.

Source edits retain managed admission with their immutable payload. Primary
source patches compose with earlier slide-order and topology edits. Durable
proof compares exact XML outside text slots and the four validated transform
values. Canonical appended text boxes retain their full fragment proof; their
suffix is removed before comparing the remaining property surface. Source
patch witnesses describe allowed edits and are not authenticated histories.
Malformed or unsupported transforms cannot be authored; proof leaves them
unmasked so an unrelated text, geometry or text-box addition can preserve them
exactly. Namespace aliases, duplicate owners, child order and typed transform
attributes participate in admission.

The previous goal turn made implementation progress. The rewritten base remains
an ancestor of HEAD, so no rebase was needed. No unrelated checkout or iWork
crate is part of this increment.

The full native PPTX/encryption suite passed 1,663 tests/doctests; the combined
DOCX/XLSX/PPTX/encryption facade suite passed 1,092. Eager/source tests cover
mixed properties, no-op bytes, invalid and duplicate atomic refusals, lexical
preservation, v2/v3 transport, reopened inverses, appended text boxes, malformed
unselected transforms, resource retention and cancellation. The forged-patch
tests now use changes outside the expanded bounds surface (preset geometry),
so they continue to verify confinement rather than rejecting legitimate edits.

Native and PPTX facade all-target Clippy and rustdoc passed with warnings
denied. The OfficeCLI delta generator's 14 tests and generated report check
passed. Scoped formatting and diff checks passed. The artifact audit found no
prohibited evidence or compiled binary among 673 changed/untracked paths.
Both glossary profiles passed (79 minimal and 70 full banned names). Facade
checks with no features and with PPTX alone passed. Crate boundaries passed for
66 packages, 268 internal dependency declarations and 11 explicit debt items.

The initial combined build exhausted the temporary disk quota with debug and
incremental artifacts. That turn-owned build directory was removed and both
full suites were rebuilt successfully with debug symbols and incremental
compilation disabled. The test scope was unchanged.

```sh
export CARGO_PROFILE_DEV_DEBUG=0 CARGO_PROFILE_TEST_DEBUG=0 CARGO_INCREMENTAL=0
export CARGO_TARGET_DIR=/tmp/litchi-unified-0153-target CARGO_NET_OFFLINE=true
cargo test -p litchi-pptx --features encryption
cargo test -p litchi-ops --no-default-features --features docx,xlsx,pptx,encryption
cargo clippy -p litchi-pptx -p litchi-ops --no-default-features --features litchi-pptx/encryption,litchi-ops/pptx,litchi-ops/encryption --all-targets -- -D warnings
RUSTDOCFLAGS='-D warnings' cargo doc -p litchi-pptx -p litchi-ops --no-default-features --features litchi-pptx/encryption,litchi-ops/pptx,litchi-ops/encryption --no-deps
python3 tools/check_api_glossary.py --profile pptx-minimal --target-dir "$CARGO_TARGET_DIR"
python3 tools/check_api_glossary.py --profile pptx-full --target-dir "$CARGO_TARGET_DIR"
cargo check -p litchi-ops --no-default-features
cargo check -p litchi-ops --no-default-features --features pptx
python3 tools/check_crate_boundaries.py
python3 tools/officecli_capability_delta.py --check --officecli /home/zhuhe/code/litchi/3rdparty/OfficeCLI
python3 -m unittest discover -s tools -p test_officecli_capability_delta.py
git diff --check
```

Managed admission is conservative and precedes semantic no-op detection; an
all-equal request can still be refused under memory pressure. Wider shape CRUD,
the 1k-operation performance gate, remaining formats, Python/CLI and complete
acceptance remain open. No throughput or latency improvement is claimed.

Both staged-artifact guards printed nothing. Scoped staging failed because Git
could not create `/home/zhuhe/code/litchi/.git/worktrees/unified-ops-api/index.lock`
on the read-only Git filesystem; no commit was created. No approval workaround
was attempted. All implementation changes remain in this worktree.

All verification processes finished before cleanup. Cleanup removed the three
temporarily copied corpus fixtures after verifying their recorded SHA-256,
the turn-owned build output and logs, generated Python caches and root lockfile.
No temporary worktree was created. Inherited files and the base checkout are
preserved.
