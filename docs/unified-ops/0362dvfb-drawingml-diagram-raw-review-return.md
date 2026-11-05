# 0362dvfb — DrawingML diagram raw review return

- Status: source repair, regressions and focused capped validation are complete
  for coordinator review. ABBA measurement, native builds and the final
  coordinator gate remain intentionally unclaimed.
- Parent source: `9786dcf45056009768f66579eb1e90faac8145c3`.
- Validated source tree: `f0c60af152e1a6d667a57f140062feb92a111ba6`.
- This packet has no API signature change, so existing DOCX, PPTX, ops and
  SmartArt-harness consumers remain migrated to the canonical limits APIs.
- The earlier `e980` packet's 245-test result was diagnostic only. The final
  `9786` run below passed 252 `litchi-drawingml` unit tests plus every
  integration suite.

## Concrete review findings and repairs

The `inventory_font_cache_limits` re-review identified seven parser/authoring
gaps. The source repair addresses each at its owning path:

1. `diagram/model.rs::SmartArt::parse_data_xml` now counts raw `dgm:pt` and
   `dgm:cxn` records against `Limits::points_value()` and
   `Limits::connections_value()`, reports the typed observed/max limits, uses
   `data_text_bytes_value()` for point text, and checks reconstructed semantic
   depth against `node_depth_value()` before descending.
2. A paired empty `dgm:dataModel` now requires the direct `dgm:ptLst` grammar;
   the existing self-closing-root rejection remains. The parser rejects
   nested/duplicate data-model roots and misplaced `pt`, `ptLst`, `cxn` and
   `cxnLst` elements instead of silently ignoring them.
3. `RawConnection` retains `srcOrd`. Parent and root sibling vectors are
   sorted by that source order with a stable point-order tie-breaker, so valid
   sibling order survives parsing.
4. Empty content points are retained, allowing an empty parent with children
   to remain in the semantic tree. A `parOf` connection whose source or
   destination is not the document/content graph now returns `Invalid` rather
   than promoting or dropping the child.
5. `diagram/data/codec.rs::DiagramDataModel::parse` rejects a nested
   `dgm:dataModel` root explicitly.
6. `diagram/edit.rs::read_drawing` requires exactly one direct `dsp:spTree`,
   so a paired `dsp:drawing` without a shape tree is rejected while a valid
   empty shape tree remains readable.
7. `diagram/authoring.rs::Spec::parse_outline` checks custom node-count,
   semantic-depth and node-text limits before each `Node` allocation; invalid
   XML text is also rejected before insertion.

## Regression coverage

The packet adds behavior tests for:

- independent points, connections, data-text and semantic-depth limits;
- paired-empty data models, misplaced structural elements, nested roots and
  dangling parent connections;
- empty-parent retention and `srcOrd` sibling ordering;
- the paired drawing-root/no-`spTree` case;
- nested typed `DiagramDataModel` roots; and
- authoring node/depth/text policies before tree growth.

No fixture, binary, performance artifact or `3rdparty/` reference was added.
No Cargo command was run for this source-only repair before a new validation
slot was granted.

## Root-order follow-up

The follow-up review found one deterministic-order defect in the first repair:
root entries were sorted only by `srcOrd`, so equal document-level orders could
be reordered by `sort_unstable`. `diagram/model.rs` now uses the same
`(srcOrd, point-index)` key for roots and child vectors. A regression supplies
three roots whose points and connections are in different orders but whose
`srcOrd` values are equal, and asserts stable point order.

The immediately following audit found that an unconditional second
`dataModel` guard accidentally rejected every normal paired root. That guard
was removed; nested-root rejection remains in the depth-aware root branch. A
new regression now generates a nonempty SmartArt data part with a paired root,
parses it successfully and checks its node text. The correction is committed
at `9786dcf450` and is covered by the final 252-test run below.

## Final locked validation packet

The validation ran on 2026-10-24 with the repository memory-capped cargo shim,
`LITCHI_MEM_MAX=20G`, one build job, offline dependencies, `--locked`, and the
external target
`(deleted-coord-artifact)`.
The target occupied `1,199,374,430` bytes at the end of the last command, well
below the 20 GB lane cap. No owned Cargo, rustc or rustdoc process remains.

The root lock was not regenerated. The pre-existing ignored lock was preserved
at
`(deleted-coord-artifact)/Cargo.lock.before-6d7895e6056c931fc19780ff4d0a031230e2ac79b9b30a9060bc8b6fad6747da`:
139,592 bytes, mode `664`, SHA-256
`6d7895e6056c931fc19780ff4d0a031230e2ac79b9b30a9060bc8b6fad6747da`.
The installed root lock is the verified 52ed acceptance lock: 139,591 bytes,
mode `664`, SHA-256
`52ed0cacac15b5e440349bfc0f4aca2665bcf50d59d4321624e596bb410783c6`.
The standalone harness retained its separate tracked lock, 30,851 bytes,
SHA-256 `7e80d6b7256ae348253671373934172c9139cece5b9c7ff14a20d20367e938d9`.
The complete pre/post lock receipt is outside the repository at
`.../drawingml-diagram-raw-validation-20261024/source-lock-prepost.txt`.

The exact commands were:

```text
env PATH=(deleted-coord-artifact):$PATH LITCHI_MEM_MAX=20G CARGO_NET_OFFLINE=true CARGO_TARGET_DIR=(deleted-coord-artifact) TMPDIR=(deleted-coord-artifact) CARGO_BUILD_JOBS=1 CARGO_INCREMENTAL=0 CARGO_PROFILE_DEV_DEBUG=0 CARGO_PROFILE_TEST_DEBUG=0 RUST_TEST_THREADS=1 cargo test -p litchi-drawingml --all-features --offline --locked --lib --tests --no-fail-fast --jobs 1
env PATH=(deleted-coord-artifact):$PATH LITCHI_MEM_MAX=20G CARGO_NET_OFFLINE=true CARGO_TARGET_DIR=(deleted-coord-artifact) TMPDIR=(deleted-coord-artifact) CARGO_BUILD_JOBS=1 CARGO_INCREMENTAL=0 CARGO_PROFILE_DEV_DEBUG=0 CARGO_PROFILE_TEST_DEBUG=0 RUST_TEST_THREADS=1 cargo check -p litchi-docx -p litchi-pptx -p litchi-ops --all-features --offline --locked --lib --jobs 1
env PATH=(deleted-coord-artifact):$PATH LITCHI_MEM_MAX=20G CARGO_NET_OFFLINE=true CARGO_TARGET_DIR=(deleted-coord-artifact) TMPDIR=(deleted-coord-artifact) CARGO_BUILD_JOBS=1 CARGO_INCREMENTAL=0 CARGO_PROFILE_DEV_DEBUG=0 CARGO_PROFILE_TEST_DEBUG=0 RUST_TEST_THREADS=1 cargo check --manifest-path tools/perf-baseline/unified-smartart/Cargo.toml --all-features --offline --locked --jobs 1
```

Results and external log artifacts:

All three logs are in
`(deleted-coord-artifact)`.

| Check | Result | External log (SHA-256 of final file) |
| --- | --- | --- |
| `litchi-drawingml --all-features --lib --tests` | 252 unit tests and all integration suites passed | `drawingml-tests.log` — `f85db1067116ea610bdc33ae3d3bade344a0bb7db207943fec543dd70609bb0a` |
| `litchi-docx`, `litchi-pptx`, `litchi-ops` all-feature library check | passed | `consumers-check.log` — `0d4fff932051de2fd48205e5a6443e25bd5a0836b4630e2601d0f103a0310cc8` |
| standalone `unified-smartart` all-feature check | passed | `smartart-harness-check.log` — `4ef97d659b340bb0e2bbb69a3d5a48e0dd9e41b1cd1f95a571ee2313660b18c7` |

The logs, source/lock receipt and target are intentionally outside the
repository; no large evidence file or binary is committed. The source-lock
receipt SHA-256 is
`133afb0404a38a5422f1a0441e7ffaacf1acf4e47c480c08f275330cf3e8158e`.
The earlier `025ae`, `8f49re` and `e980` root-lock-less checks remain useful
diagnostics, but they have unknown/old lock provenance and are not 52ed
acceptance evidence. This packet is the accepted focused validation for
`9786dcf450`.
