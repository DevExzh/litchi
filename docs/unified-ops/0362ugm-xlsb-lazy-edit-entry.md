# 0362ugm — XLSB lazy edit entry

This slice closes the XLSB edit-construction cost debt.  Package-owned edits
retain only an immutable package owner and their finite policy at entry; source
catalogs, worksheet drafts, native owner snapshots, OPC graph images and exact
before-images are prepared on the first operation or commit that needs them.
Dropping an untouched edit therefore does not scan or copy package payloads.

The deferred paths cover the structural workbook and cell-values families,
generic `part_edit::Staged` families, Calculation Chain, Data Model, Volatile
Dependencies, Custom XML Maps, connections, shared-workbook metadata, Theme,
Custom Data, and threaded comments.  Fallible accessors retain the existing
typed errors where a read-only staged projection must materialize a source
snapshot; no dummy values or panic-based lazy state are introduced.

The package cell-values entry owns a selector rather than retaining a
resolved worksheet URI. `Package::{edit_cell_values,
edit_cell_values_with_options}` therefore return a cold `Edit`; callers that
need the missing-tab capability call `Edit::resolve()`, which returns
`Ok(None)` for a missing tab and performs the catalog scan only at that first
fallible resolution. Ordinary staging operations and `commit()` resolve the
selector on demand.

The existing replay, exact no-op, inverse, signature-policy, bounded-limit and
publication behavior remains on the same first-use/commit paths.  Constructor
and drop tests prove that a cold package projection remains cold for the
structural and generic staged owners.

## Validation receipt

The repair and consumer assertions are committed at
`51c86284467a8dd834086e428acf8692781fbe9c`. The functional implementation for
every command below is this source commit. The standalone harness run
overlapped the two-file documentation-only overlay described below; it is not
claimed to have run in an exactly clean checkout. The other commands completed
before that overlay. The assigned
environment was
`/home/zhuhe/code/litchi-worktrees/.coord/xlsb-lazy-entry/env.sh` (SHA-256
`47251384c135d438f1fc4324038ecc11259b652d21d3c1d97ef4e0f5b60f2686`), with
`CARGO_BUILD_JOBS=1`, `RUST_TEST_THREADS=1`, `CARGO_INCREMENTAL=0`, offline
locked resolution, and the external target directory it defines.  The
toolchain was `rustc 1.95.0 (59807616e 2026-04-14)` and `cargo 1.95.0
(f2d3ce0bd 2026-03-21)`.  The workspace lock SHA-256 was
`6d7895e6056c931fc19780ff4d0a031230e2ac79b9b30a9060bc8b6fad6747da`; the
standalone `tools/perf-baseline` lock SHA-256 was
`4def018e9c1d5f409b8c60d6cc0b8f114c486ab3f8fed2c7a4e9c7baf21c5184`.

The final target occupied 12,379,273,678 bytes, below the assigned target
cap of 20,000,000,000 bytes (decimal 20 GB).  This is distinct from the
separate `LITCHI_MEM_MAX=20G` process-memory cap.

| Command (all rows use source `51c86284467a8dd834086e428acf8692781fbe9c`) | Result | Closed log and SHA-256 |
|---|---:|---|
| `cargo test --offline --locked -p litchi-xlsb --all-features` | 1,355 passed, 0 failed, 11 ignored (including 40 ordinary and 2 compile-fail doctests) | `.coord/xlsb-lazy-entry/receipts/full-xlsb-all-features-r7.log`, 128,061 bytes, `456525dc4bceb536d7acb6e70831d4acb13b8490b973cabdbb1e2dc9dbba8a4a` |
| `cargo clippy --offline --locked -p litchi-xlsb -p litchi-ops -p litchi -p litchi-py --all-targets --all-features -- -D warnings` | exit 0 | `.coord/xlsb-lazy-entry/receipts/strict-xlsb-consumer-clippy-r2.log`, 18,503 bytes, `98430e245113fdd1f00cff9dbd00296964d4a24394f0dde088c8f45edbbf05c4` |
| All 14 XLSB `litchi-ops` integration targets | 126 passed, 0 failed | `.coord/xlsb-lazy-entry/receipts/affected-xlsb-consumers-r3.log`, 12,774 bytes, `d0f8391866f9cb26e6d8ea17dd6fae8840e67907ace4bbd462c45f40c2330f9e` |
| `cargo test --offline --locked -p litchi --all-features --test xlsb_facade` | 23 passed, 0 failed | `.coord/xlsb-lazy-entry/receipts/focused-litchi-xlsb-facade-r2.log`, 1,998 bytes, `eca6496ca5217b6bff558b38765b8a417b5eede7a9f66f256b1cc5dc422e5707` |
| `cargo test --offline --locked --manifest-path tools/perf-baseline/Cargo.toml --features xlsb-crud --bin xlsb_crud` | 9 passed, 0 failed | `.coord/xlsb-lazy-entry/receipts/standalone-xlsb-crud-r1.log`, 9,284 bytes, `fe04c5c9095b77ab5ded6482a758f3d87f1593c8c96821ed576df06677ffe50d` |

The exact 14-target consumer command was:

```text
cargo test --offline --locked -p litchi-ops --all-features \
  --test xlsb_batch_remove_coalesce \
  --test xlsb_cell_writes \
  --test xlsb_create \
  --test xlsb_features \
  --test xlsb_p0_writes \
  --test xlsb_remove \
  --test xlsb_sheet_add \
  --test xlsb_sheet_batch \
  --test xlsb_sheet_name_dependencies \
  --test xlsb_sheet_names \
  --test xlsb_sheet_order \
  --test xlsb_sheet_remove \
  --test xlsb_visibility \
  --test xlsb_writes
```

The `affected-xlsb-consumers-r3` and `focused-litchi-xlsb-facade-r2`
receipts explicitly bind those runs to `51c8628...`; neither was inferred
from a later documentation commit or a dirty overlay.  An earlier historical
r7 sidecar was captured before its final `cargo_exit=0\n` footer (128,048
bytes; SHA-256 beginning `099650...`).  That pre-footer identity remains
historical; the closed log and table above intentionally use the final
128,061-byte file and its final SHA-256.

The coordinator observed the slicer/timeline documentation overlay during the
standalone harness run: the two source files were modified at 02:54:15 UTC,
and its log closed at 02:59:21 UTC. Strict Clippy had completed at 02:52:17 UTC.
The exact compiler-read timing was not captured. The overlay changes only
comments about lazy construction and is committed in `c7aecdb536`; its diff
from `51c8628446` over `src/slicer/edit.rs` and `src/timeline/edit.rs` is
1,912 bytes, SHA-256
`4f52dc54271dcfc1630c27c87119c79d276ec25de21d7c31ff485ef3ebcf6c13`. Regenerate with
`git diff 51c8628446 c7aecdb536 -- crates/litchi-xlsb/src/slicer/edit.rs crates/litchi-xlsb/src/timeline/edit.rs`.

The standalone harness emitted its existing `xls_comment_opaque_stream_name`
dead-code warning; it did not affect the zero exit status.  Scoped rustfmt,
`git diff --check`, and the repository large-evidence/binary guards were also
clean.  The logs and generated target remain outside the repository.
