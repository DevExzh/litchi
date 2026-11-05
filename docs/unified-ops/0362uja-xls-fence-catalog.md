# 0362uja — Share the XLS catalog without weakening source fences

The final XLS gate exposed two version observations for a warm
`/sheet[1]` read: the resolver's path lookup called the package catalog, then
the resolver's required trailing fence called `Xls::check_source` again. The
trailing fence is part of the source-change precedence contract and remains in
place. The repair keeps the package projection checked on every direct access,
while allowing nested adapter reads in one active fence to reuse the immutable
catalog already captured by that checked access.

`litchi_xls::Package::sheet_catalog` and `Workbook::sheet_catalog` now return a
shared `Arc<[SheetMetadata]>`. The eager raw reader stores its parsed catalog
in that same immutable representation; parsing uses `Arc::make_mut` only while
the reader is still being assembled, and the package's first catalog access
shares the reader allocation instead of cloning every descriptor. The
source-backed package converts its lazily obtained metadata once and shares it
through the package and adapter. `XlsState` keeps an adapter-local lazy
`OnceLock` of that Arc. It is consulted only while the handle's read fence is
active; direct package/native catalog getters still perform their source check
on every call. Constructors remain O(1), with no eager catalog scan, reopen, or
unchecked getter.

`SheetHandle` owns the shared catalog and its position, so it remains valid
for the duration of a read without borrowing a catalog that could be replaced.
The source-backed handle's cold catalog read, repeated direct native getter,
source change before the first read, source change at the trailing check, and
stale path-miss precedence are covered by focused tests. No source fence or
error mapping was relaxed.

Validation on this leaf:

```text
cargo fmt --all -- --check                                      PASS
cargo test --offline --locked -p litchi-ops --all-features \
  --lib format::xls::tests:: -- --nocapture                    PASS (14/14)
cargo test --offline --locked -p litchi-xls --test package_owner \
  --no-fail-fast -- --nocapture                                PASS (27/27)
cargo clippy --offline --locked -p litchi-xls --all-targets \
  --all-features -- -D warnings                                 PASS
cargo clippy --offline --locked -p litchi-ops --all-targets \
  --all-features -- -D warnings                                 PASS
```

The commands used one capped Cargo process at a time with `CARGO_BUILD_JOBS=1`,
`RUST_TEST_THREADS=1`, `LITCHI_MEM_MAX=20G`, and an external target under
`(deleted-coord-artifact)`; no build artifact or executable is committed. Exact
logs remain outside Git:

```text
(deleted-coord-artifact)/ops-xls-tests.log
  sha256 eeb76b42fc1f4b8c13893e2a6b37c0622ae0f6b251661437ccf8db85628cbdab
(deleted-coord-artifact)/xls-package-owner-tests.log
  sha256 05b8e7e057ef16ba313da8fec17acad2e748c0d276e32a70545c3fd0bd50d663
(deleted-coord-artifact)/xls-clippy.log
  sha256 159c95084c74469e58793af212409793e85a7eb33ab10bebe1f7928ea6fcf10f
(deleted-coord-artifact)/ops-clippy.log
  sha256 0d97a524fdb8906fe7a98e531573b75111d961771f4d8050e76b4d3e34560946
```

Native and ABBA qualification remain coordinator work. The full workspace gate
is not claimed by this leaf.

## Full native validation receipt

At the coordinator's request, the complete native XLS suite and
doctests were rerun on `dfe8281fb36a920a02b6ed150bf604663ea602fe` with the
same one-job capped environment. Both commands passed:

```text
env PATH=(deleted-coord-artifact):$PATH \
  LITCHI_MEM_MAX=20G CARGO_BUILD_JOBS=1 RUST_TEST_THREADS=1 \
  CARGO_INCREMENTAL=0 CARGO_PROFILE_DEV_DEBUG=0 CARGO_PROFILE_TEST_DEBUG=0 \
  CARGO_PROFILE_RELEASE_DEBUG=0 CARGO_NET_OFFLINE=true \
  CARGO_TARGET_DIR=(deleted-coord-artifact) \
  TMPDIR=(deleted-coord-artifact) \
  cargo test --offline --locked -p litchi-xls --all-features \
    --all-targets --no-fail-fast
  status: 0; 1,327 unit tests and 556 integration tests passed; examples passed
  one existing evidence-workbook writer test remains ignored

env PATH=(deleted-coord-artifact):$PATH \
  LITCHI_MEM_MAX=20G CARGO_BUILD_JOBS=1 RUST_TEST_THREADS=1 \
  CARGO_INCREMENTAL=0 CARGO_PROFILE_DEV_DEBUG=0 CARGO_PROFILE_TEST_DEBUG=0 \
  CARGO_PROFILE_RELEASE_DEBUG=0 CARGO_NET_OFFLINE=true \
  CARGO_TARGET_DIR=(deleted-coord-artifact) \
  TMPDIR=(deleted-coord-artifact) \
  cargo test --offline --locked -p litchi-xls --all-features \
    --doc --no-fail-fast
  status: 0; 23 doctests passed, 1 expected ignored, 1 compile-fail doctest passed
```

The combined log is retained outside Git at
`(deleted-coord-artifact)/xls-full-tests-doctests.log`
(172,749 bytes, SHA-256
`1fa5a7299d317b4f8f6920977dde1eab5b4a977c1af9fcb83b83ff84208939fd`). The
external target measured 4,531,008,467 bytes (4.3 GiB), below the 20 GiB cap.
The locked `/home/zhuhe/code/litchi-worktrees/uo2-final-xls-fence/Cargo.lock`
was 139,591 bytes with SHA-256
`52ed0cacac15b5e440349bfc0f4aca2665bcf50d59d4321624e596bb410783c6`.

This validates native correctness after moving the parsed catalog into its
shared representation. The first-access benefit is structural—the eager
package no longer clones the catalog—but the allocation/conversion placement
and the Arc handle cost remain pending ABBA qualification; this receipt makes
no performance acceptance claim.

The coordinator independently verified the combined log's size and SHA-256.
Across both commands, 1,907 tests pass and two existing cases remain ignored:
`write_every_feature` writes an evidence workbook, and the builder filter
example is an ignored doctest. No new ignored case was introduced.
