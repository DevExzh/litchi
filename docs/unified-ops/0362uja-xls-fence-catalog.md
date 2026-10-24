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
`.coord/xls-fence-target`; no build artifact or executable is committed. Exact
logs remain outside Git:

```text
/home/zhuhe/code/litchi-worktrees/.coord/xls-fence-tmp/ops-xls-tests.log
  sha256 eeb76b42fc1f4b8c13893e2a6b37c0622ae0f6b251661437ccf8db85628cbdab
/home/zhuhe/code/litchi-worktrees/.coord/xls-fence-tmp/xls-package-owner-tests.log
  sha256 05b8e7e057ef16ba313da8fec17acad2e748c0d276e32a70545c3fd0bd50d663
/home/zhuhe/code/litchi-worktrees/.coord/xls-fence-tmp/xls-clippy.log
  sha256 159c95084c74469e58793af212409793e85a7eb33ab10bebe1f7928ea6fcf10f
/home/zhuhe/code/litchi-worktrees/.coord/xls-fence-tmp/ops-clippy.log
  sha256 0d97a524fdb8906fe7a98e531573b75111d961771f4d8050e76b4d3e34560946
```

Native and ABBA qualification remain coordinator work. The full workspace gate
is not claimed by this leaf.
