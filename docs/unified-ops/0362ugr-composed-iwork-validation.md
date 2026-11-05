# 0362ugr — composed iWork validation

- Date: 2026-10-24.
- Status: source and correctness validation receipt; no native performance
  qualification or acceptance is claimed.
- Final source: `a0bf134246bb19f171d34f61822aa072d30a3fd7`.

This receipt covers the composed Pages, Numbers and Keynote lazy-entry source,
the affected iWork package readers, the selected unified facade consumers, the
operations check, the standalone fuzz package and the minimal unified-batch
harness.  It is intentionally a source-only packet.  Native ABBA results,
memory qualification and the final workspace gate remain coordinator work.

## Source boundary

The composed integration started at
`3ae5fdf0ff2afeeb0c4084a1f0f6875e01d7803b`.  The source was then advanced by
the following reviewed consumer and validation fixes, in order:

```text
ebf6dd0514  fix(pages): preserve shared footnote patch arc
8a1afeff07  fix(numbers): remove composed test qualification
177d007092  test(numbers): use a valid merge region in lazy-entry coverage
b22e1a9cea  fix(keynote): map deferred selector limits in movie captions
131bac621b  test(keynote): separate composed edit test modules
4c323ae060  test(iwa): resolve deferred header footer selectors on access
dfdd87ca91  fix(pages): retain lazy ledgers and redact table names
57b14d3f2a  fix(pages): reuse appearance and merge edit ledgers
5ac9b25891  fix(ops): update OLE metadata benchmark consumer
a0bf134246  fix(fuzz): handle lazy slide text position errors
```

The final source pin is the complete checkout at `a0bf134246...`, not a
partial diff or a generated source archive.  The Pages, Numbers, Keynote and
IWA all-target runs necessarily finished at their respective source pins
before the later unrelated consumer-only fixes: Pages at
`57b14d3f2a95081a07512feb8c7da38efe00c47b`, Numbers at
`177d0070928d70f5c9055b0a0f6608658385921c`, Keynote at
`131bac621b21a6182fb439c00ca9794fc650de15`, and IWA at
`4c323ae060bb0755080ced56cf6ebc6b49a294b1`.  The unchanged package suites were not repeated after the later consumer-only
fixes; the Pages-only strict Clippy, operations, fuzz and harness checks below
cover the later deltas where they apply.

The `5ac9b25891` OLE metadata example matches the APIs present in this
assembly: DOC property access is fallible, while XLS and PPT property access
is direct. The coordinator checked these actual signatures during integration.
The queued PPT lazy-entry packet makes PPT property access fallible and must
update that consumer when it merges; XLS remains direct. This is not a native
qualification claim.

## Reproducible environment

Every Cargo command below began with this exact source command, in the stated
worktree:

```sh
source /home/zhuhe/code/litchi-worktrees/.coord/pages-lazy-entry/env.sh
```

The environment file is SHA-256
`5f9a0886c7b7cd7f52c52942bf696b5748e2fa99c273523b1a10a2be3503b7e1` and sets
`CARGO_BUILD_JOBS=1`, `RUST_TEST_THREADS=1`, `CARGO_INCREMENTAL=0`, offline
Cargo, zero debug information, `LITCHI_MEM_MAX=20G`, target directory
`/home/zhuhe/code/litchi-worktrees/.coord/pages-lazy-entry/target`, and
`TMPDIR=/home/zhuhe/code/litchi-worktrees/.coord/pages-lazy-entry/tmp`.

The pinned toolchain is Rust `rustc 1.95.0 (59807616e 2026-04-14)` and Cargo
`1.95.0 (f2d3ce0bd 2026-03-21)`, from `rust-toolchain.toml`.  Configuration
identity checks against the final source were:

| File | SHA-256 |
| --- | --- |
| `Cargo.toml` | `f5ec007892023b38f1deaef1ce20714db58b58448a3bea088be3b847b84570de` |
| `Cargo.lock` | `6d7895e6056c931fc19780ff4d0a031230e2ac79b9b30a9060bc8b6fad6747da` |
| `rust-toolchain.toml` | `e3a213e0d222e94d213cafbc20932eb3f76c643b4dd63756acf95192df2aa310` |
| `.cargo/config.toml` | `50b044c83828f2c448ee59c362afaa79de2dd47fbd390f0cd2888391f1d1119d` |
| `clippy.toml` | `f11517bc5872d96aab42ba6519db4e78f2fdcd44512ac70a1b45adbd9929b50c` |

The minimal harness lock used by the final unified-batch commands is tracked
at `tools/perf-baseline/unified-batch-closure/Cargo.lock`, SHA-256
`95aed16c860cd6082cd50a39aabf29779b89bdb004c8263674e042d84a007870`.
The standalone fuzz manifest has no tracked lock.  Its offline, unlocked
check generated an ignored lock at
`crates/litchi/fuzz/Cargo.lock`, SHA-256
`c499d5efe6bba3991672c52f75f78a39a294a5b80e98f62a6ee2bb28dd2a2eb6`; it is
not part of this commit and no dependency lock upgrade was committed.

The sparse checkout was expanded only to expose the already-tracked fixture
`media/litchi_logo.png` (711,747 bytes,
`324062cd9a3e4375fda6fe80438162d1978e2b83`).  No `3rdparty/` path was used.
The capped target directory closed at exactly **9,078,733,643 bytes** by
`du -sb`, below the assigned 20 GB target ceiling.

## Successful package tests

Each result count is the complete output count for the command, including
ignored tests.  The recorded log paths are outside the repository and are
closed artifacts.

### Pages

```sh
source /home/zhuhe/code/litchi-worktrees/.coord/pages-lazy-entry/env.sh
cargo test --offline --locked -p litchi-pages --all-features --all-targets --no-fail-fast --message-format=short -j 1
```

Source: `57b14d3f2a95081a07512feb8c7da38efe00c47b`.  Result: 61 result
blocks, **719 passed, 0 failed, 0 ignored**.  Log:
`/home/zhuhe/code/litchi-worktrees/.coord/pages-lazy-entry/combined-pages-tests-v4.log`
(75,020 bytes, SHA-256
`d8e6eb3cad4471a315298822089e83fb0cdc814becfac54ff17708ce01208ba9`).

### Numbers

```sh
source /home/zhuhe/code/litchi-worktrees/.coord/pages-lazy-entry/env.sh
cargo test --offline --locked -p litchi-numbers --all-features --all-targets --no-fail-fast --message-format=short -j 1
```

Source: `177d0070928d70f5c9055b0a0f6608658385921c`.  Result: 87 result
blocks, **1,081 passed, 0 failed, 4 ignored**.  Log:
`/home/zhuhe/code/litchi-worktrees/.coord/pages-lazy-entry/combined-numbers-tests-v3.log`
(119,073 bytes, SHA-256
`3fe2a3a7f2b6c984af4b9dde1ea0d8852f8ded6a01edc6d437c39e6d5a46b91c`).

### Keynote

```sh
source /home/zhuhe/code/litchi-worktrees/.coord/pages-lazy-entry/env.sh
cargo test --offline --locked -p litchi-keynote --all-features --all-targets --no-fail-fast --message-format=short -j 1
```

Source: `131bac621b21a6182fb439c00ca9794fc650de15`.  Result: 103 result
blocks, **1,115 passed, 0 failed, 1 ignored**.  Log:
`/home/zhuhe/code/litchi-worktrees/.coord/pages-lazy-entry/combined-keynote-tests-v3.log`
(118,228 bytes, SHA-256
`3edc37a582684d85471570534021fa79132f7305de614b39add6202b993305d6`).

### IWA

```sh
source /home/zhuhe/code/litchi-worktrees/.coord/pages-lazy-entry/env.sh
cargo test --offline --locked -p litchi-iwa --all-features --all-targets --no-fail-fast --message-format=short -j 1
```

Source: `4c323ae060bb0755080ced56cf6ebc6b49a294b1`.  Result: 225 result
blocks, **1,780 passed, 0 failed, 0 ignored**.  Log:
`/home/zhuhe/code/litchi-worktrees/.coord/pages-lazy-entry/combined-iwa-tests-v3.log`
(244,195 bytes, SHA-256
`3f044fd6b705327e87ee57abfd3db839ed95f4598c4043b837d28811186aa19e`).

### Published doctests

```sh
source /home/zhuhe/code/litchi-worktrees/.coord/pages-lazy-entry/env.sh
cargo test --offline --locked -p litchi-pages -p litchi-numbers -p litchi-keynote -p litchi-iwa --all-features --doc --no-fail-fast --message-format=short -j 1
```

Source: `4c323ae060bb0755080ced56cf6ebc6b49a294b1`.  Result: 4 crate result
blocks, **96 passed, 0 failed, 2 ignored**.  Log:
`/home/zhuhe/code/litchi-worktrees/.coord/pages-lazy-entry/combined-iwork-doctests-v1.log`
(11,634 bytes, SHA-256
`dc1fc9fb12d4981b1abf4bf4f395c7ab74032917f5a654f5cf04e766c5413c56`).

## Strict checks and consumers

### Combined iWork Clippy

```sh
source /home/zhuhe/code/litchi-worktrees/.coord/pages-lazy-entry/env.sh
cargo clippy --offline --locked -p litchi-pages -p litchi-numbers -p litchi-keynote -p litchi-iwa --all-targets --all-features -j 1 -- -D warnings
```

Source: `4c323ae060bb0755080ced56cf6ebc6b49a294b1`.  Exit 0 with 22
check/compile lines.  Log:
`/home/zhuhe/code/litchi-worktrees/.coord/pages-lazy-entry/combined-iwork-clippy-v1.log`
(1,832 bytes, SHA-256
`88c527ae5845ae71097d667812aad3d6cf6a6bf906bdce8d934d2cb46e3642ed`).

The later Pages ledger/redaction delta was checked separately at the final
Pages source:

```sh
source /home/zhuhe/code/litchi-worktrees/.coord/pages-lazy-entry/env.sh
cargo clippy --offline --locked -p litchi-pages --all-targets --all-features -j 1 -- -D warnings
```

Source: `57b14d3f2a95081a07512feb8c7da38efe00c47b`.  Exit 0.  Log:
`/home/zhuhe/code/litchi-worktrees/.coord/pages-lazy-entry/pages-clippy-v2.log`
(899 bytes, SHA-256
`a66a5d2bdf7f429c1a44a4dec732ee3feaf6be5e0ae95a8688fd266c4369e9ea`).

### Unified facade

```sh
source /home/zhuhe/code/litchi-worktrees/.coord/pages-lazy-entry/env.sh
cargo test --offline --locked -p litchi --all-features \
  --test iwork_file \
  --test iwork_merge_readers \
  --test iwork_merge_transactions \
  --test iwork_pages_table_deletion \
  --test keynote_facade \
  --test numbers_facade \
  --test numbers_semantic_values \
  --test numbers_semantic_view \
  --test pages_facade \
  --test semantic_facade \
  --test semantic_reader_facade \
  --no-fail-fast --message-format=short -j 1
```

Source: `4c323ae060bb0755080ced56cf6ebc6b49a294b1`.  Result: 11 result
blocks, **50 passed, 0 failed, 0 ignored**.  Log:
`/home/zhuhe/code/litchi-worktrees/.coord/pages-lazy-entry/combined-litchi-facade-tests-v1.log`
(23,412 bytes, SHA-256
`08f58f21e6206801dfbd84b5d02b8d3efcfd5908e72e4d5d4ac033abfd59d629`).
This selected consumer run predates the final Pages ledger-only and fuzz-only
commits; it is not presented as a post-`a0bf` whole-facade run.

### Operations target

```sh
source /home/zhuhe/code/litchi-worktrees/.coord/pages-lazy-entry/env.sh
cargo check --offline --locked -p litchi-ops --all-targets --all-features -j 1
```

Source: `5ac9b25891b1c4512f328aa0005c2d3f31bb756b`.  Exit 0 with one check
line.  Log:
`/home/zhuhe/code/litchi-worktrees/.coord/pages-lazy-entry/ops-check-v2.log`
(167 bytes, SHA-256
`44a638ee4929b8c234f838bec2c11390a17202abb05ec2c8118fc715529f0e0b`).

### Standalone fuzz package

The fuzz package has no tracked lock, so this command intentionally omits
`--locked`; it remains offline and single-job:

```sh
source /home/zhuhe/code/litchi-worktrees/.coord/pages-lazy-entry/env.sh
cargo check --offline --manifest-path crates/litchi/fuzz/Cargo.toml --all-targets -j 1
```

Source: `a0bf134246bb19f171d34f61822aa072d30a3fd7`.  Exit 0.  The only
warnings are the three pre-existing unused media constants
`REPLACED_MOVIE_BYTES`, `REPLACED_POSTER_BYTES` and `REPLACED_AUDIO_BYTES` in
`fuzz_targets/keynote_slide_media_data_seed.rs`; the lazy slide-text `Result`
warning was fixed in `a0bf`.  Log:
`/home/zhuhe/code/litchi-worktrees/.coord/pages-lazy-entry/litchi-fuzz-check-v3.log`
(1,049 bytes, SHA-256
`801da79e5e730b965f07c9fecfcb61ab9ab08901c7cd9177b265d1d6bc4a778c`).

### Minimal unified-batch harness

```sh
source /home/zhuhe/code/litchi-worktrees/.coord/pages-lazy-entry/env.sh
cargo check --offline --locked --manifest-path tools/perf-baseline/unified-batch-closure/Cargo.toml --no-default-features --features pages,keynote,numbers --all-targets -j 1
```

Source: `a0bf134246bb19f171d34f61822aa072d30a3fd7`.  Exit 0.  Log:
`/home/zhuhe/code/litchi-worktrees/.coord/pages-lazy-entry/unified-batch-iwork-check-v2.log`
(60 bytes, SHA-256
`dede40800e26c8c8f3fbde8303ac3d28a9c607bcd68871b4b59a36596409a3cf`).

The corresponding strict check was:

```sh
source /home/zhuhe/code/litchi-worktrees/.coord/pages-lazy-entry/env.sh
cargo clippy --offline --locked --manifest-path tools/perf-baseline/unified-batch-closure/Cargo.toml --no-default-features --features pages,keynote,numbers --all-targets -j 1 -- -D warnings
```

Source: `a0bf134246bb19f171d34f61822aa072d30a3fd7`.  Exit 0.  Log:
`/home/zhuhe/code/litchi-worktrees/.coord/pages-lazy-entry/unified-batch-iwork-clippy-v1.log`
(210 bytes, SHA-256
`65716b8aae6840c1087a81b20a3f1e66bea35b46bd1eaed9e76d5bd6ba753f74`).

## Diagnostic attempts and resolved issues

The following are retained as diagnostics only; none is a final failure.
Each was followed by a narrow source or fixture correction and a successful
receipt above.

| Diagnostic log | Size / SHA-256 | Diagnosis and resolution |
| --- | --- | --- |
| `/home/zhuhe/code/litchi-worktrees/.coord/pages-lazy-entry/combined-pages-tests.log` | 1,042 bytes / `f65799a70d798235d9024cf678a45a2723c83e06b91528a6ab9b5bec4502a4dd` | Pages body-footnote Arc type mismatch; fixed by retaining the shared `Arc<Footnote>`. |
| `/home/zhuhe/code/litchi-worktrees/.coord/pages-lazy-entry/combined-pages-tests-v3.log` | 890 bytes / `c15fbfd0ccef3b31be775ff9e6248e2cb295c2580d685c418450a79dc9a571ba` | Lazy-ledger tests accessed private `WireBudget` fields; assertions were changed to the typed payload-work accessors. |
| `/home/zhuhe/code/litchi-worktrees/.coord/pages-lazy-entry/combined-numbers-tests-v2.log` | 119,322 bytes / `f25737c0fcbc6b72754083c6929b1a070be06c84d5d04e1aa8bab78b5a718040` | The new retry test used a single-cell `Region::new(0, 0, 1, 1)` where a merge region was required; the test now uses a valid 1×2 region. |
| `/home/zhuhe/code/litchi-worktrees/.coord/pages-lazy-entry/combined-keynote-tests-v1.log` | 894 bytes / `5f1ac0c11e622c7b25dc89def87697c5e3803e0dcef4bd85e3716a51586febb1` | Missing `SelectorBytes` match in the movie-caption limit mapping; fixed in the Keynote consumer. |
| `/home/zhuhe/code/litchi-worktrees/.coord/pages-lazy-entry/combined-keynote-tests-v2.log` | 321 bytes / `f0fd0bae8993bab19f7dc178572f98429c8e65e427fcdf20b39300c72384b8f1` | Composed edit test modules both declared `tests`; separated the modules. |
| `/home/zhuhe/code/litchi-worktrees/.coord/pages-lazy-entry/combined-iwa-tests-v1.log` | 1,153 bytes / `cb9a5f06fd291ebca4258415a4049756edba5f62ff3b7a30a650c680f3399ff2` | Sparse checkout omitted the already-tracked `media/litchi_logo.png`; the sparse path was restored without adding a fixture. |
| `/home/zhuhe/code/litchi-worktrees/.coord/pages-lazy-entry/combined-iwa-tests-v2.log` | 244,892 bytes / `0db9095ec43908da55259ba8c8f524a8663bc1bc89d976dc81580833530478e6` | Existing header/footer assertion still expected eager construction to reject an invalid selector; it now resolves the deferred selector on access. |
| `/home/zhuhe/code/litchi-worktrees/.coord/pages-lazy-entry/ops-check-v1.log` | 12,727 bytes / `63b5831de816f8529a65c7964961055fac75452eec8536aa6690545ca1625f38` | The OLE example called `.expect()` on the still-infallible XLS/PPT getters; `5ac9` removed those calls. The queued PPT packet must restore its own fallible access when merged. |
| `/home/zhuhe/code/litchi-worktrees/.coord/pages-lazy-entry/litchi-fuzz-check-v1.log` | 279 bytes / `05aec1489b55fdccb30b9e8a8ac138785890090e04d9190ff79893267fb7dd49` | `--locked` correctly refused because the standalone fuzz manifest has no tracked lock; the final check used offline unlocked dependency resolution and kept its generated lock ignored. |

The old fuzz v2 check was successful but emitted the three pre-existing media
constant warnings and one wave-introduced unused `Result` warning.  Its log is
retained for provenance at
`/home/zhuhe/code/litchi-worktrees/.coord/pages-lazy-entry/litchi-fuzz-check-v2.log`
(5,464 bytes, SHA-256
`4521fff1dedc537d208ebae33c335c5581ed373400ee3c61312d18ed458afd9a`).
Only the three media warnings remain in the final v3 check.

## Boundary and disposition

No benchmark, allocation profile or native timing result is included here.
The redaction audit (memo item 22) remains the sole declared audit deferral;
the changed diagnostics and errors retain content-free output.  The coordinator must still migrate PPT metadata access with the queued PPT
packet, run the facade checks required after assembly, and perform native
qualification and the workspace gate on the final integrated source.
