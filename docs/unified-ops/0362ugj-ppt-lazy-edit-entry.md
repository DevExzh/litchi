# 0362ugj — PPT lazy edit entry and consumer receipt

This receipt closes the PPT lazy-entry validation and the two standalone
consumer gaps found while assembling the final evidence.  It makes no native
timing, ABBA, allocation, or performance-qualification claim.

## Source boundary

The final clean source tip is
`c7931891205aed2af0336c121848c95f90c3ba87` (`uo2/ppt-lazy-entry`).  Its
history retains the validated PPT implementation and consumer fixes:

* `e4bc2f967e` — compiler migration for lazy PPT edit construction;
* `e527fff7e7` — deferred projection refusal and deterministic source-fence
  regressions;
* `fa2abab399` — strict Clippy migration for fallible projections;
* `ca125d2ba1f52692f42efc17adc4b0990dd41fac` — final PPT facade and ops
  consumer migration;
* `2274f4cff95160397e304f5805d8fa4cbee258ca` — standalone PPTX slide-boundary
  consumer migration;
* `1cb09b9c4ac3513bf628fa90a23e5485cdb68690` — ODG joined-probe consumer
  migration and this receipt.
* `c7931891205aed2af0336c121848c95f90c3ba87` — refreshed ODG joined-probe
  source bindings.

The PPT crate and its facade consumers are byte-identical between `ca125d2ba1`
and the final tip.  Therefore the PPT tests, doctests, Clippy check, and
focused facade tests below remain pinned to `ca125d2ba1`; the later commits
change only `tools/perf-baseline`, and those affected standalone consumers are
checked at the final tip below.  No PPT suite was rerun after the consumer-only
commits.

## Reproducibility identities

All Cargo commands sourced
`/home/zhuhe/code/litchi-worktrees/.coord/ppt-lazy-entry/env.sh` (487 bytes,
SHA-256
`5dc0dc3da7cb494e3b5356360f19e973b9b4611e3b9d1d8d927b2055df652676`).  The
effective settings were the memory-capped Cargo shim, `LITCHI_MEM_MAX=20G`,
`CARGO_BUILD_JOBS=1`, `RUST_TEST_THREADS=1`, `CARGO_INCREMENTAL=0`, all three
Cargo debug settings set to `0`, `CARGO_NET_OFFLINE=true`, and
`TMPDIR=/home/zhuhe/code/litchi-worktrees/.coord/ppt-lazy-entry/tmp`.
The assigned target was
`/home/zhuhe/code/litchi-worktrees/.coord/ppt-lazy-entry/target`; its final
size was 6,143,757,097 bytes.  No target, overlay, binary, or raw log is
tracked.

The toolchain was `rustc 1.95.0 (59807616e 2026-04-14)`
(`x86_64-unknown-linux-gnu`, LLVM 22.1.2) and
`cargo 1.95.0 (f2d3ce0bd 2026-03-21)`.  Repository input identities were:

| Input | Bytes | SHA-256 | Git object where applicable |
| --- | ---: | --- | --- |
| `.cargo/config.toml` | 470 | `50b044c83828f2c448ee59c362afaa79de2dd47fbd390f0cd2888391f1d1119d` | `c0e59805cf480523635677e333124885d257dcdd` |
| `rust-toolchain.toml` | 86 | `e3a213e0d222e94d213cafbc20932eb3f76c643b4dd63756acf95192df2aa310` | `8934eb4a5fabfca2502492140bfb51e89f9e0394` |
| root generated `Cargo.lock` (ignored) | 139,592 | `6d7895e6056c931fc19780ff4d0a031230e2ac79b9b30a9060bc8b6fad6747da` | — |
| `tools/perf-baseline/Cargo.lock` | 57,158 | `4def018e9c1d5f409b8c60d6cc0b8f114c486ab3f8fed2c7a4e9c7baf21c5184` | `a7f5a032102f8871ae1aaf208f95673f7b740f5e` |

The root lock was generated offline and is ignored by this checkout; the
standalone perf-baseline lock is tracked.  Both were used with `--locked`.

## Final commands and raw receipts

The raw logs remain outside the repository under
`/home/zhuhe/code/litchi-worktrees/.coord/ppt-lazy-entry/`.

The eight focused facade tests were invoked by this exact loop (each child
used the sourced environment and `--offline --locked`):

```text
for test in ppt_coalesced_set ppt_create ppt_elements ppt_slide_add \
  ppt_unified_read ppt_unified_structure ppt_unified_text ppt_unified_write; do
  cargo test --offline --locked -p litchi-ops --all-features --test "$test" --no-fail-fast
done
```

| Source and command | Status and counts | Raw log (bytes; SHA-256) |
| --- | --- | --- |
| `ca125d2ba1`; `cargo test --offline --locked -p litchi-ppt --all-features --all-targets --no-fail-fast` | exit 0; 75 result blocks, 1,568 passed, 0 failed, 3 ignored | `validation-ca125d2ba1/ppt-all-features-all-targets.log` (143,116; `2dd4ff4b3c9cb94afbe4b2d632331320bf84f488aeb95eb960cee5fd87f3a3f7`) |
| `ca125d2ba1`; `cargo test --offline --locked -p litchi-ppt --all-features --doc` | exit 0; 22 passed, 0 failed, 8 ignored | `validation-ca125d2ba1/ppt-doc.log` (3,373; `b1ff27b11f2faf483f83415c4902014fab1da12a3c3ff79364dffa2b7a760a05`) |
| `ca125d2ba1`; `cargo clippy --offline --locked -p litchi-ppt --all-features --all-targets -- -D warnings` | exit 0 | `validation-ca125d2ba1/ppt-clippy.log` (60; `af4e5b96b7371217a02d70fbcc05b2e509e7f1405a929155127fbde243d8686f`) |
| `ca125d2ba1`; `cargo check --offline --locked -p litchi-ops --all-features --all-targets` | exit 0 | `validation-ca125d2ba1/ops-check-all-targets.log` (60; `324aa901adc3088abe865e226e09b351320e843be5bb1a7daa6f4596c16f76ad`) |
| `ca125d2ba1`; eight focused `litchi-ops` PPT tests: `ppt_coalesced_set`, `ppt_create`, `ppt_elements`, `ppt_slide_add`, `ppt_unified_read`, `ppt_unified_structure`, `ppt_unified_text`, `ppt_unified_write` | exit 0; 8 result blocks, 71 passed, 0 failed, 0 ignored | `validation-ca125d2ba1/ops-ppt-consumers.log` (17,589; `56661f1f5974bdf45014c984338a3b42827a851d298f242f39a0b4fdf571730a`) |
| `1cb09b9c4a`; `cargo check --offline --locked --manifest-path tools/perf-baseline/Cargo.toml --all-targets` | exit 0; all standalone perf-baseline targets compiled; no tests or measurements run | `validation-1cb09b9c4a/perf-baseline-check.log` (60; `29c3e9b160b9fb9ec744e071d1facb9f3dc123610367df6ea4e28abadc5dc128`) |
| `1cb09b9c4a` plus the external candidate overlay below; `cargo check --offline --locked -p litchi-odg --release --example litchi-perf-odg-joined-qualification` | exit 0; compile-only; no probe execution | `validation-1cb09b9c4a/odg-candidate-probe-check.log` (166; `ab5c30ae0ec05c8a097f98ac57468392421360bc5cad4e192cb49c92ab665a1f`) |

The static checks `rustfmt --edition 2024` on the changed standalone Rust
sources, `git diff --check`, and the repository large-evidence/binary guards
also passed.  The final checkout was clean.

## Binding follow-up

The qualification bindings now match the reviewed tracked probe sources at
the final tip.  `common` is 17,863 bytes with SHA-256
`cdbe48f55e51ba6a35285c0b0a321e36c4d07c616053be685c9865d79fa3b9c5`,
`baseline` is 11,351 bytes with SHA-256
`8b3adac40dee59156fa667d6478779f7c63a28c60c1253f4e0187140a57a5a40`, and
`candidate` is 13,180 bytes with SHA-256
`ba4dab289b8786c83882925d3afb50ff131805a575fced4ea6c735d1f0256fbc`.
The allocator binding is unchanged.  The test manifest factory and reviewed
manifest checks consume these same fixed identities, so generated manifests
cannot retain the superseded probe bytes.

The build-free contract test command was:

```text
python3 -m unittest -v tools.test_odg_joined_qualification
```

At source content committed as `c793189120`, it exited 0 with all 26 tests
passing and no failures or skips.  The raw log is retained outside the
repository at
`/home/zhuhe/code/litchi-worktrees/.coord/ppt-lazy-entry/validation-bindings-0686/odg-joined-tests.log`
(4,971 bytes, SHA-256
`3e4067a47107b23e486a2f811c44b7c73c7e11a06a66739c2df450d2d93c7b52`).  No
Cargo command or native qualification was run for this follow-up.

## ODG candidate overlay provenance

The ODG command copied exactly these tracked final-tip sources into the
temporary example location `crates/litchi-odg/examples/`, compiled the example,
and removed those temporary files before returning.  The retained external
overlay is
`/home/zhuhe/code/litchi-worktrees/.coord/ppt-lazy-entry/odg-probe-overlay-1cb09b9c4a`:

| Overlay source | Bytes | SHA-256 |
| --- | ---: | --- |
| `odg_joined_qualification_common.rs` | 17,863 | `cdbe48f55e51ba6a35285c0b0a321e36c4d07c616053be685c9865d79fa3b9c5` |
| `litchi-perf-odg-joined-qualification.rs` | 13,180 | `ba4dab289b8786c83882925d3afb50ff131805a575fced4ea6c735d1f0256fbc` |

The overlay bytes and hashes are copied from the final tracked files at
`1cb09b9c4a`; the compile used the normal candidate adapter and did
not run `--mode verify`, calibration, timed, allocation, native ABBA, or any
other qualification route.

## Superseded diagnostic logs

The following external logs are retained for auditability but are not
acceptance evidence.  The first perf-baseline attempt exposed stale
`slides()`/`source()` consumers; the ODG attempts successively exposed the
custom-cfg, typed-error, proof-borrow, route-lifetime, and helper-visibility
issues fixed by the final consumer commit.  The accepted rows above are the
only final claims.

| Diagnostic log | Bytes | SHA-256 |
| --- | ---: | --- |
| `validation-ca125d2ba1/perf-baseline-check.log` | 10,994 | `7723d60b7864ca3fbd72cf392b337d26c64c0195da181f9aecf2641737101b9d` |
| `validation-2274f4cff9/odg-candidate-probe-check.log` | 12,564 | `4ca802ceb23ff909a42859ee992734ea1f43e35dceeba3369ea3560272130c1d` |
| `validation-2274f4cff9/odg-candidate-probe-check-v2.log` | 7,096 | `ebf1e0a4101b575281872ada7039cdd06a18d6978447eaef7ac41a5d87496ae7` |
| `validation-candidate-consumer/odg-candidate-probe-check.log` | 1,894 | `d48b657ed261a8266e410b86275d4088e717b71ac43be755382e9ce83daa5843` |
| `validation-candidate-consumer/odg-candidate-probe-check-v2.log` | 11,560 | `9533f16519334272f23ce05b1d39a35c1725749cd931d83d3e6c428f4b90f300` |
| `validation-candidate-consumer/odg-candidate-probe-check-v3.log` | 166 | `6e2b1412895f0cd5a6872586ef26d60f9225f1b77436d8a40c5cd780f6f46d2e` |
