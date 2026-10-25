# 0362ul44 — corrected ODF validation receipts

This is a bounded execution receipt, not a native, performance, or acceptance
claim.  The historical 6672 source plan in `0362ul44a-odf-validation-plan.md`
is preserved unchanged.  All source and Cargo work below was run only after
the owner granted the ODF slot; no native or profile command was run.

## Frozen source and policy

The final source endpoint is a detached, clean checkout:

| item | value |
| --- | --- |
| commit | `32253785a50b6515f69682f7eab903d6dba63fba` |
| tree | `e3dbfbf4a451eabf0e0a1498b6de93f8a50a5b29` |
| endpoint | `/home/zhuhe/code/litchi-worktrees/.coord/0362ul44-odf-validation-32253785a5/source` |
| integration basis | `30aaafbf98e097aeec8cf6fe3df6caaae62a37d6` |
| root lock | `4100b1fbb272dd24842f33a2d6f353a3f3a97dfa7e435c4649a0e1a029bfaea1` |
| perf-baseline lock | `eac86284f67136aa856a8dcb6f383178b89ca92b26434a3de1d7408d68f2690c` |
| native-resave lock | `51c25d3d6f36e20b50f0b19e39b1aa9c7b49777154122ecc95d3acd05c041378` |

The root lock is the reviewed ignored lock copied from the final capture
endpoint.  The endpoint contains all workspace crates, ODF fixtures, consumer
fixtures, the bounded runner, and exactly the 34 direct `docs/report` inputs
referenced by the selected ODS formula tests (4,626,860 bytes total).  It is
deliberately sparse; native runners were not invoked and their sparse-source
guard was not weakened.

Pinned setup files are unchanged from the reviewed plan:

```text
Cargo.toml                                      f5ec007892023b38f1deaef1ce20714db58b58448a3bea088be3b847b84570de
rust-toolchain.toml                             e3a213e0d222e94d213cafbc20932eb3f76c643b4dd63756acf95192df2aa310
.cargo/config.toml                              50b044c83828f2c448ee59c362afaa79de2dd47fbd390f0cd2888391f1d1119d
rustfmt.toml                                    5b08a5db07e6acaee62319a09314b9302c8cbc0f8e6b88b5ee8695f6303383d3
clippy.toml                                     f11517bc5872d96aab42ba6519db4e78f2fdcd44512ac70a1b45adbd9929b50c
tools/api_glossary.json                         b98cf5ae4d5fe4cf3b6430896535e905f6a1f05e15b867b6cf822f2ba819d659
run_safe_gate.py                                ad24903c0aab9dffd8e99779df57ecda56aae263bb271fdf1dcaf96f31922215
safe_chunked_tests.py                            3e03aa2227ac0002c3052cf5080d032cd2e9150f584c926b9aa0d499175ec67f
tooling/bin/cargo                               495e24c6f60b238378386d86ac4158cba904992ce8870a9494f2c4a0eedd256d
```

Rust is `1.95.0` (`rustc 1.95.0 (59807616e 2026-04-14)`, LLVM 22.1.2;
`cargo 1.95.0 (f2d3ce0bd 2026-03-21)`).  The runner enforced offline/locked
Cargo, one build job, one test thread, incremental/debug disabled, a
20,000,000,000-byte hard target cap and a 19,000,000,000-byte live target
guard.  These are decimal byte limits (20 GB and 19 GB), not GiB labels.

## Passing rows

The rows use the proven `0362ul29` runner.  The target was safely reused from
the preceding same-policy source endpoint
`/home/zhuhe/code/litchi-worktrees/.coord/0362ul44-odf-validation-2f7a009832c/f5/target`;
each row still recorded the final source and lock fence.

| row | command | tests passed | seconds | target bytes | receipt (bytes, SHA-256) |
| --- | --- | ---: | ---: | ---: | --- |
| changed-all-targets | `cargo check --offline --locked --all-targets --all-features -p litchi-odf-common -p litchi-odf-formula -p litchi-odp -p litchi-ods -p litchi-odt -p litchi-odf -p litchi-ops -p litchi` | — | 627.27 | 825127447 | `/home/zhuhe/code/litchi-worktrees/.coord/0362ul44-odf-validation-32253785a5/f5-state/00-changed-all-targets-attempt-32253785a5-0.log` (8484, `5291270c05583c038ad2ea87c0ceb3aaa624c123ad643ff532b0a419d459b4c8`) |
| changed-clippy | `cargo clippy --offline --locked --all-targets --all-features -p litchi-odf-common -p litchi-odf-formula -p litchi-odp -p litchi-ods -p litchi-odt -p litchi-odf -p litchi-ops -p litchi -- -D warnings` | — | 440.09 | 1116777171 | `/home/zhuhe/code/litchi-worktrees/.coord/0362ul44-odf-validation-32253785a5/f5-state/01-changed-clippy-attempt-32253785a5-1.log` (8483, `f1f100a1a24cbedef8e4646006ffc34ac913de4a0cee66f202b29376d267e1b1`) |
| focused-odf-libs | `cargo test --offline --locked --all-features -p litchi-odf-common -p litchi-odf-formula -p litchi-odp -p litchi-ods -p litchi-odt --lib` | 2135 | 105.16 | 1500344349 | `/home/zhuhe/code/litchi-worktrees/.coord/0362ul44-odf-validation-32253785a5/f5-state/02-focused-odf-libs-attempt-32253785a5-2.log` (199331, `67fee47e92a747bcd9a22f10c556b190b4ea4e070fbf74b840f90d866b28e426`) |
| focused-strict-xml | `cargo test --offline --locked -p litchi-odf-common --test resolved_reader -- --nocapture` | 17 | 2.91 | 1503119608 | `/home/zhuhe/code/litchi-worktrees/.coord/0362ul44-odf-validation-32253785a5/f5-state/03-focused-strict-xml-attempt-32253785a5-3.log` (1741, `d769afffd59a5948ba58a714b5fa36a9d56cbf554136fb55b51f0857f443dc37`) |
| focused-script-envelopes | `cargo test --offline --locked -p litchi-odt --lib document_scripts::tests -- --nocapture` | 8 | 0.76 | 1503119608 | `/home/zhuhe/code/litchi-worktrees/.coord/0362ul44-odf-validation-32253785a5/f5-state/04-focused-script-envelopes-attempt-32253785a5-4.log` (1043, `677518cf5ab5fb9ab3f20a4c9530ff48486ab94c01c0ba5034665f1c73aebdbd`) |

The authoritative row manifest is
`/home/zhuhe/code/litchi-worktrees/.coord/0362ul44-odf-validation-32253785a5/f5-state/rows.json`
(48667 bytes, SHA-256
`a8cb7053fe96845fb8800c4a34e8d8db7908a8877ffe0300ca758ec917345909`).  All
five rows have status `passed`; the manifest, Cargo-target traces, source
fences, and process guards are authoritative, and no pass row was hand-authored.

### Target and runner disposition

The historical 6672 plan required a fresh target for every phase.  For these
final 322 correction rows, the owner accepted a narrower target-reuse
disposition: the final source endpoint was a fresh clean checkout, the three
reviewed lock fingerprints were unchanged, `CARGO_INCREMENTAL=0` was enforced,
Cargo target content fingerprints and the process guards were recorded, and the
`f5-state` row manifest was new.  Only the compiled target was reused from
`/home/zhuhe/code/litchi-worktrees/.coord/0362ul44-odf-validation-2f7a009832c/f5/target`.
This is not a fresh-target claim and does not alter the historical plan.

The executed rows used the persisted `run_safe_gate.py` runner at the frozen
source endpoint (SHA-256
`ad24903c0aab9dffd8e99779df57ecda56aae263bb271fdf1dcaf96f31922215`) and its
persisted `safe_chunked_tests.py` helper (SHA-256
`3e03aa2227ac0002c3052cf5080d032cd2e9150f584c926b9aa0d499175ec67f`).  The
authoritative invocation and outcomes are the command arrays and guards in
`rows.json` plus the five hashed row logs above; no separate custom-driver
artifact was persisted, so this receipt does not invent or hash one.

## Preserved failed attempts

Earlier failures remain outside the repository and were not rewritten:

- `/home/zhuhe/code/litchi-worktrees/.coord/0362ul44-odf-validation-2f7a009832/f5-state/00-changed-all-targets-attempt-2f7a009832-0.log`
  (21796 bytes, `cc5e806bb11228db8283811765a37506e20241510e6136e9cbc60232c47d28bd`)
  recorded the first sparse omission (`ods-formula-aggregates` evidence).
- `/home/zhuhe/code/litchi-worktrees/.coord/0362ul44-odf-validation-2f7a009832b/f5-state/00-changed-all-targets-attempt-2f7a009832b-0.log`
  (21440 bytes, `db503b5cc6a9c71cd5218c607e32d02a4a6e99982b71f676c95f034d4ec7b267`)
  recorded the second sparse omission (`ods-formula-trigonometry` evidence).
- `/home/zhuhe/code/litchi-worktrees/.coord/0362ul44-odf-validation-2f7a009832c/f5-state/01-changed-clippy-attempt-2f7a009832c-1.log`
  (2308 bytes, `1f61efd5636b4efc18022262ce71ad170754d45bcf2b693ddba8b0ebc5dc9c81`)
  recorded the pre-fix `double_must_use` lint.
- The 232/e690 source and fixture failures, including the list-style consumer
  drift, remain in their original validation roots.

The final source correction was the minimal removal of the redundant bare
`#[must_use]` on formula `serialize`; its `&Limits` policy and all tracked
change/formula consumers remain unchanged.  No native qualification, timing,
heap/resource profile, or acceptance result is asserted here.
