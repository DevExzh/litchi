# 0362r — DOCX variables raw codec tree

This slice starts from `a18593385c` on `uo2/w8-docx-variables`. It closes
the document-variables row as a raw codec model; it does not claim the other
DOCX snapshot, source-backed, or model rows complete.

## Contract mapping

Before this slice, `litchi_docx::variables` publicly exposed a detached
`Edit`/`Commit` staging family, and its XML parser and serializer selected the
format constants implicitly. That shape made the detached result look like a
package edit while the value was only a settings codec projection.

The public raw surface is now `variables::raw::{Limits,Tree}`. `Tree` owns the
admitted settings bytes and typed `Variables`, stages `set`, `remove`,
`clear`, `replace`, and closure-based typed mutations, and serializes only on
`to_bytes(&self, Limits)`. An unchanged tree returns the retained bytes after
rechecking the caller's output limit. A changed tree performs one bounded
source rewrite, preserving opaque settings content, source-order variables,
exact no-op bytes, and atomic failed mutations. The old `variables::Edit` and
`variables::Commit` remain crate-private engine types for the existing
source-backed implementation; no public detached publication family remains.

`variables::Snapshot` remains a read-only source projection and
`variables::Patch` remains the source-checked reversible raw patch used by the
owner implementation. Their public XML ingress now requires `raw::Limits`, as
do `Variables::from_xml`, `Variables::to_xml`, `parse_variables`, and
`write_entries`. Existing package/source-backed callers pass an explicit
format-default variables policy at their boundary because DOCX had no
separate retained variable budget before this slice; this is an explicit
owner policy choice, not a hidden parser default. The source-backed snapshot
retains that policy for later edits and rewrites.

The typed `Variables` collection remains available as a data model. Its
convenience `insert`/`validate` methods retain the format hard bounds; raw
ingress, mutation, and output through `raw::Tree` use the caller's `Limits`
before allocating or publishing. `Tree::from_xml` and its base clone use
fallible admission; output checks include unchanged and rewritten paths.

## Validation

Commands were run from `/home/zhuhe/code/litchi-worktrees/uo2-w8-docx-variables`
with `source /home/zhuhe/code/litchi-worktrees/.coord/w8-file/env.sh`,
`CARGO_BUILD_JOBS=2`, `RUST_TEST_THREADS=2`, and `LITCHI_MEM_MAX=20G`:

```text
cargo test --offline --locked -p litchi-docx --lib variables:: --no-fail-fast
cargo test --offline --locked -p litchi-docx \
  --test source_backed_variables --test source_backed --test source_backed_sanitize \
  --no-fail-fast
cargo clippy --offline --locked -p litchi-docx --all-targets --all-features -- -D warnings
```

The final variables run passed 15 tests. The retained source-backed run passed
18 + 9 + 6 tests. The final all-target/all-feature Clippy run finished with
`Finished dev profile` and no diagnostics. The first all-target attempt is
retained because it caught the now-migrated integration consumer that called
the private engine `try_edit`:

| receipt | bytes | SHA-256 |
| --- | ---: | --- |
| `.coord/w8-docx-variables/variables-tests-final.log` | 1,803 | `ae4d456f9e518056ef6d6a189c829dfc25d390de0f597191144d11fad14216f0` |
| `.coord/w8-docx-variables/source-backed-tests.log` | 3,330 | `a729196178ff3a759f68186448599ef756849d50ec936cc48e02a073ef942468` |
| `.coord/w8-docx-variables/clippy-all-targets-final.log` | 222 | `090bcbbcf75d5bd00eae6f2e5da226cb83fe56f1eed34286f4365010c35cd282` |
| `.coord/w8-docx-variables/clippy-all-targets-failed.log` | 759 | `08c2b42f95a91680e62184308fd4770b6a8c885d85185ae2b2cac89f5ff37ec5` |

The external receipts are intentionally outside the repository. No benchmark
or native performance claim is made here; the batching behavior is covered by
the raw-tree unit tests, while coordinator integration still needs the final
workspace inventory and performance qualification.
