# 0362ra — DOCX variables raw codec tree

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
before allocating or publishing. `Tree::from_xml` and `Snapshot::from_xml`
take an owned `Vec<u8>` so these APIs do not silently copy an oversized slice
before admission. Their base clone and output buffers use fallible
reservation. An unchanged tree retains the parsed source depth and typed
collection limits, so a caller cannot bypass a stricter output policy by
choosing the no-op path.

Resource failures use the crate's typed `settings::LimitKind` through
`Error::SettingsLimit`: XML bytes, XML depth, variable count, name characters,
and value characters are distinct observations. Invalid policy construction
is still a request error, while allocation failures retain their allocator
source.

`Tree` retains a private summary of the admitted collection limits. A repeated
`set` under the same or a looser count/name/value policy validates only the new
entry; a genuinely tighter policy performs one full collection check before
the mutation. The `repeated_sets_reuse_an_admitted_collection_policy` test
stages 1,000 replacements in a 1,000-variable tree and observes zero full
collection validations, then observes one validation for a tighter value
policy. `to_bytes` still rechecks the complete collection, source depth, and
source XML bytes for every caller policy, including an unchanged tree.

`Tree::set` takes one `Into<Cow<'a, str>>` representation for each input and
validates the actual `Cow` view before calling `into_owned`. This removes the
old `AsRef<str> + Into<String>` split view: a type can no longer validate one
string and insert another. The focused
`cow_values_are_checked_before_ownership_conversion` test implements both
views with different contents; the oversized owned `Cow` is refused and the
legacy owned conversion is never selected. A separate
`owned_cow_values_move_into_the_collection_without_copying` test checks that
already-owned `String` values retain their allocations when admitted.

This follow-up supersedes the earlier dual-view input wording in this record;
the raw model remains bounded and source-neutral, while the source-backed
owner's existing `Edit::set` keeps its owner-level `Into<String>` convenience
surface.

## Validation

Commands were run from `/home/zhuhe/code/litchi-worktrees/uo2-w8-docx-variables`
with `source (deleted-coord-artifact)`,
`CARGO_BUILD_JOBS=2`, `RUST_TEST_THREADS=2`, and `LITCHI_MEM_MAX=20G`:

```text
cargo test --offline --locked -p litchi-docx --lib variables:: --no-fail-fast
cargo test --offline --locked -p litchi-docx \
  --test source_backed_variables --test source_backed --test source_backed_sanitize \
  --no-fail-fast
cargo clippy --offline --locked -p litchi-docx --all-targets --all-features -- -D warnings
```

The original variables run passed 18 tests. The follow-up Cow-input run passed
7 raw-tree tests, and the strict all-target/all-feature Clippy run finished
with no diagnostics. The retained source-backed run passed
18 + 9 + 6 tests. The final all-target/all-feature Clippy run finished with
`Finished dev profile` and no diagnostics. The first all-target attempt is
retained because it caught the now-migrated integration consumer that called
the private engine `try_edit`:

| receipt | bytes | SHA-256 |
| --- | ---: | --- |
| `(deleted-coord-artifact)/variables-tests-policy-final.log` | 1,949 | `5d59760add22d8f08833b31675f315c4dbc0650ad1e414ac05c0817622297a17` |
| `(deleted-coord-artifact)/source-backed-policy-final.log` | 3,544 | `1530fe6a39af01563de08c54d2e2af34e442d0b98ee84e41abf95a6550d0344e` |
| `(deleted-coord-artifact)/clippy-policy-final.log` | 170 | `94bc541edd44e5a8d1c7b24de1635ab4ee6176a54cdacd4ddfc8268282b5c835` |
| `(deleted-coord-artifact)/clippy-all-targets-failed.log` | 759 | `08c2b42f95a91680e62184308fd4770b6a8c885d85185ae2b2cac89f5ff37ec5` |
| `(deleted-coord-artifact)/cow-focused-v2.log` | 1,042 | `3249e71314292f9c1a1be5c50337a86e0ebb80d9039b6487e4082ccd81582e77` |
| `(deleted-coord-artifact)/cow-clippy-v2.log` | 222 | `146dda73c8b4c0c0716465987e56c785509e16557712e15a4cb3c7d8cd3552fe` |

The external receipts are intentionally outside the repository. No benchmark
or native performance claim is made here; the batching behavior is covered by
the raw-tree unit tests, while coordinator integration still needs the final
workspace inventory and performance qualification.
