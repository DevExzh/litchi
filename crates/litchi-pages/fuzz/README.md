# Pages body drawable-order fuzzing

`pages_body_drawable_order` exercises the focused Pages package owner through
the public selector-first API. Each input is a bounded fixture descriptor,
not an opaque native package: the harness builds a small exact `Index/Document.iwa`
package with zero to sixteen drawable objects and, when selected, a native
body-storage object occupying a structural slot in the z-order list. This keeps
the body-storage exclusion and slot-preservation contract reachable in every
campaign while still letting mutations vary physical object order, drawable
permutations, unknown fields, preview payloads, and the requested movement.

The target runs `Package::from_bytes_with_limits`, reads the semantic handles,
stages either a selector move or an exact handle permutation, reopens the
candidate, checks native order and body-slot invariants, and applies the
source-bound inverse when a commit succeeds. It also probes arbitrary bounded
ZIP bytes and deterministic malformed descriptors (duplicate root references,
missing z-order objects, wrong message types, zero/missing/non-canonical
identifiers, external references, and malformed root wire) to ensure owner
failures remain atomic and finite.

Descriptor inputs are capped at 64 KiB; generated packages contain at most 16
drawables and a fixed small number of ZIP members. Checked-in seeds under
`corpus/pages_body_drawable_order/` use `hex:` descriptors and intentionally do
not contain native package bytes.

List and type-check the target from this directory:

```sh
cargo +nightly fuzz list
cargo +nightly fuzz check pages_body_drawable_order
```

Run a bounded sanitizer smoke with mutable corpus, artifacts, and build output
outside the checkout:

```sh
fuzz_root="$(mktemp -d "${TMPDIR:-/tmp}/litchi-pages-drawable-owner-fuzz.XXXXXX")"
fuzz_corpus="$fuzz_root/corpus"
mkdir "$fuzz_corpus" "$fuzz_root/artifacts"
cleanup_fuzz_corpus() {
  if [ "${KEEP_FUZZ_CORPUS:-0}" = 1 ]; then
    printf 'retained temporary fuzz root: %s\n' "$fuzz_root"
  else
    rm -rf "$fuzz_root"
  fi
}
trap cleanup_fuzz_corpus EXIT
cp corpus/pages_body_drawable_order/*.hex "$fuzz_corpus/"
CARGO_TARGET_DIR="$fuzz_root/target" cargo +nightly fuzz run \
  pages_body_drawable_order "$fuzz_corpus" -- \
  -artifact_prefix="$fuzz_root/artifacts/" -runs=100 -max_len=65536 \
  -timeout=10 -rss_limit_mb=2048
```

## `pages_body_table_hidden_axes`

`pages_body_table_hidden_axes` exercises the focused body-table hidden-axes
owner using bounded descriptors that construct a strict two-table IWA graph.
Every input first reaches a valid selector/read setup and stages an edit; the
same descriptor then selects ownerless creation (with and without a strict
PackageMetadata registry) or an existing owner, empty/row/column/both-axis
state, set/clear/reset/no-op,
name/index selectors, filtered or pivot markers, lock state, shared
Package/COW ownership, shared-owner locality, valid non-identity UID-map
permutations, and exact field-path metadata ownership. A second probe mutates
the native graph or wire framing to cover missing, duplicate, dangling,
aliased, wrong-type, unknown-UUID, UID-map, active-state, direction,
duplicate-singular, wrong metadata, unsupported filter dependency, wrong-wire,
truncation, group, non-canonical, and invalid-UTF-8 cases. Registry-aware
probes also cover dangling and duplicate UUID bindings, near-maximum
watermarks, physical identifier collisions, missing metadata resources, and
metadata messages routed through the wrong package member.

Malformed focused reads and edits assert the precise public error class:
duplicate names are `AmbiguousTableName`, invalid filter metadata is
`UnsupportedDependency`, count-overflow metadata is a `WireFields`
`LimitExceeded` refusal before allocation, and the remaining malformed
root/graph/wire/UID/metadata cases are `InvalidSource` when physical package
ingress admits the fixture. A missing `Index/Metadata.iwa` remains a supported
metadata-free creation shape; watermark exhaustion reports a payload-object
`LimitExceeded` error.

Successful reads are checked for canonical sorted/deduplicated positions and
table bounds. Existing-owner edits are reopened and checked for exact target
semantics, one-time patch application, target conflict on replay, exact
inverse restoration, inverse involution, no-op byte identity, diagnostics,
preview invalidation, and unselected-graph-object/member preservation.
Admitted ownerless nonempty sets publish four helper objects and preserve the
source through their inverse patch; malformed metadata and graph edits fail
atomically. A data-dependent physical limit
profile runs on every input, while max/max+1 boundary sweeps and strict
nested codec fields/work/allocations/retained/scratch checks run once per fuzz
process.

Run a bounded sanitizer smoke with mutable corpus, artifacts, and build output
outside the checkout:

```sh
fuzz_root="$(mktemp -d "${TMPDIR:-/tmp}/litchi-pages-hidden-axes-fuzz.XXXXXX")"
fuzz_corpus="$fuzz_root/corpus"
mkdir "$fuzz_corpus" "$fuzz_root/artifacts"
cleanup_fuzz_corpus() {
  if [ "${KEEP_FUZZ_CORPUS:-0}" = 1 ]; then
    printf 'retained temporary fuzz root: %s\n' "$fuzz_root"
  else
    rm -rf "$fuzz_root"
  fi
}
trap cleanup_fuzz_corpus EXIT
cp corpus/pages_body_table_hidden_axes/*.hex "$fuzz_corpus/"
CARGO_TARGET_DIR="$fuzz_root/target" cargo +nightly fuzz run \
  pages_body_table_hidden_axes "$fuzz_corpus" -- \
  -artifact_prefix="$fuzz_root/artifacts/" -runs=100 -max_len=65536 \
  -timeout=10 -rss_limit_mb=2048
```
