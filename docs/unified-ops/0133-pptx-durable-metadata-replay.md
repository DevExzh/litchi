# 0133 — PPTX durable metadata replay

## Implementation

Native PPTX patches now use the explicitly versioned `LPTX0003` wire format.
They retain bounded before/after metadata images, derive changed relationship
IDs and content-type Default/Override keys using OPC's native fragment catalogs,
and distinguish relationship-member presence from an empty graph. The decoder
refuses earlier wire versions. Complete borrowed framing and count checks run
before decoded payload or metadata-image allocation.

Application validates the original package, prepares metadata against that
source, applies the resource closure to a private candidate, installs the
prepared metadata through guarded OPC replacement, and validates the result
before publication. An exact source uses its retained target image directly.
Transport onto a divergent source uses guarded lexical fragments and preserves
unselected destination entries, gaps and spelling. Root/context changes remain
opaque whole-image conflicts. Ordinary payload-only patches capture no metadata
images. Native bounded relationship decoding is public without granting any
source-replacement authority.

The diagnostic from 0128 is now a permanent passing regression: adding a part
and relationship preserves both manifest sentinels through durable forward
replay, publishes the same archive as direct native editing, and restores the
original metadata through inverse replay. Disjoint relationship IDs on a shared
owner can join alongside disjoint content-type mappings. Joins require common
source images, have deterministic ordering, and preserve unrelated destination
overlays. Joined inverses survive their own durable encode/decode cycle.

Whole metadata images are provenance, not the ordinary conflict domain. Joined
groups keep separate images; their inverse groups consequently share the
original before image as their common *after* image. Wire validation admits
this representation while joins still require common before images. Physical
member removal conflicts with another edit needing that member.

Decoded metadata must agree with declared relationship changes. Part topology
changes also require content-type provenance. If an existing default already
covers a new part, capture retains an unchanged whole-manifest proof; this is
permitted only with a companion content-type dependency and is conservatively
guarded as an exact image. Even unrelated destination manifest comments cause
this unchanged-proof case to refuse atomically; allowing such overlays requires
binding-specific dependency guards in a follow-up. Absent relationship images must use canonical empty
XML; package-root relationship members cannot be absent.

Merge admission traverses borrowed relationship outcomes and computes counts
and encoded lengths before cloning the combined write set. Metadata owner
counts are capped during capture, image lengths obey wire widths, and tighter
joined limits revalidate native catalogs. Selected three-way outcomes are also
admitted before their final clones. The relationship merge avoids repeated
vector insertion/removal. No performance acceptance claim is made.

The original frozen digest transcript remains unchanged in the test source.
A separate version-three fixture records the deliberately changed wire hashes
and wire-size-dependent refusal paths. The test additionally compares all
prior semantic revisions, publication bytes and retention outcomes.

## Verification

The native PPTX suite passed 978 tests with encryption enabled and one existing
ignore. The final focused patch suite passed all 23 tests, including forged
metadata omissions, canonical absence, different-base conflicts, serialized
joined inverses, exact wire bounds, metadata-only spelling edits and atomic
refusal of overlays against an unchanged content-type proof. The ordinary
payload capture test also asserts that its metadata set is empty.

OPC passed 1,126 tests/doctests across 47 result groups, with one existing ignore.
The PPTX/encryption facade passed 533 tests/doctests across 26 result groups.
Both native crates passed all-target Clippy and documentation with warnings
denied. Both PPTX glossary profiles passed (79 minimal and 70 full banned
names checked). Read-only review's final remaining policy question was resolved as the
explicit unchanged-proof refusal above and covered by a regression.

Verification commands were offline and serialized:

```sh
CARGO_NET_OFFLINE=true cargo test -p litchi-opc
CARGO_NET_OFFLINE=true cargo test -p litchi-pptx --features encryption --lib
CARGO_NET_OFFLINE=true cargo test -p litchi-pptx --features encryption --lib opened::patch::
CARGO_NET_OFFLINE=true cargo test -p litchi-ops --no-default-features --features pptx,encryption
CARGO_NET_OFFLINE=true cargo clippy -p litchi-opc -p litchi-pptx --features litchi-pptx/encryption --all-targets -- -D warnings
CARGO_NET_OFFLINE=true RUSTDOCFLAGS='-D warnings' cargo doc -p litchi-opc -p litchi-pptx --features litchi-pptx/encryption --no-deps
CARGO_NET_OFFLINE=true python3 tools/check_api_glossary.py --profile pptx-minimal --target-dir target
CARGO_NET_OFFLINE=true python3 tools/check_api_glossary.py --profile pptx-full --target-dir target
```

## Remaining scope

Metadata-coupled divergent three-way conflict resolution is explicitly refused;
package-root relationship three-way planning retains its previous refusal.
Compatible disjoint metadata joins are supported. Lexical expansion of a
self-closing metadata root is still conservatively opaque, so independent
first insertions into that owner can conflict. These restrictions need further
work before claiming the complete metadata merge design.

Public slide insertion, remaining native/facade API coverage, performance
acceptance and the full all-format goal remain open. This checkpoint closes
the forward metadata preservation failure from 0128, not the overall goal.

## Repository checkpoint

The rewritten base remains an ancestor of this branch; no rebase is needed.
Staging failed with:

```text
fatal: Unable to create '/home/zhuhe/code/litchi/.git/worktrees/unified-ops-api/index.lock': Read-only file system
```

No commit was created and no permission workaround was attempted. Both required
staged evidence guards printed nothing. A changed/untracked artifact audit found
no oversized restricted evidence, binary-directory files or ELF executables.
After all verification handles were terminal, checkpoint scratch logs, the
generated root lockfile and build output were removed. No temporary worktree
was created. Inherited source changes and the intentional 0128 diagnostic were
preserved.
