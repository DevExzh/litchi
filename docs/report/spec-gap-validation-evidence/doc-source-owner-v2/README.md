# DOC SourceOwner publication validation v2

This validation bundle records the four-file DOC SourceOwner overlay in
`/var/tmp/litchi-doc-source-owner-publication-validation-20260910`, based on
publication base `ce0be842b60a28db142b795389e24be5bde8ae07`. The candidate has
exactly these four tracked source changes:

- `crates/litchi-doc/src/embedded_object/transaction/edit.rs`
- `crates/litchi-doc/src/embedded_object/transaction/patch.rs`
- `crates/litchi-doc/src/embedded_object/transaction/snapshot.rs`
- `crates/litchi-doc/src/embedded_object/tests.rs`

The source blobs are byte-identical to the immutable integration candidate at
proof base `6da91778c4003b6c9819e5c25168cfe9a847359e`; their hashes are bound
by `publication-root-source-v1.json` and `receipt.json`.

## Source and gate bases

The bases are kept separate because they answer different questions:

- The **measured base** is `c4aa19e1930b79bec7db499a9248fe711ac3cf43`.
  The fresh bounded latency and allocation reports under `fresh-latency/` and
  `fresh-allocation/` use the c4aa control plus the historical four-file
  SourceOwner overlay. They are carried forward as measurement evidence only.
- The **historical validation base** is `6da91778c4003b6c9819e5c25168cfe9a847359e`.
  `historical-integration-root-gates-v2.json` records the clean integration
  context: 1,204 tests and 14 doctests, strict Clippy/rustdoc, formatting and
  whitespace checks. Its four source blobs match this publication exactly.
- The **publication base** is `ce0be842b60a28db142b795389e24be5bde8ae07`.
  `publication-root-final-gates-v1.json` records the final later-base gate:
  1,226 tests and 14 doctests, 2 and 12 ignored respectively, strict
  all-target/all-feature Clippy, warning-denied rustdoc, source-hash checks,
  pinned formatting and diff checks.

The publication gate is a correctness/tooling proof for ce0. It does not
retrospectively change the measured c4aa performance base or create a new
performance claim for ce0.

## Bounded performance evidence and tradeoff

The fresh root latency proof runs release locked `+1.95.0` harnesses with
six alternating rounds, five warmups, twenty samples, CPU affinity 0, and 34
untimed correctness checks per side. It retains JSON reports and source/fixture
verification only; targets, linked worktrees and binaries are not included.
For the large `embedded_object_snapshot_open` case, the candidate/control
paired median ratio is 1.0779031303440x, or about **+7.8% direct-open time**.

The same large `embedded_object_snapshot_open_noop_finish` case has a
candidate/control paired median ratio of 0.8069461435314x, or about
**19.3% lower compound open-plus-no-op-finish time**. This is the bounded
c4aa measurement input that substantiates the compound workflow tradeoff; it
is not a ce0 performance claim.

The separate fresh allocator proof checks 32 correctness booleans and records
240 untimed rows per side across eight operations, two profiles and fifteen
samples. It also retains reports only. For large
`embedded_object_snapshot_open`, allocated bytes are 12,412,349 in control
and 10,131,933 in candidate, a delta of -2,280,416 bytes and a ratio of
0.8162784497922x, or about **18.4% fewer counted allocated bytes**. These are
paired, bounded allocator counters; they are not RSS, peak process residency,
latency, or whole-workflow claims.

The recorded decision is the explicit allocation/compound-workflow tradeoff:
direct first-open work increases on the measured large case while the
open-plus-no-op-finish workflow and counted allocations decrease. All values
are inherited from the c4aa measurement inputs and are not attributed to the
ce0 publication base. No new performance measurement was run for ce0.

## Portable 100-member publication bundle

`source-owner-portable-100-payload-members.tar.gz` is a deterministic archive with
101 regular members: the exact original `archive-manifest.json` wrapper and
the 100 payload members listed in that wrapper. The wrapper is
self-excluded from its own `members` list, as required by the publication
format, and is required by `reproduce.py` before source verification/build.
The outer `portable-bundle-manifest.json` records the 100 payload members and
the 101-member archive. The exact wrapper is also retained beside this
README with SHA-256
`c0476ed94b8d48cbe0a8165299b5c20aaf0a06b171620006455da1e3254c4c83`.

Every payload member was checked against the immutable input manifest before
archiving; tar member order follows that manifest after the wrapper, metadata
is normalized, and gzip uses mtime zero. The archive contains no build target directories or compiled binaries; its
immutable fixture documents are retained because the reproduction harness
verifies them. Verify and load it with:

```sh
sha256sum source-owner-portable-100-payload-members.tar.gz
mkdir -p /tmp/source-owner-portable-check
tar -xzf source-owner-portable-100-payload-members.tar.gz -C /tmp/source-owner-portable-check
find /tmp/source-owner-portable-check -type f | wc -l  # 101
(
  cd /tmp/source-owner-portable-check
  PYTHONDONTWRITEBYTECODE=1 python3 -c \
    'import reproduce; print(reproduce.verify_archive(__import__("pathlib").Path(".")))'
)
```

`portable-bundle-manifest.json` provides the complete payload member list, byte
counts, and individual SHA-256 values. It also binds the original c4aa source
and fixture manifests and the input publication manifest hash.

## Gate logs and replay inputs

The `publication-*` compressed lock and logs bind the final ce0 gate. The
`historical-integration-*` compressed lock and logs preserve the prior 6da
integration proof. `portable-archive-extraction-root-verification.json` is
root's independent extraction check: it records the `ee6f6ab9d6634740b763bf9d3abd1d13a7bef1f7fcf6be9ab17fa29a659e402a` archive,
101 regular members, 100 checked payload members, the exact wrapper hash, and
no bytecode. `receipt.json` records raw and compressed hashes, commands,
source-base distinctions, the archive wrapper, and the 101-member archive.

The actual publication gate commands were run with Rust 1.95.0 and the
publication candidate as the working tree:

```sh
cargo +1.95.0 test -p litchi-doc --all-features
cargo +1.95.0 clippy -p litchi-doc --all-targets --all-features -- -D warnings
RUSTDOCFLAGS="-D warnings" cargo +1.95.0 doc -p litchi-doc --all-features --no-deps
rustup run 1.95.0 rustfmt --check --edition 2024 --config skip_children=true crates/litchi-doc/src/embedded_object/tests.rs crates/litchi-doc/src/embedded_object/transaction/edit.rs crates/litchi-doc/src/embedded_object/transaction/patch.rs crates/litchi-doc/src/embedded_object/transaction/snapshot.rs
git diff --check
```

These recorded gate commands intentionally used neither `--offline` nor
`--locked`; the source and gate receipts retain the dedicated target and all
pass results. The pinned rustfmt check covers the four source paths; the diff check covers
the candidate working-tree diff.

For an optional lock-bound replay, restore the retained publication lock and
use a fresh target. The `--locked` form below is a replay recipe and was not
part of the executed publication commands:

```sh
gzip -dc docs/report/spec-gap-validation-evidence/doc-source-owner-v2/publication-Cargo.lock.gz > Cargo.lock
replay_target=$(mktemp -d /var/tmp/litchi-doc-source-owner-replay-target-XXXXXX)
CARGO_TARGET_DIR="$replay_target" cargo +1.95.0 test --locked -p litchi-doc --all-features
CARGO_TARGET_DIR="$replay_target" cargo +1.95.0 clippy --locked -p litchi-doc --all-targets --all-features -- -D warnings
CARGO_TARGET_DIR="$replay_target" RUSTDOCFLAGS="-D warnings" cargo +1.95.0 doc --locked -p litchi-doc --all-features --no-deps
```

The fresh latency and allocation subtrees intentionally contain no build
outputs. Their JSON reports include toolchain, source/fixture hashes, command
metadata, correctness results, raw samples, allocator rows, and the bounded
comparison; rerunning requires caller-supplied source checkouts and temporary
build targets as described by the retained portable publication inputs.
