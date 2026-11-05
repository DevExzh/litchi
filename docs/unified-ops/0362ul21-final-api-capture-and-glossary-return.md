# 0362ul21 — final 472b API capture and glossary return

Date: 2026-10-24. This record binds the fresh public-API capture to source
`472b6705475b073e7fe08f6e11ee17fb1e29f838`. It records API evidence only; it
does not claim the full gate, native qualification, corpus, ABBA, or final
unified-API acceptance.

## Fresh all-68 capture

The detached source is clean at tree
`ebc1cee05f124c0c63977aff07a23288c19e7d00`. Its root `Cargo.toml` is
`f5ec007892023b38f1deaef1ce20714db58b58448a3bea088be3b847b84570de`; the
139,606-byte `Cargo.lock` is SHA-256
`4100b1fbb272dd24842f33a2d6f353a3f3a97dfa7e435c4649a0e1a029bfaea1`.
The capped cargo shim is
`(deleted-coord-artifact)`, SHA-256
`1b22a69fd625d57565041bf6c634d8ebd1a82be8b3d0006390472c86324230d8`.

All 68 packages ran fresh (`reused=0`) with hidden-item rustdoc, all features,
offline/locked Cargo, one build job, and a 20 GB target/RSS ceiling. Every
receipt and its rustdoc artifact is present, parseable, and status `ok`; the
committed [receipt manifest](results/0362ul21/receipt-manifest.tsv) preserves
each source commit, lock hash, artifact/log path, size and SHA-256. The
external raw output is 1,099,012,277 bytes and remains outside the repository;
the rustdoc artifacts total 1,098,606,755 bytes. Peak reported RSS is
2,996,432,896 bytes and the final target is 3,748,156,727 bytes, below the
20,000,000,000-byte cap. The reduced [capture receipt](results/0362ul21/capture-verification.json)
contains the exact paths and run hashes.

The capture command was:

```sh
env -u LITCHI_CAPPED -u LITCHI_BLESS \
  PATH="(deleted-coord-artifact):$PATH" \
  LITCHI_MEM_MAX=20G CARGO_BUILD_JOBS=1 RUST_TEST_THREADS=1 \
  CARGO_INCREMENTAL=0 CARGO_PROFILE_DEV_DEBUG=0 CARGO_PROFILE_TEST_DEBUG=0 \
  CARGO_PROFILE_RELEASE_DEBUG=0 CARGO_NET_OFFLINE=true LC_ALL=C \
  CARGO_TARGET_DIR="$TARGET" TMPDIR="$OUTPUT/tmp" \
  python3 "$SOURCE/tools/final_api_inventory.py" --root "$SOURCE" \
  --output "$OUTPUT" --target-dir "$TARGET" \
  --cargo (deleted-coord-artifact) --no-reuse
```

## Inventory and reviewed dispositions

The fresh [84-row inventory](results/0362ul21/inventory-472b.tsv) is byte
identical to the independent recheck and to the prior 0362ul14 inventory:
38 `mut_draft`, 31 `edit_type`, 13 `transaction_type`, and 2 `mut_owner`
rows. The captured blank scaffold remains external at
`(deleted-coord-artifact)/disposition.tsv`
(12,634 bytes, SHA-256
`57d486415c0e940160a3e0ed29bee08b79b14b013773973a06827e2b17445529`). Its
84 rows are `UNASSIGNED_REQUIRED_REVIEW`; it is explicitly not accepted as a
disposition.

The already reviewed 84-row disposition remains
[0362ul14/reviewed-disposition-082.tsv](results/0362ul14/reviewed-disposition-082.tsv),
SHA-256
`3b0d48c38a4cc53e319c720d5e5ef4a1378631b10de5ec13190b03f5d3ad1a3e`.
The final independent read-only receipt
[independent-review-472b.json](results/0362ul21/independent-review-472b.json),
SHA-256
`6bd44bf07454ce459f9b0cfb27c91a0718bd12a2b9bc02b08a011e3c3ed2b2f3`,
revalidates all 84 role identities against this capture. It finds no new
inventory rows; the reviewed decisions are retained without inferring an
approval from the fresh blank scaffold.

## Supplemental glossary sweep

The [supplemental sweep](results/0362ul21/glossary-sweep.json) ran all 59
registered profiles against the fresh all-feature JSON without Cargo or new
rustdoc: 58 passed and one failed. The only finding is `keynote-minimal`
(`litchi-keynote`): the all-feature JSON exposes the internal-iwork-source
helper `__invalidate_slide_preview` with banned `Limits`. `keynote-full`
passes. This is the known all-feature/minimal-profile mismatch; it is not a
source waiver and this sweep makes no minimal-feature claim. The
feature-specific glossary gate remains required.

A first attempt is preserved externally at
`(deleted-coord-artifact)`.
All 59 commands exited 2 because its generated path omitted
`/home/zhuhe/code/litchi-worktrees`; its results and provenance hashes are
listed in the sweep receipt. It is diagnostic-only and was not used as
acceptance evidence. The corrected run is the only supplemental result
reported above.

Nothing here modifies `PROGRESS.md`, source, or historical capture artifacts;
nothing is pushed. The raw capture and glossary outputs remain in their
external paths.
