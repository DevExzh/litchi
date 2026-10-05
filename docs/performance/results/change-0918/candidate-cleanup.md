# 0918 candidate cleanup custody

The workspace lock generated for offline quality checks is retained outside the
candidate source tree before any later cleanup:

- pre-cleanup source-tree copy: `/home/zhuhe/litchi-performance-0918/candidate/tree-v2/Cargo.lock`
- retained copy: `/home/zhuhe/litchi-performance-0918/candidate/Cargo.lock.workspace-generated-copy`
- bytes: `138201`
- SHA-256: `00ef62deed9b94a8b4ed0b3e2497923ae6e1a092975c2311b654407738881684`
- original workspace-lock copy remains at
  `/home/zhuhe/litchi-performance-0918/candidate/Cargo.lock.workspace-root-copy`
  with SHA-256 `cc7d3578cf7570a96351db585288209dd9510fa4872c7b3514eb896d6b4500dd`.

The standalone probe lock is separate and remains bound to
`candidate-source.json`; it remains available in the retained source archive and
the committed probe tree. The source and final-build receipts are immutable.

Post-capture verification is recorded in
`docs/performance/results/change-0918/postcapture-source.json`: 10,898 source
rows matched, all eight probe files matched the candidate receipt and root
checkout, both candidate binaries retained their hashes, and the exact root
baseline `package.rs` hash was `a04da643b093ec9724c9ac57ed8fbfdf2c2fedfc16eade105c91c32df7c9a446`.

After that verification, the owned candidate source export, superseded export,
target caches, and redundant temporary file lists were removed. The detailed
path and byte record is retained at
`/home/zhuhe/litchi-performance-0918/candidate/cleanup-after.json`:

- removed source export: `tree-v2`, 272,454,632 bytes, 10,898 files;
- removed superseded export: `superseded-v1`, 13,894,486,329 bytes, 470,513 files;
- removed target caches and sidecars: 8,186,244,987 bytes as itemized in the
  cleanup receipt;
- removed redundant lists: `head-tracked-files-v2.txt` (56,035,871 bytes) and
  `tree-v2-files.txt` (590,641 bytes).

The two final binaries remain at their original paths under
`target/release/`, with hashes bound by the post-capture receipt. The archived
source, all source manifests, both workspace lock copies, build and quality
logs, and the fingerprint/config evidence remain external. The retained
fingerprint/config copy is
`/home/zhuhe/litchi-performance-0918/candidate/fingerprint-config-evidence/`.
Post-cleanup stability verification passed for the retained source artifacts,
all 61 raw log entries, both binaries, and the target-release directory. Its
receipt is `/home/zhuhe/litchi-performance-0918/candidate/postcleanup-stability.json`
with SHA-256
`c5394d3126f22a6f78176e9193ce0e2fac91d35f877a062b05f3442c9c11fa77`.
