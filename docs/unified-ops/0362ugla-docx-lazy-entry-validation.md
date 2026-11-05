# 0362ugla — DOCX lazy edit entry validation

This is the validation-only successor to `0362ugl-docx-lazy-edit-entry.md`.
The final source pin is `0cf1a17174`; the runtime suites ran on the equivalent
source pins `e5bc620a13` (DOCX and facade) and `609523b702` (the signed
empty-property correction).  The two commits after those suites are
publication-consumer migration and formatting-only changes.  No source
behavior was changed after the signed correction.

The signed focused test covers all three required boundaries: a writable
empty assignment remains a no-op, a signed valid target preserves the exact
signed bytes, and a signed invalid target still returns `PathNotFound` rather
than being hidden by the signed policy.  The unretained source case remains
`UnsupportedCapability`.

## Checks

All commands used Rust 1.95.0 (`rustc 1.95.0 (59807616e 2026-04-14)`,
`cargo 1.95.0 (f2d3ce0bd 2026-03-21)`), `--offline`, and the assigned target
`/home/zhuhe/code/litchi-worktrees/.coord/docx-lazy-entry/target`.

* `litchi-docx --all-features --all-targets`: 2355 tests passed, 0 failed.
* `litchi-docx --all-features --doc`: 168 doctests passed, 31 ignored, and 2
  compile-fail doctests passed.
* `litchi-ops --features docx --all-targets`: 739 tests passed, 0 failed.
* `litchi --features docx --all-targets`: 23 tests passed, 0 failed.
* The focused signed regression test passed.
* Strict `-D warnings` Clippy passed for `litchi-docx` and for
  `litchi-ops --features docx --all-targets`.
* The DOCX and litchi-ops standalone fuzz manifests passed `cargo check
  --all-targets`; the three affected perf-baseline binaries passed
  `cargo check`.
* `cargo fmt --check` passed for the DOCX, DOCX fuzz, litchi-ops and
  litchi-ops fuzz manifests, and `git diff --check` passed.

The final external target directory is 4,803,301,769 bytes, below the 20-GB
lane cap.  No native timing, ABBA measurement or performance qualification is
claimed here; the perf-baseline command was a compile check only.

## External reduced receipts

These receipts remain outside the checkout and were not committed:

| command | external receipt | bytes | SHA-256 |
| --- | --- | ---: | --- |
| DOCX all-target tests | `/home/zhuhe/code/litchi-worktrees/.coord/docx-lazy-entry/receipts/docx-all-target-tests-e5bc620a13.log` | 222427 | `7d62612f1383859166970cff183a9542b5a7c0775779f74eaadec51c0563f3fe` |
| DOCX doctests | `/home/zhuhe/code/litchi-worktrees/.coord/docx-lazy-entry/receipts/docx-doctests-e5bc620a13.log` | 23348 | `31bb84388896c2e55f831d44d04c5590f230165f71b5e663a8ab27b128feba5c` |
| facade all-target tests | `/home/zhuhe/code/litchi-worktrees/.coord/docx-lazy-entry/receipts/litchi-docx-tests-e5bc620a13.log` | 11996 | `e05bff5fed9a5d0af1b1f06a4b67ca03e766d2edae0c8e1409625aac2738e8ec` |
| litchi-ops all-target tests | `/home/zhuhe/code/litchi-worktrees/.coord/docx-lazy-entry/receipts/litchi-ops-docx-tests-609523b702.log` | 117399 | `d54a30657fbcc0da12bd7edd634783acf5775bfa124a23bcbde79a8af316cd3b` |
| signed empty-property test | `/home/zhuhe/code/litchi-worktrees/.coord/docx-lazy-entry/receipts/litchi-ops-signed-empty-correction.log` | 515 | `6d57b6e7b372cd8a8cfa13edac88637ceb2cc616a2c7c257b5c4420c95d9ae21` |
| DOCX strict Clippy | `/home/zhuhe/code/litchi-worktrees/.coord/docx-lazy-entry/receipts/docx-clippy-e5bc620a13.log` | 168 | `519c93b7d35693e5e7c13590d69d3ce59114fb765c0172aa57a8444cb4fb2100` |
| litchi-ops strict Clippy | `/home/zhuhe/code/litchi-worktrees/.coord/docx-lazy-entry/receipts/litchi-ops-docx-clippy-609523b702.log` | 166 | `5174c1853f86a3fe8335d9c9d30c5b5eed97a597cfd4e0a7d4d20a2ba40da819` |
| DOCX fuzz manifest check | `/home/zhuhe/code/litchi-worktrees/.coord/docx-lazy-entry/receipts/docx-fuzz-check-final2-609523b702.log` | 177 | `60e0d59bcb018462f6efa1882ad7622c9e01df31cf984440295cf1f877d7827a` |
| litchi-ops fuzz manifest check | `/home/zhuhe/code/litchi-worktrees/.coord/docx-lazy-entry/receipts/litchi-ops-fuzz-check-472d21297b.log` | 951 | `ee209e1ceeddedbb6534a614ddb8da221234e5c14320bb056f161aef9e0888af` |
| perf-baseline consumer check | `/home/zhuhe/code/litchi-worktrees/.coord/docx-lazy-entry/receipts/docx-perf-baseline-check-472d21297b.log` | 8549 | `256d140ede2bd24db2c2536bdfdd7046436cc8152ce4ad74a9832a3cccf71268` |
