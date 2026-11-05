# 0362ul65 — RTF remaining raw boundaries

Date: 2026-10-25. This source leaf is cut from integration
`ef1acb5ef0b776c9a6066029a5cac35daeb0b5eb` on
`uo2/rtf-remaining-raw-0362ul65`. It closes the remaining RTF raw-boundary
surface identified by the 0351 contract. The source, every checked-in
consumer, glossary registration and this record move together; the coordinator
owns the integration merge and `PROGRESS.md`.

## Explicit raw validation

`ValidationReport::{from_str, from_bytes}` now take the caller's required
`read::OpenOptions` directly. The free `validation::{validate, validate_str}`
functions have the same explicit profile argument. The old default methods and
their `_with_options` twins are deleted; no alias or deprecated compatibility
entry remains. `raw::Tree` already had this required-options shape and is
unchanged.

The standalone RTF tests use explicit `OpenOptions` values, with a test-only
helper for the ordinary default profile. The performance baseline passes a
prebuilt `OpenOptions::default()` outside its timed validation closure. The
unified operations validator does not reopen published bytes with a fresh
profile: `File::validate` now asks the opened RTF owner for its retained
package validation report, preserving the source policy selected at open.

## Bounded compressed transport

`transport::compress(data, compress, CompressionLimits)` now requires a
caller-selected output profile. `CompressionLimits` exposes `new` and
`max_output_bytes`, with a finite 256 MiB default matching the decompression
profile's default ceiling. The encoder checks the complete framed size before
each temporary compressed-output growth and again before the final frame
reservation. Exceeded budgets return the existing typed `RtfError::LimitExceeded`
variant; failed reservations remain `AllocationFailed`. The compressor does
not grow an unbounded vector and all LZFu/MELA tests, examples, operations
fixtures, detection probes and performance-baseline calls supply a profile.

The exact-limit and one-byte-under regression covers the final framed output,
including its header, before allocation. Wire bytes and the LZFu checksum,
dictionary behavior and round-trip paths are otherwise unchanged.

## Public registration and source checks

The `rtf-minimal` API glossary profile registers required typed signatures for
`ValidationReport::{from_bytes, from_str}`, `validation::{validate, validate_str}`
and `transport::compress`. The RTF
README, feature matrix and human glossary describe both explicit transport
profiles and the required validation options. No raw method, output capability,
error classification or redaction rule was weakened.

Source-only checks completed in this leaf:

* `python3 -m json.tool tools/api_glossary.json`;
* `python3 -m unittest -q tools.test_check_api_glossary` (55 tests);
* `git diff --check`;
* repository search confirming no remaining default RTF validation calls or
  two-argument compressed transport calls.

## Validation lock correction and preserved diagnostic receipt

The source was frozen at `7a760a3a46541f35e3ce5acdf58c6c849b3be9b4`, but the
first isolated f5 receipt used an unauthorized regenerated root lock and is
not an acceptance result. The worktree had no root `Cargo.lock` in `HEAD`;
the ignored file was created before that run by:

```text
env CARGO_NET_OFFLINE=true CARGO_BUILD_JOBS=1 CARGO_INCREMENTAL=0 \
  CARGO_TERM_COLOR=never /home/zhuhe/.cargo/bin/cargo generate-lockfile --offline
```

Cargo selected the locally available latest compatible transitive versions,
producing root-lock SHA-256
`558fafc4fdaf4d517aef588586b8fcde52efa11f683b54e5e96a6fa7a82e2ab0` (139,607
bytes). The byte-for-byte preserved copy is outside the checkout at
`(deleted-coord-artifact)/Cargo.lock.generated-558fafc4fdaf4d517aef588586b8fcde52efa11f683b54e5e96a6fa7a82e2ab0`.
The approved lock is
`(deleted-coord-artifact)`,
SHA-256
`4100b1fbb272dd24842f33a2d6f353a3f3a97dfa7e435c4649a0e1a029bfaea1` (139,606
bytes). Complete TOML comparison found six registry substitutions: `cc`
1.9.0→1.10.0, `crypto-common` 0.1.7→0.1.6, `ctutils` 0.4.3→0.4.5,
`generic-array` 0.14.7→0.14.9, `pulldown-cmark` 0.13.4→0.13.5 and `smallvec`
1.16.3→1.16.4; affected dependency arrays consequently reference the newer
`smallvec` (and, for `cipher`/`digest`, `crypto-common`). No source manifest
change authorized those resolutions.

The prior receipt remains preserved at
`(deleted-coord-artifact)/result.json`
(SHA-256
`13d068950fc9bf167a91679d2f01734aa854f7ce8a9baa6dbcbe2baaedb2d561`). Its
RTF tests and scoped checks passed, but those results are diagnostic only
because the root lock was 558, not the approved 4100. Its standalone
`tools/perf-baseline` check also remains a failed exit-101 receipt from the
known historical PPT caller at
`crates/litchi-ppt/src/presentation/package.rs:90,283`; that lane is skipped
until the coordinator composes the owning PPT repair.

The earlier broad `tests-leaves` attempt is preserved at
`(deleted-coord-artifact)`.
No new Cargo command is recorded in this correction.

## Approved-lock f5 rerun

The ignored root lock was replaced byte-for-byte from the approved 4100
artifact before the run. The three lock identities were verified before and
after every row: root
`4100b1fbb272dd24842f33a2d6f353a3f3a97dfa7e435c4649a0e1a029bfaea1`
(139,606 bytes), `tools/perf-baseline/Cargo.lock`
`eac86284f67136aa856a8dcb6f383178b89ca92b26434a3de1d7408d68f2690c`
(57,227 bytes), and `tools/native-resave/Cargo.lock`
`51c25d3d6f36e20b50f0b19e39b1aa9c7b49777154122ecc95d3acd05c041378`
(56,975 bytes). The source was clean at commit
`c0efda769d8fc66224a1d74c3cad4056c0520232`, which includes the preceding
lock-provenance correction. The capped runner was
`1904177040fc7c523582433028d9bba56efc30fe0d9d96beb554a2fab343691e`, with
Cargo shim
`495e24c6f60b238378386d86ac4158cba904992ce8870a9494f2c4a0eedd256d`,
offline/locked Cargo, one build job, one test thread, no incremental or debug
artifacts, and 19,000,000,000-byte live/20,000,000,000-byte hard target
guards. The target was reused from the preserved diagnostic run and remained
well below the guards (maximum 4,832,336,502 bytes).

The new receipt directory is
`(deleted-coord-artifact)`.
Its `result.json` has SHA-256
`9c4f31ac7ce10820854e35ad5ec02d46aa7df0c74214c6881d21faacd5922917`; the
manifest and row receipts have SHA-256
`55411e19c32f4bf8bc2fc951ef6da9de6bfe7a8865fbe4844dd89c873462f825` and
`e08d89428924b29b9655b2d089469dd43b01fc2be60e772eb9e05c853c935a22`.
All nine selected rows passed: the all-feature/all-target RTF suite ran 190
test suites with 1,539 passed and zero failures; the exact/one-byte-under
compression-bound test and retained-source/decompression-boundary test each
passed; and RTF, `litchi-ops` with `rtf`, and facade with `rtf` all-target
checks and `-D warnings` clippy each passed. The per-row receipt log hashes
are, in order: `f04ec450d9a64636522e50bd039439b0d0256b3608d748c870495c5b7be9bb1b`,
`bf5321cd626409aebcf92f9043b141605c42914df6f6487dedcc1841a7df1e46`,
`9b1e62661e2a520b642684612cd84a6e037a2a5a07238b3e7fe8a541766c506e`,
`0f49e6beaa83c557560068bd562fe666cf1a8d0c3128a3dc51ccaf2c0a3c6eac`,
`c5ec1a7414fc07913e8bdf81a3698ff0be47923655a8c29f6f16cf542f3b7fc6`,
`8e0cc24ded715b8d53f2e6461b705d8ec5a9a689f19e6b0351e8011e7ca65926`,
`b84b732da48322406ab8a62e36ac02743cc0225387cb4e54be96c16c278e5777`,
`54e1d2fb6ea0b8ce21153be448e69a00239b71f3c09643d12fedace081651770`,
`154b954f36539b0f8a66b8ea86b1cb99e393bfaf30f771beb8ed95d895bb5d6d`.

The standalone `tools/perf-baseline` check is intentionally not repeated in
this approved-lock receipt: it remains the preserved exit-101 diagnostic
from the known historical PPT caller at
`crates/litchi-ppt/src/presentation/package.rs:90,283`, pending the
coordinator's owning PPT repair. That composed standalone check is queued for
the coordinator's post-repair validation; it is not represented as green in
this leaf. The earlier 558-lock receipt remains diagnostic only, and both it
and the broad `tests-leaves` state are preserved unchanged. This approved-lock
receipt covers the RTF source and its affected ops/facade consumers;
coordinator integration and the separate PPT repair remain outside this leaf.
