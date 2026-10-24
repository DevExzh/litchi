# 0362ul37 legacy consumer API repair

This repair closes the consumer/documentation gaps found during the
final-scope audit. It is based on the clean repair leaf
`uo2/corpus-gate-repairs-2` at
`3431ba566c0c0a89f73f292f7b8b5d638c7c0944`. No production API was added,
aliased, deprecated, or renamed by this record.

## Consumer and documentation changes

* The RTF `parse_rtf` fuzz target now calls
  `transport::decompress(data, DecompressionLimits::new(MAX_INPUT_BYTES))`.
  The same `MAX_INPUT_BYTES` value remains the fuzzer's input skip bound and
  the decompressed-output bound; the removed `transport::Limits`/
  `decompress_with_limits` spelling was not retained as an alias.
* RTF's README and feature matrix describe the canonical `decompress` API and
  its `DecompressionLimits` argument.
* The CLI API table names `batch::from_bytes` with its required
  `SourceLimits`, and `File::refresh_with_options`.
* The operations fuzz README names `batch::from_bytes` with a required
  `SourceLimits` profile.
* The DOC README example uses the public `SourceLimits::builder()` setters and
  `OpenOptions::source_limits`.
* The PPT embedded-font guide is unchanged in the final repair. Its public
  record/parser limit contract is being repaired separately in 0362ul39;
  this record does not classify those public entry points as internal.
* The XLSB feature matrix uses `Package::{xml_maps, edit_xml_maps,
  edit_xml_maps_with_options, apply_xml_maps}`, `Builder`'s canonical XML-map
  methods, and the public writer `Worksheet` name.

## Gate registration and provenance

The existing 53-step bounded runner keeps the existing `fuzz` step and adds
one command to it:

```
cargo check -q --manifest-path crates/litchi-rtf/fuzz/Cargo.toml --bins
```

The step still runs the existing core, operations, facade, and smoke checks in
their original order; no new gate row or fuzz framework was introduced. The
runner's `lock_hashes` list now also observes
`crates/litchi-rtf/fuzz/Cargo.lock` when the coordinator supplies that ignored,
offline lock. At the initial source-only endpoint both `Cargo.lock` and
`crates/litchi-rtf/fuzz/Cargo.lock` are absent, so there is no lock digest to
claim here and no lock was generated or copied during this documentation-only
turn. The eventual locked check must prepare that lock from the pinned offline
cache and record its SHA-256 before launch.

The planned bounded validation, to be run only after the coordinator grants a
quiet Cargo slot and supplies the lock, is:

```
env -u LITCHI_BLESS -u LITCHI_CAPPED \
  PATH=/home/zhuhe/code/litchi-worktrees/uo2-corpus-gate-repairs-2/docs/unified-ops/results/0362ul29/tooling/bin:/home/zhuhe/code/litchi-worktrees/uo2-corpus-gate-repairs-2/.coord/bin:$PATH \
  LITCHI_MEM_MAX=20G CARGO_BUILD_JOBS=1 CARGO_NET_OFFLINE=true \
  CARGO_INCREMENTAL=0 CARGO_PROFILE_DEV_DEBUG=0 CARGO_PROFILE_TEST_DEBUG=0 \
  CARGO_TARGET_DIR=/home/zhuhe/code/litchi-worktrees/.coord/rtf-fuzz-consumer-3431/target \
  /home/zhuhe/code/litchi-worktrees/uo2-corpus-gate-repairs-2/docs/unified-ops/results/0362ul29/tooling/bin/cargo \
  check --offline --locked -q --manifest-path \
  /home/zhuhe/code/litchi-worktrees/uo2-corpus-gate-repairs-2/crates/litchi-rtf/fuzz/Cargo.toml \
  --target-dir /home/zhuhe/code/litchi-worktrees/.coord/rtf-fuzz-consumer-3431/target --bins
```

The wrapper and the explicit `--offline --locked` flags are intentional. The
external target must remain below 20,000,000,000 bytes, with the wrapper's
20 GiB memory policy and one build job. This command was not run in this
isolated documentation turn.

## Fuzz-consumer audit

A tracked-tree scan covered every `crates/*/fuzz/**` consumer for removed
`*_with_limits`/`decompress_with_limits`/`transport::Limits` names, facade
`*_with_limits` names, `Mutable*` names, and banned transaction vocabulary.
The RTF call above and the operations README were the only real API consumer
gaps. The iWork protocol fuzz helpers named `options_with_limits` and
`execute_with_limits` are local protocol-bound helpers, and the Keynote fuzz
target's `UnknownMutable*` strings are synthetic wire variants; neither is a
removed public API reference. No other fuzz target was changed.

Read-only validation completed before this record was written:

* `git diff --check` passed.
* Python AST parsing of the changed gate runner passed.
* Read-only plan construction reports 53 unique step names and one `fuzz`
  step (zero-based step index 41, the 42nd step); the new RTF command is
  present in that step.
* No Cargo command, lock generation, native run, corpus replay, or blessing
  was performed in this source-only turn.

The historical gate outputs and prior source/lock receipts remain immutable;
the actual bounded check is recorded separately below.

## Follow-up bounded validation

The coordinator later granted the isolated RTF check on commit
`39a417576f4931f136f68bb57c3371c05cd202a7` in the same leaf. The first
offline metadata attempt was intentionally retained as a diagnostic: the
sparse checkout did not yet materialize `crates/litchi-codepage`, so it exited
101. Its empty JSON has SHA-256
`e3b0c44298fc1c149afbf4c8996fb92427ae41e4649b934ca495991b7852b855` and its
stderr has SHA-256
`2713acaf0fe44b9d4c8ba62b93d20d34fd82659845eb77e6ccc8405c35ce84a8`.
Materializing that path and its `crates/litchi-core` dependency made the
second metadata run pass: 87 packages, JSON SHA-256
`3d2f9d4e2cd4eac396915bd8d271fad7dd7d3210c0fb3ed4af60276a9401ce19`.

The approved root lock remained SHA-256
`4100b1fbb272dd24842f33a2d6f353a3f3a97dfa7e435c4649a0e1a029bfaea1`; offline
metadata generated the ignored standalone fuzz lock with SHA-256
`f040accf51bb28293d71218f497018039566c40f5feea1841f08b7927741f3f2`. Its 83
registry rows were a complete subset of the approved root lock's 518 rows.
The pinned Cargo wrapper was SHA-256
`1b22a69fd625d57565041bf6c634d8ebd1a82be8b3d0006390472c86324230d8`.

With `LITCHI_CAPPED` and `LITCHI_BLESS` unset, offline and locked mode, one
build job, incremental compilation disabled, `MemoryMax=20G`, no swap, and an
external target, `cargo check --bins` passed with exit 0 in 72.066 seconds.
The compile result is SHA-256
`e3ff2f3f60861718b79d22ccd3095785a32327769294e41cc9ada80b65dda734`; the
target peaked at and ended at 118,970,506 bytes, below the 20,000,000,000-byte
guard, with no guard errors. The supervisor's post-exit samples contained no
owned Cargo, rustc, or timeout process. The complete external receipt is
`rtf-fuzz-39a417-20261024/receipt.json`, 5,060 bytes, SHA-256
`ef1fc17ff1e95bab311bacc130e06dfeb292fcf86090a3d61c5071bc6b4c4f3a`.
This validation ran only metadata and the authorized bounded check; no native
run, corpus replay, or blessing was performed.
