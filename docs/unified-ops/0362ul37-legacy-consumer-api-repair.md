# 0362ul37 legacy consumer API repair

This source-only repair closes the consumer/documentation gaps found during the
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
* The PPT embedded-font policy documents the public package ingress route,
  `OpenOptions::source_limits`; record/font `*_with_limits` entry points are
  crate-private implementation details.
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
offline lock. At this source endpoint both `Cargo.lock` and
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
  step (the current source order places it at row 42); the new RTF command is
  present in that step.
* No Cargo command, lock generation, native run, corpus replay, or blessing
  was performed.

The historical gate outputs and prior source/lock receipts remain immutable;
this record is only the source-ready consumer repair and validation handoff.
