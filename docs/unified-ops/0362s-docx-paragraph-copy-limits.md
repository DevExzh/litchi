# 0362s — DOCX paragraph-copy durable limits

Status: source and consumer migration complete on `uo2/w8-docx-paragraph-copy`.
This slice is based on `b96137acb4` (`api: focus docx source-backed edit surfaces`)
and is intentionally separate from the coordinator's later replay and metadata
assembly.

## Contract change

The detached DOCX paragraph-copy records are raw durable records. Their public
codec methods now require the caller's policy:

- `source_backed::paragraph_copy::raw::Patch::to_bytes(&self, limits: Limits)`
- `source_backed::paragraph_copy::raw::Patch::from_bytes(bytes, limits: Limits)`
- `source_backed::paragraph_copy::raw::RemovalPatch::to_bytes(&self, limits: Limits)`
- `source_backed::paragraph_copy::raw::RemovalPatch::from_bytes(bytes, limits: Limits)`

Encoding checks the caller's `max_durable_bytes` before reserving the output,
then keeps the embedded edit policy check. Decoding checks the caller's input
ceiling before parsing and continues to enforce the embedded policy after it is
read. The change therefore preserves the operation policy carried by a patch
while preventing an independent raw caller from silently selecting the default
64 MiB durable budget. No compatibility alias or `*_with_limits` twin was
added.

Every repository consumer was migrated in this slice. The paragraph-copy and
paragraph-removal integration tests pass explicit limits and now cover a caller
policy one byte below the encoded patch. The two perf-baseline DOCX consumers
pass their active limits to both paragraph-copy codecs; the same consumers also
use the canonical `tail_append::EditOptions::new(Limits)` path. The paragraph
index test was migrated from the removed positional options constructor to the
canonical limits carrier. The public plain-paragraph wrappers retain their
existing owner API and document the required limits argument for durable
serialization.

The raw records remain partial, source-bound patches. The operation's embedded
limits, source fingerprint, inverse direction, and exact byte validation are
unchanged; only the caller admission policy is made explicit before allocation.

## Validation

The source seed diff was retained externally as
`(deleted-coord-artifact)/paragraph-copy-limits.diff`
(2,622 bytes, SHA-256
`e8af4ae839dcbb3ded8bc78b13520331003d95f98f042b1ac5dce0eaf14266d5`).
The final worktree source is `b96137acb4` plus the changes in this record.

Commands used the capped environment from
`(deleted-coord-artifact)`, with
`CARGO_TARGET_DIR=(deleted-coord-artifact)`
and `CARGO_BUILD_JOBS=2 CARGO_TEST_JOBS=2` unless stated otherwise.

- Focused raw suites:
  `cargo test --offline --locked -p litchi-docx --test source_backed_paragraph_copy --test source_backed_paragraph_removal`
  passed 12 copy and 6 removal tests. Log:
  `(deleted-coord-artifact)/w8-docx-paragraph-tests.log`
  (6,060 bytes, SHA-256
  `ecc49b3e83831f5260893284d60a729937106cd9dd847e433c458379f7eb3f25`).
- Full DOCX targets:
  `cargo test --offline --locked -p litchi-docx --all-targets`
  passed 2,270 tests across 85 suites, with no failures or ignored tests. Log:
  `(deleted-coord-artifact)/w8-docx-paragraph-all-tests.log`
  (215,145 bytes, SHA-256
  `05f48cabaec1f769aecaafbd7a99e7e1960998ea7df146ab0df9a86d3d9a37c3`).
- DOCX all-target compilation:
  `cargo check --offline --locked -p litchi-docx --all-targets`
  passed. Log:
  `(deleted-coord-artifact)/w8-docx-paragraph-docx-check-v2.log`
  (175 bytes, SHA-256
  `7a3ed482f5621dd447a8d50cc3ea13f59fedbb3b1ff6fe807bc14d92e9e2736`).
- Perf-baseline all-target compilation:
  `cargo check --offline --locked --manifest-path tools/perf-baseline/Cargo.toml --all-targets`
  passed after migrating the raw arguments and canonical tail options. Log:
  `(deleted-coord-artifact)/w8-docx-paragraph-perf-check-v4.log`
  (60 bytes, SHA-256
  `9ae36bb22d8f0f810a77044760863f70b81405396a32cce477ec08eecd4e6d59`).
- Formatting:
  `cargo fmt --all -- --check` passed.
- Before the layout repair, strict DOCX library Clippy was also run with
  `cargo clippy --offline --locked -p litchi-docx --lib --all-features -- -D warnings`.
  It reports one pre-existing `clippy::large_enum_variant` diagnostic for
  `source_backed::tail_append::TailChange` (the 297-byte `Forward` variant),
  outside this slice. This pre-repair receipt is retained only to show the
  original diagnostic; log:
  `(deleted-coord-artifact)/w8-docx-paragraph-clippy-lib-allow-existing.log`
  (1,839 bytes, SHA-256
  `e1e8ba2349ad62512ee1f362b2fad85f209e2cf262ef6591ffb1b789e528eb92`).
  The corresponding pre-repair all-target run with only that existing lint
  allowed also passed; log:
  `(deleted-coord-artifact)/w8-docx-paragraph-clippy-all-allow-existing.log`
  (7,818 bytes, SHA-256
  `18db4aa40e46ca5bddd70c19e53da29e19ff5d6f0202c5b2d57c14b204a47437`).

No performance measurement or compiled artifact is part of this commit. All
large logs, targets, and binaries remain outside the repository.

## Follow-up layout repair

The original deferred change representation used a large `Forward` enum
variant and boxed every non-noop inverse. The follow-up stores the forward
fields in one private `TailChange` struct and toggles a private inverse bit.
This keeps the public API and forward field layout unchanged, removes the
inverse box allocation, and leaves inverse publication and exact no-op
behavior unchanged. It also removes the `large_enum_variant` diagnostic
without a lint suppression or an added hot-path allocation.

The final unfiltered all-target/all-feature Clippy command was:

```text
cargo clippy --offline --locked -p litchi-docx --all-targets --all-features -- -D warnings
```

It passed. The log is
`(deleted-coord-artifact)/w8-docx-paragraph-clippy-final.log`
(7,818 bytes, SHA-256
`265e0b39a6ce4d23c34acaa99d21551a3684aab67fb5b93ff22d8bb3aa8ea762`). The
focused post-repair suites were:

```text
cargo test --offline --locked -p litchi-docx \
  --test source_backed_tail_append \
  --test source_backed_tail_append_stream \
  --test paragraph_index_reuse
```

All 87 tests passed; the log is
`(deleted-coord-artifact)/w8-docx-paragraph-focused-final.log`
(7,724 bytes, SHA-256
`6574babf41545784bfe6baa91b3121bcba64c15861083e27d148ec21b52b914e`). The
full all-target run immediately before the focused receipt passed 2,270 tests
across 85 suites; its log is
`(deleted-coord-artifact)/w8-docx-paragraph-tests-final.log`
(217,628 bytes, SHA-256
`bb1f0f455bf2ea89fddc8a0a95a66d5682679ca65cf9e47be6fb3365611f34f8`). The
earlier allow-only Clippy logs remain historical evidence for the original
diagnostic and are superseded by the strict run above.
