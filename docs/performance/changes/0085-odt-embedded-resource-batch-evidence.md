# Change 0085: ODT embedded-resource batch evidence

Date: 2026-08-12

Production capability: `312d66d1eebed680c3b835ed8a9f62dcacb0424d`

Status: selectable matched baseline; no comparative performance claim

## Scope and matched operation

This change adds harness, documentation, and CI coverage for existing ODT
embedded-resource replacement. It changes no production ODT code and does not
broaden the API's capability boundary.

Both cases consume the same deterministic
`litchi-odt-embedded-resource-batch-publication-v1` corpus: 200 paragraphs,
eight retained incompressible 2 MiB package resources, and 64 fixed existing
package-backed image owners with unique frame names, paths, and 4 KiB payloads.
Both replace all 64 owners onto corresponding fixed same-length target paths
without owner insertion, removal, rename, or reorder in one transaction and
one commit. The displaced source payloads remain packaged under the production
replacement contract.

- `odt_embedded_resource_scalar_replace_save` calls
  `replace_embedded_image` 64 times.
- `odt_embedded_resource_batch_replace_save` submits the same base-snapshot
  positions to one `edit_embedded_resources` call.

The immutable snapshot, replacement descriptors and payloads, expected-output
construction, sink reservation, and complete correctness oracles are outside
timing. Each interval includes transaction creation, scalar staging or one
bounded batch stage, one commit, and one complete copy to a pre-reserved
bounded sink. Because every operation is a replacement, scalar selector
shifting is irrelevant.

This is matched selectable evidence, not an optimization result. No latency,
instruction, allocation, peak-heap, RSS, I/O, or materialization improvement
is claimed without frozen CPU-pinned balanced ABBA evidence.

## Deterministic evidence and correctness oracle

The corpus has 272 logical workload entries, 77 ZIP members, and 17,061,898
archive bytes. Its SHA-256 is
`7b0ddd1c00ef91d24e60f30bf4a0ca0045807d537329e213f2f03020dfb0750b`.

The scalar output is 17,336,931 bytes with SHA-256
`2da19ec3aff1f8cf76a2690a498bb9582b604c0aab25cd40c3b688efa5888a1d`.
The batch output is 17,336,924 bytes with SHA-256
`fa71c846111de90d5cfed8e6a95493126baad291f4ef4d9f4905bf65fc54e896`.
The regression and one-sample CI smoke pin every value. Each sample must report exactly
one sink write, with accepted-byte and largest-write counters both equal to the
complete case-specific output size.

Physical scalar and batch package identity is not part of the equivalence
contract. When their deterministic byte streams differ, the retained
case-specific hashes disclose that fact. The harness independently requires
both outputs to satisfy the same complete semantic specification:

- fully reopen all 200 paragraphs and require their exact text;
- fully reopen exactly 64 packaged images and require `Content` ownership,
  fixed frame names and target paths, `image/png` declaration and manifest
  types, and
  every exact replacement payload digest;
- require all 64 displaced source paths, manifest types, and payload digests;
- require all eight retained media paths, types, and payload digests;
- require raw ZIP identity for every source member except `content.xml`, the
  manifest; this includes all 64 displaced source image members;
- require scalar commit results to be exactly 64 `Unit` values and the
  replace-only batch result to be exactly `Indices([])`;
- apply each volatile patch, apply its inverse, and reject replay on the
  already changed snapshot; and
- deterministically serialize and parse each durable patch, repeat exact
  forward/inverse checks, and reject stale replay.

Every measured publication is byte-compared with its case-specific untimed
oracle, so retained samples inherit the complete reopen, patch, preservation,
and digest proof.

## Counters and capability boundary

The current ODT transaction API accepts owned package bytes and exposes no
positional source facade or logical-Part materialization diagnostics. The
result therefore retains real bounded sink counters and deliberately omits
source and materialization fields.

Production correctness remains authoritative for the bounded API. The focused
`embedded_resource_batch` suite covers deterministic add/replace/remove
ordering, base-snapshot positions, mixed owner types, no-op/change behavior,
the exact 256-change limit and aggregate resource-byte bound, duplicate and
conflicting selectors, missing owners and paths, malformed/unsafe sources,
signed/encrypted refusals, atomic failure, durable replay/inverse, stale and
foreign sources, and partial-write behavior. The harness adds no fallback.

## Reproduction

```sh
cargo run --release --locked --manifest-path tools/perf-baseline/Cargo.toml -- \
  --warmup 10 --samples 100 \
  --case odt_embedded_resource_scalar_replace_save,odt_embedded_resource_batch_replace_save \
  --json target/perf/odt-embedded-resource-publication.json
```

For a future comparison, freeze identical harness revisions around the
production change and retain raw CPU-pinned balanced control A, candidate A,
candidate B, control B evidence. A scalar-versus-batch ratio from these
selectable cases alone is not a latency claim.

## Verification

```sh
cargo fmt --manifest-path tools/perf-baseline/Cargo.toml -- --check
cargo check --locked --manifest-path tools/perf-baseline/Cargo.toml
cargo test --locked --manifest-path tools/perf-baseline/Cargo.toml \
  media_rich_odt_scalar_and_batch_resource_replacements_are_matched -- --nocapture
cargo clippy --locked --manifest-path tools/perf-baseline/Cargo.toml \
  --all-targets -- -D warnings
cargo test --locked -p litchi-odt --test embedded_resource_batch
```

CI runs a one-sample debug smoke and a 15-sample scheduled/manual release
baseline for both cases. Real-producer documents, positional source I/O,
allocation/memory counters, filesystem-atomic save, resource addition/removal,
and comparative ABBA evidence remain outside this change.
