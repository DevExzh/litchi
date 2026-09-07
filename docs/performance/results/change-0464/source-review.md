# 0464 frozen source review

This review covers the final 0464 harness source epoch recorded by
`checks/harness-clippy-r3.json`: custody record
`a35f4507a74e7678f29f91dd6857d07a87579ba872785a1eaf19e296540debbe` (7,032
Rust/TOML/lock files). The relevant file hashes are:

| file | SHA-256 |
| --- | --- |
| `tools/perf-baseline/src/pptx_pair_lifecycle.rs` | `17e96bcd64c7971bf25a0cdc152ce193077eaa95511d7bbb37dccb404855d222` |
| `tools/perf-baseline/src/bin/pptx_semantic_inventory.rs` | `94edaab83256b0b83ef0bea8db08922b87cac6fd06f5b672466305e1e4d3af80` |
| `tools/perf-baseline/src/lib.rs` | `320d3acb1204507e3ad723322bc1af66c08443a260ba3ae8a436b20cd36f91ab` |
| `tools/perf-baseline/src/main.rs` | `5658a7afeb934fc7b85673f1ac53b6a0d19f6a186354144f308085e4d8e0c201` |
| `tools/perf-baseline/src/bin/litchi-perf-baseline-alloc.rs` | `c10f0c79d672a7d544c4190357809437a3c3a996af02f0e5700b33ea7df2a16f` |

The review is limited to source correctness, custody, bounds, oracle scope,
and the available receipts. It does not make a timing, speedup, or retention
claim. The runner is harness-only and does not change production PPTX code.

## Review result

No source blocker remains in the frozen candidate after the input-read,
archive-limit, provenance, and raw-oracle fixes. The runner now rejects an
oversized or changing input before retaining more than the configured bound,
binds the operation to a closed provenance vocabulary and exact source
revision, and applies finite limits to both OPC parsing and the raw ZIP oracle.
The generated relationship parts produced by a valid copy are accounted for
separately from copied payload parts. The independent Python driver still
requires its package closure and relationship oracle after every successful
Rust report.

The retained `checks/pair-pilot.json` is intentionally a pre-fix receipt. It
failed with `added member ppt/slides/_rels/slide3.xml.rels has no source payload
identity` under the earlier Rust source. It is superseded by the fresh
post-fix smoke reports and is not a final positive result.

## Input custody and provenance

`read_bounded` in `pptx_pair_lifecycle.rs:484` checks that the path is a
regular file, rejects a metadata size above the manifest limit, reads through
`File::take(limit + 1)`, and rejects a concurrent growth beyond the limit.
`load_inputs` then checks both exact byte counts and SHA-256 values against the
manifest (`:645`). The final `checked_file_identity` path reuses that bounded
reader after all iterations (`:1500`), so input replacement, shrinkage, or
growth cannot be silently accepted.

Manifest normalization (`:515`) validates lowercase 64-character input and
output digests, rejects equal inputs unless the provenance kind is `self`, and
accepts only `self`, `derived`, `same-source-derived`, or `independent`.
`source_revision` is mandatory and must be exactly 40 lowercase hexadecimal
characters (`:602`); the optional CLI override is validated and must equal the
manifest value (`:802`). The retained pair correctly declares
`same-source-derived` and `independent_producer_claim: false`, so the fixture
does not acquire an independent-producer claim from the runner.

The Python capture driver independently binds the pair, binary, build receipt,
source custody record, manifest, protocol, and oracle hashes. It invokes the
external oracle only after a successful Rust report and fails the lane when the
oracle returns nonzero (`capture.py:346-351`). It also checks pair input and
source-record identity before and after each lane. These checks complement the
Rust path rather than being inferred from the Rust report.

## Parsing and resource bounds

The pair runner's `read_limits` (`pptx_pair_lifecycle.rs:849`) binds the
compressed input ceiling to `input_limit` and binds ZIP entry, aggregate ZIP,
single-part, and aggregate-part expansion to the configured memory ceiling. It
also supplies finite member-name and metadata limits, content-type and
relationship XML limits, XML attribute and relationship-target limits, and
the package's default finite count/graph/event ceilings. The source-backed
semantic pre-read uses this same profile (`:992`), so it does not silently
fall back to unbounded OPC defaults.

The raw member oracle has a second explicit ZIP profile
(`archive_limits` at `:1028`) and passes it to both `PreservationIndex` and
`ArchiveReader` (`:1040`). It bounds member count, names, central metadata,
compressed member size, per-entry expansion, and aggregate expansion before
materializing payloads. Its scratch buffer is fixed-size; raw-oracle work is
outside the timed API phases and outside the report's owner budget by design.

The semantic inventory applies finite input, ZIP aggregate, ZIP entry,
materialized-part, and aggregate-part limits (`pptx_semantic_inventory.rs:216`).
Its input reader uses metadata plus `take(MAX_INPUT_BYTES + 1)` and fallible
reservation (`:226`), while slide/image counts, descriptor strings, retained
semantic bytes, and serialized report bytes have explicit ceilings. The eager
and source-backed views receive the same checked `ReadLimits` profile
(`:477-483`), and output creation is exclusive and synced (`:454-464`).

These are parser and retained-artifact bounds, not a claim that the complete
process has at most `memory_bytes` resident bytes: the runner intentionally
loads caller-owned input copies and performs semantic/raw oracles outside the
timed owner contexts. The report exposes that timing and allocation scope.

## Raw and semantic correctness

The Rust semantic check derives source and destination slide expectations before
timing, then reopens the output and compares slide order, names, text, and
direct internal image bytes (`pptx_pair_lifecycle.rs:992-1025`). The raw check
requires every destination member to remain present and byte-identical except
the three declared metadata owners. Added non-relationship members must carry
the payload of a source member; added `.rels`/`/_rels/` members are counted as
regenerated relationship parts and are not incorrectly required to equal an
old relationship payload (`:1082-1139`).

That Rust raw check deliberately does not claim dependency-closure mapping:
payload equality against any source member is weaker than source-part-to-output
mapping. The retained `pair-oracle.py` performs the independent slide closure,
relationship retargeting, content-type, inheritance, metadata splice, and
copied-payload checks. The capture driver makes that oracle mandatory, so a
successful formal lane cannot rely on the weaker generic Rust oracle alone.

The semantic inventory emits source-backed image/name/relationship observations
and an eager ordered-text observation, but it does not assert source/eager
equivalence or archive-byte preservation. Its module documentation states that
scope. That is suitable as an inventory artifact; it must not be described as
an independent semantic proof.

## Lifecycle, timing, and failure behavior

The source and destination owners use separate execution contexts and
caller-owned providers. The four timed regions begin immediately around source
open, destination open, planning, and publication; adapter creation, sink
reservation, diagnostics, semantic/raw oracles, artifact writes, and drops are
outside the clocks (`pptx_pair_lifecycle.rs:1214-1303`). Operation-scoped
allocator counters use the same boundaries, and the normal binary omits those
fields. Publication checks sink accounting, output identity, semantic
readback, raw preservation, source immutability, and the expected output
identity before lifecycle teardown. Final teardown checks that the source and
destination `Memory`, `Objects`, and `Depth` budget resources return to zero
(`:1431-1443`).

The current protocol is a serialized one-worker positive lifecycle. It does not
exercise cancellation refusal, malformed-package refusal, an over-budget
negative lane, or parallel scaling; those remain test-plan gaps rather than
proofs supplied by the positive capture. The `owner_context` cancellation
source is intentionally not exposed to the workload, so no cancellation claim
should be attached to a report.

## Receipts and focused evidence

- `checks/harness-clippy-r3.json` is PASS with warnings denied across all
  targets/features, and its source-before/source-after record is the final
  `a35f4507…` epoch above.
- `checks/normal-build-r1.json` and `checks/allocator-build-r1.json` are PASS
  against the same final source record. `normal-final-binding.json` binds both
  retained binaries and the inventory binary to that record.
- The pair runner's focused source tests cover malformed digests, same-path
  rejection, selector retention, unknown provenance, missing revision, and
  malformed revision (`pptx_pair_lifecycle.rs:1640-1708`).
- The fresh post-fix smoke directories each contain four successful positive
  reports (`normal-bytes`, `normal-range`, `allocator-bytes`, and
  `allocator-range`). Every report reproduces the 55,891-byte output with SHA
  `550e8d8e…`, and exercises both the normal and operation-scoped allocator
  binaries. The initial `smoke.py` check receipts are marked failed because two
  negative-case assertions expected the old diagnostic strings; the
  observed failures were the correct bounded `ArchiveMetadataBytes` and
  `OutputBytes` errors. This is a receipt-driver wording defect, not a source
  correctness failure, and remains visible in the evidence bundle. After
  correcting those assertions, `checks/smoke-r2.json` passes all four positive
  configurations and seven rejection cases; `smoke-results-r2.json` retains
  the successful checks and artifact identities.
- `checks/oracle-tests.json` is PASS. Its four negative cases reject wrong
  insertion order, copied-payload mutation, relationship retargeting, and an
  unexpected output member; the positive derived pair also passes. This is
  independent package-oracle evidence, with the provenance limitation stated
  above.
- `checks/capture-r1.json` and `checks/capture-r2.json` are PASS. Each repeat
  completed all four lanes in its declared order; each lane retained 30
  samples after three warmups, returned oracle exit code 0, and preserved both
  input files and the final source custody record. Thus the final bundle has
  two repeats, eight reports, 240 retained samples, and 55,891-byte outputs.
- The earlier `harness-clippy.json` and `harness-clippy-r2.json` failures are
  retained historical receipts; the final r3 receipt is the applicable source
  check.

The candidate is source-reviewable and bounded, and the post-fix formal capture
is complete for the declared derived pair. The bundle still makes no native
Office acceptance or independent-producer claim, and the derived pair's
provenance remains narrower than an independently authored native PPTX pair.
