# Native DOC owner/public-reader phase attribution

Date: 2026-08-16

Status: correctness and phase-attribution evidence only

## Change

`litchi-doc` now has an opt-in `performance-diagnostics` feature. Under that
feature, `Snapshot::open_bounded_profiled` and `Edit::commit_profiled` emit a
finite ordered stream of content-free phase events for strict revision-owner
validation, the independent public-reader validation, exact-source retention,
in-memory owner rendering, exact no-op detection, and patch construction.
The format crate owns no clock, global recorder, source content, package
identity, or timing policy. Ordinary open and commit APIs are unchanged.

The performance harness adds the opt-in `doc_owner_public_phases` selector. It
uses the exact deterministic tiny, large, and payload-heavy native DOC writer
bytes and records the observer intervals together with edit construction,
replacement staging, outer operation, output materialization, and a checked
unattributed remainder. `Finish` is specifically the in-memory revision-owner
render; it is not external publication or file saving. Synchronous observer
overhead is included in the measured intervals.

The selector takes the current matrix from 302 to 303 names. The historical
default tranche remains 36 cases and 198 records.

## Acceptance gates and timing placement

Each corpus must pass:

- complete semantic readback through both the revision owner and the public
  package reader for the source and changed candidate;
- exact no-op source identity;
- forward/inverse patch replay plus stale-source refusal;
- malformed-source and typed out-of-range edit refusal;
- deterministic source, candidate, and output hashes;
- exact event ordering and cardinality for successful harness samples; and
- byte identity for every untouched CFB stream.

Checked arithmetic requires phase-attributed time not to exceed its outer
interval and requires attributed plus unattributed time to equal the complete
measured lifecycle for every retained sample. Semantic, no-op, patch,
refusal, hash, and preservation gates are untimed. Successful event validation
runs after each named outer interval but before the lifecycle timer stops, so
its recorder work is included in checked unattributed lifecycle time. Separate
format tests prove that strict-owner, public-reader, and finish errors close
their started event with an error outcome.

## Current evidence

Focused tests cover tiny and payload-heavy corpora, and an unoptimized debug
smoke covered all three shapes with one retained sample each. Those runs prove
selector dispatch, schema, arithmetic, and the gates above. Their timing values
are intentionally not retained as performance evidence.

No latency ranking, optimization, speedup, physical-I/O, allocation, peak-heap,
RSS, cold-cache, filesystem, or real-producer claim is accepted. A current
clean release distribution is required before selecting any production
optimization from these phases.

## Reproduction

Correctness and schema smoke:

```bash
cargo test -p litchi-doc --features performance-diagnostics \
  body_text::tests::profiled -- --nocapture
cargo test --manifest-path tools/perf-baseline/Cargo.toml \
  doc_owner_public_phase_case -- --nocapture
cargo run --release --manifest-path tools/perf-baseline/Cargo.toml -- \
  --case doc_owner_public_phases \
  --writer-shape tiny,large,payload-heavy \
  --warmup 0 --samples 1 --json /tmp/litchi-doc-owner-phases-smoke.json
```

The next acceptance run should use a clean release revision, a declared pinned
CPU and host, at least 10 warmups and 100 retained samples for all three
shapes, and a machine-readable result artifact. The result must keep every
phase vector and the checked unattributed remainder; it must not infer physical
I/O, allocation, or memory behavior from elapsed time.

## Remaining work

- Collect and independently validate the clean release distribution.
- Select a production optimization only if a phase is dominant and the change
  retains both independent validation layers.
- Add representative Word and LibreOffice producer corpora separately; do not
  treat the deterministic writer corpus as producer coverage.
- Gather allocation, peak-memory, RSS, filesystem, and physical-I/O evidence
  with the dedicated profilers before making those claims.
