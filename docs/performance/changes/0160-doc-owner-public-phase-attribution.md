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
smoke covered all three shapes. A clean release build of exact revision
`ab333008d31b1f63ee0a84c6087fee0de48895d1` then ran on the named AMD EPYC
9575F host. Four fresh processes per shape were pinned to CPU 2; each used 20
warmups and retained 200 samples. The worktree was clean. All untimed
case-level semantic, patch, refusal, hash, and untouched-stream gates passed in
all 12 reports; all 2,400 timed samples passed arithmetic, event, and output
checks.

| Shape | Complete lifecycle p50 / p95 | Initial + final public-reader validation p50 | Patch p50 | Replacement staging p50 |
|---|---:|---:|---:|---:|
| tiny | 0.081 / 0.098 ms | 0.016 ms | 0.026 ms | 0.014 ms |
| large | 1.157 / 1.246 ms | 0.598 ms | 0.165 ms | 0.174 ms |
| payload-heavy | 44.227 / 46.712 ms | 20.721 ms | 8.413 ms | 7.470 ms |

This accepts a phase ranking only for the exact named deterministic release
distribution. The complete public-reader validations are the largest grouped
named phase for the large and payload-heavy shapes; patch fingerprinting is
largest for tiny. Across-process lifecycle p50/mean spread is at most 2.98% /
3.76%; the p50 spread of the three named phase groups in the table is at most
3.42%. Tiny public-reader and replacement means cross the predeclared 5%
review trigger at 6.45% and 5.57%, but their phase rank does not change and no
cross-shape aggregate is formed. This is not a control/candidate comparison,
so no optimization or speedup is accepted. Physical-I/O, allocation,
peak-heap, RSS, cold-cache, filesystem, and real-producer claims also remain
open.

The complete vectors and environment are retained in twelve compressed raw
reports identified by the [SHA-256 manifest](../results/doc-owner-public-phases-0160.sha256),
with a [machine-readable summary](../results/doc-owner-public-phases-0160-summary.json).
The clean release binary has SHA-256
`4ebdfec6c70e4a5d40936824a8e1acc80a3579fa04dfd3bd3054ad9955466c61`.

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

Accepted attribution run:

```bash
for rep in 1 2 3 4; do
  for shape in tiny large payload-heavy; do
    taskset -c 2 litchi-perf-baseline \
      --case doc_owner_public_phases \
      --writer-shape "$shape" \
      --shape tiny --payload compressible \
      --semantic-shape tiny --rtf-variant plain --workers 1 \
      --warmup 20 --samples 200 \
      --json "/tmp/doc-owner-phases-r${rep}-${shape}.json"
  done
done
```

The recorded result keeps every phase vector and the checked unattributed
remainder. It does not infer physical I/O, allocation, or memory behavior from
elapsed time.

## Remaining work

- Evaluate a private sharing/copy-elision mechanism against this distribution;
  retain both independent validation layers and require a balanced clean
  release control/candidate comparison before accepting a latency result.
- Add representative Word and LibreOffice producer corpora separately; do not
  treat the deterministic writer corpus as producer coverage.
- Gather allocation, peak-memory, RSS, filesystem, and physical-I/O evidence
  with the dedicated profilers before making those claims.
