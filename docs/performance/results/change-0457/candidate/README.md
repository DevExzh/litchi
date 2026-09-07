# 0457 source-tail candidate capture

This bundle describes the opt-in `odp_source_tail_append_lifecycle` lane. It
uses the same deterministic ODP source corpus as the frozen 0457 control:
64, 4,096, and 8,192 source slides; one appended title/body slide; normal and
allocator binaries; 30 retained samples after three warmups; CPU 2; one
worker; and two serialized repeats. R1 is forward order and R2 is reverse
order.

The lane measures the specialized
`SourceBackedPackage::from_read_at + SourceBackedTailAppendEdit::plan +
SourceBackedTailAppendPublicationPlan::write_to` lifecycle. It has a different
publication-report and retained-result contract from owned `Snapshot`/`Commit`/
`Patch`; the capture cannot support an ordinary CRUD or Commit/Patch speedup
claim.

`corpus-bindings.json` is derived from the frozen control reports but is used
as an independent candidate oracle input. It binds the source archive and
content hashes, corpus manifest, member names and identities, opaque payload,
semantic/order/text projections, and expected append request. The control
output is used for semantic/order/text/count references only. The style-free
candidate page can therefore have a different `content.xml` spelling and ZIP
framing. Before formal capture, `bind-output.py` retains one source/output
archive pair per shape and runs `native/verify-output.py`; its immutable
`output-binding.json` records the native receipt, candidate archive/content
hashes and bytes, and every source/output member identity. The report oracle
requires those independently proven candidate identities while requiring all
untouched members to equal the source. It also requires source/output manifest
gates, source immutability, stale-source refusal, proof fields, stable
semantic/order/text projections, and the returned publication report length.

The untimed `proof_source_version_id` and `proof_source_version_revision` are
retained as fixture metadata. Every retained timed sample must carry aligned
`runtime_proof_source_version_id`/`runtime_proof_source_version_revision`
vectors equal to the corresponding publication-report vectors. The oracle
keeps those identities separate and does not compare a fresh timed provider ID
with the fixture scalar.

The generic source observer exposes only logical `ReadAt` call count and
returned bytes. Requested lengths, largest ranges, request histograms, and
compressed/decompressed/recompressed boundaries are deliberately required to
remain `unavailable`; the oracle never substitutes source bytes for them.
Sink accepted bytes, accepted write calls, largest accepted write, and bucket
counts are required to be measured and consistent with the publication output.
Allocator mode requires operation-scoped vectors and live-byte balance. The
ambient `MALLOC_*`/`GLIBC_TUNABLES` observation must be empty, and `RUSTFLAGS`
must be unset or empty so reports retain `rustflags: null`.

The shared in-process operation envelope records
`latency_claim: comparable_timed_operation`; the source-tail summary keeps the
specialized publication scope and explicitly withholds an ordinary
Commit/Patch comparison or speedup claim.

After the candidate build has passed its source-custody check, retain both
candidate executables and bind them before the first capture. The binder does
not build or execute anything and refuses to overwrite an existing binding or
any receipt after capture has started:

```sh
mkdir -p /tmp/litchi-goal-0457/candidate
cp tools/perf-baseline/target/release/litchi-perf-baseline \
  /tmp/litchi-goal-0457/candidate/litchi-perf-baseline
cp tools/perf-baseline/target/release/litchi-perf-baseline-alloc \
  /tmp/litchi-goal-0457/candidate/litchi-perf-baseline-alloc
python3 -B docs/performance/results/change-0457/candidate/bind.py \
  --build-receipt docs/performance/results/change-0457/checks/candidate-build.json \
  --repo-root /home/zhuhe/code/litchi \
  --normal-source tools/perf-baseline/target/release/litchi-perf-baseline \
  --allocator-source tools/perf-baseline/target/release/litchi-perf-baseline-alloc \
  --normal-copy /tmp/litchi-goal-0457/candidate/litchi-perf-baseline \
  --allocator-copy /tmp/litchi-goal-0457/candidate/litchi-perf-baseline-alloc
```

The candidate build/probe must first emit the raw source and generated
style-free output archives for each shape. Bind those files once, before R1;
the helper copies them into the candidate bundle, runs the independent native
oracle, and updates the pending output-binding hash in `protocol.json`:

```sh
python3 -B docs/performance/results/change-0457/candidate/bind-output.py \
  --repo-root /home/zhuhe/code/litchi \
  --native-oracle docs/performance/results/change-0457/native/verify-output.py \
  --protocol docs/performance/results/change-0457/candidate/protocol.json \
  --fixture tiny /path/to/tiny-source.odp /path/to/tiny-output.odp \
  --fixture medium /path/to/medium-source.odp /path/to/medium-output.odp \
  --fixture large /path/to/large-source.odp /path/to/large-output.odp
```

The command fails if a retained fixture, output binding, protocol hash, or
capture receipt already exists. Keep `output-fixtures/`, `output-binding.json`,
the native oracle receipt, and the final protocol immutable with the formal
reports.

Then run the two phases serially, preserving any failed attempt directory:

```sh
python3 -B docs/performance/results/change-0457/candidate/capture.py \
  --phase R1 --attempt formal \
  --repo-root /home/zhuhe/code/litchi \
  --protocol docs/performance/results/change-0457/candidate/protocol.json \
  --custody-driver docs/performance/results/change-0457/check.py
python3 -B docs/performance/results/change-0457/candidate/capture.py \
  --phase R2 --attempt formal \
  --repo-root /home/zhuhe/code/litchi \
  --protocol docs/performance/results/change-0457/candidate/protocol.json \
  --custody-driver docs/performance/results/change-0457/check.py
```

An individual retained report can be replayed independently after capture:

```sh
python3 -B docs/performance/results/change-0457/candidate/oracle/verify-report.py \
  --report docs/performance/results/change-0457/candidate/runs/R1/formal/R1-normal-tiny-r1.json \
  --mode normal --shape tiny
```

The driver records the exact protocol, driver, binding, source snapshot,
argv, binary identity, ambient allocator environment, workload/resource logs,
oracle output, and per-lane failure receipts. It stops at the first failed
lane and never overwrites evidence. Use `-B` and keep the protocol/oracle,
binding, raw reports, and receipts immutable once R1 starts.
