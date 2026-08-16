# Change 0165: DOC lazy fingerprints and same-lineage patch replay

Date: 2026-08-17

Status: accepted correctness and attribution evidence for the named native-DOC
lifecycle and workflow comparison. The positive-faster deltas are descriptive;
this is not a speedup, generic DOC, memory, I/O, or CRUD-completeness claim.

## Scope and mechanism

The native `litchi-doc::body_text::Snapshot` now keeps its diagnostic FNV-1a
fingerprint in an inline `OnceLock<u64>`. Opening a snapshot, cloning an
uninitialized snapshot, constructing a patch, and applying a patch from the
same immutable source lineage do not scan the complete artifact merely to
populate that diagnostic value. A clone copies an already initialized value;
an uninitialized clone keeps an independent lazy cache.

`Patch::new` no longer computes complete before/after fingerprints eagerly.
`Patch::is_noop` and `Patch::apply` first use the immutable `Arc<[u8]>` pointer
and length as a same-lineage fast path. A same-lineage no-op returns a cloned
source snapshot, while a same-lineage changed patch returns the retained
after snapshot. For an independently reopened or otherwise non-identical
allocation, the normal lazy fingerprint check is followed by exact byte
comparison; the FNV value is not an authorization boundary. Collision,
stale-source, foreign-source, inverse, and output checks remain exact and
failure-atomic.

The public `source_fingerprint` and `target_fingerprint` accessors are now
non-`const` because their first call may initialize the lazy cache. No public
constructor could create a `Patch` in a const context, so this removes a
misleading capability while making the deferred work explicit. The change
adds no unsafe code, global state, dependency, source ownership shortcut, or
unbounded cache.

## Corpus and correctness closure

The existing deterministic tiny, large, and payload-heavy native-DOC writer
corpora are retained. The established `doc_owner_public_phases` lifecycle
boundary remains unchanged: strict owner validation, complete public-reader
validation, source retention, edit construction, replacement staging, owner
rendering, patch construction, and output materialization remain attributed in
the original vectors. Same-lineage apply and the first source/target
fingerprint demand are explicit post-lifecycle workflow extensions rather than
hidden inside `measured_total_ns`.

Case-level preflight checks cover semantic reopen, source immutability, inverse
restoration, stale-source refusal, malformed/typed refusal, and untouched CFB
streams. Focused crate tests separately cover foreign-source and synthetic
fingerprint-collision rejection. Every retained sample checks the independently
computed source and target FNV-1a values, same-lineage replay, output identity,
and workflow arithmetic. The harness computes expected fingerprints independently
of the production helper. The complete raw reports, per-leg statistics,
binary/source identities, and SHA-256 manifest are linked from the [machine-readable
summary](../results/doc-lazy-fingerprint-0165-summary.json) and [release
manifest](../results/doc-lazy-fingerprint-0165-manifest.json).

## Final exact revisions and binaries

The clean control source revision is
`d6818e290aa77fd7666b7b16ee6908319d0f332b`; the clean candidate source
revision is `5dd813b1e108e253457ccb6c504c125c2becc1c6`. The corresponding
release binary SHA-256 values are:

| Role | Source revision | Binary SHA-256 |
|---|---|---|
| control | `d6818e290aa77fd7666b7b16ee6908319d0f332b` | `344c0504c254109ee6b4361e375599d187f8a12333abb44f207d837af259ef8c` |
| candidate | `5dd813b1e108e253457ccb6c504c125c2becc1c6` | `c95e6c6004cbd725c789597566a81c0897ab6915ecd7c274deab222d134b3fd3` |

Both release builds were clean exact-revision builds. The final evidence uses
Linux 6.8.0-101-generic on the named AMD EPYC host, Rust 1.95.0, system
allocator, CPU affinity 2, and one execution worker.

The harness records the clean invoking worktree's runtime Git HEAD in every raw
report. Both binary roles were invoked from the candidate worktree, so that
environment field is `5dd813b1e...` even for A1/A2 control legs; it is not the
binary provenance field. Control/candidate identity is instead bound by the
separate source revisions and verified release-binary SHA-256 values above and
in the machine-readable summary.

## Balanced release result

Clean release binaries ran in strict CPU-2 `A1 control, B1 candidate, B2
candidate, A2 control` order. Each leg used 20 warmups and 500 retained
samples for each of the three shapes, for 6,000 retained lifecycle samples.
Positive percentages mean the candidate is faster. Each cell is
`p50 / mean / p95`, with `A1 -> B1 / A2 -> B2` in order.

| Shape | Positive-faster lifecycle delta | Evidence decision |
|---|---:|---|
| tiny | +33.78% / +35.19% / +38.94%  /  +33.21% / +34.76% / +39.67% | retained |
| large | +12.27% / +12.59% / +17.53%  /  +13.81% / +13.55% / +11.68% | retained |
| payload-heavy | +17.33% / +17.09% / +16.58%  /  +17.80% / +17.75% / +16.25% | retained |

The same-implementation lifecycle drift is retained as a host-stability
disclosure. Control A1 to A2 p50/mean drift is `-1.18%/-1.41%` tiny,
`+0.26%/-0.42%` large, and `+0.47%/+0.72%` payload-heavy. Candidate B1 to
B2 drift is `-0.33%/-0.75%`, `-1.51%/-1.51%`, and `-0.11%/-0.08%`,
respectively. The paired directions remain positive, and the final
distributions are substantially tighter, but the result still should not be
generalized beyond the named host and corpus.

## Explicit workflow accounting

The edit patch and same-lineage apply extensions span approximately
99.6-99.99% across the reported p50/mean/p95 deltas versus the eager-fingerprint
control for their isolated operations. The
first deferred fingerprint demand is intentionally visible: the control's
eager work is about 20-170 ns at that boundary, while the candidate's first
source-plus-target demand is about 25.7 us for tiny, 164 us for large, and
8.37-8.39 ms for payload-heavy artifacts. These are workflow observations,
not hidden work removal; callers that request the diagnostic values pay for the
scan at that point.

With the immediate fingerprint demand included in the explicit workflow, the
candidate's descriptive p50/mean/p95 positive-faster delta is +14.56% / +16.34% / +22.24% and
+13.90% / +15.80% / +21.90% for tiny, +4.49% / +4.82% / +10.24% and
+5.82% / +5.64% / +4.26% for large, and +6.55% / +6.41% / +6.26% and
+7.08% / +7.08% / +6.33% for payload-heavy in the two paired directions.

## Guardrails and resource observations

The mandatory DOC guards remain favorable or neutral: exact no-op p50 has a positive-faster delta of
`+78.84%/+79.89%` for tiny and `+71.08%/+70.40%` for large; one-edit p50
improves `+37.23%/+40.81%` and `+20.45%/+19.79%`; and DOC open is near-neutral
at `-3.52%/+0.13%` tiny and `+0.55%/-1.80%` large. The neighboring XLS
one-edit and open guards are mostly neutral or improved: one-edit p50 is
`-0.20%/+2.42%` tiny and `+2.19%/+2.05%` large, while open is
`-3.78%/+0.10%` tiny and `+2.97%/+5.10%` large. XLS exact no-op remains
directionally noisy (`-14.50%/-6.87%` tiny and `+14.87%/-26.42%` large), so
it is disclosed rather than used as a regression gate.

Representative three-sample-plus-preflight payload heaptrack totals record
exactly 50,677 process allocation calls in both control and candidate, peak
heap 128.28M in both, and profiler RSS 145.14M versus 142.81M. A 30-sample
`/usr/bin/time` boundary records maximum RSS in A1/B1/B2/A2 order as
`138160 / 138024 / 138028 / 138032 KiB`. These are whole-process,
descriptive boundaries with setup and profiler overhead; they are not
operation-only allocation, total-memory, or peak-RSS attribution.

## Reproduction

Run the focused production and harness gates:

```sh
cargo test -p litchi-doc
cargo test --locked --manifest-path tools/perf-baseline/Cargo.toml \
  doc_owner_public_phase_case_is_opt_in_and_checked -- --nocapture
```

The release comparison uses the existing selector and writer shapes. Build
clean control and candidate binaries, pin each fresh process to CPU 2, and run
the four legs with 20 warmups and 500 samples:

```sh
taskset -c 2 <control-binary> \
  --case doc_owner_public_phases \
  --writer-shape tiny,large,payload-heavy --warmup 20 --samples 500 \
  --json /tmp/doc-lazy-fingerprint-0165-a1.json
taskset -c 2 <candidate-binary> \
  --case doc_owner_public_phases \
  --writer-shape tiny,large,payload-heavy --warmup 20 --samples 500 \
  --json /tmp/doc-lazy-fingerprint-0165-b1.json
```

Repeat the candidate and control commands for B2 and A2 in that order. The
summary and manifest record the exact clean-build commands, source/binary
hashes, raw artifact names, and resource-probe invocations. Verify the final
binary identities directly with:

```sh
sha256sum <control-binary> <candidate-binary>
```

The expected hashes are the control and candidate values in the table above.
The heaptrack and `/usr/bin/time` observations are descriptive process probes
only.

## Claim boundary

This accepted correctness and attribution record is limited to the exact deterministic native-DOC
owner/public-reader lifecycle, same-lineage patch replay, and explicit
fingerprint-demand workflow on the named CPU-2 release host. It does not claim
physical I/O, cold-cache behavior, filesystem or remote-range behavior,
real-producer coverage, generic DOC performance, total-memory reduction,
operation-only allocator/RSS reduction, speedup, or completion of the CRUD matrix.
Formatting-rich, malformed/security-heavy, encrypted/signed, broad producer,
topology-changing, and other native DOC transactions remain separately
scoped. The lazy fingerprint cache preserves exact-byte authorization and the
full independent validation layers; it is not a reason to remove those guards.
