# Change 0395 OPC case-fold index evidence

This directory retains the complete matched evidence for Change 0395, which
adds an allocation-free ASCII-folded order index to large unmanaged
`SourceBackedPackage` catalogs. The index is built only at the measured
2,048-part threshold; smaller catalogs, managed execution contexts, and
mutable `OpcPackage` paths keep their existing lookup behavior.

Two independent captures are retained:

- `probe-*` is the transparent unthresholded-index probe. It is rejected
  because the 256-part source lookup regressed (+31.305% / +33.451% p50 in
  normal A1->B1 / A2->B2; +27.417% / +26.018% in the allocator harness).
  The same probe improved 2,048- and 16,384-part lookups, and source-open
  overhead stayed within the observed +4.04% p50 and +3.74% mean maximum.
- `final-*` is the thresholded candidate accepted for a scoped claim. It
  reports normal and allocator-enabled harness legs in matched `A1/B1/B2/A2`
  order; normal legs provide the authorized latency measurement, while
  allocator legs are retained for exact allocation evidence and corroboration.

Every report has 12 results (four selectors at 256, 2,048, and 16,384 parts),
30 retained samples, five warmups, one worker, and CPU affinity `2`. The fixed
lookup vector contains nine query classes repeated 16 times (144 lookups per
sample). The corpus catalog sidecar is schema 2, with three deterministic
corpora and 12 timed bindings.

`probe-summary.json` and `final-summary.json` contain the derived latency
statistics and ABBA math. `latency-metrics.tsv` is a flat projection of all
192 retained result rows. `allocation-metrics.json` retains every allocator
sample vector for all allocator-enabled reports. `adjudication.json` records
the accepted scope, exact allocation tradeoff, rejected probe, and withheld
claims. `evidence-manifest.json` binds every compressed and textual artifact
to its SHA-256 and byte count.

`performance_claim: scoped`; `claim_authorized: true`. The authorized timing
claim is limited to the normal benchmark constructor for unmanaged source
lookup p50 on the fixed 2,048- and 16,384-part vectors. It does not claim eager `OpcPackage` latency, managed
execution-context behavior, mutable-package behavior, typical OOXML/general
workloads, RSS, physical I/O, cold-cache behavior, decompression, throughput,
or broad scaling.

## Integrity and binding checks

The report and catalog streams are unchanged before compression. Check every
compressed frame and validate every pair without retaining decompressed files:

```sh
set -e
check_dir=$(mktemp -d /tmp/litchi-0395-evidence-check.XXXXXX)
trap 'find "$check_dir" -depth -delete' EXIT
for run in probe final; do
  for leg in a1 b1 b2 a2; do
    for profile in normal allocator; do
      report="$check_dir/$run-$leg-$profile.json"
      catalog="$check_dir/$run-$leg-$profile-catalog.json"
      zstd -q -t "docs/performance/results/change-0395/$run-$leg-$profile.json.zst"
      zstd -q -t "docs/performance/results/change-0395/$run-$leg-$profile-catalog.json.zst"
      zstd -q -d -c "docs/performance/results/change-0395/$run-$leg-$profile.json.zst" > "$report"
      zstd -q -d -c "docs/performance/results/change-0395/$run-$leg-$profile-catalog.json.zst" > "$catalog"
      python3 tools/validate_perf_corpus_binding.py --report "$report" --catalog "$catalog"
    done
  done
done
```

The expected result is 16 lines of
`validated schema-2 corpus catalog: 3 corpora, 12 bindings`. Compare all
compressed and raw SHA-256 values with `evidence-manifest.json` using
`sha256sum` or an equivalent implementation.

No Cargo invocation is required to inspect this package. The capture used the
explicit Rust/Cargo 1.98.1 toolchain; the repository's pinned 1.95 toolchain
was not used for these reports.
