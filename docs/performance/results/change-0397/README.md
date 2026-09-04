# Change 0397 OPC owned-open evidence

Change 0397 is an accepted scoped production change. The authorized claim is
limited to the normal, non-allocator `OpcPackage::from_vec(owned)` operation on
the fixed stored OPC corpora and the matched A1/B1/B2/A2 protocol below. The
candidate removes one redundant eager ZIP validation/index pass before the
real `PhysPkgReader`; this is not a claim that only one ZIP index exists
overall.

The package retains all eight report/catalog pairs from
`/tmp/litchi-0397-VYmAwG/results`: four normal and four allocator legs, each
with 30 samples and five warmups. The compressed members are lossless zstd
frames under `raw/{normal,allocator}/`, with matching `a1-control`,
`b1-candidate`, `b2-candidate`, and `a2-control` stems. The sidecar catalogs
contain four deterministic corpora (256, 2,047, 2,048, and 16,384 ordinary
Parts, each with a 32-byte payload).

## Authorized result

Positive values below mean the candidate is faster. Pooled p50 is computed by
sorting the concatenated 60 control or candidate samples for each corpus.

| Parts | A1→B1 normal p50 | A2→B2 normal p50 | pooled normal p50 |
| ---: | ---: | ---: | ---: |
| 256 | +8.617829% | +8.204676% | +8.452941% |
| 2,047 | +8.298670% | +8.719476% | +8.356702% |
| 2,048 | +8.945417% | +8.268274% | +8.490980% |
| 16,384 | +4.648655% | +4.348226% | +4.645459% |

`summary.json` records all leg metadata, corpus/source/output oracles, report
statistics, sample-vector digests, and recomputed ABBA math. The flat
`latency-metrics.tsv` projection contains all 32 retained result rows.

Allocator elapsed time is observational only. Exact candidate-minus-control
allocation/deallocation call reductions are `-1,038 / -8,202 / -8,206 /
-65,550`, and allocated/deallocated transient-byte reductions are
`-152,024 / -1,212,620 / -1,212,888 / -9,699,800`, in the same corpus order.
Reallocations and each per-sample net-live after-before vector are exactly
unchanged: `192,122 / 1,526,677 / 1,527,162 / 12,207,482` bytes. Raw global
live-before/after baselines are not cross-run metrics. The exact allocation
vectors remain in the compressed allocator reports and are projected with
digests and unique values in `allocation-metrics.json`.

The reports retain shared metadata for nine exact/ASCII-alias/miss query
classes and 16 repetitions, but this `owned_open` selector has
`lookup_count: 0`; it does not time lookups. Owned part-count, payload,
exact-source, output, and malformed equivalent-name error oracles are retained
and verified. Same-implementation drift diagnostics are at most 1.395% p50,
1.267% mean, 1.927% p95, and 3.893% p99; these are diagnostics, and no p99
claim follows.

The scope excludes allocator latency, validation-constructor timing,
path/from-reader/session timing, public `OwnedPhysPkgReader` timing, RSS or
peak operation memory, physical I/O, cold-cache behavior, throughput, other
formats/facades, and generalization beyond this constructor and corpus.
Earlier captures made with runtime rustc 1.95 and one busy-CPU normal run were
deleted and are not evidence.

## Provenance

Production commit `f275d45660de15711edd04b8d0205eaf4e620e68` (short
`f275d4566`) and measured candidate revision
`f20d3f417edc3f3da07bf515676b8e71285ad76f` share tree
`0b3187f9f85373e4ec72042ca2fff472ae581f23` and parent control
`6e98db9ece29c1e50241cf3e84c9410ce71dd748`. Candidate source hashes are
recorded in `adjudication.json` and `summary.json`.

## Integrity and binding checks

The eight raw report/catalog pairs were validated before compression with the
named validator flags. To repeat frame and binding checks after checkout:

```sh
set -eu
check_dir=$(mktemp -d /tmp/litchi-0397-evidence-check.XXXXXX)
trap 'find "$check_dir" -depth -delete' EXIT
for profile in normal allocator; do
  for stem in a1-control b1-candidate b2-candidate a2-control; do
    report_zst="docs/performance/results/change-0397/raw/$profile/$stem.json.zst"
    catalog_zst="docs/performance/results/change-0397/raw/$profile/$stem.catalog.json.zst"
    zstd -q -t "$report_zst"
    zstd -q -t "$catalog_zst"
    zstd -q -d -c "$report_zst" > "$check_dir/report.json"
    zstd -q -d -c "$catalog_zst" > "$check_dir/catalog.json"
    python3 tools/validate_perf_corpus_binding.py \
      --report "$check_dir/report.json" --catalog "$check_dir/catalog.json"
  done
done
```

The expected validator output is eight lines of
`validated schema-2 corpus catalog: 4 corpora, 4 bindings`. Verify every
compressed/raw byte count and SHA-256 in `evidence-manifest.json`; parse the
text/JSON projections with `python3 -m json.tool`. No Cargo invocation is
required for this evidence package.
