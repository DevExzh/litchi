# Change 0396 OPC case-fold evidence

Change 0396 production experiments are **REJECTED**. The accepted deliverable is
harness observability and auditable negative evidence; this directory does not
authorize a production latency or allocation claim.

The compressed raw members under `raw/` retain 48 report/catalog pairs (96
zstd members) from `/tmp/litchi-0396-abba-XkZQH6/results`. Each report is paired
with the catalog emitted for that capture and was validated with the named
flags below before compression.

Evidence groups:

- `mapless-binary-search/` retains the `normal/` and `allocator/` full ABBA
  runs for the mapless/binary-search alternative.
- `linear-probe-full/` retains its full normal and allocator ABBA runs;
  `linear-probe-probe/` retains the transparent preliminary probe.
- `scalar-vec-probe/` retains the transparent scalar `Vec` probe from the
  `pre-inline-probe` root.
- `std-prehashed-map-probe/` retains the standard-library prehashed-map probe.
- `hashbrown-hashtable-probe/` retains the direct `hashbrown::HashTable`
  probe, while `hashbrown-hashtable-exact-1000/` retains the high-sample exact
  ABBA (1,000 samples and 20 warmups).
- `shared-packuri-probe/` retains the preliminary shared-PackURI `Arc` probe;
  `shared-packuri-arc-final/` retains the final full 30-sample normal and
  allocator ABBA runs (A1/A2 control, B1/B2 candidate).

The catalogs cover the four deterministic corpus sizes 256, 2,047, 2,048,
and 16,384 parts. Full ABBA reports have 30 samples and five warmups;
preliminary probes have five samples and one warmup; the exact-only ABBA has
1,000 samples and 20 warmups. `rejected-summary.json` contains the key
source-operation projections and ABBA/probe math for every rejected group;
`final-summary.json` contains the corresponding final Arc ABBA projection.
`latency-metrics.tsv` is the flat projection of all 1,088 retained result
rows. `allocation-metrics.json` contains compact allocator summaries for 440
source-operation records; the exact per-sample vectors remain in the
compressed allocator reports.

The source-open timing case uses the normal unmanaged
`SourceBackedPackage::from_read_at` constructor. The lookup timing cases use a
fixed pre-open unmanaged package and time the lookup operation only.
`SourceBackedPackage::from_read_at_for_validation` is correctness-tested only
and was not benchmarked. Allocator-profile timing is observational and is
retained for allocation vectors/corroboration only.
The adjudication and all withheld claims are machine-readable in
`adjudication.json`.

For the final Arc gate, the pooled normal p50 deltas are +6.09%/+6.96% for
exact lookup at 2,048/16,384 parts, +3.38%/+4.30% for source-open, and
-0.59%/-0.50% for the mixed operation. The allocator source-open vectors add
three allocation calls, +2,051/+16,387 deallocation calls, and net-live
deltas of -65,536/-524,288 bytes at those sizes. Full per-leg values remain in
`final-summary.json` and the compressed reports.

## Integrity and binding checks

Validate the compressed frames, decompress each pair into a temporary directory,
and run the repository validator with named flags:

```sh
set -eu
check_dir=$(mktemp -d /tmp/litchi-0396-evidence-check.XXXXXX)
trap 'find "$check_dir" -depth -delete' EXIT
for report_zst in docs/performance/results/change-0396/raw/*/*.json.zst; do
  case "$report_zst" in
    *.catalog.json.zst) continue ;;
  esac
  catalog_zst="${report_zst%.json.zst}.catalog.json.zst"
  test -f "$catalog_zst"
  zstd -q -t "$report_zst"
  zstd -q -t "$catalog_zst"
  zstd -q -d -c "$report_zst" > "$check_dir/report.json"
  zstd -q -d -c "$catalog_zst" > "$check_dir/catalog.json"
  python3 tools/validate_perf_corpus_binding.py \
    --report "$check_dir/report.json" \
    --catalog "$check_dir/catalog.json"
done
```

The expected result is 48 validator lines: 24 lines for 28-binding reports,
20 lines for 20-binding probes, and four lines for the 4-binding exact-only
ABBA. Verify every manifest byte count and SHA-256 (including decompressed raw
hashes) with `evidence-manifest.json`. Parse all JSON files with
`python3 -m json.tool`; no Cargo invocation is required for this evidence
package.
