# Controlled performance regression comparator

`tools/perf_compare.py` compares two reports produced by the standalone release
performance harness. It is a control-plane gate, not a benchmark and not a
substitute for the balanced ABBA, profiling, allocation, and RSS evidence used
to accept an optimization.

The checked policy is
[`perf-regression-policy-v1.json`](perf-regression-policy-v1.json). Policy
schema 2 pins the release Linux tool identity, all 36 default case names and
the required count of 198 case/corpus records, the default
corpus/writer/semantic shape selections, range settings, warmup count and
filesystem flags, a 15-sample minimum, build identity fields, and explicit
upper regression thresholds:

- p50 latency: 5%
- p95 latency: 10%
- p99 latency: 15%
- allocation, RSS, and work counters when present: 5%

Allocator-only evidence uses the separate
[`perf-regression-policy-allocator-v1.json`](perf-regression-policy-allocator-v1.json).
That policy requires the `litchi-perf-baseline-alloc` binary identity and
`system_allocator_operation_scoped` instrumentation identity. The comparator
validates and compares allocation vectors under that policy while withholding
all elapsed-latency comparisons; the normal policy's `binary` and
`instrumentation` identity still rejects allocator reports.

The schema-2 metric-class `presence` field distinguishes required and optional
counters and is mandatory for every class. Required classes must match at least
one metric in every result, and every required path must be present and valid in
both reports. Missing or malformed required vectors fail closed as invalid
input. Optional classes may be absent from both reports; if an optional path is
present, it must be present and valid on both sides. The checked policy's
counter classes cover allocation calls/bytes, RSS, instructions, cycles, faults,
source reads,
sink writes, materializations, copied, decompressed and recompressed bytes,
work units, and maximum in-flight work. Output byte counts and maximum write
size are excluded because corpus identity and semantic validation already
require them to be exact correctness values; smaller is not inherently better.
A numeric sample vector is reduced to its p50 and must independently meet the
15-sample floor. Deterministic scalar counters are compared directly.

Policy schema 1 is intentionally rejected; the existing policy filename is
retained for the workflow path, but its document now carries schema 2. Any
other policy must likewise migrate to schema 2 and declare each counter class's
presence explicitly before comparison.

Latency p50, p95, and p99 are always required. Each reported percentile must be
finite, agree with the nearest-rank (median for p50) value computed from its
finite positive sample vector, and remain non-decreasing. The three latency
percentiles are compared independently against their policy thresholds.

## Fail-closed contract

The comparator exits with status 0 only when every matched metric is within
policy. It exits 1 for a valid comparison with one or more regressions and 2
when comparison is unsafe. Unsafe input includes malformed or unsupported
schema, unexpected tool identity, dirty reports, build or configuration
identity differences, identical reference/current revisions, changed corpus
identity, duplicate/missing case-corpus keys, absent required cases, an
unexpected result count, too few samples, non-finite, overflowing, or negative
metrics, a missing or reported percentile inconsistent with its samples,
missing or asymmetric required metrics, and asymmetric optional metrics.

The policy also requires a SHA-256 digest of all 198 exact `(case, canonical
corpus JSON)` keys. Keys are sorted, then hashed as UTF-8 case name, a zero
byte, compact canonical corpus JSON, and a newline. This prevents a reference
and candidate from silently agreeing on the same replacement corpus. The
checked default manifest digest is
`3b57c3b5aef77f5149d520fd885194d1fd8734460b28bff9d317d1cd840c246f`.

The digest was derived from the current harness's default `Case::DEFAULT`
selection and a fresh deterministic one-sample, zero-warmup report (the
sample count is irrelevant to this identity-only calculation). The emitted
198 keys decompose into:

| harness branch | cases | corpus selection per case | records |
| --- | ---: | --- | ---: |
| ZIP/OPC and CFB/OLE2 substrate | 18 | 4 archive shapes × 2 payload kinds | 144 |
| fresh DOC/XLS/PPT writers | 3 | 3 writer shapes | 9 |
| XLSX semantic matrix | 15 | 3 XLSX shapes | 45 |
| total | 36 | — | 198 |

The archive shapes are `tiny`, `many-small`, `few-large`, and `wide-root`; the
payload kinds are `compressible` and `incompressible`; writer shapes are
`tiny`, `large`, and `payload-heavy`; and XLSX shapes are `tiny`, `medium`,
and `dense-wide`. The exact identity-only manifest is recorded in
[`perf-regression-default-manifest-v1.json`](results/perf-regression-default-manifest-v1.json);
it contains no latency samples, resource counters, or output measurements.

Git revisions are recorded in the result but deliberately differ between the
reference and candidate. The compared build identity consists of Rust version,
visible logical CPU count, allocator, Rust flags, Cargo target, Linux perf
event policy, OS, kernel, CPU model, total memory, and page size. A field may be
JSON `null` only for Rust flags and the Cargo target, and only when both reports
record it that way; presence and exact value must still match. All other pinned
identity values are typed and non-empty/positive. The complete benchmark
configuration and exact corpus objects must also match.

The default `execution_workers` list is derived from the recorded positive
logical CPU count using the harness's `1,2,4,8,available` selection and must
match exactly on both sides.

The machine-readable JSON report is always required. A short human summary is
also written when `--summary-out` is supplied, and is always printed to stdout.
Exit 2 is used if the required JSON output or a requested summary output cannot
be written.

## Additive corpus catalog

The harness can emit a schema-2 corpus catalog sidecar with
`--corpus-manifest PATH`.  The report keeps schema 1 and receives only an
optional `corpus_catalog` reference, so the existing 198-key identity digest
and all older reports remain compatible.  The catalog's
`content_set_sha256` identifies the exact corpus/member set; its
`catalog_sha256` additionally covers generator, producer, security, malformed
input, limits, and provenance metadata.  See
[`CORPUS_MANIFEST_V2.md`](CORPUS_MANIFEST_V2.md) for the migration and
truthfulness rules.

## Exact invocation

Capture reference and candidate reports in the same controlled environment
with the checked release harness and default matrix. This comparison is
enabled only after the checked policy contains the reviewed reference's exact
case/corpus manifest digest:

```sh
cargo run --release --locked --manifest-path tools/perf-baseline/Cargo.toml -- \
  --samples 15 --json target/perf/reference/container-baseline.json
cargo run --release --locked --manifest-path tools/perf-baseline/Cargo.toml -- \
  --samples 15 --json target/perf/current/container-baseline.json
python3 - <<'PY'
import json
from pathlib import Path
from tools.perf_compare import report_result_key_manifest_sha256

report = json.loads(
    Path("target/perf/reference/container-baseline.json").read_text()
)
print(report_result_key_manifest_sha256(report, 198))
PY
python3 tools/perf_compare.py \
  --policy docs/performance/perf-regression-policy-v1.json \
  --baseline target/perf/reference/container-baseline.json \
  --current target/perf/current/container-baseline.json \
  --json-out target/perf/comparison/perf-regression.json \
  --summary-out target/perf/comparison/perf-regression.txt
```

Run the performance-tooling unit suites independently with:

```sh
python3 -m unittest \
  tools.test_perf_compare \
  tools.test_perf_abba_summary \
  tools.test_perf_abba_package \
  tools.test_perf_resource_profile
```

## CI scope

Pull requests and ordinary pushes run the performance-tooling unit suites and
the existing deterministic performance correctness smoke. They do not apply
latency thresholds to hosted-runner smoke data.

The `Performance baseline` workflow exposes an optional manual
`reference_run_id`. When supplied, the full release job first produces the
current 15-sample matrix, then a separate reference-gated hosted comparison
downloads the named prior full-run artifact and the current artifact, applies
the checked schema-2 policy, and uploads both summaries. A regression or any
identity/input defect fails that manual job. Leaving the input empty records
the scheduled/manual baseline without a latency gate.

GitHub-hosted runner labels and the report identity fields are compatibility
checks; they do not prove identical CPU frequency, host model, thermals, or
background load. A manual hosted comparison is therefore a conservative alert
gate, not the controlled environment needed for an acceptance claim. Teams
making release decisions should run the same command on a pinned host and
retain balanced ABBA evidence.

No reference report is committed or silently selected. The checked digest only
authorizes the case/corpus key identity; it is not a latency baseline and does
not authorize any performance claim. A qualifying controlled reference report
must still be supplied separately, must satisfy the minimum sample count and
build/configuration identity checks, and must have a distinct revision from
the candidate. The hosted job therefore remains a conservative comparison
gate, not evidence that a baseline has passed or that any optimization has
improved performance.
