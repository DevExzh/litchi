# Controlled performance regression comparator

`tools/perf_compare.py` compares two reports produced by the standalone release
performance harness. It is a control-plane gate, not a benchmark and not a
substitute for the balanced ABBA, profiling, allocation, and RSS evidence used
to accept an optimization.

The checked policy is
[`perf-regression-policy-v1.json`](perf-regression-policy-v1.json). Policy
schema 1 pins the release Linux tool identity, all 36 default case names and
the required count of 198 case/corpus records, the default
corpus/writer/semantic shape selections, range settings, warmup count and
filesystem flags, a 15-sample minimum, build identity fields, and explicit
upper regression thresholds:

- p50 latency: 5%
- p95 latency: 10%
- allocation, RSS, and work counters when present: 5%

The optional counter classes cover allocation calls/bytes, RSS, instructions,
cycles, faults, source reads, sink writes, materializations, copied,
decompressed and recompressed bytes, work units, and maximum in-flight work.
Output byte counts and maximum write size are excluded because corpus identity
and semantic validation already require them to be exact correctness values;
smaller is not inherently better. A numeric sample vector is reduced to its
p50 and must independently meet the 15-sample floor. Deterministic scalar
counters are compared directly. A counter present on only one side makes the
inputs incomparable.

## Fail-closed contract

The comparator exits with status 0 only when every matched metric is within
policy. It exits 1 for a valid comparison with one or more regressions and 2
when comparison is unsafe. Unsafe input includes malformed or unsupported
schema, unexpected tool identity, dirty reports, build or configuration
identity differences, identical reference/current revisions, changed corpus
identity, duplicate/missing case-corpus keys, absent required cases, an
unexpected result count, too few samples, non-finite or negative metrics, a
reported percentile inconsistent with its samples, and asymmetric optional
metrics.

The policy also requires a SHA-256 digest of all 198 exact `(case, canonical
corpus JSON)` keys. Keys are sorted, then hashed as UTF-8 case name, a zero
byte, compact canonical corpus JSON, and a newline. This prevents a reference
and candidate from silently agreeing on the same replacement corpus.

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

Both outputs are always produced when the comparator itself can write files:
a versioned machine-readable JSON report and a short human summary. Exit 2 is
used if either output cannot be written.

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

Run the comparator regressions independently with:

```sh
python3 -m unittest tools.test_perf_compare
```

## CI scope

Pull requests and ordinary pushes run only the comparator unit tests and the
existing deterministic performance correctness smoke. They do not apply
latency thresholds to hosted-runner smoke data.

The `Performance baseline` workflow exposes an optional manual
`reference_run_id`. When supplied, the full release job first produces the
current 15-sample matrix, then a separate reference-gated hosted comparison
downloads the named prior full-run artifact and the current artifact, applies
policy v1, and uploads both summaries. A regression or any identity/input
defect fails that manual job. Leaving the input empty records the
scheduled/manual baseline without a latency gate.

GitHub-hosted runner labels and the report identity fields are compatibility
checks; they do not prove identical CPU frequency, host model, thermals, or
background load. A manual hosted comparison is therefore a conservative alert
gate, not the controlled environment needed for an acceptance claim. Teams
making release decisions should run the same command on a pinned host and
retain balanced ABBA evidence.

No reference report is committed or silently selected. At the time this
mechanism was added there was no qualifying current controlled reference, so
`expected_result_keys_sha256` is deliberately `null`. The comparator treats
that state as invalid and the hosted reference job cannot pass until a
reviewed controlled reference manifest digest is committed. This document
makes no passing-baseline or performance-regression claim.
