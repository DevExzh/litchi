# change-0577 evidence packet: what the OOXML open's relationship reads are for

Change record:
[`docs/performance/0577-ooxml-open-relationship-parts.md`](../../0577-ooxml-open-relationship-parts.md).
Disposition: **design only**, no production change. `performance_claim: none`.

## Contents

| Path | What it is |
| --- | --- |
| `make_contract_fixtures.py` | Generates seven minimal OPC packages, each isolating one open-time decision the relationship-part reads are responsible for. All members stored, byte-deterministic. |
| `contract-probe.rs` | The throwaway probe binary: opens each generated package through **both** ingress modes (`SourceBackedPackage::from_read_at` and `OpcPackage::from_bytes`) behind a counting `litchi_core::ReadAt`, and reports each open's verdict verbatim. Nothing under `crates/` was touched. |
| `rels_decomposition.py` | Reproduces `walk_relationship_graph` and the post-classification orphan fallback from package bytes alone, and reports which mechanism reads each relationship part. Validated against change 0572's eleven measured structural-member counts. |
| `structural_runs.py` | Parses each fixture's local records and computes the contiguous runs the structural members form, today's per-member read cost, and the modelled cost of one bounded read per run. |
| `results.json` | The retained machine-readable result: environment, revision, binary hashes, corpus and contract-fixture SHA-256s, the seven contract verdicts, the measured open-phase counts at HEAD, the delayed-transport medians, the window-and-model comparison, the two-capture repeatability check, the stored-member read cost, the 167-fixture decomposition and the 167-fixture run-coalescing model. |

## Headline

```
open phase, exact policy, zero-delay transport, HEAD 32d25e088
  ConditionalFormattingSamples.xlsx   89 of 222 requests   23,261 bytes   40.1% of the scenario
  shape-soft-edges.pptx               84 of  87 requests   12,111 bytes   96.6% of the scenario
  shapes.pptx                         47 of  49 requests    9,755 bytes   95.9% of the scenario

the open is UNCHANGED from change 0572's capture on all eleven fixtures;
change 0573 halved the strict-layout proof around it, so its SHARE rose.

relationship-read decomposition, 167 OOXML fixtures
  relationship parts present     1,237
  read by the graph walk         1,070   <- gates part admission; mandatory
  read by the orphan fallback        0   <- could be deferred; never fires
  never read at open                 0
  (1,237 - 1,070 = 167 = one `_rels/.rels` per fixture, read before the walk)

contract experiments, 7 packages, both ingress modes agreeing error-for-error
  a malformed relationship part the caller never names FAILS THE OPEN,
  including one belonging to an orphan part reached only by the fallback loop;
  and the SAME untyped member opens as tolerated junk or refuses with
  ContentTypeNotFound depending only on whether a deep relationship names it.

modelled run-coalesced prefetch (NOT IMPLEMENTED, NOT MEASURED)
  ConditionalFormattingSamples.xlsx   89 -> 10 requests, 23,261 -> 26,796 bytes
  all 167 fixtures                 3,665 -> 969 open requests (73.6% fewer),
                                   structural bytes 481,943 -> 726,329 (1.507x),
                                   0 fixtures costing more requests
```

## What is a measurement here and what is a model

Four things are **measured**:

- the HEAD re-capture of change 0572's arm matrix, through 0572's own retained
  probe and classifier (eleven fixtures, 132 arms, five repeats), taken **twice**
  at host load averages 7.38 and 0.62 with **zero divergent arms**;
- the seven contract-experiment verdicts, byte-identical across both runs;
- the corpus and contract-fixture hashes;
- the delayed-transport medians on change 0572's 1 ms-per-request model transport,
  retained in `results.json` under `delayed_transport_medians`. These are
  sleep-driven arithmetic over a deterministic request count — a model transport,
  not a device.

Two things are **models** computed from package bytes with no help from the
library:

- the relationship-read decomposition (`rels_decomposition.py`). It is validated
  in one direction — adding content types, the package `_rels/.rels`, the walk's
  reads and the format's main part reproduces all eleven of change 0572's
  measured structural-member counts exactly. It does not reproduce the open's
  limit checks, so a package that would refuse at open is not modelled correctly.
  Its orphan-fallback branch is not dead code: run against the contract fixtures
  it reports 1 walk read and 1 fallback read for `e3-orphan-rels-malformed`,
  which is the package built to reach a relationship part that way.
- the run-coalescing prediction (`structural_runs.py`). Its *today* columns
  reproduce the measured open request count **and** the measured structural read
  bytes exactly on all eleven measured fixtures, with no free parameter; its
  *coalesced* column is unvalidated, because no implementation exists, and the
  167-fixture totals extrapolate it to fixtures whose open was never captured.

No latency *improvement* is claimed anywhere in this packet: the design that
would produce one is not implemented, so there is no after-figure to compare.

## Replay

The build must be pinned. Change 0572 had to discard a capture because its probe
linked another agent's uncommitted change through path dependencies; the same
applies here.

```sh
rev=32d25e08806d93f792ffd4954d83acc9db9c5301
work=$(mktemp -d)
git -C . archive "$rev" crates Cargo.toml rust-toolchain.toml | tar -x -C "$work"

# The probe crate is change 0572's, reused unchanged so the figures compare.
cp -r docs/performance/results/change-0572/probe "$work/probe"
sed -i "s|\.\./\.\./\.\./\.\./\.\./crates|$work/crates|g" "$work/probe/Cargo.toml"
mkdir -p "$work/probe/src/bin"
cp docs/performance/results/change-0577/contract-probe.rs "$work/probe/src/bin/contract.rs"

CARGO_TARGET_DIR="$work/target" cargo +1.95.0 build --release -j 8 \
  --manifest-path "$work/probe/Cargo.toml"

# 1. contract experiments
python3 docs/performance/results/change-0577/make_contract_fixtures.py "$work/gen"
"$work/target/release/contract" "$work/gen"

# 2. open-phase request counts at HEAD, through change 0572's own instruments
"$work/target/release/litchi-0572-probe" . "$work/capture.json" "$rev"
python3 docs/performance/results/change-0572/classify_zip_requests.py \
  --capture "$work/capture.json" --out "$work/attribution.json" --repo . \
  --raw-sequence-for '*:exact'
python3 docs/performance/results/change-0572/summarize.py \
  --attribution "$work/attribution.json" --capture "$work/capture.json" \
  --json-out "$work/summary.json"

# 3. the two models
python3 docs/performance/results/change-0577/rels_decomposition.py \
  $(find test-data/ooxml -maxdepth 2 -type f \
      \( -name '*.docx' -o -name '*.xlsx' -o -name '*.pptx' \) | sort)
python3 docs/performance/results/change-0577/structural_runs.py \
  test-data/ooxml/xlsx/ConditionalFormattingSamples.xlsx
```

Check the retained verdicts and hashes without rebuilding anything:

```sh
python3 -B -c "
import json
d = json.load(open('docs/performance/results/change-0577/results.json'))
for e in d['contract_experiments']:
    print(f\"{e['fixture']:32s} {e['source_backed_open'][:44]}\")
print()
print({k: v for k, v in d['relationship_read_decomposition'].items()
       if k not in ('per_fixture', 'note')})
for f in d['corpus']:
    print(f['sha256'][:16], f['size'], f['path'])"
```

## What is not here

The intermediate `capture.json`, `attribution.json` and `summary.json` are **not**
retained: they are regenerable by the replay above, and every figure the change
record cites is in `results.json`. Change 0572's `results.json` already retains
the full ordered request sequences for the `exact` policy at the pre-0573
revision, and this packet's open-phase counts are identical to them, so nothing
is lost.

No filesystem, cold-cache, physical-device or cross-platform capture. No
allocation, peak-RSS, instruction-count or syscall measurement. No ABBA
comparison, because the design this record freezes is not implemented and there
is nothing to compare it against. The counting source holds each fixture in
memory, so nothing here measures the page cache.
