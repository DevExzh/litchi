# 0454: source-preserving unnamed PPTX copy

This bundle retains the progression from a missing-name refusal, through the
relationship-XML and authored-XML publication refusals, to the source-proven
publication candidate. Intermediate inventories and failed checks are historical
evidence. A receipt marked `pass` proves only its recorded command and source
epoch; it does not by itself certify the batch or the full performance goal.

The source records, binary hashes, exact commands, corpus identities, raw
samples, and output oracles are separate evidence. The baseline lifecycle binary
is the preserved final 0453 candidate, with its original build revision and
source manifest retained. The 0454 candidate receives its own source manifest;
the revision label alone is insufficient to identify either build.

The external fixture is an unmodified LibreOffice QA archive, reopened separately
as source and destination. Its [provenance](fixture-provenance.md) identifies the
immutable upstream artifact and its hashes. This does not establish an original
producer/save chain, an independent document pair, or an application roundtrip.
The external comparison is descriptive because the baseline refuses that input.

The [protocol](protocol.md) records the canonical-workload regression controls,
provider settings, CPU affinity, warmups, samples, uncertainty calculation, and
metric limitations. [Review findings](review-findings.md) and
[validation history](validation-notes.md) preserve problems found before final
acceptance. [Design](design.md) explains why source-bound XML publication is
required instead of normalizing the original document.

For a fresh execution, build the standalone tools from the intended source
revision:

```sh
RUSTUP_TOOLCHAIN=1.98.1 CARGO_BUILD_JOBS=4 CARGO_INCREMENTAL=0 \
  cargo +1.98.1 build --locked --release \
  --manifest-path tools/perf-baseline/Cargo.toml --features allocator-metrics \
  --bin litchi-perf-baseline --bin pptx_external_cross_copy \
  --bin pptx_native_copy_probe
```

This command matches the `final-candidate-build-r2` toolchain, four build
jobs, disabled incremental compilation, allocator feature, and three binary
targets. The sealed bundle is an evidence record: `build.py`, `check.py`, and
`capture.py` refuse existing receipt, log, run-directory, or report paths, so
capture must not be rerun in this checkout. A new capture needs a fresh
checkout and a fresh results destination; retain this bundle for read-only
verification and replay of its recorded attestations.

The native probe takes `source-path source-position destination-path
destination-position`. The external harness takes `fixture output-json samples
warmups source-revision bytes|range`; it verifies the fixture hash before doing
any work and retains its first measured output beside the report for the
independent Python byte oracle. Use fresh caller-owned output paths. The
historical capture drivers deliberately refuse to overwrite retained evidence.

The full non-iWork goal remains open. The next measurement hypothesis and its
limitations are recorded in [next-hypothesis.md](next-hypothesis.md).

Two narrow amendments are explicit. `fuzz-source-amendment.json` proves that
only the standalone fuzz driver changed after production measurements; its
own sanitizer, strict and format checks use the amended source manifest.
`derivation-amendment.json` retains the frozen `derive.py` and identifies
`derive-final.py` as its display-only counter-key correction. The captured
protocol identity is preserved. `custody-corrections/renderer-r1/restoration.json`
discloses reconstruction of 18 receipt protocol hashes after an incorrect
post-capture update, and retains the intermediate bytes. Raw reports and timing
inputs were unchanged. The verifier checks both amendments and that restoration.

The verifier is the release boundary for this bundle. While capture is in
flight, a plain custody audit reports `status: partial` and preserves failed
attempts; it does not certify the batch. After the candidate source epoch,
formal provider lanes, external formal and pilot lanes, inventories,
measurements, and all required release and fuzz gate receipts are present, run
the final checks from the repository root:

```sh
python3 -B docs/performance/results/change-0454/verify.py --precleanup
python3 -B docs/performance/results/change-0454/verify-mutations.py
```

The precleanup receipt must exist before owned artifacts are removed. The
cleanup inventory and retired probe-example helper are then run once, in this
order:

```sh
python3 -B docs/performance/results/change-0454/cleanup.py inventory
python3 -B docs/performance/results/change-0454/remove-probe-example.py
python3 -B docs/performance/results/change-0454/cleanup.py cleanup
python3 -B docs/performance/results/change-0454/fuzz-control.py cleanup
python3 -B docs/performance/results/change-0454/verify.py --cleanup
```

`remove-probe-example.py` records and removes only its enumerated release
example and fingerprint files. Its `retired-example-*.json` receipts are
separate from the binary and external-output cleanup receipts, so they must be
retained with the sealed bundle. The canonical patch transport is retained at
`candidate/canonical-memory.patch.txt`; its `/tmp` copy was removed and
`checks/canonical-memory-patch-removal.json` records that identity. This
ordinary custody attestation is separate from source manifests and final gate
receipts. A clone can replay the retained external report oracles read-only with
`verify.py --cleanup` after raw output removal; `--sealed` may be added only
when the final `SHA256SUMS` file is present. These modes verify retained
attestations and do not recapture, rebuild, or reparse deleted raw output.

`verifier-adapter-amendment.json` records the outer verifier's post-cleanup
path and cleanup-receipt bookkeeping corrections. Failed replay logs and two
intermediate verifier snapshots remain available. The frozen external oracle
and capture receipts are unchanged by this amendment. A separate-copy replay
can be run with `portable-verify.py`; its receipt records the actual command,
exit status, output and removal of its temporary directory.

Final post-cleanup and separate-copy replay both pass. See
`checks/postcleanup-verification.json` and `checks/portable-verification.json`.
All owned temporary build files and generated outputs are removed; the bundle
retains the failed attempts, source snapshots and cleanup identities.
