# 0896 QName expansion qualification packet

This packet qualifies the borrowed validated QName expansion candidate in
`litchi-ooxml-common` against restored production. It is a measurement
protocol, not an adoption decision. The candidate remains an overlay until
the root coordinator records a passing quality/build review and a paired
performance result.

The packet imports the retained 0894 experiment module and redirects its
packet and schema globals. The strict 0893 worksheet guard, report oracle,
allocator checks, command builder, and child validator are therefore shared
without modification. `experiment.py` is the only workload entry point and
must be run by the root coordinator.

## Frozen protocol

The plan has eight native blocks and four allocator blocks. Each native block
runs all four cases as an alternating before/after or after/before pair. Each
allocator block runs the primary XLSX case in the opposite order. This gives
72 fresh sequential children and 3,320 measured scans:

| Lane | Children | Warmup per child | Scans per child | Measured scans |
| --- | ---: | ---: | ---: | ---: |
| Native | 64 | 3 | 50 | 3,200 |
| Allocator | 8 | 3 | 15 | 120 |

The primary case is
`xlsx_producer_medium_source_full_cell_scan`; controls are the marker-free
XLSX, DOCX, and PPTX cases inherited from the retained qualification. The
filesystem is warm, children are pinned to CPU 12, and capture is fail-fast
with no retries or replacements. The primary p50 gate is a candidate/baseline
ratio at most `0.98` with the bootstrap 95% upper bound below `1.0`.
Bootstrap seed `896` and 10,000 draws operate on whole paired child ratios.
Control and spread flags remain descriptive review evidence.

## Freeze input

`freeze.py --bindings BINDINGS --output plan.json` accepts a JSON object with
either a `bindings` object or binding names at the top level. Required names
are:

`source_manifest`, `source_identity`, `candidate_freeze` (normally
`candidate-source.json`), `baseline_build`, `candidate_build`,
`baseline_quality`, `candidate_quality`, and `guard_reference`. Optional
`source_review`, `candidate_review`, `methodology_review`, and report-review
bindings are retained when supplied; this keeps a frozen plan independent of
a later review receipt that itself binds the plan hash.

Each entry has a regular non-symlink `path`, `kind`, and `role`; its optional
`bytes`, `sha256`, `mode`, and `executable` fields are checked against the
file. Additional review or report entries are retained as descriptor-bound
bindings. The candidate build receipt must expose normalized `native` and
`allocator` binaries (directly or under `binaries`) with `role`, `name`,
`features`, `build_receipt`, `depfile`, and `elf_build_id`. The baseline
native and allocator binaries are inherited byte-for-byte from the retained
0894 plan.

The candidate metadata must contain or derive:

```json
{
  "candidate": {
    "hypothesis": "mce-expand-borrowed-validated-qname-0896",
    "paths": ["crates/litchi-ooxml-common/src/mce/stream.rs"]
  },
  "evaluation_head": "3ff493b638b7aa5860d97858dae8bd82bac0f2d5"
}
```

The source/build receipts are checked for the restored baseline revision
`506aa2edae6666d9c76ae6827c2f5941a8e47896`; the candidate is always frozen
with `production_adoption: false` and `performance_claim: null`.

## Capture and offline replay

After the plan exists, the root coordinator runs:

```text
python3 -B experiment.py --plan plan.json --output-root capture-v1
python3 -B reduce.py --plan plan.json --experiment-root capture-v1 --output reduction.json
python3 -B verify.py --plan plan.json --experiment-root capture-v1 --output verification.json
```

The reducer reopens every receipt and report and delegates paired statistics
to the retained 0887 reducer through the local 0896 experiment module. The
verifier launches no workload and repeats plan, command, artifact, schedule,
guard, and semantic checks. A reduction or gate result never authorizes
production adoption by itself.
