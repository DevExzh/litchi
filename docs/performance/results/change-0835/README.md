# 0835 filesystem route/cache baseline evidence

All 72 formal commands / 2,160 measured samples pass strict validation and the
independent numerical/custody audit. Seven separate qualification reports
contain sixteen samples. All nine reader mutation checks pass, and no planned
20% block-spread flag is triggered. Production is unchanged.

Cleanup removed 1,784 files / 1,318,938,046 logical bytes from the two owned
roots. Qualification, full capture, mutation checks, analysis, and audit all
replay successfully after cleanup. The one failed offline cardinality check
remains retained; all eighty recorded commands (one build and 79 workloads)
succeeded without retry.

This packet measures the committed harness at the revision in `origin.json`.
It changes no Rust source. The six selectors cover OPC eager/source open,
OPC eager/source one-Part atomic save, and PPTX eager/source selected-slide
lifecycle. iWork and the three unrelated worktree files are excluded.

## Protocol and provenance

`measurement-plan.json` fixes six alternating forward/reverse blocks of twelve
case/cache pairs. Each command uses CPU 12, thirty measured fresh child
processes and three untimed warmups. Each measured operation has a separate
priming child. Warm and verified-cold observations remain separate. The full
matrix is 72 reports / 2,160 measured samples; qualification is seven additional
reports / sixteen samples and is never pooled into those distributions.

`origin.json`, `host.json`, `freeze-baseline.json`, `build-baseline.json`, and
`commands/` retain source, toolchain, executable, command, log and timing
identity. `quality-reuse.json` binds the byte-identical source to the committed
0834 quality record: the full library suite passed 569 tests with one ignored
before two helper amendments; the final helper passed twelve focused tests and
format/check/Clippy/rustdoc/boundary gates. This packet does not claim another
full-suite run. The fresh release build and all seven qualification commands
must pass before capture admission.

The strict reader checks corpus/output identities, statistics and chronological
sample mapping, process isolation, raw logical-read evidence, PPTX phase
classification, and OPC EOCD-comment proofs. The independent audit checks
command/source custody and recomputes descriptive statistics from raw reports.
Offline validation commands and any failed attempts are retained in
`validation/`; native commands live in `commands/`.

## Interpretation

This is a current-route/cache baseline on one host and fixed synthetic corpora,
not a before/after production optimization. OPC eager open includes package
destruction in its timer; source open retains the package for later diagnostics.
PPTX source logical reads are from an untimed replay. Verified cold describes
observed page-cache residency and process `read_bytes`, not physical-device
reads. Allocator instrumentation is disabled. Thirty samples per block give a
nearest-rank p99 equal to the block maximum. Six-block spread and all route
ratios remain visible; no historical timing samples are pooled.

Only the two exclusively created, marker-bound roots may be removed by
`cleanup.py`. Their executable descriptors survive for offline replay.
`seal.py` binds the exact owned paths and can verify their committed Git blobs.

## Offline replay in the recorded checkout

The qualification reader's first execution rejected the correct 16-child
manifest because its count check expected 14. `reader-versions/v1.py`,
`qualification-validation.json`, and the failed validation receipt retain that
attempt. `reader-correction.json` binds the successful correction. The admitted
qualification result is **`qualification-validation-v2.json`**. The native
qualification workloads were not repeated.

From this packet directory, after capture and cleanup:

```sh
python3 -B reader.py --qualification qualification.json --write qualification-validation-v2.json --check
python3 -B reader.py --capture capture.json --check
python3 -B reader-tests.py --check
python3 -B analyze.py --check
python3 -B audit.py --check
python3 -B seal.py verify --committed
```

The command drivers enforce the recorded source inventory and checkout state;
they are not restart commands for a later revision or for deleted temporary
roots. A future measurement needs a new packet and new marker-owned roots.
