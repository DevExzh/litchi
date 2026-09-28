# 0827 — matched ordinary-save effect of the shipped PPTX compaction proof

Completed: **216 reports / 4,488 samples**, with all twelve native and observer
rows independently replayed. On the admitted PPTX file, the paired full-lifecycle
p50 ratio is **0.975058 [0.973040, 0.982246]** and edit ratio is
**0.884765 [0.882642, 0.887019]**. Lifecycle and edit each save 298 allocation
calls and 21,688 allocated bytes, with unchanged net-live and peak-above-entry
medians. No frozen p50 regression, allocation increase, or paired RSS review
threshold is triggered. There are 30 native and 16 observer metric spread flags;
all are retained in [metrics.md](metrics.md). These results establish only the
scoped complete-save effect described in the [report](../../0827-ordinary-save-compaction-effect.md).

This fresh trial compares the exact before/after transaction and XML archives
from the adopted 0824 change in the unchanged ordinary-save runtime harness.
It makes no new adoption decision. All other production source, harness source,
locks, inputs, and settings remain identical. The committed after source is
restored before comparative capture. The earlier 0825 attempt remains aborted;
no historical timing is reused.

The protocol includes twelve selectors: three real DOCX/XLSX/PPTX files ×
lifecycle, edit, atomic path publication, and counting publication. The lifecycle
includes open + edit + default full-durability save; edit has no publication;
publication owners are prepared outside the clock. PPTX counting publication
materializes `to_bytes`, not a streaming save. Timed owner drop boundaries are
recorded in `source-review.md`. Independent phase medians are not additive.

Before capture, each leg exports six corpora under five policies and passes
independent XML/OPC plus ZIP preservation admission. Historical admitted bytes
provide byte oracles only. Both legs then pass twelve one-sample qualification
reports and separate qualification-admission gates. Comparative native timing
uses six counterbalanced blocks × thirty samples with three warmups. Observer
capture uses two blocks × three samples with no warmup. All child processes
run serially on CPU 12; allocator/process observations are diagnostic and kept
separate from native latency. Planned totals are 216 reports / 4,488 samples;
completed lane receipts confirm those collected counts.

The 0826 allocation-vector helper is frozen as an execution input. Preflight
exercises report schemas against historical fixtures before the new protocol is
frozen. Captures additionally check source/corpus/publication identities and
require zero failed allocations under this trial's policy. The helper alone is
not artifact admission. Native report vectors must remain explicitly unavailable.
Prepared reader corrections resolve accepted admission paths from descriptors
and compare ZIP preservation as canonical JSON.

Raw receipt order labels use B=before and A=after: `BA` means [before, after],
`AB` means [after, before]. Full-word plan arrays are authoritative.

Root owns all Cargo, exporter, qualification, timing, observer, reader, and Git
execution. Static agents may prepare/review assigned files. Before correction,
`run_reader.py` preserves complete reader source snapshots, command, exit state,
and log for every preflight or numerical-reader attempt. Frozen drivers may not
be edited after freeze. The target and scratch roots are owned only by 0827.

Quality uses six fresh harness gates and explicitly verifies exact-source reuse
of the two sealed 0824 PPTX quality receipts. Fresh harness tests total 641 passed /
zero failed / one ignored. Reused PPTX receipts have 1,241 before and 1,253 after
tests passed, with three ignored per leg; these are not fresh 0827 checks.
No iWork work is included.

Execution order (all scripts below are relative to this packet):

```text
prepare.py
quality.py
run_reader.py driver_preflight.py --write
run_reader.py admission_preflight.py --write
run_reader.py reader_preflight.py --write
# Bind the three passing attempt receipts in preflight-suite.json.
freeze.py
install_source.py before
build.py before
capture.py artifacts before
admission.py artifacts before
capture.py qualification before
admission.py qualification before
install_source.py after
build.py after
capture.py artifacts after
admission.py artifacts after
capture.py qualification after
admission.py qualification after
capture.py native
capture.py observer
run_reader.py analysis.py --write
run_reader.py raw_audit.py --write
run_reader.py validate.py
cleanup.py
run_reader.py validate.py --final
```

Use `python3 -B` for each command. Completed outputs are immutable. The sequence
above documents execution and must not be rerun into this completed packet.

Fourteen `reader-attempts` directories retain complete source snapshots and
logs, including seven failures. Two prefreeze failures concerned publication
hash cardinality and the native unavailable-allocation envelope. Four later
primary-reader failures concerned provenance normalization, transition subset
comparison, qualification instrumentation, and case keys. One raw-reader
failure concerned the nested `rows` container. Final readers pass with those
integration repairs; frozen drivers and all raw evidence remain unchanged.
No Cargo or workload command failed.

Cleanup verified all six binaries and removed only the owned target and
filesystem scratch roots: 7,015 files / 8,056,907,786 logical bytes. Final
validation confirms the committed production source and unrelated files remain
unchanged. The broader performance goal remains open.

Offline replay (no build, profiler, or workload execution):

```sh
python3 -B docs/performance/results/change-0827/analysis.py --check
python3 -B docs/performance/results/change-0827/raw_audit.py --check
python3 -B docs/performance/results/change-0827/validate.py --final
python3 -B docs/performance/results/change-0827/seal.py --check-head
```

`analysis.json` retains every native/observer process, quantile, confidence
interval, allocation vector, process diagnostic, and spread flag. `native.csv`
and `analysis.md` provide compact views. `raw-audit.json` records independent
arithmetic replay; `seal.json` binds the packet, report, and five indexes to
the exact staged/committed file set.
