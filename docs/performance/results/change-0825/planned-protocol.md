> Unexecuted protocol retained for provenance. Qualification failed; see README.md and abort_validate.py for the actual outcome and replay. The counts and commands below describe the original plan.

# 0825 — complete ordinary-save effect of the PPTX compaction proof

This packet measures the already-adopted 0824 optimization in the existing
ordinary-save harness. The prior edit-only trial does not establish a benefit
for complete save. Fresh matched builds reconstruct the pre-optimization PPTX
transaction/XML owners and compare them with current HEAD. All other production
and harness files, dependency locks, corpus inputs, and execution policy stay
identical. The current implementation is restored after baseline qualification;
this batch proposes no further production change.

Twelve workflows cover the three admitted real DOCX, XLSX, and PPTX inputs,
each in lifecycle, edit, atomic path publication, and counting publication.
Lifecycle includes the default full-durability path save. Edit performs no
publication. Atomic publication prepares the edit outside the timed region.
PPTX counting publication materializes `to_bytes`; it is not sequential
streaming evidence. Phase measurements are independent and cannot be added or
subtracted to attribute costs. DOCX/XLSX serve as unaffected-format controls.

The unchanged exporter produces six corpus cases and five policy outputs per
leg. Fresh independent admission checks exact inputs, allowed edit closures,
XML/relationship semantics, output identities, and untouched ZIP metadata,
compressed payloads, member order, and archive comments. Historical admitted
artifacts provide byte oracles only; no historical timing samples are pooled.
Current artifacts are not externally resaved in Office.

Before measurement, root runs six fresh harness quality commands and checks
exact-source reuse of the sealed 0824 complete PPTX quality results for both
legs. The full source census, 35 normative inputs, locks, toolchain, host,
candidate source archives, unrelated files, and drivers are frozen. Root owns
all execution and Git mutations; agents prepare and review static files only.

Native and observer binaries run separately in fresh serial processes on CPU
12. Both legs must pass qualification before comparative capture. Native uses
six counterbalanced blocks with thirty samples and three warmups; observer
uses two blocks with three samples and no warmup. Qualification contributes
twenty-four reports/samples, native 144 reports/4,320 samples, and observer
48 reports/144 samples: 216 admitted reports and 4,488 samples in total.

Offline readers retain all twelve rows, quantiles, means, process RSS,
allocation metrics, process-counter diagnostics, paired bootstrap intervals,
and spread flags. The frozen bootstrap uses 10,000 resamples, seed 825825,
and sorted zero-based endpoints 250/9749. No aggregate hides regressions.
The full-word leg arrays in `plan.json` define execution order. Receipt labels
use `BA` for `before, after` and `AB` for `after, before`; they abbreviate the
leg names rather than assigning conventional A/B experiment labels.
This is attribution of a committed change, not a new candidate-adoption gate.
Results cannot establish universal, cross-format, cold-cache, tail, RSS,
producer-wide, or historical gains.

The owned target and filesystem scratch roots are
`/home/zhuhe/code/litchi-target-0825` and
`/home/zhuhe/code/litchi-fs-0825`. Cleanup must verify retained executable
identities and ownership before removing either root. Original failed commands,
if any, retain their logs and source witnesses before correction. Output drivers
refuse overwrites; completed packets are replayed offline, not recaptured in
place.

## Execution and replay

The root-owned sequence, in a fresh packet at the pinned base, is:

```text
python3 -B prepare.py
python3 -B quality.py
python3 -B freeze.py
python3 -B install_source.py before
python3 -B build.py before
python3 -B capture.py artifacts before
python3 -B admission.py artifacts before
python3 -B capture.py qualification before
python3 -B admission.py qualification before
python3 -B install_source.py after
python3 -B build.py after
python3 -B capture.py artifacts after
python3 -B admission.py artifacts after
python3 -B capture.py qualification after
python3 -B admission.py qualification after
python3 -B capture.py native
python3 -B capture.py observer
python3 -B analysis.py --write
python3 -B raw_audit.py --write
python3 -B validate.py
python3 -B cleanup.py
python3 -B validate.py --final
```

Script paths above are relative to this packet. Each executed command retains
its console log; `validation-precleanup.log` is required before cleanup. The
preparation/quality phase also rechecks the preceding committed 0824 seal.
The write commands refuse completed outputs. Once sealed, use `analysis.py
--check`, `raw_audit.py --check`, `validate.py --final`, and `seal.py --check-head`
for offline replay. The final seal binds the packet, main report, and five
performance indexes to the exact commit; production has no net change.
