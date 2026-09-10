# 0500 evidence

This batch investigates managed DOCX multi-paragraph template filling through
`Edit::replace_body_paragraph_texts`. The base revision is `c797d04c7`.
The before executable's repeated scalar route is the historical control;
the after executable measures both that route and the newly supported batch
route. Batch versus repeated is an explicit API-choice comparison with equal
semantic and exact-publication oracles, not a historical same-method claim.

The deterministic harness records full open/edit/commit/sequential-publication
latency and separate edit latency. Fixture/provider setup and independent
output/preservation/inverse verification are outside those clocks. Files are
warm; no controlled-cold or native-producer claim follows from this corpus.
Charged Work may differ because the batch removes reconstruction passes;
monotonic consumption and zero released memory/object reservations remain
required. Caller-visible source counters are not physical disk-I/O counts.

`plan.md` records the hypothesis, scope, and preservation constraints.
`before-source.json` records production hashes before edits. The same benchmark source is frozen in both phases. Raw capture receipts retain
commands, executable hashes, output hashes, timings, and cleanup results.
Whole-child RSS and supplementary perf counters include setup and verification
and do not supply operation-local allocation or CPU attribution.

Checks use `run-gate.py` and `gate-commands.json`, Rust 1.98.1, offline locked
release builds with debug information disabled, incremental compilation off,
two build jobs and the dedicated cache target. CPUs 16–31 are for gates;
measurements use CPUs 0–7 on the same shared host. Affinity does not isolate
memory, storage, or other host activity.

All thirty ADR hashes match 0499. Sixteen unrelated files remain protected by
`protected-work.json`. The full non-iWork goal remains open; this work does
not close other semantic domains, native producers, broad history, or the
program-wide tenfold objective.

The dedicated build target was relocated to `/tmp/litchi-goal-0500-target`
and linked from its original cache path when root filesystem capacity became
constrained. `target-relocation.json` records this build-storage change;
fixture and retained executable paths stayed unchanged. Final cleanup accounts
for both owned paths and deduplicates allocated blocks by device and inode.

Completed captures contain 12 before and 24 after children (2,160 measured
samples and 216 warmups). `compare.py` reproduces `comparison.json` and
`comparison.md`; `verify-evidence.py` independently checks row identities,
oracles, metrics, source and binary hashes, profiles, and gate receipts.
`summarize-profiles.py` and `inspect-adverse.py` retain supplementary counter
and adverse-phase diagnostics. `production-review.md` records independent
source review and bounded metadata-accounting limitations.

The managed and unmanaged benchmark preflights verify XML patch forward and
inverse behavior outside the timer. The focused integration suite additionally
checks complete-artifact inverse publication; that is not a per-sample
benchmark assertion. Untouched-member checks compare media and opaque member
payloads, while the full output is also byte-identical to its reference route.

The downstream facade gate is a DOCX-only library check. Initial all-target
attempts exposed existing example feature declarations: `comprehensive_docx_test`
imports `ooxml_common` without enabling it; adding that feature exposes
`core_props_office`, which also requires PPTX and XLSX. Those failed commands
and full logs are retained as downstream attempts. No example manifest was
changed, and no downstream all-target success is claimed.
