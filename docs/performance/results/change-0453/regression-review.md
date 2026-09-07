# Regression and retention review

Retain shared PPTX decoded payloads. Both formal media allocation repeats remove
16,777,408 bytes of planning allocation (32.413%) and retained live growth
(33.271%). Plan allocation calls fall 3,404 to 3,388. Plain allocation counters
are unchanged. No positive normal median/phase/RSS trigger exceeds 5%.

The two primary plain/bytes p99 increases (10.883%/10.174%) remain visible. Each
candidate process contains one roughly 0.33 ms spike, first in planning and then
in publication. Their cause is not established. A separately frozen two-ABBA
investigation of 240 samples does not reproduce the trigger: block p99 changes
are +1.508%/+0.001%; pooled descriptive p50/p95/p99 changes are
+0.132%/+0.122%/+1.082%. Thirty-sample process tails and the separate combined
distributions do not establish a stable population p99. The primary matrix is
not replaced. See confirmation-summary.md and its full source/binary custody.

Normal bytes/media API p50 improves 2.995%/3.217%; simulated range/media changes
are -0.054%/-0.066%, effectively unchanged at this scale. Plain API medians
change +0.053%/+0.296% for bytes and +0.003%/+0.077% for range. No pooled
normal/instrumented latency, physical-network, cold-I/O or scaling claim is made.

Source and destination read calls/bytes and source work are identical in every
paired lane. Source publication data reads stay zero. Actual managed sharing
keeps the source plan reservation at 33,622,602 bytes; destination planning
reservation increases 16,826,487 to 16,826,615 bytes (+128) for inline handles.
Full decoded fallback allowance remains. After plan drop destination reservations
are zero and source is 16,807,458 bytes; after all drops all owners release.

Publication allocated bytes and peak growth are unchanged. Its absolute live
peak falls 238,097,429 to 221,320,019 bytes because the duplicate decoded copy
is absent at entry. The two-byte entry difference also appears in plain/open
regions; it is not a decoded-payload effect. Process RSS includes fixture setup
and is not evidence of timed allocation peaks. The benchmark allocator counts
requested layouts, excludes allocator-internal realloc overlap, and is not
physical memory measurement. No new hardware-counter/CPU-stack attribution ran.

All twelve primary review flags are individually retained below. No repeat
trigger exceeds 5%. The ten negative allocator flags agree in both repeats.

| Kind / lanes | Metric | Change | Disposition |
|---|---|---:|---|
| paired 0→2 | api_p99_ms | +10.883% | Primary p99 2.227084 to 2.469448 ms is retained. Each candidate process has one approximately 0.33 ms spike, in planning for R1 and publication for R2; no cause is established. Separate fixed two-ABBA confirmation gives block p99 changes +1.508%/+0.001%, with pooled descriptive p50/p95/p99 +0.132%/+0.122%/+1.082%. The >5% tail increase was not reproduced. Retain for deterministic allocation benefit; no stable population-p99 or universal latency claim. |
| paired 15→13 | api_p99_ms | +10.174% | Primary p99 2.244439 to 2.472789 ms is retained. Each candidate process has one approximately 0.33 ms spike, in planning for R1 and publication for R2; no cause is established. Separate fixed two-ABBA confirmation gives block p99 changes +1.508%/+0.001%, with pooled descriptive p50/p95/p99 +0.132%/+0.122%/+1.082%. The >5% tail increase was not reproduced. Retain for deterministic allocation benefit; no stable population-p99 or universal latency claim. |
| paired 17→19 | plan_allocated_bytes | -32.413% | 51760790 to 34983382 bytes. Removing the staged decoded copy reduces actual planning allocation by 16,777,408 bytes; both repeats pass the frozen allocation gate. |
| paired 17→19 | plan_region_peak_live_bytes | -7.585% | 220706178 to 203964628 bytes. Absolute planning-region live peak falls with decoded sharing. This includes preexisting live owners and is not RSS or a lower admission limit. |
| paired 17→19 | plan_live_growth | -33.271% | 50425902 to 33648494 bytes. Managed sharing removes duplicate retained decoded storage; both repeats show the same 16,777,408-byte reduction. Source reservations remain intact. |
| paired 17→19 | plan_peak_growth | -33.192% | 50438807 to 33697259 bytes. Planning transient growth above region entry falls; the difference need not equal payload size because the maximum occurs at a different operation boundary. |
| paired 17→19 | publication_region_peak_live_bytes | -7.046% | 238097429 to 221320019 bytes. Absolute publication-region peak falls because less memory is already live at entry. Publication peak growth remains exactly 17,404,156 bytes; publication itself does not allocate fewer bytes. |
| paired 22→20 | plan_allocated_bytes | -32.413% | 51760790 to 34983382 bytes. Removing the staged decoded copy reduces actual planning allocation by 16,777,408 bytes; both repeats pass the frozen allocation gate. |
| paired 22→20 | plan_region_peak_live_bytes | -7.585% | 220706178 to 203964628 bytes. Absolute planning-region live peak falls with decoded sharing. This includes preexisting live owners and is not RSS or a lower admission limit. |
| paired 22→20 | plan_live_growth | -33.271% | 50425902 to 33648494 bytes. Managed sharing removes duplicate retained decoded storage; both repeats show the same 16,777,408-byte reduction. Source reservations remain intact. |
| paired 22→20 | plan_peak_growth | -33.192% | 50438807 to 33697259 bytes. Planning transient growth above region entry falls; the difference need not equal payload size because the maximum occurs at a different operation boundary. |
| paired 22→20 | publication_region_peak_live_bytes | -7.046% | 238097429 to 221320019 bytes. Absolute publication-region peak falls because less memory is already live at entry. Publication peak growth remains exactly 17,404,156 bytes; publication itself does not allocate fewer bytes. |
