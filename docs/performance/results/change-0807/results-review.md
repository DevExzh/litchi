# 0807 results review

I reviewed the terminal report against `native-analysis.json`,
`profile-analysis.json`, `perf-analysis.json`, `root-frame-counts.json`,
`root-scan-costs.json`, and the post-capture event assembly and offset
receipts. The retained replay logs pass after cleanup. I did not run a build,
capture, benchmark, or heavy replay.

The report's numerical claims are internally consistent. The native lane has
36 reports and 1,080 measured samples. Its six-block p50 medians and bootstrap
intervals match `native-analysis.json`; the six recorded process-spread flags
remain visible. The Callgrind totals match all six positive dumps and their
empty termination dumps. The large pass-0 details also match the raw edges:
`inspect_element` calls `CheckedAttributes` 152,410 times with 37,176,356
inclusive edge Ir, and `scan_processed_xml` calls `NsReader::process_event`
282,612 times with 46,146,741 inclusive edge Ir. The raw incoming count of
152,412 for `CheckedAttributes` includes two calls from the presentation
reject helper, so it is not a report discrepancy.

The frame-pointer counts match both independent readers: 3,164/3,169 whole
process stacks, 1,036/1,033 exact-owner stacks, and 189/190
`NsReader::process_event` leaf samples. One unresolved interior frame remains
in each repeat. The report correctly treats these as observed nested samples,
keeps overlapping rows separate, and makes no native phase-fraction or
instruction-cost claim. The post-capture offset census also matches the
retained frames: offsets `0x43`, `0xbe`, and `0x118` have 54/67/45 and
47/55/59 observations. Its binary identity and disassembly receipts bind the
follow-up, while the report correctly leaves sampling skid, code-generation
differences, and causality unresolved.

The wrapper latency table is a diagnostic control comparison. Its large-leg
3.285% difference is explicitly identified as wrapper/code-generation
perturbation, and no production improvement is inferred. The scope limits are
clear: warm generated in-memory PPTX shapes, one opened-capture boundary,
guest Callgrind work, and exact-owner native ancestry. The packet does not
claim native Office, cold or remote input, concurrent workloads, comprehensive
CRUD coverage, production speedup, or causal benefit from the proposed event
handoff hypothesis. The 0806 candidate remains rejected.

No material results or limits error was found. The next event-handoff trial
still requires the stated buffered scanner oracle and refusal/namespace
regressions before any paired workflow measurement.
