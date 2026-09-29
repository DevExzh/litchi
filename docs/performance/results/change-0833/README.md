# 0833 — large-package filesystem qualification stopped

Status: qualification failed; **zero formal performance reports**. Production
and Rust harness source remain unchanged at
`eefaca16e39ace3c1e6219aedbac4fe1c0cd060a`.

The [frozen design](design.md) and [schedule](measurement-plan.json) planned
six OPC/PPTX routes and 72 formal reports, conditional on qualification. Five
one-sample warm/cold qualification commands succeeded. The source-backed PPTX
command failed its payload-range classification, so no formal schedule ran.
Three separately labelled diagnostics show warm-only PPTX success, cold-only
PPTX failure, and a combined cold OPC save output-parity failure. Six retained
JSON reports contain eleven qualification/diagnostic samples; none is a
performance baseline. Failed commands retain exact terminal receipts and logs.

[Coverage priorities](coverage-priority.md), [protocol review](protocol-review.md)
and [diagnosis](diagnosis.md) distinguish source findings, observed failures and
unproven causes. The [PPTX optimization follow-up](pptx-follow-up.md) remains
unmeasured. It is separate from repairing this harness admission failure.

Root executed `driver.py prepare`, `build`, and `qualify`, followed by
`diagnose.py`. The driver records 9,388 source/normative inputs, the executable,
compiler, host, environment, exact commands and logs. All outputs are exclusive;
no failed qualification was replaced. `commands_pass` would mean only zero
subprocess exit codes, not full oracle admission; the actual status is `failed`.

`reuse_quality.py` checks the eight sealed 0832 after-source gates against all
9,388 current inputs. These are reused test results, not fresh 0833 tests.
`driver.py quality` was not executed. It remains an optional command for a new
packet and must not overwrite this reuse witness. The release executable was
freshly built. Unexecuted formal-capture/statistics drafts were removed after
qualification failed; the frozen measurement plan remains as intent, not data.

The cold OPC outputs differ: source-backed publication preserves the private
1,776-byte EOCD alignment comment, while borrowed eager reconstruction returns
the unpadded output length. Warm outputs agree. The existing combined-selector
check rejects the cold difference. Do not silently substitute hash equality,
drop comment bytes, or declare a production preservation bug from this record.
The PPTX failure lacks retained failing request vectors; its exact overlap still
needs diagnostic proof before changing its classifier.

Existing counter scopes matter: OPC source-backed counters run inside timing;
PPTX positional counters belong to an untimed replay; eager logical counters
are unavailable. OPC eager open uses the borrowed constructor and drops its
materialized package inside timing, while source open retains its package
through post-timer diagnostics. No route latency comparison is admitted here.
Verified-cold evidence proves only per-file page-cache/procfs observations.
Default durable atomic publication is unchanged. No allocation, hardware-counter,
physical-request, copied-byte, concurrency, or speedup claim is made.

Offline replay commands and cleanup/seal status are recorded in the final
performance report. Reproduction on another host/revision needs a new packet,
source/environment binding and qualified cold support; never overwrite this
record. Complete generated corpus manifests are embedded in the reports.
