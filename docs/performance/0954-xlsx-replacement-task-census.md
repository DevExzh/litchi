# 0954 — exact XLSX replacement compression tasks

The medium one-percent workload clears the existing cold compression-pool
threshold; dense-sparse does not. This closes the missing payload-size evidence
from [0953](0953-xlsx-frame-pointer-publication-attribution.md), whose qualified
profiles identified compression as the largest sampled publication cost.
No production change or speedup is adopted in this batch.

The retained 0904 exporter artifacts under
`/home/zhuhe/litchi-performance-0947/inputs/artifacts` contain complete source and
output ZIPs. Both primary source/output SHA-256 pairs exactly match the current
0953 census's 0950 benchmark identities. Their age therefore does not require
rebuilding a workload to recover these exact bytes. The independent Python
census verifies archive hashes, every selected XML sidecar, decoded member
lengths against central-directory metadata, exact changed-member membership,
and compressed payload equality for every untouched member. ZIP decoding also
checks member CRCs. This does not independently replay the Rust semantic audit.

| Replacement member | Medium 1% bytes | Dense-sparse 1% bytes |
| --- | ---: | ---: |
| `xl/workbook.xml` | 527 | 527 |
| `xl/worksheets/sheet1.xml` | 63,294 | 462,569 |
| `xl/worksheets/sheet2.xml` | 68,488 | 31,398 |
| `xl/worksheets/sheet3.xml` | 68,488 | 8,115 |
| `xl/worksheets/sheet4.xml` | 68,488 | 6,352 |
| Total | 269,285 | 508,961 |
| Largest | 68,488 | 462,569 |
| Total minus largest | 200,797 | 46,392 |

All five changed members use Deflate. Twelve untouched compressed payloads are
byte-identical in each primary. The retained medium one-edit control changes
workbook and sheet 1: 63,821 total bytes, 63,294 largest, and only 527 remainder;
fifteen untouched compressed payloads match. Its identities are bound to the
0904 manifest; no new 0953 one-edit run is claimed.

`PreservationPlan::deflate_wave` caps width by workers, in-flight tasks, task
count, and any per-task floor after the remainder passes the session threshold.
At the existing zero per-task floor, a cold local session requires 65,536 bytes
of remainder; an existing session/facility requires 8,192. With adequate budgets,
medium's advisory widths for 1/2/4/8 workers are 1/2/4/5 in either state.
Dense-sparse remains width 1 cold, and permits 1/2/4/5 warm. The one-edit control
remains width 1 in either state. These are source-policy calculations, not
observed parallel execution, CPU balance, latency, or speedup. Warm refers to
pool/session lifecycle, not cache warmth. Logical byte balance does not measure
compression cost balance.

The existing managed XLSX benchmark hardcodes one worker and one in-flight task,
and creates a new source-backed package per iteration. Raising a global CLI
worker argument alone cannot exercise this route. A subsequent opt-in experiment
should use explicit finite execution and memory budgets, starting with two and
four workers for medium. Keep dense-sparse and one-edit as serial-fallback
controls. The policy's per-worker memory estimate is 655,360 bytes, additional
to package/payload/output needs; this is a reservation model, not measured RSS.
A managed lifecycle replay must cover inverse/no-op byte identity, source
freshness, cancellation and refusal boundaries as well as output equality and
reopen semantics before native performance qualification. A warm-facility
experiment would be a separate workload contract.

[The evidence packet](results/change-0954/README.md) contains the reducer,
exact sizes/hashes, source-policy bindings, input revalidation and replay receipt.
All 35 normative input hashes remain unchanged from 0953. No build, native
capture, production edit, or broad correctness-suite rerun was needed. Prior
performance gates remain binding, including the one-edit tail regression flag.
The overall performance goal remains active; iWork remains outside this work.
