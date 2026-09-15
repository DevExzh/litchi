# Log sections for change 0597

Four paragraphs for the coordinator to merge, one each into `HOTSPOTS.md`,
`GOAL_AUDIT.md`, `REPORT.md` and `ADR_COMPLIANCE.md`, in the style of their
newest sections. The record's own link is
[`0597-xlsx-selected-cell-ineligibility-gate.md`](../../0597-xlsx-selected-cell-ineligibility-gate.md)
and the packet is [`results/change-0597/README.md`](README.md); both live at the
top level of `docs/performance/`, not under `changes/`.

---

## HOTSPOTS.md

## 0597 — XLSX selected-cell ineligibility gate, frozen at its design

Survey item XML-2 priced and frozen. On a source-backed one-cell read of an ineligible worksheet, `raw::worksheet::selected::scan` is 140,852,422 Ir (63.80% of the whole child) on the marker-stripped control fixture and 152,947,518 (13.81%) on the real one, after which `SourceWorksheet::store` re-parses the same part; 0587's falsification threshold of 10% is not met. A bounded 8 KiB `<cols>` pre-gate removes that scan for −63.19% and −14.08% of the two reads and is a byte-identical no-op across the whole 4,606-row real-corpus transcript, but it moves error identity on malformed input — two refusals become acceptances, three change typed variant, five change message — so it stops at a frozen design with its patch retained, pending a limit question and an ADR 0005 clarification. Stopping at the first mark instead is not implementable from `litchi-xlsx`: `invoke_active` disables the observer and lets the MCE driver run to EOF by design. The brief's oracle also found a live defect, fixed here and separately logged. No speedup, latency or resource claim. OLE2/OOXML remain active; ODF is deferred until completion and iWork excluded. [Change and limitations](0597-xlsx-selected-cell-ineligibility-gate.md); [retained evidence](results/change-0597/README.md).

---

## GOAL_AUDIT.md

## 0597 — XLSX selected-cell ineligibility gate, frozen at its design

Rank-6 survey item XML-2 is now priced (−63.19% of a source-backed one-cell read on the control fixture, −14.08% on the real one, measured as callgrind isolation pairs) and frozen: the gate moves which typed error a malformed `<cols>`-bearing worksheet reports, so GOAL's "capture BEFORE measurements, make the smallest coherent change, never trade a typed refusal for a partial result" line puts it behind a frozen design rather than in this batch. The same brief's oracle closed a correctness gap instead: `SourceWorksheet::cell`/`cells` refused 434 reads over 24 of 180 fixtures with `worksheet mergeCells appears before sheetData`, a question answered from state the scanner stops maintaining after `mark()`. After the fix the source-backed path agrees with the mandatory materialized parser on all 3,948 comparable rows (372 disagreements at the base). 1,297 tests pass. The unmeasured XML-1 codec expansion, the harness's lack of any marker-bearing or ineligible worksheet corpus, and the gate's two open prerequisites keep the non-iWork performance goal open. OLE2/OOXML remain active; ODF is deferred until completion and iWork excluded. [Change and limitations](0597-xlsx-selected-cell-ineligibility-gate.md); [retained evidence](results/change-0597/README.md).

---

## REPORT.md

## 0597 — XLSX selected-cell ineligibility gate, frozen at its design

Mixed outcome, `performance_claim: none`. Implemented: the selected-cell scanner's post-mark merge placement check no longer asks whether `<mergeCells>` "appears before sheetData", a question it answered from a `seen_sheet_data` flag that its own post-mark early return prevents from advancing; 434 reads over 24 of 180 `.xlsx` fixtures were refused for a placement the bytes contradict, and after the fix the path agrees with the fully materialized `litchi_xlsx::Workbook` on all 3,948 comparable rows. Frozen, not implemented: a bounded 8 KiB `<cols>` pre-gate that skips the whole semantic stream, worth −63.19% and −14.08% of a source-backed one-cell read on the two survey fixtures and observationally a no-op on the real corpus, but which moves eight error classes on a 60-case first-error matrix in the 0541 style; its patch, matrix and prize are retained. Paired timing of `xlsx_file_selected_cell`, `xlsx_range_source_first_cell` and `xlsx_narrow_column_range_scan` moves nothing outside an A/A floor that runs from 0.874% to 24.011% at p50 across the seven scenario/shape pairs; the two deltas above 5% are reported and both are below their own floor. No speedup, latency, cold-cache, physical-I/O, allocation or RSS claim follows. OLE2/OOXML remain active; ODF is deferred until completion and iWork excluded. [Change and limitations](0597-xlsx-selected-cell-ineligibility-gate.md); [retained evidence](results/change-0597/README.md).

---

## ADR_COMPLIANCE.md

## 0597 — XLSX selected-cell ineligibility gate, frozen at its design

ADR 0005 places mandatory structural validation at open and loads semantic payloads lazily; the worksheet payload's mandatory validation is the materialized parser's, and change 0362 already required every `NotEligible` result to reach it ("`NotEligible` is not worksheet semantic validity"). The landed fix routes to that parser instead of around it and touches no fence: `with_verified_decoded_reader`, the CRC/size/source/execution fences of changes 0363 and 0365, `selected_stream_limits`, the MCE and x14ac observers and `Scanner::finish` are unchanged, the stream still reaches XML/MCE/x14ac EOF on every worksheet, and both placement helpers stay private. No new `unsafe`, no new dependency, no weakened limit, no public API change. The frozen gate is frozen precisely on ADR grounds: it would let a `<cols>`-bearing worksheet reach only the materialized path's limits rather than the stream's `max_event_bytes`/`max_input_bytes`, and it would turn two synthetic refusals into acceptances that `litchi_xlsx::Workbook` already grants on the same bytes — a convergence or a weakened defence depending on which reader ADR 0005 means to own malformed-input refusal for a lazily loaded payload, which this record does not decide. 1,297 tests pass; `cargo fmt --all --check`, `cargo clippy -p litchi-xlsx --all-targets`, `cargo doc -p litchi-xlsx --no-deps` are clean. Accepted ADR hashes unchanged. See [Change 0597](0597-xlsx-selected-cell-ineligibility-gate.md); `performance_claim: none`.
