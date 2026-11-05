# XLSX eager policy allocation repair

The focused source-family ABBA driver is
`tools/perf-baseline/src/bin/xlsx_source_family_abba.rs`. Its canonical source
for the diagnostic pair is SHA-256
`165de1ef06c37811f49e6ec66bb85a42a72ff33cac9b35b80646922589e1b7e0`.
The earlier protected candidate and baseline driver binaries were
`663e6904ea114a3ceccdf46a1aa4c8c266d7b65a97f19e2ccded11db17fe749a` and
`78222adfbacf846b31112d7f1cdd6d6a4705500950e263e8e600f4a4ccf1bb28`.
Those earlier binaries were built in distinct target directories; the baseline was rebuilt after
an earlier stale-target attempt was discarded.

The v3 six-family receipt is outside the repository at
`/tmp/litchi-w4-source-family-abba-v3/report.json`. Its per-block data showed
small eager-side increases across all six routes. The focused receipts are
`/tmp/litchi-w4-source-family-abba-v5-page_margin/report.json` and
`/tmp/litchi-w4-source-family-abba-v5-row_visibility/report.json`.
Page-margin eager legs were positive in 17/20 blocks (median delta
+0.880%); row-visibility eager legs were positive in 19/20 (median +1.193%).
The corresponding source-backed medians were -1.052% and +0.336%.
These are retained diagnostic results, not acceptance measurements.

The common eager open path previously cloned `OpenOptions` into every
`Workbook::Inner` and allocated a `Policy` `Arc` even when the package used
strict default semantic limits and had no encryption provenance. `Workbook`
now keeps the policy as `Option<Arc<Policy>>`: strict defaults remain in
`Inner::options`, while explicit semantic-limit overrides and encrypted
provenance still install and share a policy. Derived snapshots preserve the
same override. Worksheet form-control reads use the owning default only when
there is no override. The `OpenOptions` move into `Inner` also removes the
unnecessary full-policy clone.

The eager snapshot now keeps a compact `WorkbookOptions` containing only
leniency, semantic limits and (when enabled) encryption limits. The removed
`cache_limits` and `execution_context` are source-owner policies and the
owning package explicitly ignores them. The removed physical `SourceLimits`
remain on the validated `OpcPackage::read_limits()` value; every eager edit
and codec limit check reads that package value, and OPC clone/edit propagation
is covered by its package tests. This distinction follows
`crates/litchi-opc/src/limits.rs` (physical ceilings for reading one OPC
package) and ADR 0030 (`docs/adr/0030-lazy-opc-part-decode.md`): an owning
edit does not perform a new generic output census, while each derived reader
still applies the retained physical profile. The XLSX regression
`derived_owned_edit_keeps_the_physical_part_limit` opens at the largest
admitted part boundary, commits an oversized derived cell, checks source
immutability and limit retention, and verifies that reopening the resulting
physical part under the same ingress ceiling is refused. The source-family
regression in `tests/custom_data_limits.rs` performs an unrelated eager edit
and then starts the connections reader under a retained depth ceiling; both
the committed owner and an independent reopen return the same typed depth
limit refusal. This exercises a real derived-family admission path rather than
using read-limit equality as a proxy.

Validation for this repair completed with:

- `cargo check -p litchi-xlsx --all-targets`
- `cargo check -p litchi-xlsx --all-targets --features encryption`
- `cargo test -p litchi-xlsx --lib derived_owned_edit_keeps_the_physical_part_limit` (1/1)
- `cargo test -p litchi-xlsx --test custom_data_limits --features encryption` (29/29)
- `cargo test -p litchi-xlsx --test form_control_owner_read --features encryption` (39/39)
- `cargo test -p litchi-xlsx --test calculation_metadata --test custom_data_limits` (7 + 29)
- `cargo clippy -p litchi-xlsx --all-targets --all-features -- -D warnings`
- `cargo fmt --all -- --check` and `git diff --check`

The compact-policy change and its retained physical-limit regression passed
the post-hold validation above. No native ABBA was run in this validation
window; the recorded v5 binaries remain diagnostic only.

The same validation also covers the deferred source-family replay repair in
the W4 tree. `RestoreEndpoint` retains an exact opened artifact or a compact
deferred publisher pair, authenticates the recipient's immutable before
artifact, and streams the after artifact through the recipient's own source
limits and execution context. The source families provide authenticated
overlay writers so an already authenticated replay does not perform a second
whole-archive publication prepass. Their merge source metadata is shared per
owner while rewritten worksheet bytes remain separate. The focused source
family integration set (auto-filter, calculation metadata, cell values,
conditional formatting, data validation, defined names, source file,
hyperlinks, merge, page-breaks, page-margins, page-setup, print options, row
visibility, selected payloads, sheet protection, and tab state) passed 210/210
tests with encryption enabled. The owning-facade page-break conversion keeps
its explicit materialization refusal; source-backed page-break replay uses the
same authenticated recipient path as the other families.

The recorded v5 binaries predate this final claim-string edit and are retained
only as diagnostic receipts. A new native ABBA is required after the repair and
with the committed driver; no timing was run concurrently with the
coordinator's later quiet windows.

## Final 2a-to-fee repair attribution

The protected repair pair is the 2a74117-to-fee2cdf pair, not the original
pre-W4 acceptance baseline. The source-family orchestrator source is
`tools/perf-baseline/src/bin/xlsx_source_family_abba.rs`, SHA-256
`165de1ef06c37811f49e6ec66bb85a42a72ff33cac9b35b80646922589e1b7e0`. The
protected orchestrator binaries are `832d592f...` (before) and
`bc61708b...` (after); the child binaries are `80741228...` and `b2eab1f...`.
Full hashes, per-leg samples, and complete child fidelity results are outside
the repository under
`/tmp/litchi-w4-source-batched-abba-20261021-v1/`.

The final focused source/eager route was 20 ABBA blocks with order
before/after/after/before. Each child used two warmups and 64 measured samples,
with every elapsed sample and complete child result retained. The batched
source legs therefore exceed 150 ms of measured work even though the original
one-sample child legs were only a few milliseconds:

- `page_margin` source: 80 legs, 182.900--240.639 ms per leg, pooled after vs
  before mean +0.4179%; paired block deltas ranged from -5.594% to +8.548%.
  Every output hash was `2b50a470...`, with 16,786,883 accepted bytes and 375
  logical source reads totalling 16,793,599 bytes per sample.
- `row_visibility` source: 80 legs, 1,369.205--1,461.450 ms per leg, pooled
  after vs before mean -0.3180%; paired block deltas ranged from -2.752% to
  +2.606%. Every output hash was `59b17e47...`, with 4,222,186 accepted bytes
  and 163 logical source reads totalling 4,228,351 bytes per sample.
- `row_visibility` eager: the same medium shape and 64-sample protocol, with
  696.028--706.606 ms per leg and pooled after vs before mean +0.1278%.
  Eager output, sink bytes, and the no-source-read contract matched the source
  route.

The first row batch omitted the required `--xlsx-row-visibility-shape medium`
selector and returned both medium and large child results before accepting a
leg. It is retained as the excluded attempt
`excluded-row-visibility-unconstrained-v0.json` in the same external
receipt directory; it is not combined with the corrected medium-shape data.
The corrected route has the same output hash, sink bytes, and source read
counts as the short medium-shape route.

The compact route used the protected merge children for 20 ABBA samples with
two warmups. It retained 160 merge/unmerge observations. Merge p50 was
140,561 ns before and 139,971 ns after; unmerge p50 was 128,951 ns before and
124,731 ns after. All observations passed exact one-part overlay, inverse
complete-package restoration, stale-source refusal, and foreign-source
refusal. The compact report is outside the repository at
`/tmp/litchi-w4-source-family-repair-compact-page-margin-final-v2/report.json`.

The paired allocator route used 128 retained history steps and four warmups.
Exact forward and inverse checks were 128/128 for both binaries. Retained
history allocated bytes fell from 4,433,237 to 4,165,697, and peak live bytes
from 450,079 to 448,851. Independent verification allocated bytes fell from
215,493,056 to 212,529,856, and peak live bytes from 887,051 to 636,280.
The common replay aggregate nevertheless rose from 58,341,504 to
110,317,056 allocated bytes; this remains an unaccepted attribution result,
not a performance exception. The complete pair is outside the repository at
`/tmp/litchi-w4-source-allocation-2a-fee-v1/`.

These measurements are repair attribution only. The original pre-W4 frozen
baseline has no complete source-state provenance, so the original-trunk to
final acceptance comparison is still pending a reproducible build from the
exact pre-W4 integration state.

## Returned defined-name source reparse repair

The completed defined-name route in
`/home/zhuhe/code/litchi-worktrees/.coord/w4-xlsx/e63-source-family-abba-batched-v3`
is retained as an unaccepted diagnostic receipt. Its five source blocks were
all slower after the W4 migration: the source means were 1.009538x,
1.010027x, 1.016369x, 1.011511x and 1.012955x (paired block order as recorded
in `route-defined_names.json`). The source output hash, accepted bytes,
largest write, and source read shape were unchanged; the receipt was not
accepted as timing evidence because the enclosing run was later stopped by an
unrelated host build.

Source phase counters are unavailable for this route, so the code audit found
the attributable extra work. The generic deferred publisher parsed the
rewritten workbook catalog once to populate `defined_names()` and
`active_sheet_position()`. The defined-name snapshot had already parsed that
same catalog while authoring and reopening the target. The old stream publisher
performed only its source-state check before streaming. Tab-state had the same
duplicate catalog construction.

The repair adds a parsed-catalog handoff to the deferred change contract.
Defined-name and tab-state snapshots retain the catalog fields they already
parsed, and their patches pass those fields directly to the deferred owner.
The handoff is used only when the recipient's `SourceLimits` exactly match the
limits retained by the target snapshot. A stricter or otherwise different
recipient falls back to parsing the target XML under its own limits, so a
changed target that exceeds recipient admission is refused before publication.
The exact-artifact replay path applies the same check: it reuses the retained
catalog only for an equal policy and reparses the recorded target XML for a
different recipient policy before constructing the restored owner.
Other families retain the XML fallback and all recipient source, execution,
signature, and bounded-stream checks. Focused validation passed 7 defined-name
tests, 12 tab-state tests, and the cross-policy target-admission regression with
`cargo test --offline --locked -p litchi-xlsx --test source_backed_defined_names
--test source_backed_tab_state`. A fresh protected e63-to-repaired-candidate
native receipt is still required; no timing result from the old receipt is
accepted by this record.
