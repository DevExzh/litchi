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
