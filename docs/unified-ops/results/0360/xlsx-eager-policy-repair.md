# XLSX eager policy allocation repair

The focused source-family ABBA driver is
`tools/perf-baseline/src/bin/xlsx_source_family_abba.rs`. Its canonical source
for the diagnostic pair is SHA-256
`7dd57fce2b97f3c044c3698916b73a8bad33db0c28973702b7cc3cb2661820c7`.
The protected candidate and baseline driver binaries were
`663e6904ea114a3ceccdf46a1aa4c8c266d7b65a97f19e2ccded11db17fe749a` and
`78222adfbacf846b31112d7f1cdd6d6a4705500950e263e8e600f4a4ccf1bb28`.
They were built in distinct target directories; the baseline was rebuilt after
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

Validation for this repair completed with:

- `cargo check -p litchi-xlsx --all-targets`
- `cargo check -p litchi-xlsx --all-targets --features encryption`
- `cargo test -p litchi-xlsx --test form_control_owner_read --features encryption` (39/39)
- `cargo test -p litchi-xlsx --test calculation_metadata --test custom_data_limits` (35/35)
- workspace rustfmt check and `git diff --check`

A new native ABBA is required after the repair and is not claimed by this
record; no timing was run concurrently with the coordinator's later quiet
windows.
