# Pre-accounting baseline

This directory preserves the ODS sheet-metadata profile captured before the
selector lookup/staging Work-unit and cancellation accounting fix. The source
was the staged Git index at that point; the focused source files are copied
under `source/`, and the exact staged addition is
`staged-ods-sheet-metadata.patch`. The copied `report.md`, `receipts/`, and
`harness/` retain the zero-Work observations and their replay inputs.

The baseline is intentionally kept separate from the post-accounting report.
It is an absolute profile only; no performance improvement is inferred from a
later comparison. To reproduce it, restore the staged patch into a clean
checkout, use `harness/Cargo.toml`, and run the commands recorded in the
copied report.
