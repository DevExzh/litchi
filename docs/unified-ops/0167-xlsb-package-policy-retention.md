# XLSB package policy retention

`Workbook::into_package()` now consumes an already validated workbook into its
physical `Package` owner. It moves the existing OPC graph without serialization
or reparsing, retains its source lineage and limits, and carries external-link
limits, drawing projection, worksheet leniency and the open-time tolerance
report. The report is cloned; worksheet caches and semantic projections are not
retained in the package.

`Package` now retains drawing and leniency policies in addition to its existing
external-link limits and tolerance report. Borrowed and consuming workbook
projections use those retained policies. Default package ingress still selects
eager drawings and strict leniency; converting an explicitly opted-in workbook
does not change the default for other callers.

Package candidate publication now uses a shared reconstruction method that
retains those policies. Slicer, timeline, custom-data, calculation-chain,
data-model, volatile-dependency and explicit signature-removal paths use it.
The original candidate OPC graph is retained after validation, preserving
source/signature metadata; it is not replaced with a serialized reconstruction.

## Regression coverage

- A workbook with both referenced and unrelated malformed drawing parts is
  admitted through the explicit skipped projection, converted to a package,
  projected, reconstructed and consumed again. All part payloads remain exact;
  ordinary eager reopening still refuses the malformed drawing.
- A signed workbook converts and projects with exact original archive bytes
  and its signature intact.
- Explicit OPC input limits and external-link limits survive conversion and
  consuming projection.
- Opt-in cached NaN handling still yields the expected recorded `#NUM!` repair
  after conversion and package reconstruction. Default ingress still rejects
  that cell when read.
- A runnable native example demonstrates the consuming conversion and default
  policy introspection.

This is the policy-preserving owner conversion prerequisite identified in
0166. It does not yet move the structural edit entry points onto Package or add
the canonical package member to structural commits. The next wave must reuse
validated replay results, preserve exact no-op owners and remove superseded
Workbook edit/apply entry points as consumers migrate. Full-image staging and
source-backed publication costs remain unresolved; no performance acceptance
or unified XLSB write support is claimed here. The all-format goal stays open.

Independent review found no missed Package reconstruction path. It identified
an adjacent remaining gap: detached cell-value, cell-watch, sparkline and
structural transactions use `Workbook::reparse_candidate_with_policy`, which
carries drawing/external-link policy but not leniency. Their transaction state
must retain leniency before the complete owner migration is claimed. This wave
only repairs the Package conversion and reconstruction paths.

## Verification

Default native tests and doctests passed 1,072 cases before adding the final
runnable example. The final `sign,vba-inspection` suite passed 1,084 cases,
including that example. All 489 XLSB-only facade tests and doctests passed.
Native and facade all-target Clippy passed with warnings denied. Warning-denied
minimal/full-feature rustdoc JSON builds and both glossary profiles passed
(17 banned names each). Scoped formatting and whitespace checks passed.

Validation uses `CARGO_PROFILE_DEV_DEBUG=0 CARGO_PROFILE_TEST_DEBUG=0
CARGO_INCREMENTAL=0 CARGO_NET_OFFLINE=true` and an external temporary target.
Full-feature checks use a temporary workspace lockfile copy with `cc` 1.6.0,
resolved offline using Cargo's unstable lockfile-path option. The inherited
workspace lockfile remains unchanged.

Reproduce native tests with `cargo test -p litchi-xlsb` and
`cargo test -p litchi-xlsb --features sign,vba-inspection`; facade checks use
`cargo test -p litchi-ops --no-default-features --features xlsb`. Clippy uses
the latter native and facade feature sets with `--all-targets -- -D warnings`.
Build XLSB rustdoc JSON under minimal and full features with warnings denied,
then use `tools/check_api_glossary.py` profiles `xlsb-minimal` and `xlsb-full`.

The crate-boundary checker passed for 66 packages and 268 internal dependency
declarations. Temporary build targets, rustdoc JSON, logs, scripts and the copied
lockfile were removed after verification. No worktrees or borrowed fixtures were
created. The rewritten base remains an ancestor.
