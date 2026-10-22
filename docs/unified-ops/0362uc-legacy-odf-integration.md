# 0362uc — Assemble legacy and ODF raw-policy repairs

This integration combines `532cd8634b` with the legacy/ODF branch through
`8d1199f45e`. The incoming history includes the legacy owner and OLE
property-policy work, typed ODF reader errors, and required raw limits for
ODF styles, masters and drawing resources. Records 0359b–p and 0362n–nf
retain their individual implementation, review and measurement evidence.
This record does not supersede their outstanding performance returns.

The three textual conflicts were in documentation and the performance
harness. The glossary keeps both the XLSB contract entries and the OLE
property-policy entry. The curated benchmark matrix keeps the union of
both route lists without changing existing dispositions. The harness keeps
the XLSX source-history feature together with the legacy/current W5 owner
and property feature combinations. No production source required a textual
conflict resolution.

The first all-features harness check exposed an existing paired-probe
configuration gap: enabling `legacy-xlsx-source-history` selected the
removed `SourceBackedEditor` API even on the current source. The new default
`current-xlsx-source-history` feature takes precedence when both are enabled.
Historical baseline invocations select the legacy feature with
`--no-default-features`. This applies to both the tab-state and merge-history
drivers; neither adapter's operation body changes. The XLSB CRUD driver also
passes its fixed missing-marker reason directly to the crate's now-static
`InvalidFormat` variant. All three failed harness attempts and the repaired
consumer checks are retained in the receipt.

The locked workspace check with all targets and all features passes on the
assembled source. The locked all-target/all-feature performance harness and
all-target native-resave checks pass after the consumer repairs, as do the
benchmark-matrix and crate-boundary checks. Exact commands, failed attempts
and external log hashes are recorded in
[`results/0362uc/validation.json`](results/0362uc/validation.json). This is
compilation and consumer integration evidence; it does not replace semantic
integration tests, the final combined gate or native performance qualification.

The incoming 88 commits were checked against the history-rewrite map before
integration; none reintroduces a pre-rewrite commit. Large logs, captures and
compiled binaries remain outside Git. The required repository artifact
checks apply to the complete staged merge, including incoming files.
