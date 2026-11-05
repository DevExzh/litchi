# 0360d — Assemble W4b with W4 and File options

- Inputs: W4/File candidate `eafeed22a4` and W4b `b2e21c0ab4`.
- Branch: `uo2/coordinator-w4-integration`.
- Status: integration candidate; final performance qualification and trunk merge pending.

The assembly retains XLSB's immutable Package, read-only Workbook projection,
package-owned edit families and deferred source-backed metadata replay. The
OPC conflict resolves to W4b's clone-local content-type compatibility memo and
its invalidation tests, retaining the shared lexical byte projections used by
both spreadsheet families. The superseded content-types presence accessor is
removed. Benchmark parsers and generated rows are combined, and all nine
XLSB harness lockfiles retain both required shared dependencies.

The Custom Data adapter now wraps the exact shared XML limit dimension in
`custom_data::LimitKind::Xml`. Data Model proof and rename errors wrap the
exact XLDM dimension in `data_model::LimitKind::Model`. Both retain observed
and maximum amounts; they no longer classify depth, work or item failures as
a generic payload limit. XLDM's bounded rename failures now identify output
bytes and table-name text bytes explicitly. Existing rename tests verify the
one-byte-under output ceiling and oversized-name dimensions, and XLSB tests
cover preservation through the package adapters. Unrepresentable arithmetic
failures still follow their existing invalid/unproven-source path; this
assembly does not invent observed counts for them.

All-target validation found facade raw-record and sparkline consumers that
had not migrated with W4b, plus borrowed comparisons against the new static
error strings. Those tests and the structure-dump example now use explicit
policies and the current error representation. The W4b File benchmark probe
now consumes the result with `into_file()`. An initial coordinator attempt
used `into_package()` there; its compilation failure is retained with the
other failed attempts. The facade spelling is explicitly required by 0351
section 2.5, rule 7.

The first complete lint run also found 410 incoming XLSB diagnostics: 382
unnecessary lazy error constructors after the static-error conversion, 21
redundant field initializers, three unused enumeration indexes, three error
matches that hid known variants and one redundant guard. The mechanical
fixes preserve the constant error values; the three matches now enumerate
all current variants so future additions require review. Formatting exposed
24 unnecessary argument braces left by the automatic fix; these and the
remaining equivalent constant-expression blocks were simplified. No lint
suppression was added. The failed runs and automatic-fix log are retained.

The [validation receipt](results/0360d/validation.json) records external log
sizes, SHA-256 values, commands, exit codes and test counts, including failed
attempts. Nine affected packages pass all-target/all-feature compilation.
The OPC/XLDM library command passes 824 tests; the full XLSB command passes
1,317 tests with 11 ignored doctests. Tool tests, corpus, capability delta,
benchmark-matrix regeneration and crate boundaries pass. All-target/all-feature
Clippy passes for eleven affected packages, as does the XLSB-only ops profile.
The facade suite passes 23 tests and the unified XLSB filter passes 36.
The main benchmark harness and all nine XLSB harnesses pass locked offline
all-target checks. Both XLSB glossary profiles and both formatting checks pass.
All 26 incoming commits were checked against the history rewrite map and
contain no purged IDs. The source assembly is committed as `933ce43de3`;
the later results use the same Rust source and lockfiles.

This assembly does not accept the later XLSX source-owner memory residual or
metadata no-op composition restriction. Those repairs remain on the worker
branch. The remaining replay Commit migrations, raw ODF XML policy work,
standalone XLSB raw-codec policy propagation, whole-workspace
inventory/glossary closure, native performance qualification
and final 53-check gate remain required. Nothing has been pushed.
