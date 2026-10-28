# iWork reference-line parity order disposition

This record distinguishes the frozen failing source
`39e4bf101ca51f269e9352f8fed868cb84799113` from the isolated repaired source
`de367c1b08c1b3dcbabef21a97e1921361e72b0c` on
`uo2/iwa-w1c-reference-order` (the record-only follow-up commit is separate).
It addresses only the row-33 Tier-B failure in
`reference_line_cases::pages_chart_reference_lines_allocate_parity`; it does
not change a host or leaf writer.

## Finding

The host chart writer collects managed reference-line object identifiers in a
`HashSet` (`crates/litchi-iwa/src/charts/reference_line/archive.rs:714-741`)
and appends the set's iteration order to `MessageInfo.object_references`
(`:661-665`). That order is process-seed dependent. The shared leaf engine
retains the extension's ordered identifiers (`crates/litchi-iwa-style/src/reference_line.rs:372-394`). The chart extension is the ownership source (`TSCH.ChartArchive.reference_lines`, field 10005), while the aggregate list is dependency membership: core permits retained identifiers to be reordered under an explicit transition (`crates/litchi-iwa-core/src/archive.rs:3676-3687`).

## Narrow comparator policy

`Case::reference_line_allocation_metadata(chart_id)` opts in only the three
Tier-B reference-line allocation cases and binds the policy to that selected
chart object. `ReferenceOrderPolicy::ReferenceLineAllocation` activates only
for that chart's message type 5021. It derives owned IDs from both payloads'
field-10005 extension, requires the extension-owned ID sequence and all nested
required reference/UUID fields to be valid, and requires every owned aggregate
reference exactly once. It compares owned references as a multiset; all
residual references retain exact source order and multiplicity. The default
policy remains exact for every other case, object, and message.

The targeted payload comparison masks only known lower/upper UUID fields of
reference-line items whose object IDs are absent from the source snapshot.
Existing item UUID submessages, including their unknown fields, must remain
byte-identical to the source in both outputs. Unknown fields in newly
allocated UUID submessages remain part of the normalized payload, as do all
other unknown chart fields; changed unknown payload bytes fail the targeted
policy. Missing or duplicate aggregate references, unrelated reorder, wrong
extension IDs, duplicate extension IDs, wrong-chart selection, existing-UUID
corruption, and malformed wire payloads are refusals, not allowed
differences.

The intentional Tier-B difference is reported as:
`chart reference-line-owned dependency references differ only in order; all
unrelated references remain ordered and exact`.

## Evidence boundary

The preserved pre-repair gate failure is external to the checkout:

- `full-39e4bf101c/33-tests-workspace-excluding-ods-ops.out`, 8,134,497 bytes,
  SHA-256 `903b805d5951fd345d658669ff02007d6c1449adcc33bfbf7c77538c30b1e9cc`;
- row 33 was the sole reported failure at the time of this repair; native
  acceptance remains queued and is not claimed here.

The repeated Pages allocation parity runs and focused negative controls are
process-local differential evidence only; they do not grant native E3/E4
evidence. Final full-gate qualification remains coordinator-owned on the
integrated trunk commit.

## Verification

The affected reference-line parity suite passed ten non-cost tests, the
comparator negative controls passed, the policy opt-in/default/wrong-chart
assertions passed, and the Pages allocation case passed in three independently
started test processes. The end-to-end changed-unknown-payload case is
expected to fail the candidate comparison and asserts that refusal. Exact
commands and exit statuses for the repaired source commit are:

- `rustfmt --edition 2024 --check crates/litchi-iwa/tests/parity/archive_diff.rs crates/litchi-iwa/tests/parity/harness.rs crates/litchi-iwa/tests/parity/reference_line_cases.rs` — 0;
- `CARGO_BUILD_JOBS=2 cargo test -p litchi-iwa --test parity archive_diff::tests::reference_line_policy -- --nocapture` — 0 (6 passed);
- `CARGO_BUILD_JOBS=2 RUST_TEST_THREADS=1 cargo test -p litchi-iwa --test parity reference_line_cases:: -- --nocapture` — 0 (10 passed, 1 ignored cost probe);
- `CARGO_BUILD_JOBS=2 RUST_TEST_THREADS=1 cargo test -p litchi-iwa --test parity reference_line_cases::pages_chart_reference_lines_allocate_parity -- --exact --nocapture` — 0 in each of 3 independently started processes;
- `CARGO_BUILD_JOBS=2 cargo clippy -p litchi-iwa --test parity --no-deps -- -D warnings` — 0;
- `RUSTC_BOOTSTRAP=1 python3 tools/check_api_glossary.py --profile iwa` — 0;
- `python3 tools/iwa_exit_ledger.py --check` — 0;
- `python3 tools/check_crate_boundaries.py` — 0;
- `python3 -m unittest tools.test_no_3rdparty_references tools.test_gate_runner` — 0.

No production writer, ledger row, macOS queue entry, native evidence, or
performance claim changed in this test-only repair; coordinator requalification
must run on the final integrated source.
