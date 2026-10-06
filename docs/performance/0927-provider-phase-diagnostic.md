# 0927 — observe section-layout provider reads at phase boundaries

The previous batch established two complete source fingerprints but could not
assign the remaining provider reads to workflow phases. This batch adds a
test-only diagnostic to the performance harness. It uses the existing DOCX
section-layout corpus and the paragraph-64 columns edit, with fresh providers
that return at most 64 KiB or 4 KiB per call. Both providers are unpaced.
They model positional read boundaries, not network or physical storage.

The diagnostic snapshots logical and delegated counters immediately after
open, query, edit, commit and publication. Checked differences produce phase
counts; checked sums must reproduce the operation totals. It reports
successful nonempty reads at the existing observer boundaries. It does not
measure metadata calls, failed reads, CPU cycles or elapsed time. Verification
and inverse restoration occur after the publication counter snapshot.

All helper code is inside the harness's test module. Production libraries,
the normal benchmark runner, and its timing/report schema are unchanged.
This is a separate count diagnostic, not a recapture or repair of 0925.

## Observed phase counts

Both successful routes return exactly the same bytes in every phase. The
short provider increases calls and caller-requested buffer volume, while its
delegated requests remain capped and equal returned bytes.

| Phase | Full-return calls | 4 KiB calls | Returned bytes, either route | 4 KiB logical requested bytes |
| --- | ---: | ---: | ---: | ---: |
| Open | 14 | 14 | 1,692 | 1,692 |
| Query | 34 | 515 | 2,102,808 | 33,132,693 |
| Edit | 0 | 0 | 0 | 0 |
| Commit | 0 | 0 | 0 | 0 |
| Publication | 117 | 1,078 | 4,203,561 | 50,962,086 |
| Operation | 165 | 1,607 | 6,308,061 | 84,096,471 |

Subtracting the source-derived fingerprint model from query leaves one
1,381-byte read on either route. Subtracting one fingerprint from publication
leaves 84 full-return calls or 564 short-provider calls, returning 2,102,134
bytes. Together with opening, these account for the 2,105,207-byte residual
identified in 0926. The publication remainder still combines preservation
validation and source copying; the phase observer does not distinguish them.

Query has 513 short reads and publication has 993; opening has none. No
provider reads occur during the isolated edit and commit in this scenario.
This does not imply those phases have no CPU or allocation cost. Likewise,
the full-return query's single non-fingerprint read does not mean its XML
parsing work is cheap. Counts alone do not rank native CPU opportunities.

Both routes emit 2,101,446 bytes with SHA-256
`5e87e9ca9fd6b9a98933c36d9aee1e848bc97eb7693229f7949514830121c4ec`,
match the existing semantic/preservation oracle, and restore the exact
2,101,427-byte source through inverse publication. The largest sink write is
65,536 bytes. Source versions remain equal around the diagnostic and checks.

## Native CPU evidence boundary

The retained 0923 native data contains phase wall times, not CPU samples that
separate SHA-256, XML/layout and ZIP work. The older 0915 Callgrind profile
establishes guest-instruction work; it cannot establish native CPU shares on
the current binary. Its publication snapshot subtree also excludes later ZIP
emission and output hashing.

A new capability probe, `perf stat -e cycles:u -- true`, succeeds with perf
7.0.14 and `perf_event_paranoid=1`. The receipt and raw output are retained in
`results/change-0927/perf-capability-probe/`. This proves only that this
process-scoped event could be opened for the probe. It does not establish a
document performance result, sampling quality or phase attribution.

The next native profile can use a source-bound binary with explicit workflow
boundaries. A whole-process profile would include corpus construction,
preflight, reopen and inverse checks; those must remain distinct from the
query and publication intervals before assigning CPU cost or selecting an
optimization. Generic providers still require fresh complete fingerprints;
the owned immutable memo must not be extended on version tokens alone.

## Validation and reproduction

Base source is `d2ce19788234036e1db30a7f3cf682af1fb77c7d`; the only code
change is the test module in `tools/perf-baseline/src/lib.rs`. Its final SHA-256
is `23ce48a33ffb9afd05571738642c78a3511338d34ce361f75ab3a710bfd74ab9`.
The exact patch, emitted JSON, combined Cargo/test log, toolchain and binary
identities are retained under `results/change-0927/receipts/`. The receipt
records the actual coder command, distinct from the earlier planning record.
The debug test binary remains outside the checkout; it is not performance
build evidence.

```sh
cargo test --manifest-path tools/perf-baseline/Cargo.toml \
  docx_section_layout_read_diagnostic_conserves_full_and_short_phases -- --nocapture
cargo test --manifest-path tools/perf-baseline/Cargo.toml \
  'docx_section_layout_' --lib
```

The focused diagnostic and all three section-layout tests pass. The existing
section-layout test retains its generic negative-case controls; the new
diagnostic does not claim new route-local negative coverage. Source formatting
and source/documentation diff checks pass; hash-bound raw logs, patches and
the archived reviewer script retain their original whitespace. Strict library
Clippy passes. Including tests triggers the
unchanged `selected_route_tests.rs` `result_large_err` warning; the same check
passes with only that lint allowed. This limitation is not a clean strict
tests-inclusive Clippy result.

Independent offline review verifies the emitted counters, checked sums,
source/output identities and preservation flags. A root execution of the
copied binary also reports one passing test and emits identical JSON. Its
wrapper initially expected the marker at the start of a line, whereas the
single-threaded test runner prefixed it with the test name. Offline parsing
of that same stdout succeeds; no replacement run was made. `root-replay.json`
records the wrapper failure and the child exit code that was not persisted.
The primary successful coder run remains separately bound.
