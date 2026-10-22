# 0359h — OLE property-set overlay limits and error boundary

- Date: 2026-10-21.
- Base: `604822b441` on `uo2/w5-ole`.
- Status: implementation and focused validation are complete on the slice
  branch; coordinated performance evidence remains pending a native
  measurement grant.

The inert user-defined OLE property overlay now has one required-limit shape:
`user_defined::Properties::new(&Section, Limits)` and
`user_defined::Edit::new(&mut Section, Limits)`. The removed `with_limits`
constructors and the old defaulting `new` forms have no aliases. Their public
fallible methods, `LimitsBuilder::build`, `LinkBase::new` and `Hyperlink::new`
use `litchi_ole_common::Error`, preserving CFB failures while classifying the
overlay budgets as `UserDefinedBlobBytes`, `UserDefinedLinks`,
`UserDefinedStringUnits` and `UserDefinedTotalUtf16Units`.

The property-set compound-file editor also requires an explicit
`OleFileLimits` in `property_set::codec::Edit::new`. DOC, PPT and XLS package
transactions, common codec tests, summary-property tests, tolerance tests and
the property-set integration tests pass that policy explicitly. The DOC,
PPT and XLS user-defined callback boundaries and DOC hyperlink adapter were
migrated to the shared OLE error in the same slice, so typed overlay failures
are not converted back to a lossy `OleError` boundary.

The overlay still validates before changing a section, writes only bounded
wire values, preserves unknown properties and retained hyperlink hashes, and
keeps malformed wire errors distinct from configured-limit failures. The
existing DOC hyperlink, property-set and common OLE suites remain the required
correctness checks. The OLE common suite passed 226 unit tests and all
integration targets; DOC, PPT and XLS all-target suites passed, including 1,314
XLS unit tests. All four package clippy checks passed with `-D warnings`, and
the API glossary checker passed 49 profiles, 3,302 banned names, 499 edit
families, 85 immutable owners and 30 family-method profiles. The workspace
format check still reports pre-existing formatting in untouched OLE stream
files; every Rust file changed by this slice is rustfmt-clean. No performance
result is claimed until a clean baseline and candidate measurement is
available.

## Driver preflight

The property routes use the standalone `ole_w5_probe` source at
`fdb139197d`. The paired source is `604822b441` (owner migration complete,
user-defined constructors still on the old boundary) → `fdb139197d` (typed
required-limit constructors). The `h` baseline and corrected candidate were
built in distinct absolute target directories under the capped W5 environment;
their binary SHA-256 values are respectively
`f5520c923c7ec1d763c96bd58e5ad837c08dcc31ee231727fa9ad17535e82e71` and
`44b3e0427841cbb0ba3fed146f7da67926f096f05cf266ea9f40ec71f9498700`.
The baseline uses the current owner with `legacy-property-api`; the candidate
uses the current API. Untimed one-sample preflights passed for `read` and
`edit` on both binaries, each preserving the six-property fixture result.

The quiet acceptance pair is `A1, B1, B2, A2`, with 20 warmups and 500
retained samples per route. Run `read` and `edit` with the `h` baseline as
`A` and the corrected candidate as `B`; retain each leg in a distinct CSV path
and record the source and binary hashes above. The native run remains pending;
the preflight outputs are outside the checkout at
`/tmp/w5-ole-preflight-20261021T164402Z`.
