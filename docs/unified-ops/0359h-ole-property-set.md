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
