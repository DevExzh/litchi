# 0362ul94hb — XLSX join classifier coverage follow-up

This source-only follow-up is composed after 0362ul94ha at
`dc530b726b440d4b7ec7426f9b80981df9a7ad98`.  It closes the coverage return
without adding a public constructor, test-only export, compatibility alias, or
dependency from the XLSX crate to `litchi-ops`.

## Layered coverage

The private `litchi-ops` join-source classifier is now tested with both typed
operation-boundary leaves: `SourceChanged` maps to `io_error` and `Internal`
maps to `internal_error`.  The existing public cross-snapshot edit test still
constructs a real `JoinError`, converts it through the XLSX and unified error
boundaries, and verifies the no-source `corrupt_file` fallback.

The XLSX transaction test constructs the existing crate-internal `JoinError`
shape for each operation-boundary leaf.  It verifies `normalization_source()`,
the `JoinError` standard-error source, and the `litchi_xlsx::Error::from`
wrapper all retain the same typed leaf.  This test is about source-chain
preservation; it does not claim that an ordinary public workbook operation can
manufacture an internal invariant failure.  Together the tests cover both new
classifier branches and the actual public wrapper path without changing the
production API.  At this point the delegated retained-source cases were not
yet covered; record 0362ul94hc adds those cases and corrects the historical
fallback boundary.

## Validation status

No Cargo, build, rustdoc, runtime, or performance command was run.  The
coordinator must rerun the composed OOXML consumer and ops gates.
