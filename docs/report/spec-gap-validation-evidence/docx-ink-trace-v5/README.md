DOCX Ink trace validation

Adds streaming validation of regular trace values against the active trace format, with channel-local difference state and source spans. It enforces Office integer-millisecond T values while retaining ignored intermittent metadata as opaque XML. It does not allocate a decoded point vector.

The six-path candidate is based on 42cce54f0. Root verified its exact Git closure and primary preimages, then ran the full all-features DOCX suite, strict all-target Clippy, warning-denied rustdoc, and scoped rustfmt. Counts and raw-log hashes are in root-validation.json. The preparer archive includes the full patch, focused tests, and external conformance harness.
