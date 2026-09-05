# ODT facade test correction

Baseline `f22917bbe` exposed three stale ODT/polyglot assertions. The accepted
policy is content-derived: ADR 0006 says format selection is never inferred
from a filename, and ADR 0009 assigns packaged ODF detection to
`litchi-odf-common`. Existing detector coverage in
`detection_smart::{functions,detected}` already accepts a valid ODF package
with a malformed `[Content_Types].xml` extra member as ODT/ODS after the
bounded OOXML probe.

The failing facade assertions came from older history: the malformed-catalog
owned test was added by `2bf6a7447`, the filesystem counterpart by
`9b717bb08`, and the suffix-limit expectation by `463e0b99b`. The later
catalog arbitration work (`463e0b99b`, followed by `094d92601`) established the
current fallback and content-derived suffix behavior.

The corrections are limited to the two ODT facade test files:

- Owned malformed-catalog input now asserts `DocumentImpl::Odt` and reads the
  expected ODT text.
- Filesystem malformed-catalog input now asserts `DocumentImpl::OdtSource` and
  reads the expected ODT text.
- An ordinary ODT stored at a `.DOCX` path now asserts successful native ODT
  opening and the `bounded` semantic text. Both valid OOXML-catalog polyglot
  cases, including the case-variant catalog spelling, retain their DOCX
  input-limit assertions unchanged.

No production behavior or error boundary was weakened. No Cargo, build, test,
profiling, or CPU command was run while making this correction; the diff was
reviewed statically and checked for whitespace errors.
