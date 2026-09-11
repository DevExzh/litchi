# Chart formatcode2 implementation boundary

The audit names the exact `http://schemas.microsoft.com/office/drawing/2015/06/chart`
namespace in MS-ODRAWXML §2.44. Both the global `formatcode2` element and the
qualified `formatcode2` attribute use `ST_Xstring`. They carry custom number
formatting information with language/culture information. The namespace schema
defines no complex types and supplies no universal containing chart owner.

The implementation must distinguish XML lexical text from the escaped-string
semantic domain. Successful XSD validation alone cannot prove escaped-string
decoding, lossless editing, owner placement, or application acceptance.

ECMA-376 §22.9.2.19 requires escaped UTF-16 code-unit decoding and protection
of literal escape-looking text. MS-OI29500 §2.1.1747 adds context-sensitive
Office output rules: CR is escaped; tab/LF remain literal in element content
and are escaped in attributes. Attribute XML normalization and numeric
references must be distinguished before decoding the semantic string. Tests
must exercise these rules independently of the XSD string restriction.

Shared code should own bounded checked values and source-preserving XML
operations. Package relationships, chart selection, transactions, and source
publication remain with their format owners. An element-only fragment owner
must not be reported as complete element-and-attribute coverage. A host path
must be established from normative grammar or a documented compatibility
profile before automatic placement is added.
An attribute start-tag helper must accept caller-supplied inherited namespace
bindings, preserve the selected raw source, and distinguish the qualified
attribute from an unqualified lookalike. A default namespace does not qualify
an attribute name.

The offline validator extracts MS-ODRAWXML §5.42 and resolves its imported
`c:ST_Xstring` through the equivalent ECMA shared simple type. This is an
explicit schema packaging adaptation, not an invented format-code grammar.
It accepts generated element XML and reports the vendored schema/archive hashes.
Run it with `--spec-root` when an isolated worktree lacks the untracked local
ECMA archive. It does not claim qualified-attribute host placement validation.

A bounded root corpus search inspected chart XML members no larger than 2 MiB:
25 chart members across 179 `test-data/ooxml` packages, and 1,117 chart members
across 4,721 `3rdparty` OOXML packages. The latter scan skipped 61 unreadable
packages. Neither scan found the exact `formatcode2` token. This negative search
is discovery evidence only; it does not establish absence from all files or
native application support. No Office round-trip claim follows from it.

The required verification layers are semantic regression tests, malformed XML
and resource-boundary tests, source-preserving no-op/scalar-edit tests, generated
element XSD checks, independent review, and source-bound allocation/runtime
measurements. Native application acceptance remains unproven without a native
fixture and a real application round trip.
