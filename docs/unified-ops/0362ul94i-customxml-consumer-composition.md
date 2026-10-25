# 0362ul94i — Shared OOXML fidelity and consumer composition

This private candidate composes three independently reviewed source chains
onto `5ec78a48589c7ab4396ee90cb1f6d41cf7f1aaa6`:

- Custom XML 94c–94ce through
  `44e0be0633a28b01da8d762294ad06ed229c9412`, based on `0db7ff3a6a`.
- DOCX 94g/94ga/94gb through
  `5d8aec545a36bd10007b64915094b11e9454fd77`.
- XLSX 94h/94ha/94hb/94hc through
  `cef014a61f8ae19616c84f0c7770fcd7b06339bf`.

All patches apply without conflict. Their shared ops test file contains both
sets of consumer regressions. Source-only formatting normalizes the exact
composed preimages, retained outside the repository under
`.coord/ooxml-combined-0362ul94/composition-94i-format/`.

Custom XML preserves original prefixes, quotes and opaque markup while
applying exact bounded lexical edits. Selected MCE branches map directly to
source offsets, avoiding marker XML inflation and duplicate retained output.
Paired schema-reference indexing, early source limits and typed event/depth/
attribute errors are repaired. Missing/occupied/ambiguous authored targets
use fixed refusals; source and invariant failures retain their proper kinds.
DOCX graph-cleanup callers propagate relationship resolution errors.

The DOCX consumer moves the owned VBA payload without a copy and preserves
properties refusals, patch conflicts, source changes and internal errors.
`LimitKind::Properties(properties::LimitKind)` retains the exact family kind,
observation and maximum in both conversion directions. Its ops regression
covers all six current property limit dimensions.

XLSX join normalization retains its typed source. The unified classifier
now delegates every retained normalization error to the existing leaf
classifier; only a join with no source retains the ordinary corrupt-file
fallback. This includes refusal kinds, patch conflicts, typed limits and
nested source changes. Layered tests cover private classification and leaf
source chains; a public cross-snapshot join tests the real fallback path.
No public constructor or dependency cycle is added for testing.

Earlier returns remain documented: 94g lost limit details; 94ga projected
resource kinds inaccurately; 94h lacked wrapped ops classification; 94ha
only tested the ordinary fallback; 94cd still had two offset type errors.
Root caught and returned those gaps before this composition. Scoped source
reviews pass on the final immutable chains; compiler, Clippy, runtime,
glossary and performance qualification remain required.

The prior downstream capture on `5ec78a4858` passed formatting and OPC
compilation, then failed the four-format compiler row. Evidence remains at
`.coord/ooxml-combined-0362ul94-f5-consumers-5ec78a4858`:
`result.json` is 17,338 bytes, SHA-256
`c01eedfe86b84f78a6d302914a269fa99d2d7664fc4da5304a1275b77470331a`.
Root verified its complete command, source, three-lock, log/trace, guard and
terminal ownership receipts. This composition is not runtime acceptance.
