# DOCX Ink semantic projection validation

This batch selects recognized contexts, traces, brush properties and source/destination
links before typed decoding. Ignored ancestry remains source-preserved and absent
from metadata; XML structural and resource limits apply to all content. Local
reference closure rejects duplicate, dangling, external and cross-owner references.
The generic DrawingML reader retains its existing semantic behavior.

Base: `6da91778c4003b6c9819e5c25168cfe9a847359e`. The eight source hashes and
candidate identity are in `source-manifest.json`; exact root commands, dependency
lock and compressed-log hashes are in `receipt.json`. Root validation passed
1,666 tests and 77 doctests (1 and 31 ignored), strict Clippy, warning-denied
rustdoc, pinned formatting and diff checks. Independent final review is clear.

Full imported trace lexical syntax, EMMA first-child/mode ordering, broader brush,
context/link/channel semantics and native Office acceptance remain unimplemented
or unverified. This batch does not claim complete MS-ODRAWXML conformance. Source
lexicals and opaque extension bytes remain exact; no rendering or execution occurs.
