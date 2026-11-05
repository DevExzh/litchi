# 0362ul89ab — PPTX test-only Clippy return

The isolated seven-row qualification on
`2da172ece2298107daa4feebe9b514b30eabe509` passes formatting (393 paths,
2.835 seconds) and the minimal all-target compiler check (58.533 seconds).
Strict all-feature/all-target Clippy clears the previous 76 diagnostics and
then stops on three test-only diagnostics (81.023 seconds, exit 101).
No runtime or glossary row runs.

The return moves `bounded_output_tests` after the production items without
changing its contents, and replaces `19.min(1024)` with the identical value
`19` in the changes and revision policy assertions. Production logic,
selected limits and assertions remain unchanged.

Retained external state:
`(deleted-coord-artifact)`.

- `rows.json`: 35,905 bytes, SHA-256
  `fc3a80c91ef8c91b8a5c52fed521c03578e2779ea090b74a75a935b533545381`.
- `result.json`: 37,806 bytes, SHA-256
  `c9eb8ac7323e4ed0fc8a6516328ca9115e9de8e8ef3deb69229385de0803785e`.

Root verifies the exact row commands, source/three-lock fences, log/trace
hashes, target guards and terminal ownership samples. The persistent session
terminates with exit 1; the target ends at 307,178,232 bytes. This follow-up
has source formatting and diff checks only; fresh qualification is required.
