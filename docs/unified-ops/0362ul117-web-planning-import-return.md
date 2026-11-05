# 0362ul117 — remove the obsolete Web Extensions import

The joint compiler packet on `11be7c2e19b8520a61447df968eb184e173e591d`
passes its 508-file formatting row, then minimal PPTX compilation rejects an
unused `BytesStartExt` import in common Web Extensions planning. Record108a's
conformance inspection now reads the already parsed `Node` attributes, so
the raw event attribute extension trait is no longer used in this module.

Remove that import. Parsing, relationship-profile selection, signed no-op
behavior, limits and all consumers remain unchanged. No lint suppression is
introduced.

Root verifies the failed terminal capture before editing. The external packet
is `(deleted-coord-artifact)` under the
worktree parent. Its `rows.json` is 38,321 bytes, SHA-256
`677803d5b07a7a37c6708dd9e6443837c2fce33c1193049d05a2c3bf93b1d4ed`;
`result.json` is 40,248 bytes, SHA-256
`90ce7e5177361ff1ec23029085e1c2154c3e49da7ed59b424812d873198d69a2`.
Formatting takes 3.780 seconds; compilation exits101 after 35.874 seconds,
with 132,010,477 bytes in the target. The other six rows do not run.

The edited file passes pinned workspace rustfmt and `git diff --check`;
commit size and binary guards pass. Fresh compiler qualification remains
required, followed by runtime/glossary and final workspace/performance gates.
