# 0109 — XLSX remaining typed font properties

The unified cell/range writer now admits `font.outline`, `font.shadow`,
`font.condense`, `font.extend`, `font.charset`, `font.family` and `font.scheme`.
The cell read vocabulary exposes the same properties through the existing
native typed formatting model, including explicit false, zero and `none`.
The existing base-format/default-font read omissions remain in force.

The native `FontChange` retains typed assignments for each property. Null
removes the authored element; boolean false and scheme `none` remain authored
values. Native joins, conflict checks, replay/inverse and immutable source
fences apply to these assignments and their payload companions. Font edits
remain barriers to payload-only coalescing.

Family authoring accepts 0–5 and charset accepts 0–255. Both the property
registry and native builders enforce these domains. The local Open XML SDK
schema at `3rdparty/Open-XML-SDK/data/schemas/`
`schemas_openxmlformats_org_spreadsheetml_2006_main.json` defines those bounds
in `CT_FontFamilyNum/family` and `CT_ByteProperty/charset`, respectively.
Unlike OfficeCLI's raw long-tail child writer, unified authoring validates
these typed domains. Existing out-of-domain source values remain preserved
when unrelated fields change. Scheme uses the shared `none`, `major`, `minor`
token table. No undocumented bare-name aliases are introduced.

The stylesheet editor preserves unrelated XML and original attribute quote
styles. It compares decoded values for lexical no-ops and inserts properties
in the native font child order for both empty and nonempty source elements.
Broader cell formatting, format coverage and full performance acceptance
remain open.

The shared cell vocabulary appends the new keys, preserving existing XLSB
positional selections. XLSB still lacks native font projection in its unified
adapter; its capability report lists all seven new property gaps explicitly
with the existing font-projection follow-up.

Validation to date:

- 936 combined DOCX/XLSX/encryption facade tests and doctests passed offline.
- 36 native stylesheet and 37 font-focused tests passed, including authoring
  bounds, preservation of out-of-domain source values, ordered insertion,
  lexical no-ops, conflict handling and replay/inverse.
- Independent source-produced XML fixtures verify reads and bool/integer/scheme
  query predicates for eager and source-backed owners.

The final DOCX/XLSX/XLSB/encryption run passed all 967 tests and doctests.
Both plain and encrypted capability runs passed seven tests. The generated
OfficeCLI delta is current, and its 14 generator tests passed. Scoped
formatting and whitespace checks passed. Warning-denied all-target Clippy
passed with DOCX/XLSX/XLSB/encryption enabled.

Final independent review found two existing helper defects: acceptance of
uppercase XML boolean tokens inconsistent with the native parser, and an
overestimate of the byte count when expanding an empty font. The editor now
uses the same boolean tokens as the parser and computes the exact retained
prefix, child payload and closing-element length. After these corrections,
all 38 stylesheet, 38 font-focused and 39 cell-write regressions passed.

The rewritten base remains an ancestor of this branch; no rebase replay is
needed against the current base.


Final warning-denied Clippy and scoped formatting passed after the helper
corrections. The 463-path artifact audit found no oversized evidence or
compiled binaries; both mandatory staged-file checks printed nothing.

The borrowed Clippy fixture, root Cargo.lock, Python cache, temporary logs
and 29.6 GiB of build outputs were removed. No temporary worktree was created.

Staging failed because shared Git metadata is read-only: Git could not create
`/home/zhuhe/code/litchi/.git/worktrees/unified-ops-api/index.lock`. No commit
was created. Verified changes remain in the working tree; the full goal
remains active.
