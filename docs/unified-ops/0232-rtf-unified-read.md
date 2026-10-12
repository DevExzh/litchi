# 0232 — Unified RTF paragraph reads

RTF now dispatches `root`, `get`, `find`, `query`, child iteration and generic
node views through a concrete adapter. Root/body/paragraph paths use Word-family
semantic spelling, one-based ordinals and `last()`, without OOXML wire identities.
The adapter also supplies text, annotated, outline and statistics document views.

Nodes hold borrowed native `text::Paragraph` values. Contiguous paragraph text
is borrowed; fragmented text is materialized only when requested. Child/query
walks retain the native iterator, avoiding a second paragraph model or repeated
prefix scans. Direct positional lookup follows the native iterator to the
requested position. Structured JSON builds its requested output directly;
text views stream native paragraph displays.

Paragraph boundaries use native structural events, so inline line breaks remain
inside a paragraph and explicit empty paragraphs survive. The handle's paragraph
count now uses the same cached native count rather than the older flattened-text
projection. This fixes disagreements on inline breaks and explicitly empty
paragraphs, including the retained `chtoutline.rtf` fixture.

The declared text property is wired into query comparison while remaining a
node text member rather than a duplicated formatting entry. Capabilities report
read support and explicit pending runs, tables, headers/footers, notes, comments,
pictures/shapes, hyperlinks, styles and numbering. These concepts are not
reported as absent from RTF. Unified mutations remain pending.

## Verification

The all-features facade suite passes 1,626 tests/doctests. After the final
paragraph-count alignment, the full RTF-only suite passes 491 tests/doctests.
Six dedicated adapter tests cover paths and missing targets, typed nodes,
borrowed text, fragmented text predicates, lazy limits, child order, structured
views, capabilities, unsupported-substory classification and native paragraph
parity across four retained corpus fixtures. The existing byte-route test now
compares with the canonical native Package rather than its raw legacy model.

All-target warning-denied Clippy passes in full and RTF-only configurations;
warning-denied rustdoc and no-feature compilation pass. Crate boundaries pass
for 66 packages; 36 Python checker tests pass; the OfficeCLI capability delta is
unchanged. This wave makes no runtime performance acceptance claim.

The shared dispatch changes are committed as `9b2c3c65c1`. The owned 0232 agent
worktree/branch and its build output were removed. Root validation uses external
copied locks, leaving all three repository lockfile hashes unchanged; its owned
temporary artifacts are removed after qualification. No binary or large evidence
is committed.

## Remaining work

Unified RTF write dispatch/history, additional native node/property projections,
remaining native publication/redaction vocabulary and all other full-goal
requirements remain active. This paragraph adapter is one read milestone, not
all-format completion.
