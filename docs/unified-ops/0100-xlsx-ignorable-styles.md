# 0100 — Preserve ignorable stylesheet extensions during font edits

XLSX bold and italic assignments now accept ordinary producer stylesheets
with `mc:Ignorable`, including Excel's `x14ac:knownFonts` metadata and opaque
font/XF extensions. The native span editor retains the original XML and
clones resources in their original collection scope. It does not publish a
preprocessed or canonicalized stylesheet.

The shared MCE admission proof establishes agreement between source bytes
and the semantic reader's processed view before editing and again before
publishing new resources. Both calls use the OOXML baseline profile with
bounded input, projected output, depth, namespace bindings, attributes and
directive tokens. Direct font flags remain indexed once. Existing resource
deduplication, exact repeated-write no-ops and reversible part changes remain
in place.

Alternate content, directives such as `ProcessContent`/`MustUnderstand`,
ignorable subtrees that hide understood content, and unproven nested scopes
still return typed refusals. Namespace declarations containing references
also refuse: the raw span scanner resolves literal bindings, while MCE
decodes them. This prevents the two layers assigning different font indexes.
Empty XML elements now obey the same depth limit as paired start/end tags.

Review also corrected insertion of missing bold/italic children. New flags
follow the ordering used by the [Microsoft Open XML SDK font schema](https://raw.githubusercontent.com/dotnet/Open-XML-SDK/main/data/schemas/schemas_openxmlformats_org_spreadsheetml_2006_main.json),
without reordering existing producer children. Insertion anchors are indexed
once, and flags inserted at the same offset retain bold-before-italic order.

This narrows the stylesheet-profile capability gap from record 0099. Missing
style-table creation, additional font and formatting properties, inherited
formatting, generated-resource compaction, native command coalescing and the
remaining all-format/Python/CLI/performance goal are still required. No
performance improvement or full goal completion is claimed.

Validation on the final code:

- Native stylesheet codec: 16 tests passed (cargo test -p litchi-xlsx --lib raw::styles_edit).
- Native font integration: 11 tests passed (cargo test -p litchi-xlsx --test font_edits).
- Unified DOCX/XLSX/encryption suite: 859 tests/doctests passed (cargo test -p litchi-ops --no-default-features --features docx,xlsx,encryption).
- Plain XLSX capability snapshots: 2 tests passed. Capability delta generation/check and its 14 Python tests passed.
- Clippy passed for litchi-ops and litchi-xlsx with DOCX/XLSX/encryption, all targets and warnings denied.
- Scoped rustfmt and git diff --check passed. Cargo commands ran offline.

The all-targets check required the existing base-checkout LibreOffice fixture
tdf169496_hidden_graphic.xlsx; it was copied temporarily and removed after
validation. Temporary logs, generated root lockfile, Python cache and build
outputs were cleaned. No temporary worktree was created.

The rewritten base remains an ancestor of this branch. Commit staging was
attempted but the shared Git worktree metadata is read-only (index.lock);
no commit was created. The full goal remains active.
