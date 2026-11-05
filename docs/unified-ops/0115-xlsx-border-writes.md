# 0115 — XLSX border writes and typed local formatting

This checkpoint extends cell formatting through the existing native stylesheet
transaction. Border side, line, color and direction edits use canonical native
types; the unified adapter admits the OfficeCLI property vocabulary and delegates
resource allocation and lossless rewriting to the XLSX owner.

The native `BorderChange` uses the existing `Side`, `Line` and `BorderColor`
types. Whole-side replacement supersedes earlier fields; later line/color
changes can override its fields. Whole-side requests conflict with every
mentioned field on that side, while disjoint fields and sides compose. Physical
left/right edges follow Transitional conformance; logical start/end edges
follow Strict conformance. Native interior edges and outline remain explicit
capability follow-ups for the unified property vocabulary.

The facade supports side styles/colors, diagonal directions, `border` and
`border.all`, `border.color` and `border.all.color`, and dotted `.style` aliases.
Shorthands apply before per-side overrides regardless of input order. `none`
or null removes only the line attribute, preserving existing color and opaque
source markup. A non-null color without a resulting line is rejected atomically
instead of following OfficeCLI's silent discard. Null clears a color or flag.
Each range retains one native border change; only color-only operations that
need an existing line read the effective native source format.

The codec preserves source records and clones only changed resources. Empty or
absent border collections are authored lazily. Known children precede retained
future children, and added border records precede collection extensions.
Color removal uses source byte spans; opaque content that would be discarded
causes a typed refusal. Border property indexing takes bounded passes over the
element index and retains entries only for border records. Its memory is
checked before allocation. Publication remains subject to the existing output
and source-part bounds.

Semantic staging charges a conservative fixed border entry and lowering-copy
allowance per cell, combined with retained fill payloads. Staging and join
admission run before mutation; typed resource-limit failures retain prior and
rejected edits. Whole-style replacement, reset, removal and transfer release
the corresponding counters. This checkpoint makes no new performance claim.

`format::xlsx::Node::cell_format()` exposes a borrowed native local format from
the node's existing cache. This makes complete gradient geometry, ordered stop
positions and all native color kinds available without a second formatting
model or an allocated generic format map. The accessor retains source-version
checks on cached reads. Cells without an explicit local style, missing cells and
non-cell nodes return `None`; an explicit base-style reference returns a format
with `is_base() == true`. The ordinary OfficeCLI scalar projection retains its
documented omissions.

Initial validation corrected codec byte offsets, empty-record insertion,
whole-side/facet precedence, sparse index admission and source-quote test
expectations. Seven typed read-policy tests and all 71 facade cell-write tests
passed. The native library run passed 1,619 tests, including the corpus oracle,
and exposed two incorrect assertions (quote normalization and a font-only error
message). Their corrections passed in the final focused runs: 90 stylesheet
codec tests, 64 native cell-edit tests and 3 state-projection tests. Both new
native border examples passed.

All 1,012 combined DOCX/XLSX/XLSB/encryption tests and doctests passed. After
Clippy's two mechanical match cleanups, the 90 codec and 71 facade cell-write
tests passed again. Both encrypted and plain capability configurations passed
seven checks each, including the updated typed-format capability descriptions.
The regenerated OfficeCLI delta passed its check and all 14 generator tests.
Clippy passed for `litchi-ops` and `litchi-xlsx`, all targets, with warnings
denied and the DOCX/XLSX/XLSB/encryption feature configuration. The
`xlsx-minimal` glossary profile passed. These are scoped gates, not a claim of
full workspace acceptance.

Reproduction (offline):

```sh
CARGO_NET_OFFLINE=true cargo test -p litchi-xlsx --lib raw::styles_edit
CARGO_NET_OFFLINE=true cargo test -p litchi-xlsx --lib workbook::edit::tests::cells
CARGO_NET_OFFLINE=true cargo test -p litchi-xlsx --lib workbook::edit::tests::worksheet_edit_state
CARGO_NET_OFFLINE=true cargo test -p litchi-xlsx --doc border
CARGO_NET_OFFLINE=true cargo test -p litchi-xlsx --doc BorderChange
CARGO_NET_OFFLINE=true cargo test -p litchi-ops --no-default-features --features docx,xlsx,xlsb,encryption
CARGO_NET_OFFLINE=true cargo clippy -p litchi-ops -p litchi-xlsx --no-default-features --features litchi-ops/docx,litchi-ops/xlsx,litchi-ops/xlsb,litchi-ops/encryption --all-targets -- -D warnings
CARGO_NET_OFFLINE=true python3 tools/check_api_glossary.py --profile xlsx-minimal --target-dir target
```

The existing SVG reader test's `tdf169496_hidden_graphic.xlsx` fixture was
temporarily copied from the base checkout's
`3rdparty/libreoffice-core/sc/qa/unit/data/xlsx/` for all-target Clippy.
The 472-path artifact audit found no oversized evidence or compiled binaries.
The rewritten base is still an ancestor; no rebase replay is needed.

The next substantial tier-1 step is PPTX shape-text writes through its retained
native editors, including the required glossary/default-policy migration.
XLSB scalar cell writes follow. The all-format goal and its full performance
acceptance remain open.

Scoped Rust formatting and `git diff --check` passed. Both mandatory staged
artifact checks printed nothing. The requested staging attempt failed because
Git could not create
`/home/zhuhe/code/litchi/.git/worktrees/unified-ops-api/index.lock` on the
read-only filesystem; no commit was created.

Removed this checkpoint's temporary logs, generated build directory and root
Cargo.lock, Python bytecode cache, and borrowed LibreOffice fixture. No temporary
worktree was created.
