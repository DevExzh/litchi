# 0110 — XLSX alignment writes

The unified cell/range writer now admits horizontal and vertical alignment,
wrap text, shrink to fit, text rotation, indentation and reading order. These
properties use the existing native SpreadsheetML alignment types through a
new `AlignmentChange` request and worksheet alignment method.

Native transactions combine font and alignment requests for each cell before
allocating shared stylesheet records. Payload, font and alignment assignments
therefore publish together with one native commit and retained patch history.
Alignment-only writes retain the font resource and its apply flag. Joined
transactions permit disjoint alignment fields and font/alignment changes,
while overlapping assignments and competing style replacement remain conflicts.

The facade accepts OfficeCLI's documented aliases and checked domains:
rotation is the native wire value 0–180 or 255 (stacked text), indent is 0–255,
and reading order accepts context/ltr/rtl plus their numeric and expanded text
aliases. Context clears the authored reading-order attribute, matching
OfficeCLI. Null clears any selected alignment attribute. Existing readback
omissions for false wrap/shrink, zero rotation/indent and contextual reading
order remain unchanged. Native callers can explicitly author these defaults.

Native contextual rotation 254 remains a Transitional-only capability; the
facade rejects it for OfficeCLI parity. The native editor must refuse authoring
that value into Strict SpreadsheetML. Horizontal `centerAcrossSelection`
normalizes to the existing `centerContinuous` token.

The native stylesheet editor preserves unrelated attributes, children,
namespaces and quote delimiters. Semantic repeats remain exact no-ops;
explicitly disabled apply-alignment flags are activated when an assignment
requests a value. The same limits, immutable source fences and batch barriers
used by font edits apply here. Other alignment attributes and broader cell
formatting remain outside this checkpoint; full goal and performance
acceptance remain open.

Validation to date:

- All 975 combined DOCX/XLSX/XLSB/encryption tests and doctests passed offline.
- Both plain and encrypted capability runs passed seven tests. The OfficeCLI
  delta was regenerated and checked; all 14 generator tests passed.
- Native cell-edit tests passed all 37 cases, including combined font/alignment
  edits, disjoint joins, overlap conflicts, complete-style replacement,
  reset/remove ordering and inverse patches.
- Stylesheet regressions cover apply-flag activation, opaque markup and quote
  preservation, malformed flags, extension ordering, Strict/Transitional
  rotation policy and repeated exact no-ops.

Final native runs passed all 48 stylesheet and 37 cell-edit tests. The new
exact-budget regression includes all seven alignment attributes and a long
producer namespace prefix. Initial Clippy identified unnecessary clones of
`Copy` alignment requests; these were removed. Final warning-denied all-target
Clippy passed. All 45 facade cell-write tests passed again after the final
native refactor, and the new native `AlignmentChange` rustdoc example passed.
Scoped formatting and whitespace checks passed.

The 464-path artifact audit found no oversized evidence or compiled binaries;
both mandatory staged-file checks printed nothing.
The rewritten base remains an ancestor; no rebase replay is needed against
that current base.


The borrowed Clippy fixture, root Cargo.lock, Python cache, temporary logs
and 22.6 GiB of build outputs were removed. No temporary worktree was created.

Staging failed because shared Git metadata is read-only: Git could not create
`/home/zhuhe/code/litchi/.git/worktrees/unified-ops-api/index.lock`. No commit
was created. Verified changes remain in the working tree; the full goal
remains active.
