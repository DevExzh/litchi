# 0106 — XLSX strike-through and underline writes

Cell and range assignments now admit `strike` and `underline` through the
static XLSX property registry. They use native `FontChange` operations in the
same worksheet edit as companion payloads. Range and selector bounds, source
fences, native patches and atomic publication apply unchanged. These font
assignments close a pending payload-coalescing run.

`strike` accepts the shared boolean vocabulary and the aliases `strikethrough`,
`font.strike` and `font.strikethrough`. `underline` accepts `single`, `double`,
`singleAccounting`, `doubleAccounting` and `none`, case-insensitively. Its
`font.underline` alias and textual boolean/short-token aliases normalize in
the registry. Unknown tokens refuse rather than silently disabling underline.
The underline property remains an enum input. The shared property coercer
accepts typed booleans through the same true/false aliases.

Null removes the authored property. False strike and explicit underline `none`
remain authored native values, consistent with the existing bold/italic
false-versus-null distinction. OfficeCLI instead removes underline for its
false/none inputs; this adapter preserves the native distinction. Existing
unified readback continues to omit disabled underline, while native style
inspection retains the explicit value.

The leaf editor extends its source-span font index and ordered insertions;
it does not regenerate the stylesheet from a second semantic model. Existing
font and cell-format resources remain intact, and matching resources are
reused. Namespaces, opaque markup, MCE admission, output budgets, `applyFont`
activation, staged property merging and conflict checks remain native concerns.
Resource compaction, additional formatting and full format/Python/CLI coverage
remain open; no performance acceptance claim is made here.

Review also exposed a pre-existing stylesheet document-envelope gap. Font
editing now refuses text, CDATA or references outside the root, declarations
that are not the first event after an optional BOM, and DTDs. Only XML
whitespace is accepted around the root; valid comments and processing
instructions retain their original bytes.

Validation:

- 919 combined DOCX/XLSX/encryption facade tests and doctests passed offline.
- Final native checks passed all 20 stylesheet tests and 22 font-focused
  tests, including tri-state properties, source-preserving clearing, native
  joins/conflicts, exact no-ops, replay/inverse and malformed envelopes.
- The initial broad native run completed with 1,500 passed and two failing
  new lexical assertions. The assertions/output convention were corrected;
  both failures are included in the passing final focused runs above.
- Both encrypted and plain capability snapshots were regenerated and passed
  their six-test suites. The OfficeCLI delta was regenerated and checked;
  all 14 delta generator tests passed.

The rewritten base remains an ancestor, so no rebase replay is required.
Full performance acceptance and the wider API goal remain open.


Final cell/range/coalescing regressions passed 49 tests after the envelope
change. Independent review confirmed the corrected document-envelope checks
and found no further actionable defect in the formatting delta.


Warning-denied all-target Clippy passed for ops and XLSX with combined
DOCX/XLSX/encryption features. Scoped formatting and whitespace checks passed.
The 459-path artifact audit found no oversized evidence or compiled binaries;
both required staged-file checks printed nothing.

The borrowed Clippy fixture, root Cargo.lock, Python cache, temporary logs and
18.4 GiB of build outputs were removed. No temporary worktree was created.

Staging failed because the shared Git metadata is read-only: Git could not
create `/home/zhuhe/code/litchi/.git/worktrees/unified-ops-api/index.lock`.
No commit was created; the verified changes remain in the working tree.
