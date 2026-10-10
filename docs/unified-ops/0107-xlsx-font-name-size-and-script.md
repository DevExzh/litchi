# 0107 — XLSX font name, size and script writes

The unified cell/range writer now admits `font.name`, `font.size`,
`superscript` and `subscript` through the static property registry and native
worksheet font edit. These properties share the immutable commit, source
fences and native patch history with companion payload assignments.

Font names retain owned text through the native change value. Admission
charges their bytes against semantic staging limits before ownership is
allocated. The native stylesheet editor changes source spans, preserving
unrelated attributes, namespaces and opaque markup; decoded name comparison
keeps semantically identical names as exact lexical no-ops.

Font sizes use the existing font-size length grammar: bare points or a `pt`
suffix, with the shared minimum of 0.5 points and XLSX maximum of 409 points.
Null removes the authored size or name. Superscript and subscript address one
native script property: true selects that script, false authors baseline,
and null removes the element. Supplying both properties in one assignment
refuses atomically, rather than choosing a winner. Ordinary readback continues
to omit inactive script flags.

Native font changes now retain name text and therefore clone shared ownership
instead of implementing `Copy`. Staged merge, conflict detection and stylesheet
request deduplication use the same native typed value. Font edits remain
barriers to payload-only coalescing. Additional formatting, resource compaction,
full performance acceptance and broader format/Python/CLI coverage remain open.

Source attribute delimiters are preserved when changing names. The bounded
escaping helper handles both quote kinds and writes CR/LF/TAB as character
references, so single-quoted producer attributes and Unicode/control-containing
names round trip. Repeated semantic names preserve their original lexical
encoding. Authoring limits do not prevent unrelated edits from retaining
source font sizes outside the supported authoring range.

Validation:

- 924 combined DOCX/XLSX/encryption facade tests and doctests passed offline.
- Final native runs passed all 26 stylesheet and 28 font-focused tests after
  the single-quoted attribute fix, including escape expansion budgets, source
  preservation, checked authoring, ownership, joins, replay and inverse.
- All 29 cell-write regressions passed again after that final native change.
- Both encrypted and plain capability snapshots passed their six-test suites.
  The OfficeCLI delta was regenerated and checked; all 14 generator tests passed.
- Independent review confirmed quote-safe rewrites, native size bounds and
  bounded name ownership, with no further actionable blocker.

The rewritten base remains an ancestor of this branch, so no rebase replay
is needed against the current base.


Warning-denied all-target Clippy passed for ops and XLSX with combined
DOCX/XLSX/encryption features. Scoped formatting and whitespace checks passed.
The 460-path artifact audit found no oversized evidence or compiled binaries;
both mandatory staged-file checks printed nothing.


The final native bounds regression passed after the lint-only range-expression
adjustment. The borrowed Clippy fixture, root Cargo.lock, Python cache,
temporary logs and 19.1 GiB of build outputs were removed. No temporary worktree was created.

Staging failed because shared Git metadata is read-only: Git could not create
`/home/zhuhe/code/litchi/.git/worktrees/unified-ops-api/index.lock`. No commit
was created. The verified changes remain in the working tree; the full goal
remains active.
