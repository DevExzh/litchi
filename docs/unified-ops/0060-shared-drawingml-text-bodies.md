# 0060 — Shared DrawingML text-body reads

## Scope and provenance

Integrate L-PPTX-R7's shared local text-body projection and its PPTX and
SpreadsheetDrawing consumers. Source-only candidate `9ca944f6c8` supplied the
starting implementation; its history and evidence were not merged. This wave
removes the candidate's legacy parsing profile, adapts native identities to
explicit raw access, and strengthens refusal, limits and authoring behavior.
No iWork source is changed.

The full unified-operations goal remains active. Presentation related-part
reads and unified PPTX dispatch, all-format writes, Python, complete API
migration and paired performance acceptance remain required.

## Shared projection and strict policy

`text::body` owns body properties, paragraphs, runs, breaks, cached field text,
local paragraph/run properties, fonts, spacing, tabs, bullets and hyperlinks.
Existing body-property schema defaults remain explicit; paragraph/run values
remain local and optional. List, theme, master and placeholder inheritance are
not resolved by this projection. SpreadsheetDrawing now uses the same event
builder as PPTX instead of its separate run reader.

The shared reader has one strict policy. Invalid modeled scalar values,
duplicate/competing singleton properties, malformed field GUIDs, ambiguous
relationship namespace aliases, undeclared relationship prefixes and
Strict/Transitional relationship mismatches are refused. There is no legacy
profile that drops fields or reads invalid properties as absent.
`read_with_options` accepts inherited namespace context and finite limits.
The streaming builder checks its document byte bound; element, depth and
retained-text bounds also apply. Generated break/tab characters count toward
the retained-text bound. SpreadsheetDrawing supplies a finite element budget.

Native field and relationship IDs are private model data accessed through
`text::body::raw`. Authored bodies, properties, fields, fonts, hyperlinks,
picture bullets and builders have content-free Debug output. Shared text
errors retain typed error categories without authored content; host text
errors preserve the DrawingML source. Existing broad source/semantic budget
vocabulary migration and unrelated legacy diagnostic families remain open.

## Lazy PPTX access and ownership

`Common::text_body` and `Shape::text_body` read the shape's direct body on
demand; duplicate direct bodies fail. Its physical span is available only
through `shape::raw::text_body_span`. Retained scenes and handles carry exact
scene limits into lazy decoding. Namespace declarations on roots, trees,
groups, shapes and the body are respected, including BOM and MCE output.

For bodies inheriting only root and own-shape bindings, namespace preparation
reads only those start tags. Group/mid-tree declarations use a bounded prefix
scan. This is not a general constant-time namespace index; pathological
repeated reads with mid-tree declarations still need performance evaluation.
Retained scene cache accounting includes the added record/data sizes.

## Authoring and preservation

Shared `writer::{write, write_contents}` now return typed results and validate
before touching output. Supported breaks serialize as `a:br`; a break is not
silently rewritten into ordinary text. Fields, unsupported local properties,
and incomplete projections are refused atomically.

`Properties::has_unmodeled_content` records skipped list-style/effect/extension
markup and unmodeled attributes. This makes a local inspection projection
non-authorable instead of silently losing its source content. Spreadsheet
authoring validates the marker before emitting anchors. The spreadsheet
reader also retains its existing opaque shape markup; reading a field now
keeps both its semantic cached value and its original opaque representation.
This does not add arbitrary editing support for those opaque features.

## Verification

- All-feature PPTX/XLSX suite: 3,628 tests passed. PPTX contributed
  895 unit, 473 integration and 10 doctests (1,378 total); XLSX contributed
  2,250 passing tests. The long XLSX facts-oracle corpus check completed.
- Final all-feature DrawingML/SpreadsheetDrawing suites: 428 tests passed,
  including 206 DrawingML and 29 SpreadsheetDrawing unit tests.
- Final PPTX property/body/owned-source parity regressions: 18 passed.
  Real fixture parity compares body projections as well as existing shape
  properties. Six dedicated body tests cover namespace placement, groups,
  BOM/MCE, duplicate direct bodies, retained owners and exact caller limits.
- After review fixes, focused XLSX package-shape tests: two passed.
- All-feature/all-target Clippy for DrawingML, SpreadsheetDrawing, PPTX and
  XLSX passed with warnings denied.
- Workspace library compilation passed.
- DrawingML glossary profile passed (five bans); both PPTX profiles passed
  (21 bans each). Glossary checker Python tests: 21 passed.
- Review strengthened empty-color/incomplete-body refusal and redacted writer
  spec Debug output. Action-only hyperlinks have a regression proving action
  and tooltip survive an absent relationship. Unsupported autofit scale
  attributes also prevent lossy reauthoring.

Full format-policy/semantic-budget migration is not claimed. The shared
fragment reader retains finite codec-specific limits; it does not implement
the eventual format-wide SourceLimits/SemanticLimits option envelope or
incremental execution-context charging. Existing package readers and source
fences continue to own physical ingress. Paired performance acceptance is
still pending.

## Commit and cleanup

Final warning-denied rustdoc for DrawingML, SpreadsheetDrawing and PPTX
passed. Changed Rust sources were formatted and `git diff --check` passed.
The changed/untracked artifact audit passed; both mandatory staged-file checks
printed nothing.

No temporary worktrees were created, and participating agents reported no
additional scratch files. The four base-checkout fixtures needed by existing
PPTX/XLSX full-leaf tests were removed along with only their newly created
directories. Generated lockfiles and Python cache were removed.
`cargo clean` removed 50,255 files (81.6 GiB); turn scratch was removed.

The local base tip remains `32ad64e3207655b4ba558fc2793d9a7396655cdf`, already
an ancestor of HEAD. `git add -A` failed again because
`/home/zhuhe/code/litchi/.git/worktrees/unified-ops-api/index.lock` is on a
read-only filesystem. No commit is claimed. The stale partial index must be
refreshed before an eventual commit; records 0051–0060 remain working-tree
changes.
