# PPTX slide name and visibility writes

The unified API now sets slide `name` and `hidden` through paths, queries,
transactions and batches for eager and source-backed presentations. Property
coercion, duplicate detection, target expansion and staged-byte admission happen
before publication. Names accept text, including XML character references for
tabs and line endings; clearing a name sets an empty authored name. Hidden
requires a boolean. The native primary transaction exposes `set_slide_name`
and `set_slide_hidden`; eager and focused source edits also expose a combined
`set_slide_properties` operation.

The shared lexical codec edits only `p:cSld/@name` and `p:sld/@show`.
Missing `show` means visible, and new boolean values use producer spellings
`0` and `1`. Semantic no-ops retain exact bytes. Source edits preserve
unrelated XML, relationships and parts. Eager edits refresh the staged name
index before publication, so later name selectors see renames and ambiguity.
Eager commit compaction now retains attribute character references for tab,
newline and carriage return; previously compaction could change their values
by emitting literal XML attribute whitespace.

## Durable proof

The explicit source v4 envelope now admits these two managed attributes along
with the existing top-level shape removal grammar. Older v2/v3 payloads retain
their narrower shape-only grammar. A rooted, namespace-aware scan validates
managed values, duplicate attributes and closing element names before omitting
the two managed attributes from a proof skeleton. Matching skeletons prove a
slide-only edit. Mixed edits normalize those attributes and pass through the
existing text, geometry, addition and removal proof. Retained unknown
attributes and unrelated shape names remain protected. Exact source preimages
and the existing topology checks still apply.

The envelope is a grammar version, not authentication. A transported patch can
express any admitted edit against its matching source. Tests reject stripping
the v4 envelope, malformed managed booleans and unrelated shape-name changes.

## Validation

The complete native run passed 1,717 tests/doctests and the combined
DOCX/XLSX/PPTX/encryption facade run passed 1,122. After the final parser
hardening and lint simplification, 28 native slide-property/removal regressions
and six codec tests passed again. These cover exact no-ops, escaping and
whitespace round trips, staged name lookup and ambiguity, mixed geometry and
removal, ordering, replay/inverse/reopen, unknown-part preservation, strict
coercion, atomic/best-effort batches, budget pressure and cancellation.

Native, facade and harness all-target Clippy passed with warnings denied.
Facade rustdoc, minimal-feature compilation, scoped formatting and diff checks
passed. Crate boundaries passed for 66 packages and 268 internal dependency
declarations, and the generated OfficeCLI delta is current.

The full-feature PPTX rustdoc glossary check passed all 70 bans with warnings
denied. As in 0161, it used a temporary copy of the workspace lockfile with
cached `cc 1.6.0`; the inherited `cc 1.7.0` was unavailable offline.
The tracked workspace lockfile was preserved.

## Paired measurements

[Machine-readable evidence](0162-pptx-slide-properties-results.json) retains
every sample, generated-corpus hash, source digest and binary provenance. The
harness edits one slide name and hidden state using two native setter calls
and one commit, compared with one prepared unified set call. Three ABBA rounds
provide six samples per leg, pinned to CPU 0. All ten cases produced exactly
equal native/unified bytes; the failure matrix is empty.

| Shapes in slide | Eager unified/native | Source unified/native |
|---:|---:|---:|
| 1 | 1.042 | 1.014 |
| 10 | 1.004 | 1.005 |
| 100 | 1.004 | 1.002 |
| 1,000 | 0.994 | 1.003 |
| 10,000 | 0.997 | 1.002 |

At one shape, native/unified medians were 308.1/321.1 microseconds for eager
owners and 80.7/81.8 microseconds for source owners. Ratios below one are
measurement variation, not evidence of a faster facade. This measures warmed
immutable owners and excludes opening, serialization, allocation/RSS
measurement and real-world decks. It compares the two-setter native sequence,
not the eager combined setter. Broader performance acceptance, including
optimal native batching, remains a follow-up.

The broader all-format CRUD, Python, performance and ADR migration goal remains
open. No all-format completion claim is made by this milestone.


## Integration and cleanup

The rewritten `feat/office-format-completeness` base is already an ancestor.
Git metadata became writable during this turn, so this commit includes the
previously verified but uncommitted 0159–0161 implementation and evidence as
well as this milestone. Their historical blocked-commit notes describe the
state at those earlier turns.

Temporary builds, raw logs, measurement scratch and the three owned borrowed
LibreOffice fixtures were removed after verification. The reduced JSON retains
all timing samples and the removed binary's name, size, SHA-256 and regeneration
command. No worktree was created for this milestone; existing worktrees owned
by other work were preserved.
