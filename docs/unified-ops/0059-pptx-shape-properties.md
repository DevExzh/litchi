# 0059 — Lazy PPTX shape properties and explicit raw payload metadata

## Scope and provenance

Integrate the shape-property portion of L-PPTX-R6 from candidate
`9ca944f6c8` as source-only changes, adapting its public shape to ADR 0034.
No candidate history or evidence was merged. The candidate's R7 text-body
reader and namespace reconstruction are deferred. The full unified API goal
remains active: related presentation parts, text-body projection, PPTX unified
dispatch, other formats, writes, Python and paired performance acceptance
remain required. Existing borrowed shape identities and scene selectors still
need their separate glossary migration. No iWork source is changed.

## Behavior and policy

Common/Shape reads expose lazy `NonVisual`, description, title, optional hidden
flag and typed DrawingML transform. Missing properties remain distinct from
explicit false/zero. Attribute decoding borrows unchanged text and normalizes
XML attribute whitespace and character references when needed. Transforms
preserve authored rotation, flips and group child coordinates, and reject
duplicate owned transforms, foreign children and invalid scalar values.

New physical metadata is available only through `shape::raw`: payload
relationships, their attribute/value spans, the non-visual start-tag span and
the transform span. Payload selection refuses competing embed/link attributes,
namespace aliases and payload elements using a content-free typed error.
Relationship namespaces must match the Strict or Transitional owner dialect;
undeclared relationship prefixes are refused. Foreign extension payloads do
not become a frame's payload. Mixed frame markers remain generic frames
with an ambiguity error instead of choosing one payload kind, including
inline table plus related-part payload combinations.

Common, concrete shape, scene, iterator, selector, non-visual and raw payload debug output omits authored
names, descriptions, IDs and XML. Borrowed lookup error Display/Debug also
omits authored names; its legacy name fields remain pending API migration. Transform error messages are redacted in the
shared DrawingML reader itself; PPTX retains its typed DrawingML error source.
This is not a claim that all legacy PPTX diagnostics/selectors are migrated.

Shape, text and table readers now distinguish formatting tab stops under
`a:tabLst` from text tabs. The retained record stores spans rather than decoded
property strings. Eager, owned and source scenes share this projection;
retained-scene cache weights already account for record size.

Minimal/full PPTX glossary profiles add eight bans against ordinary payload
metadata exports/accessors while allowing explicit raw access.

## Verification

- Ten property integration regressions cover lazy values, content-free
  diagnostics, transform ownership, rotation/flips/group coordinates, malformed
  values, namespace aliases and dialects, ambiguous payloads, content parts,
  BOM/MCE spans and retained lifetimes.
- Existing real-corpus owned/eager/source parity tests additionally compare
  non-visual values, transforms and raw payloads.
- `CARGO_NET_OFFLINE=true cargo test -p litchi-pptx --all-features`:
  895 unit tests, 467 integration tests and 10 doctests passed (1,372 total).
- `CARGO_NET_OFFLINE=true cargo test -p litchi-drawingml --all-features transform::`:
  six transform unit tests passed, including authored-content redaction.
- Both PPTX glossary profiles passed (19 bans each).

## Commit and cleanup

All-feature/all-target Clippy for PPTX and DrawingML with warnings denied,
both crates' warning-denied rustdoc, workspace library compilation and
minimal/full glossary checks pass after the review fixes. Review identified
table-plus-chart ambiguity and borrowed Debug leaks; regressions cover both.
Changed Rust files are formatted and `git diff --check` passes.

No temporary worktree was created. The three base-checkout fixtures needed by
existing full-leaf tests were copied temporarily and removed along with only
their newly created directories. Generated lockfiles and Python cache were
removed. The local base tip remains `32ad64e3207655b4ba558fc2793d9a7396655cdf`,
already an ancestor of HEAD.

The changed/untracked artifact audit passed, and both mandatory staged-file
checks printed nothing. `git add -A` again failed because
`/home/zhuhe/code/litchi/.git/worktrees/unified-ops-api/index.lock` is on a
read-only filesystem. No commit is claimed. The stale partial index must be
refreshed before any eventual commit; records 0051–0059 remain working-tree
changes. `cargo clean` removed 34,532 files (33.5 GiB), and turn scratch was
removed.
