# Source-backed PPTX modern comments

This wave implements the modern-comment portion of L-PPTX-R9. The complete
unified-operations goal remains active; unified PPTX dispatch and performance
acceptance are separate outstanding work.

## API and ownership

`SourceSlide::modern_comments()` acquires the selected slide's optional modern
comment list. Detached semantic handles live in `presentation::modern_comments`;
native identities, XML and physical package metadata require explicit raw access.

Acquisition follows the retained selected-slide relationship catalog and the
presentation's author relationship. It does not load unrelated comment parts.
An absent comments relationship differs from an existing empty list. Comments,
replies and assignments require resolved authors. Duplicate comment/reply IDs
are refused within the selected list; no package-wide uniqueness claim is made.

## Validation

The source physical/MCE boundary precedes the existing bounded leaf codecs.
Internal processed-input helpers reuse those codecs without a second MCE pass.
Selected relationships must be unique and internal, with exact content types;
comments and author parts must not have outbound relationships. Source and
execution fences surround acquisition and semantic projection.

Text projection uses the shared DrawingML reader, including namespace bindings
inherited through the comment list, comment, reply-list and reply elements.
Absent and present-empty bodies remain distinct. Ordinary diagnostics omit
native identifiers and authored content.

The shared comment/author codecs now preserve predefined and numeric character
references in inert fragments while rejecting undefined entities. This fixes
valid ampersand text previously refused by the codec. Existing semantic bounds
now return typed `Error::Limit` values instead of formatted invalid-input errors.
Projection moves parsed fields into detached handles and indexes author names;
no performance acceptance or complete configurable semantic-budget migration is
claimed.

## Verification

Offline validation passed:

- `cargo test -p litchi-pptx --all-features`: 1,450 tests/doctests passed,
  3 ignored.
- `cargo test -p litchi-pptx --test source_modern_comments`: 8 passed.
- `cargo clippy -p litchi-pptx --all-features --all-targets -- -D warnings`.
- `RUSTDOCFLAGS='-D warnings' cargo doc -p litchi-pptx --all-features --no-deps`.
- `cargo check --workspace --lib`.
- PPTX minimal/full glossary profiles: 59 bans each; glossary checker unit
  tests: 21 passed.
- Offline crate-boundary audit: 66 packages, 267 internal dependency
  declarations, 11 existing explicit debt items.

Tests cover detached semantic/raw access, missing/empty bodies, inherited and
rebound namespaces, MCE fallback, selected-part locality, invalid relationship
closures, duplicate IDs, unresolved author references, source revisions,
caller byte limits, diagnostic redaction and codec character-reference handling.

## Commit and cleanup

The local rewritten base `32ad64e3207655b4ba558fc2793d9a7396655cdf` remains an
ancestor of HEAD. `git add -A` still fails because the linked-worktree
`index.lock` cannot be created on the read-only filesystem. This wave and earlier
outstanding changes remain uncommitted. Both required staged checks printed
nothing; the changed/untracked artifact audit found no prohibited evidence or
compiled binaries. Removed 29.3 GiB of generated Cargo output, owned fixture
copies, generated untracked lockfiles, Python caches and checkpoint scratch
files. No temporary worktrees were created.
