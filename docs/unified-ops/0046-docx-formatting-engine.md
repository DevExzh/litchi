# 0046 — reversible DOCX formatting and body structure

## Change and provenance

The primary package-bound DOCX Edit now formats direct run and paragraph
properties, inserts typed paragraphs and runs, removes runs/paragraphs and moves
paragraphs between body anchors. The source package remains unchanged. Existing
snapshot edits use the same engine; canonical commits return a package and an
exact patch. Formatting does not resolve styles or numbering definitions.

The starting source diff was candidate
bbe04c1de62e2b7373592f72f87cd7b2fe0e33e8. Only DOCX leaf code/tests were imported;
no candidate ancestry, benchmark evidence or claimed performance was imported.
The integration uses current Error/Result names, package-bound Edit forwarding,
patch.is_noop, OpenOptions and write_to. No removed compatibility alias returns.

Changed direct properties splice only their owned markup. Unchanged properties
retain lexical spelling; new elements follow schema order. Container removal
refuses tracked or unknown content, including comments and non-element content.
Inherited namespace bindings now reach property parsing, splicing, exact-head
replay and structural dependency scans. Foreign attributes cannot satisfy Word
range/field identity. New inputs and Operation Debug omit authored content.

Six operation variants cover property-head replacement, paragraph fragment
insertion/removal and run insertion/removal. Exact inverses, durable encoding,
composition effects and source-owner overlay authorization share those variants.
A transaction carrying such operations retains its staged XML without
compaction, as revision dispositions already do. Derived layout paths retain
namespace snapshots; differential tests compare them with a full rescan.

Before property-plan cloning/generation, a conservative fixed staging estimate
bounds variable property strings and tab vectors. Generated element lengths and
output sums use checked arithmetic; large output reservations are fallible.
FormattingStagingBytes reports the 64 MiB staging cap. Existing replacement and
whole-document bounds still apply. This is not caller-selected SemanticLimits
or a claim that all legacy allocations are fallible. Managed routes refuse
before formatting work, including append_run, rather than emitting uncharged
retained output.

Durable inserted fragments independently validate their root/property XML and
attribute constraints before publication. Multi-step paragraph moves and empty
paragraph run insertion stage a clone, publishing only after all steps succeed.
Signature and source-owner publication policies remain enforced by their owners.

## Verification

The focused unit suite includes exhaustive typed-property splices, exact no-ops,
byte-exact inverse/durable replay, composition, a 1,000-edit batch and layout
rescan differentials. Across 57 fixture documents it exercises 3,618 paragraph
and 6,004 run property edits, plus structural round trips.

New review regressions cover inherited alias/foreign names, opaque comments and
text in removed containers, foreign range IDs, duplicate replay attributes,
typed oversized-property refusal, content-free Debug and canonical eager/lazy
owner commits with inverse application. Integration coverage includes OfficeCLI
property sets, a producer revision fixture, signature refusal and both owners.

Validation commands and results:

- `cargo test -p litchi-docx -p litchi-ops --all-features`: 2,533 unit/integration
  tests passed (1 existing ignored), plus 161 doctests (31 existing ignored).
- `cargo test -p litchi-docx --no-default-features --lib formatting`: 45 passed
  before the final large-no-op regression; that regression passed separately
  with all features.
- `cargo clippy -p litchi-docx -p litchi-ops --all-features --lib -- -D warnings`:
  passed.
- `cargo check --workspace --all-features --lib`: passed, including Python.
- `CARGO_TARGET_DIR=target cargo check --manifest-path crates/litchi-docx/fuzz/Cargo.toml --bins`:
  passed.
- `CARGO_TARGET_DIR=target cargo run --manifest-path crates/litchi-docx/fuzz/Cargo.toml --bin smoke_formatting`:
  6,620 cases, 42 accepted with exact inverse restoration, 6,578 refused;
  no invariant failure.
- `python3 tools/check_api_glossary.py`: passed, 11 profiles, 144 banned names,
  zero complete transaction families.
- `python3 tools/check_crate_boundaries.py`: passed, 66 packages, 267 internal
  dependencies, 11 existing explicit debt items.
- Changed Rust formatting, whitespace and repository evidence restrictions
  are checked before commit.

Unchanged large property values bypass generation admission and remain exact
no-ops. New staging errors contain structural sizes only.

Read-only review found no remaining material blocker after the preservation,
namespace, replay and allocation fixes. Small tag/delimiter allocations remain
bounded but do not all use fallible reservation helpers. Paragraph replay checks
the root and paragraph properties; producer-invalid nested run property values
remain preservable, while fresh run authoring validates each run.

## Remaining work

The unified property registry/coercions, M1/M2 facade integration, managed
formatting retention, caller semantic limits, effective relationship projection,
owner consolidation, paired performance acceptance and remaining formats are
still required. No milestone or full transaction family is declared complete.
The corpus counts are correctness evidence, not a performance claim. Fuzz smoke
is deterministic bounded exploration, not a sustained fuzz campaign.

Borrowed fixtures, generated root/fuzz lockfiles, build output and session files
are removed before commit. No temporary worktree was created; base, iWork and
pre-existing worktrees remain untouched. The rewritten base remains an ancestor.
