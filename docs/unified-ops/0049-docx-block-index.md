# 0049 — DOCX indexed blocks and generation memos

## Scope

Add the leaf prerequisite for the DOCX M1 adapter, using code-only work from
`383b27af301d51d1b33f233748231bedb1e19c66`, adapted to current immutable owners
and ADR 0034. No old branch history or measurement artifacts were merged.
The current base HEAD is already an ancestor; no further rebase was needed.

`Blocks` indexes direct visible children of document bodies, cells, content
controls, headers/footers, notes and comments. It retains namespace context,
flattens custom-XML wrappers, separates display equations from paragraphs per
plan 0010 D4, and exposes checked raw transaction positions only when proven.
Positions use `Position`. New native identity access lives under `document::raw`;
duplicate paragraph IDs resolve to the first block, matching the candidate
lookup contract. Existing legacy identity accessors remain migration debt.

Owning and physical source-backed packages memoize visible XML and block indexes
by payload generation. Effective source owners use revision-local block memos;
their visible projection is still rebuilt. One-entry caches have 16 MiB visible
and 8 MiB block ceilings, conservative namespace-inclusive weights, explicit
clean trimming, pinned-value retention, oversized bypass and failure retry.
Immutable edits keep old views readable and new indexes fresh.

Indexed parsing checks namespaces, attributes, nesting, references and duplicate
body/content containers. It refuses DTDs and processing instructions. Indexed
owned fragments now require declared namespaces. Errors added here omit authored
content and carry typed limit dimensions. These readers are not full Word schema
validators: unknown markup remains opaque, and a document without a body yields
an empty index. Fixed scan limits are not yet caller-selected SemanticLimits.

Managed semantic views remain refused. The refusal also covers an unmanaged
base with a managed donor, preventing a detached view from outliving its budget
reservation. Existing pinned document text remains available. Markdown retains
its prior strict refusal of content controls rather than silently exporting a
new block variant.

## Verification

Tests cover corpus owner
agreement, raw transaction positions, namespace aliases, control fields, malformed
XML, cache sharing/trim/oversized bypass, immutable revisions and managed donors.
The existing DOCX fuzz target now visits indexed blocks and control reads.

- `cargo test -p litchi-docx -p litchi-ops --all-features`: 2,724 unit/integration
  tests and 226 doctests passed; 32 pre-existing ignored tests remain ignored.
- `cargo test -p litchi-docx --no-default-features --lib document::blocks::tests`:
  14 passed.
- `cargo test -p litchi-docx --all-features --lib memo`: 17 passed after the
  final allocation-capacity accounting adjustment.
- `cargo check --workspace --all-features --lib`: passed, including Python.
- `cargo clippy -p litchi-docx -p litchi-ops --all-features --lib -- -D warnings`:
  passed.
- `CARGO_TARGET_DIR=target cargo check --manifest-path
  crates/litchi-docx/fuzz/Cargo.toml --bin parse_docx`: passed.
- `python3 tools/check_crate_boundaries.py`: passed, 66 packages, 267 internal
  dependencies and 11 existing explicit debt items.

- `RUSTDOCFLAGS='-D warnings' cargo doc -p litchi-docx -p litchi-ops
  --all-features --no-deps`: passed.
- `python3 tools/check_api_glossary.py`: passed, 11 profiles, 144 banned names,
  zero complete transaction families.
- `CARGO_TARGET_DIR=target cargo run --manifest-path
  crates/litchi-docx/fuzz/Cargo.toml --bin parse_docx -- -runs=100 <scratch-corpus>`:
  passed with three small repository DOCX seeds. This was an uninstrumented
  bounded smoke run, not a coverage-guided campaign.

Read-only review caught namespace, retained-memory, donor-lifetime and XML
classification/validation issues; fixes have regression coverage. Legacy ordered
block readers retain some looser malformed-input behavior; this record does not
claim their complete migration. Visible memo weight uses allocated capacity,
while indexed namespace scope weights deliberately overcharge shared scopes.

## Remaining work

This is an M1 prerequisite, not M1 closure. Concrete unified format nodes,
File get/find/query, registry and capabilities integration, Python/CLI operations,
all-format reads/writes, caller semantic budgets and paired performance evidence
remain required by the goal. No speedup or full performance acceptance is claimed.


## Hygiene

Changed Rust files are formatted. Borrowed fixtures, generated root/fuzz
lockfiles, build output and session scratch artifacts are removed before commit.
No temporary worktree was created. Base, iWork and pre-existing worktrees remain
untouched. Staged evidence size and binary restrictions are checked before commit.
