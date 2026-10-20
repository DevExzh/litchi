# 0355a — W1 DOCX gaps closed: source-backed families, typed limits, restored capabilities, batch budget

- Status: in progress on `uo2/w1-docx` on top of the W1 merge (`3e686506f3`).
- Date: 2026-10-20
- Input: the coordinator review of record [0355](0355-w1-docx-contract.md):
  items F4 (restore removed capabilities), W1b (source-backed streaming
  families follow §2.3), F2 (typed `LimitKind`), F3 (one refusal
  taxonomy), P1 (the 1k paragraph-text batch within 2.0×), and the flaky
  `source_backed_batch` cancellation test.

## 1. The source-backed family shape (for other waves to follow)

This is the shape every source-backed (lazy, `ReadAt`-backed) DOCX edit
family now has; the W2 PPTX `SourceBacked*` editors mirror it.

```text
source_backed::Package::edit_x(&self, ..) -> x::Edit          // O(1), shares the package
x::Edit::{verbs..}(&mut self, ..)                              // staging only
x::Edit::commit(self) -> Result<x::Commit, x::Error>           // validates, writes nothing
x::Commit = litchi_core::edit::Commit<source_backed::Package, x::Patch>
    .package()       // the committed package (a new owner; the source is unchanged)
    .patch()         // x::Patch: is_noop(), inverse()
    .diagnostics()   // litchi_core::edit::Diagnostics: changed()
source_backed::Package::apply_x(&self, &x::Patch) -> Result<x::Commit, x::Error>
source_backed::Package::write_to(&self, impl Write) / save(..)  // the only output
```

How a committed package works (`source_backed/deferred.rs`):

- **Commit validates, output streams.** `commit()` runs every check the
  former `publish_x_to_stream` ran before output (source conflict,
  protection, preservation, signature, encryption, limits) and returns a
  new `Package` holding the pending output: the same byte-preserving source
  splice (`write_part_overlay*`) the old publisher wrote, run when the
  package is written. A commit followed by `write_to`/`save` therefore
  streams exactly as before, with the same output bytes, memory and
  speed; untouched members stay raw-copied references into the source.
- **Reads materialize once.** Reading the committed package (or editing it
  again) writes its output into memory once and serves every later read
  from it: chained edits and reads work, at O(archive) memory for that
  package only. A caller that only saves never pays it.
- **Patch.** A family's `Patch` wraps its detached record patch (the exact
  before and after bytes of what it changed). `apply_x` on any package
  checks the package's current state against the record patch's *content*
  (bytes, selector, limits), not the source lineage, and commits it there.
  `inverse()` swaps the record states; applied to the commit's own package
  it returns the exact source package, so the inverse output is the
  original artifact byte for byte. Forward-only record patches (the two
  redactions) invert only that way. `From<record::Patch> for x::Patch`
  replays a detached record patch onto a package.
- **Refusals before output.** A changed commit of a signed or encrypted
  source is refused at commit (`OpcError::SignedSourceRequiresExplicitPolicy`),
  not at write. A source that changes between commit and write is refused
  by the write (`SourceChanged`, or the family's artifact-fingerprint
  conflict) before any byte is emitted.
- **Macro.** `source_family!(what, Change[, Error])` declares `Patch`,
  `Commit`, `Diagnostics`, `Error`, `commit_change` and `apply_change` over
  a family's `Change` (`is_noop`, `inverse`, `publish(&Package) ->
  Package`); the family adds `Edit` and `LimitKind`.

Families converted so far: `document_variables`, `section_layout`,
`hyperlink_detachment`, `hyperlink_redaction`, `story_hyperlink_redaction`,
`alt_chunk_payload`, `story_text_edit`, `glossary_text`.
