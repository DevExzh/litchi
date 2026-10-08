# 0023 — Retained source-backed topology publication

Date: 2026-10-08. Status: retained topology foundation; source owner migration open.

## Change

OPC `PreparedTopology` now implements cheap `Clone` and its `publish_to_stream`
borrows the prepared value. The source archive/catalog and complete prepared
state are shared: replacements, additions, effective relationships/content
types, omission lists, transfer authorizations and reservations. Unchanged
payloads remain lazy. Successful publication and sink failures leave the value
available for candidate readback, retry or another publication.

Every publication checks the base source and all transfer sources and execution
contexts before output, then uses the existing source-fenced, counted writer.
Exact no-op preparation still copies the original artifact and does not expose
a candidate callback. Cloning does not reopen, read payloads or duplicate
reservations. Existing source XML and precompressed reservation tests now drop
the retained preparation explicitly before asserting release.

`soapberry_zip::PreservationPlan` stores appended entries as
`Vec<Arc<RegeneratedEntry>>`. `try_append` still accepts an owned entry, while
`try_append_shared` accepts a retained handle; `appended()` exposes the shared
slice. The compression scheduler and writer borrow the same entries. Prepared
OPC publication only clones these handles into its temporary plan, avoiding
`RegeneratedEntry::Clone`'s deep copy of names and owned relationship payloads.
Generated relationship replacements also retain shared payload allocations.

Managed preparation admits the shared state holder, generated relationship
holders and appended-entry metadata before their allocations. Reservations and
transfer monitoring guards remain attached to the prepared state's lifetime.

## Verification

New tests cover repeated byte-identical publication, concurrent cloned
publication, sink failure followed by successful retry, retained candidate
readback, base and transferred source changes between publications, clone
reservation equality and final-owner reservation release. ZIP coverage verifies
that cloned/shared append plans retain the same generated entry allocation and
produce identical source-preserving output. Existing boundedness, physical
preservation, source-fence and output-accounting tests remain active.

Verification results:

- Full-feature OPC and ZIP unit/integration run: 1,686 passed, two existing
  tests ignored. Four new retained-preparation tests passed in the focused
  33-test topology target (1,690 distinct tests total).
- Updated ZIP shared-entry test: passed, including byte-equivalent output.
- DOCX SVG source/lifecycle tests: 30 passed; XLSX form-control scalar lifecycle:
  32 passed.
- Warning-denied Clippy for OPC and ZIP libraries: passed.
- Warning-denied doctests: 37 passed across OPC/ZIP, one existing ignored test.
- Full-feature facade and Python library compilation: passed.
- Glossary rustdoc: 3 profiles, 14 banned-name checks, 0 complete transaction
  families; passed.
- Crate-boundary gate: 66 packages, 267 internal dependency declarations,
  11 existing explicit debt items; passed.
- Focused formatting and `git diff --check`: passed.

Read-only review found no blocker in reservation admission, retained lifetimes,
source fences or concurrency. The intentional public ZIP accessor type change
requires migration for external callers that name `&[RegeneratedEntry]`
explicitly; ordinary inferred access and dereferencing continue to work.

## Remaining work and limits

A retained prepared topology is not yet a chained editable package. Its
`EffectiveTopology` remains a callback view over a catalog prepared against the
original source. `SourceXmlPart` deliberately refuses a second range proof on
derived XML; reusing the original physical lineage/version as an overlay
revision would authorize against the wrong state. That refusal is unchanged.

The next source-owner step needs a logical revision for each effective overlay,
reads from the effective catalog, and revision-bound XML authorization for
subsequent edits. Physical archive and transfer-source checks must remain
independent and active. Saving/reopening a replacement archive between edits
is not an acceptable substitute. Source `Package::edit` still returns the
transitional XML family from record 0022; no complete family or W2 completion
is claimed. All other original goal requirements remain open.

Prepared transfer monitors remain active until the final prepared owner drops,
as before for an unpublished preparation. A long-lived editable owner should
scope such monitors to operations to avoid imposing permanent per-read
observation costs on a shared foreign source.

The sharing tests prove ownership/accounting properties, not timing budgets.
There is one new shared holder per prepared state and per appended member;
owned append construction also allocates a holder. Paired timing and allocation
evidence remains required for the completed owner wave. No performance
improvement or completed performance gate is claimed.

## Cleanup

Removed generated build output, root lockfile, temporary native SVG fixture
copies (the same two paths/hashes recorded in 0019), and this turn's check logs
after recording their results. No new worktree was created. The base checkout,
iWork sources, pre-existing worktrees and tracked standalone lockfiles were
left unchanged.
