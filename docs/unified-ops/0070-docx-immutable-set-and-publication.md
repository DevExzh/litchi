# 0070 — Immutable DOCX set and stream publication

## Scope

This checkpoint connects the property-input/registry foundation in 0069 to the
leaf-owned DOCX Edit/Commit transactions. It introduces `File::set`,
`File::to_bytes`, and `File::write_to`, with typed path/query targets re-exported
through the facade. Structural edits, batches, atomic path saving, other format
write handlers, Python bindings and performance acceptance remain open.

`File::set` returns a new handle. It resolves the target, admits the ordered
assignments against the registry, and applies them to one isolated leaf edit.
A failure discards that edit. Successful eager and source-backed commits are
adopted directly; no ZIP serialization/reopen occurs between edits. Retention,
limits, and signature disposition follow the committed owner. Source checks
surround the entire operation, including failing paths.

## Targets and properties

Target strings distinguish structural paths from scoped queries. Excel ranges
remain paths; query combinators, pseudo-classes, wildcards and content predicates
use query parsing. Unscoped mutation selectors are rejected for strings,
parsed queries and compiled plans. Parsed paths must satisfy the structural
path contract. A union uses one global scope; multiple absolute scopes are
refused. Query compiler diagnostics are fatal to mutation, so an unknown
predicate under negation cannot silently broaden a write.

The DOCX handler supports direct main-body paragraphs and their proven direct
runs. It refuses nested/table/MCE selections without a stable raw transaction
address, and mixed paragraph/run fanout that could invalidate descendant
ordinals. Existing leaf restrictions still apply to complex content and
managed transactions. These restrictions are not claims of complete Word
write coverage.

Paragraph/run text and typed direct formatting use the same registry metadata
as capability reports. Text is write-only property metadata, appended without
changing the established read-property indices. Strict admission rejects
unknown keys, read-only properties, invalid tokens, out-of-range scalars and
unsupported color representations. Null clears supported optional properties;
ordered duplicate assignments retain their order. Errors carry property
attribution without exposing authored values in Debug output.

## Publication

Stream publication delegates to the leaf owner, including signature and
encryption policy. Signed no-op output preserves signatures; explicit unsigning
uses the leaf removal operation. Publication does not mutate the input handle.
`write_to` reports the accepted byte count on success; a failing sink retains
the typed leaf/I/O error. The unified layer adds no ZIP writer or atomic-file
implementation. `Verb::Save` remains unsupported until path saving is wired.

## Preservation findings and corrections

The new publisher exposed an existing eager-ingress defect: the facade opened
owned bytes through the borrowed OPC constructor, discarding archive retention.
It now uses the shared owned-byte constructor. Ordinary byte/path eager opens
preserve exact no-op output without holding a second semantic/source owner.

Prepared eager ingress uses an explicit source-preserving owning conversion.
An already-owned archive is shared; a positional source is copied once under
its admitted input bound, with fallible reservation and source fences. The
result remains usable after the original path is removed. The existing
non-preserving positional conversion remains available to compatibility reads
that do not need an archive-sized retention allocation.

The conversion transfers parsed content-types bindings and matching raw
relationship captures. It meters structural manifest/relationship reads when
accounting is requested. Establishing physical publication provenance performs
a bounded ZIP member scan; it does not serialize or reparse the semantic OPC
graph. Effective replacements are authored parts rather than falsely marked
as untouched source allocations. Changed conversions fail closed if physical
preservation provenance cannot be established.

A write-only `text` registry row initially shadowed the query compiler's
virtual text key; readable lookup now filters by `get`. A bare-selector error
initially echoed authored input; its message and suggestion are now static.
Regression tests cover both corrections.

Five operation-context gaps are exposed in capability JSON and the generated
OfficeCLI comparison: unproven paragraph selection, unproven run selection,
mixed paragraph/run fanout, managed transactions, and complex leaf content.
Unretained handles do not advertise `set`. This conditional coverage does not
close any of the remaining structural-write or all-format milestones.

## Verification

- Ops, all features: 917 tests/doctests passed.
- Ops, DOCX only: 545 passed; minimal feature set: 368 passed.
- OPC: 1,093 passed, one ignored.
- DOCX leaf, no default features: 2,252 passed, 32 ignored.
- All 28 advertised paragraph setters and 40 run setters are exercised in both
  source-backed and eager modes, including publication/reopen. Seventeen
  integration tests cover immutable chaining, query admission, ordered and
  empty assignments, null clearing, signature policy, source staleness,
  no-op/same-value preservation, prepared eager ingress after unlink, untouched
  member preservation and sink failures.
- All-feature and minimal ops all-target Clippy pass with warnings denied.
  OPC Clippy also passed in the leaf review.
- Ops/DOCX/OPC rustdoc passed with warnings denied. Minimal ops rustdoc passed.
- The minimal DOCX glossary gate checks 40 banned names and passes.
- The facade DOCX build and workspace library check passed.
- OfficeCLI delta generator: 13 tests passed; regeneration `--check` is clean.

The broad DOCX run initially lacked eight external corpus fixtures. They were
linked temporarily from the base checkout for verification and removed during
cleanup. An older XLSX assertion expected diagnostic counts in Debug; it now
checks the explicit error message while retaining its authored-content
redaction assertions, consistent with 0069's Error contract.

The shared eager-ingress path was exercised with all ops features. These tests
are correctness evidence, not performance acceptance; throughput, allocation
budgets, batch behavior and all-format write parity remain open.

## Commit and cleanup

The rewritten base remains an ancestor of HEAD. Both required staged evidence
checks print nothing, and the changed/untracked artifact audit finds no
prohibited evidence or compiled executables. `git add -A` fails because the
linked worktree cannot create its `index.lock` on the read-only Git metadata
filesystem. No commit was created. Removed 88.1 GiB of generated Cargo output, eight
temporary corpus fixture links, and 21 scratch/log/lock/cache paths. No
temporary worktree was created.
