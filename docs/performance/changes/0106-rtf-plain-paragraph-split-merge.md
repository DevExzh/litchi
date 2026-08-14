# Change 0106: RTF plain paragraph split/merge

Date: 2026-08-14

Status: correctness-only CRUD coverage; no performance claim

## Scope

Commit `be37096ef` adds bounded `litchi_rtf::edit::Edit` operations for
splitting one ordinary body paragraph and merging two adjacent ordinary body
paragraphs. The selector is a zero-based paragraph position. A split offset is
a UTF-8 byte offset on a character boundary; it inserts the exact canonical
`\\par ` bytes at the proven source position. A merge removes only the exact
source bytes of the selected paragraph boundary. `split_paragraph_at` and
`merge_paragraph_with_next` are naming aliases for the same checked operations.

The source map admits only a root-level, contiguous, literal-ASCII ordinary
body whose semantic text has a one-to-one source mapping and whose paragraph
boundaries are exact `\\par` controls. Formatting and paragraph properties are
not rebuilt. The terminal paragraph cannot be split at its end unless an exact
boundary already exists; merges require the second selector to be the immediate
successor. Operation, source, text, boundary and allocation sizes remain under
the existing finite edit/parse limits.

## Contract and refusal boundary

Each operation stages one immutable source-bound transaction. Commit builds a
candidate by a bounded source splice, reparses it, re-resolves the ordinary
paragraph map, and checks paragraph count and text readback before returning a
`Commit`. The existing sequential writer remains the publication path. Durable
`paragraph.split` and `paragraph.merge` operations carry the selected text (or
left/right text), offset where applicable, exact boundary bytes, and a
SHA-256 result-artifact precondition. Durable apply checks the source
preconditions and candidate result digest. Focused tests separately prove
forward replay, exact inverse restoration, and foreign/stale source refusal.

The closure fails closed for compressed or non-ASCII transport, unknown or
opaque syntax, nested groups, non-paragraph controls, binary payloads,
external/transformation/mail-merge metadata, tables, fields, drawings,
objects, pictures, shapes, form fields, bookmarks, revisions, annotations,
notes, math/custom-XML content, protection ranges, editable regions and other
body-story events. Protected documents refuse at commit. Active/external
content is therefore not executed or rewritten; signed-document verification
and preservation are not a claimed RTF capability and remain outside this
operation's proof. Noncanonical forged boundary metadata is rejected by the
durable result digest/boundary checks. Rich formatting, encoded legacy text,
tables, fields, positioned content, structural edits and cross-document
composition remain outside the closure.

## Verification

The focused `paragraph_split_merge` integration suite has six tests covering
split/merge wire preservation, sequential publication, durable replay and
inverse, foreign-source refusal, forged boundary/result rejection, adjacent
selector validation, exact boundary-byte restoration, external/unknown
metadata refusals, Unicode-boundary and terminal-offset errors, operation
limits, unsafe syntax, and protected-source failure atomicity. The current RTF
crate library gate runs with warnings and deprecations denied; related RTF
integration suites, strict Clippy, rustfmt check, and the focused diff review
are required before release. These are correctness gates only: no latency,
I/O-range, allocation, RSS, cold-filesystem, high-latency, stream-window,
producer-breadth, or general rich-RTF claim is made.
