# XLS User Names stream ownership

This batch addresses the MS-XLS §2.1.7.17 stream gap in
`docs/report/spec-gap-audit.md`. It supplies deferred workbook inspection,
bounded typed stream snapshots, metadata and membership transactions, and
source-checked reversible publication of a complete XLS CFB artifact.

The stream grammar is `CUsr UsrChk CbUsr BCUsrs *UsrInfo`. Active counts,
record sizes, user IDs, UTF-16 lengths, and revision-header GUID dependencies
are validated; fields the specification explicitly says to ignore retain
their exact source bytes, including nonzero inactive CbUsr slots and UsrChk
reserved values. Reading and unrelated edits do not normalize these fields. Accepted string option bits and original wide-character encoding survive
public record round trips. NUL characters remain inert text, so a timestamp
edit does not reject an otherwise representable source name. The package
owner replaces the existing root User Names stream and reopens the output.
It delegates physical CFB preservation and protected-container admission to
`litchi-ole-common`.

Revision Log admission now shares a borrowed grammar validator with its full
reader. Required lock order, nested revision productions, record-size limits,
code-page identifiers, and sheet-table uniqueness are checked before publication.
The RRTabId omission rule uses the containing workbook's actual BoundSheet8
count, rather than its next available sheet ID. When that count is known,
the table must contain exactly one identifier per sheet. Standalone callers can provide
that count explicitly; absent context does not establish an omission exception.

RRDHead code-page metadata accepts the 152 identifiers in Microsoft's
[CODEPG table](https://learn.microsoft.com/en-us/windows/win32/intl/code-page-identifiers),
verified against the published table on 2026-09-10. Metadata validity is
independent of whether a text decoder is available for that page.

Transactions share immutable source bytes and parsed metadata. A changed
candidate is validated when staged; commit consumes that validated state.
The focused [performance evidence](performance/README.md) describes synthetic
stream measurements and their limits.

Coverage does not establish native Excel interoperability. This owner does
not acquire shared-workbook locks, merge or replay revisions, evaluate
formulas, or execute external content. Signed/encrypted/DRM package mutation
remains subject to the common owner's existing refusal policy. Standalone
stream patches and complete-package patches remain separate in-memory APIs;
this batch does not introduce durable patch serialization or concurrent merge.

Final source-bound validation passed 1,417 tests across 73 targets (one existing
ignored doctest), all-target/all-feature strict Clippy, warning-denied rustdoc,
formatting, and diff checks. Commands, logs, and unchanged-source hashes are
recorded in `gates.json` and `gates/`. The independent allocator run is retained
in `root-profile.json` and `root-profile.log`.
