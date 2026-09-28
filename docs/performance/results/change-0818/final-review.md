# 0818 final implementation review

This is a bounded read-only review of the finalized 0818 preservation repair
against base `7cdaba587b`. It covers the DOCX main-part update in
`crates/litchi-docx/src/package/codec.rs` and the focused real-file regression
in `crates/litchi-docx/tests/real_file_preservation.rs`. It does not authorize
workload execution or make a performance claim.

## Disposition

**Pass for the bounded preservation repair; no concrete review gap remains.**
The implementation preserves the existing main OPC `Part` during ordinary
publication and retains the missing-part fallback. The final regression set
covers both unchanged relationship bytes and a genuinely changed relationship
graph.

## Implementation correctness

At `codec.rs:1229-1247`, the writer now obtains the existing main part with
`get_part_mut`, installs the rebuilt document payload with
`set_blob_shared`, and swaps the rebuilt relationship collection into that
part. A missing main part still uses the prior `add_part` path; other OPC
errors are returned. This keeps a custom `Part` implementation and its
implementation-owned metadata instead of substituting a new `BlobPart`.

The relationship collection's object-level source capture is dropped by the
swap, but this does not weaken preservation. `get_part_mut` still revokes
whole-source authorization and marks the signature graph as tracked. The
package-level retained relationship source is still available to the writer,
which compares the rebuilt binding by relationship ID, type, target spelling,
and target mode. An equal binding reuses the producer's exact relationship
member bytes and order; a changed binding uses validated serialization. Thus
lexically changed targets and added or removed edges cannot reuse stale source
XML.

The existing write rollback guard continues to restore the OPC package and
provenance state if mutation or publication fails. Signature-policy checks
therefore retain their existing explicit-policy boundary. The change does not
alter the public API, global `add_part` behavior, XML validation, or package
publication ownership.

## Regression evidence

The unchanged-source test saves a paragraph edit from
`documentProperties.docx` and compares the decoded
`word/_rels/document.xml.rels` member byte-for-byte, then reopens the output
and checks the edit marker. This exercises the fixture's producer-specific
relationship order and lexical XML rather than only comparing graph meaning.

The changed-graph test adds an external hyperlink, requires the emitted
relationship member to differ from the source and contain the new URL, then
reopens the package and checks both hyperlink text and the typed external
`HYPERLINK` relationship target. It therefore guards against incorrectly
reusing source XML after the relationship binding changes.

The caller reports formatting and checking complete, with the all-feature test
gate in progress at review time. This reviewer did not run Cargo, binaries, or
workloads. Existing OPC-level relationship reuse, signature, and rollback
coverage addresses the corresponding lower-level contracts; no additional
test is recommended without a concrete new failure.

## Scope boundary

This repair establishes preservation correctness for the affected ordinary
DOCX save path. It provides no evidence of lower latency, fewer allocations,
higher throughput, or any other optimization benefit. Any such claim requires
a separate measurement batch after the fresh quality and independent artifact
admission gates pass.
