# 0818 source review

The base is `7cdaba587b`. The previous batch demonstrated the preservation
failure on `documentProperties.docx`: an ordinary paragraph append left the
relationship edge set unchanged but reordered its XML and removed declaration
whitespace. The new byte-level regression fails on unchanged production at
the expected assertion, after successful compilation.

DOCX `package/codec.rs` reconstructs the main document's payload and relationship
map during ordinary publication. Its prior `OpcPackage::add_part` call removes
both payload and relationship source provenance because that API represents an
arbitrary replacement. The bounded correction updates an existing main Part
in place, retaining the package-level relationship source map. It leaves the
missing-Part insertion path and the global OPC authoring contract intact.

Independent read-only review confirms that `get_part_mut` still revokes exact
whole-source authorization and marks the signature graph for validation. The
OPC writer compares the current relationship binding, including IDs, types,
target spelling and modes, with the admitted source binding. Only an equal
binding reuses original relationship XML. Changing the relationship object
may drop its own source capture; exact reuse here comes from the package-level
source map, not an assumption that the rebuilt map preserves source order.

The existing `WriteRollbackGuard` restores package and provenance state after
failed mutation/publication. Updating the existing Part also retains its
implementation and metadata rather than substituting a new BlobPart. The
public API, dependency graph, signature policy, XML audit, and durability
contract do not change. This is preservation repair; no latency, allocation,
or throughput benefit is inferred without measurement.

The regression fixture intentionally has noncanonical relationship order and
CRLF after the declaration. Exact decoded member equality tests the actual
contract; equal graph sets alone would miss the failure. A changed-hyperlink
control must prove that changed bindings publish the new edge instead of
reusing stale source XML. Fresh artifact admission remains a separate
independent ZIP/XML oracle.
