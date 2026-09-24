# XLSX Custom Data caller-limit completion

This batch completes caller-limit propagation for the existing Custom Data
package owner and recognized connection `embeddedDataId` bindings. Custom Data
payloads remain inert. No query, connection, macro, or embedded payload executes.

The ordinary package API admits a finite `CustomDataLimits` profile for reads
and transactions. A transaction stages storage insertion, payload/properties
replacement, UID rename, or removal with an explicit reference disposition.
UID rename updates the recognized connection binding as one atomic candidate.
Source-checked patches preserve their limit profile and support exact inversion.

Validation must preserve source XML, unknown extension content, relationship
provenance, and unrelated members. Exact semantic no-ops retain source storage.
An owned changed payload should move into shared result storage without a byte
copy merely to satisfy a temporary XML construction limit. A refused edit must
leave the target package unchanged.

Caller limits cover XML bytes, event/depth/attribute/namespace dimensions,
UTF-16 UID units, relationships, package nodes and bytes, payload size,
construction, and aggregate publication. Package/output byte limits measure
logical uncompressed modeled OPC parts, content types, and relationship XML;
they are not compressed ZIP file-size limits or accounting for opaque non-part
ZIP members. It includes the canonical empty relationship view for an owner
without a physical relationship member. Relationship members and content types
retain independent policies. Final output accounting must use the candidate
aggregate, admitting a transaction that combines shrinking and growing members
when its final size fits. Temporary construction and final-output limits must be
checked before the corresponding allocation/publication. The temporary ceiling
bounds individual XML replacement and PI/MCE transformation buffers; it is not
an operation-wide allocator or retained-memory budget. Already retained source
XML is admitted under input limits independently of a smaller changed output.

Final evidence identifies the exact tested source, retained authored/native
inputs, measured workloads, limitations, and any baseline-only gate exceptions.
Previously recorded tests at base `180954433` are historical evidence and do not
replace validation on the current batch.
