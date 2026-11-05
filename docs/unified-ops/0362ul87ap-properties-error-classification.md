# 0362ul87ap — classify OOXML metadata transaction failures

The shared OOXML document-properties transaction now exposes the contract
error shapes required by §2.4. Unsupported metadata fields produce a
`RefusalKind::UnsupportedProperty` refusal with a fixed, content-free reason.
Rewriting a tolerated custom-properties part and removing a part with another
inbound relationship produce `UnsupportedContent` refusals. Prepared-layout,
admission-ledger, and checked accounting inconsistencies produce `Internal`;
real bounded-resource failures remain typed `Limit` errors.

Nested OPC source changes, including an `IncompleteOutput` wrapper, are
promoted to `properties::Error::SourceChanged { expected, observed }`. Source
graph, XML, and namespace failures are retained as package errors so the
unified layer reports corruption, while authored value failures remain
invalid-value errors. The DOCX classifier and the XLSB source publication
bridge preserve those distinctions; XLSX and PPTX re-export the shared error
unchanged. Metadata consumers and tests were migrated in the same change,
including core-only application-field refusal and malformed-source coverage.

Only `rustfmt` and source/diff checks were run for this repair. No Cargo,
rustdoc, glossary, native, or performance command was run.
