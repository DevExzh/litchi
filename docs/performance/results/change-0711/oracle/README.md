# 0711 `alt::scan` differential oracle

This standalone binary exercises the public `litchi_docx::alt::scan` API.  It
is intended to be run twice by the parent experiment: once with the baseline
library and once with the candidate library.  The resulting JSON files must
be byte-for-byte identical.  Every case records its source label, source
SHA-256, input SHA-256, byte length, and the exact structured result.  A
successful result includes every returned offset, relationship ID, and
`match_source` value.  An error includes both its exact `Display` text and
`Debug` representation; errors are observations, not passes manufactured by
the oracle.

The bounded matrix covers strict and transitional WordprocessingML, custom
and default namespace spellings, namespace rebinding, valid and malformed
`altChunkPr`/`matchSrc` content, duplicate and missing relationship IDs,
markup-compatibility choice and fallback branches, BOM and XML-depth/anchor
limits, truncation and other malformed XML, and text, comments, processing
instructions, CDATA, opaque child content, and general references around
anchors.  It also reads `word/document.xml` from two checked-in DOCX fixtures
through `soapberry_zip::office::ArchiveReader`; the ZIP bytes are the source
whose hash is recorded and the decoded main part is the scanner input.

The 49-case matrix contains one over-limit XML input and one over-limit anchor
input.  The baseline scanner accepts the DOCTYPE event case; it is an event-parity
probe, not a generic XML security assertion. The otherwise identical no-DOCTYPE
case also reaches the scanner's outside-anchor
text, CDATA, general-reference, comment, processing-instruction, and XML
declaration event paths.  These inputs are generated once per run and are
hashed without embedding their large payloads in the report.  No panic is
caught or converted into a successful case: a parser panic terminates the
process and leaves no valid report.

Run from the repository root:

```text
cargo run --release --manifest-path \
  docs/performance/results/change-0711/oracle/Cargo.toml -- \
  --repo-root . \
  --output docs/performance/results/change-0711/oracle/baseline.json
```

The output path is not included in the JSON, so changing it does not affect
the differential comparison.
