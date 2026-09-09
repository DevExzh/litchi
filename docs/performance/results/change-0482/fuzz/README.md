# XML bounded-auditor fuzz scope

`crates/xml-minifier/fuzz/fuzz_targets/minify_xml.rs` runs the bounded slice
auditor and both reader policies (`verify_reader` and
`verify_authored_reader`) on the same input. Inputs above 64 KiB are skipped.
The profile uses a 4 KiB token ceiling, depth 32, 1,000 attributes, 128 KiB
events, and 64 KiB aggregate text. The source adapter is an allocation-free
`BufRead` whose chunk sizes are selected from the first 256 input bytes, with
chunks from one through 97 bytes.

Successful slice and reader results must have identical reports. For valid
UTF-8 failures, the harness requires the same error category and resource
limit; token lookahead values and source-order byte offsets are allowed to
differ as documented by the streaming API. A truncated token may be reported
as malformed by the slice parser after it reaches EOF, while the reader
reports its bounded `TokenBytes` lookahead first; that specific precedence
split is accepted. The slice path prechecks complete UTF-8, while the reader
reports the first failure observed from the source, so an invalid UTF-8 input
may produce an earlier malformed, compactness, or limit error in the reader.
That precedence difference is intentionally accepted and kept in the harness
comments.

The existing producer macros are compile-time APIs. The target includes one
fixed `minified_xml_str!` smoke invocation, while runtime fuzz bytes exercise
the public audit APIs. The corpus contains accepted mixed markup, authored
ambiguous whitespace, BOM/comment raw-span behavior, truncation, malformed
references, raw end-tag spacing, an oversized text token, excessive depth, a
large attribute tag (which also exercises token-limit precedence), and invalid
UTF-8 binary input.
