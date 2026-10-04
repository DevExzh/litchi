# 0905 attribute-summary semantic parity oracle

`crates/xml-minifier/examples/attribute_summary_parity.rs` is a standalone,
deterministic differential oracle for the proposed zero/one-attribute summary.
It has no timing assertion and no dependency beyond the public `xml-minifier`
API. Build the same example once from the unchanged source tree and once from
the candidate tree, capture stdout from each binary, and compare the captures
byte for byte. A SHA-256 of each capture is a convenient second record of that
comparison; the JSON-lines output itself remains the authoritative evidence.

The corpus contains 11,520 generated cases. Its dimensions are six tag
cardinalities (`0`, `1`, `2`, `31`, `32`, and `33`), 128 deterministic lexical
styles, and 15 resource profiles. Styles cover ordinary, Unicode and
punctuation names, `xml:*`, valid and invalid `xml:space`, namespace
declarations and uses, both quote marks, empty and escaped values, whitespace
around `=`, tabs and non-ASCII whitespace, unquoted and unterminated values,
bad references, raw `<` values, duplicate tails, CDATA, malformed character
data, declarations, BOMs, and invalid UTF-8. The generated tags include both
start/end and empty-element forms. The profile set includes the default and
exact/one-under aggregate attributes, per-element attributes, depth, events,
text bytes, total bytes, and token bytes limits. For valid default-profile
inputs the exact aggregate values are taken from the public `Report`; fallback
values are fixed by the generator for failing inputs.

Every case records the slice results for `verify_source`, `verify_authored`,
and `verify`, plus chunked-reader results for `verify_authored_reader` and
`verify_reader`. Reader sources are replayed with one-byte, deterministic
uneven, and nearly-whole-input (`prefix`) buffers. A replacement record covers
identical bytes, local text and nested edits, replacement failure, replacement
budget retotaling, a complete fallback, original DOCTYPE failure, and original
malformed XML. There is no source-policy reader API in the public crate, so the
reader matrix intentionally contains the two available reader policies.

Each record is a JSON object with a stable field order. `debug` is the exact
Rust `Debug` rendering of the complete `Result<Report, Error>`,
`Result<Report, StreamError>`, or replacement result. `text` is the public
`Display` rendering for errors and a fixed report/proof text for successes.
`hash` is FNV-1a-64 over `debug`, one NUL byte, and `text`; `input_hash` and
`replacement_hash` bind the exercised bytes. JSON escaping is implemented in
the example, so no serializer version can change the output. The first and
last lines carry the schema and record count.

The root coordinator should run the example from isolated baseline and
candidate builds and use `cmp` before interpreting any performance result:

```sh
baseline/attribute_summary_parity > baseline.jsonl
candidate/attribute_summary_parity > candidate.jsonl
cmp --silent baseline.jsonl candidate.jsonl
sha256sum baseline.jsonl candidate.jsonl
```

The example was added as source for the root build supervisor; this packet does
not claim that it has been compiled or executed here.
