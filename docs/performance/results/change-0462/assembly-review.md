# 0462 assembly review note

The assembly receipts retain raw `objdump` output for four targets in each
epoch: `ElementAttrs::get`, `ElementAttrs::lookup`, `ShapeAttrs::get_known`,
and `Parser::shape_builder`.  The verifier binds each output to its receipt,
the assembly driver, and the normal binary.  A target may be recorded as
`missing_or_inlined`; that status is evidence of what the requested symbol
lookup returned, not evidence that its work disappeared.

`generated_code_evidence.known_cached_scan.eliminated` is a narrow heuristic.
It is true when the retained `shape_builder` body is present and its parsed
direct call list does not contain `ElementAttrs::lookup`.  It does not inspect
inlined instructions, calls through `ElementAttrs::get`, or the body of an
inlined `ShapeAttrs::get_known`.  The baseline receipt therefore must not be
read as proving cached-scan elimination merely because this field is `true`;
the baseline body calls the getter path and the direct-call heuristic can
miss the scan.

The production claim requires manual inspection of the four retained raw
disassemblies for both epochs.  Reviewers must compare the actual getter and
known-key bodies, follow the `Parser::shape_builder` path, and use the
reported stack frame observations only as generated-code measurements.  The
review conclusion belongs in the final decision; this note does not amend or
rewrite either assembly receipt, its raw artifacts, or the protocol.
