# ODS audit disposition after bounded literal decoding

This batch closes the `parse_string` memory follow-up identified in the
[preceding scan audit disposition](../ods-formula-scan-performance/audit-status.md).
Decoded string storage now scales with decoded content. It also closes the
ODF 1.4 Part 4 §5.4 U+0000 exclusion gap. The source-bound tests and measurements
are in [README.md](README.md).

The broader OpenFormula item in `docs/report/spec-gap-audit.md` remains partial.
The existing tokenizer recognizes the 393 normative chapter-6 function names
and bounded §5.8 references. Those lexical capabilities do not establish
expression grammar, function arity, type semantics, evaluation, or recalculation.
Array expressions (§5.13), named and host-defined expressions, and the remaining
expression productions still require specification-backed implementation and
validation. Formula-text transactions continue to preserve inert expressions.

The sheet-metadata and DDE dispositions in the preceding audit remain unchanged.
This batch adds no evaluation, external source access, or package publication API.
