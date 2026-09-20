# 0720 — DOCX structural scan qualification

This batch qualifies the duplicate DOCX body scans with a source-bound,
untimed diagnostic. It retains no production optimization and makes no speed,
allocation or memory improvement claim. The performance program remains active.

The actual owner is `DocumentBody::from_xml`, which returns
`ParsedDocumentBody`; 0719's next-experiment note mistakenly named the return
type as the owner. After stripping a BOM, it scans altChunk metadata, runs
alt-only MCE selection, scans outer paragraph/table/altChunk ranges, runs block
MCE selection, and finally captures the body. The proposed optimization would
combine the first two structural walks and keep both MCE calls and the final
body walk.

The standalone probe extends 0712's controls to 19 synthetic XML cases and
opens two complete packages: ordinary-save's 200-paragraph generated-medium
shape and the admitted NumberedList fixture. Temporary instrumentation records
actual body-boundary scanner invocations, successful event counts, consumed
source positions, raw and selected metadata/ranges, and each MCE input/output.
Standalone public offset probes and generator setup are outside representative
totals. This is a debug diagnostic, not a benchmark.

The original and both instrumented runs produce byte-identical public JSON;
the two trace stderr files are also byte-identical. Each representative body
performs two complete structural passes over the same source:

| Corpus | Main XML bytes | Events per structural pass | MCE input counts, in order | Structural source bytes summed over two passes |
| --- | ---: | ---: | --- | ---: |
| generated medium | 21,517 | 1,410 | 0, 200 | 43,034 |
| NumberedList | 4,563 | 178 | 0, 5 | 9,126 |

Event counts include EOF. These are logical parser traversal counts, not
physical I/O, CPU instructions or memory-copy measurements. MCE calls and the
final body capture are excluded from the summed structural bytes. On refusals,
the alt counter records successfully read events, while the range counter
records events admitted through its initial depth/node/classification checks;
its rejecting event is therefore excluded. Complete successful-pass counts
remain comparable.

The controls establish several requirements for a future implementation:

- With no anchors, the first MCE call has empty input; malformed MCE or unknown
  MustUnderstand can still fail the second call on paragraph/table offsets.
- A paragraph or table suppresses nested block ranges, while the alt parser
  still validates and collects nested anchors.
- Unbound fragments can match the range scanner but not the alt parser. The
  existing missing-anchor-metadata refusal must remain observable.
- A range depth violation occurs at 128, while the alt parser has a separate
  256 policy and also charges empty elements. A later missing relationship ID
  must still precede an earlier range-only depth error.
- BOM-adjusted offsets, strict/transitional namespaces, inactive branches and
  marker text remain part of the diagnostic matrix.

The [independent design review](results/change-0720/design-review.md) recommends
a private dual-state walk with separate counters, namespace predicates,
metadata validation and capture state. It requires deferred range errors until
the alt parser and its MCE call succeed. Offset unioning is excluded because it
changes marker placement and accounting. Interleaving range storage with alt
storage changes allocation lifetime; exact host allocator-exhaustion scheduling
is not established by this trace. Alt metadata is retained as exact debug
strings and hashed; the analyzer does not structurally parse those strings.
This is source-bound observation, not a complete equivalence proof.

A future pilot must first prove exact differential outcomes and source
preservation, including expanded node/anchor/offset/marked-byte and allocation
refusal controls, then run fresh native edit, lifecycle and allocator lanes.
Historical Callgrind percentages do not predict the saving. The separate
source-backed consumer and the rejected 0711 borrowing and 0715 section
collection changes remain outside this candidate.

Production source was restored byte-for-byte after tracing. The initial trace
build's SHA-256 formatting error is retained under `initial-trace-build`; it
produced no trace capture. The corrected builds and all three public probe runs
succeeded. Diagnostic formatting, Clippy and all six repository evidence
gates passed; no library-wide test suite or native pilot is claimed.

[Evidence and reproduction](results/change-0720/README.md).
