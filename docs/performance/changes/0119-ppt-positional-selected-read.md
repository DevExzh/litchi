# Change 0119: native PPT positional selected reads

Date: 2026-08-15

## Scope

This tranche adds a public immutable `SourceSnapshot::read_text(Target)` query
that shares the existing bounded positional resolver and refusal policy with
the equal-length native PPT text editor. It also lets the umbrella
`Presentation::open` path retain a validated positional native-PPT package on
Unix and Windows, including metadata projection through an immutable shared
OLE property-set reader. OLE2 classification remains content-derived, Word
keeps precedence over PowerPoint when both host streams are present, and
non-PPT inputs continue through the established fallback.

The performance harness adds three opt-in selectors:

- `ppt_source_backed_one_shape_text` pairs the existing eager selected-shape
  query-only control;
- `ppt_semantic_fresh_open_one_shape_text` measures eager open plus the same
  selected-shape query;
- `ppt_source_backed_fresh_open_one_shape_text` measures positional source
  open plus that query.

The selectable-case count is 219. The default 36 cases / 198 records are
unchanged.

## Evidence boundary

Every selector uses the same deterministic corpus, semantic target, expected
text, and SHA-256 digest. Source-backed elapsed samples use an uninstrumented
immutable source. Separate untimed instrumented replays run exactly once per
measured sample, exclude warmups, require stable nonzero logical read
calls/bytes, and retain the selected-text digest in the result JSON. Query-only
replays reset counters after source validation; fresh-open replays include
source opening and the query.

The production differential tests compare every generated shape against the
ordinary eager reader and retain stale-source, macro, unsupported-topology,
aggregate-limit, and out-of-range refusal checks. Umbrella path tests compare
native path and byte facades for core queries and metadata, preserve non-PPT
fallback behavior, and cover DOC-before-PPT OLE2 polyglots when DOC support is
enabled.

## Claims deliberately not made

This is correctness and fixture-scoped logical-read evidence. It does not
claim a latency improvement, physical-I/O reduction, allocation or RSS
improvement, cold-filesystem behavior, or end-to-end edit/publication gain.
Any performance claim still requires a frozen release build and controlled,
balanced ABBA measurement with disclosed environment and variance.
