# Change 0460 source review

## Verdict

The frozen fused traversal is semantically sound. I found no blocker in the
read-only review. The shared path keeps the existing settings, declaration,
and page state machines, feeds the source scanner from the same namespace-aware
reader, and retains source-only failures until metadata completion. The direct
`ContentSource` path drives the same incremental scanner with its own reader;
the preserved reference scanner remains available for differential checks.

## Semantic checks

- Both readers explicitly keep `check_end_names = true` and `trim_text(false)`;
  the remaining reader defaults are unchanged. Element and attribute
  classification uses the resolver scope for the current event, so aliases and
  nested prefix shadowing retain their previous behavior.
- A leading UTF-8 BOM is removed once before tokenization. Model size checks use
  the original XML length, source spans use offsets in the BOM-free body, and
  `write_prolog` restores exactly one BOM. The fused metadata path remains
  bounded by the existing 8 MiB scanners; standalone source scanning retains
  the 64 MiB source limit and its depth, page, and automatic-style limits.
- Reader/settings errors remain immediate. Declaration errors are retained
  ahead of page errors, and source scanner errors are retained behind all
  metadata finish steps. Mutable staging applies the retained source result
  after MIME and styles setup, matching the former order.

## Adversarial coverage

The frozen tests compare fused metadata and source projections with their
independent references across namespace aliases, rebinding, malformed input,
BOM spans, native fixtures, 8 MiB/page limits, and the standalone 64 MiB source
limit. Explicit cases cover source-shape errors followed by settings,
declaration, and page failures, plus a source error followed by malformed XML;
the latter confirms that the shared reader error still wins. Strict Clippy is
green in the owner checks. The reviewer did not run builds or tests.
