# 0362ul69b — Formula encoded-depth preflight return

Date: 2026-10-25. Parent: 0362ul69a. Status: source repair complete; the
coordinator still owns Cargo, rustdoc and performance validation.

The second review returned the previous depth repair because its byte walk
treated transparent AST wrappers as nesting and still undercounted encoded
MTEF levels. This return adds a separate, allocation-free depth walk that
mirrors the actual `NodeWriter` calls:

- the root LINE or direct-break PILE is entered first;
- every template, slot LINE, matrix record/cell LINE, pile row LINE and ruled
  run level is checked at the point the writer enters it;
- Row, Style, Run without rules, function arguments, accents, degraded
  wrappers and other transparent sequences do not consume a level;
- word-shaped large operators and fenced matrices follow their distinct
  template/slot paths.

The conservative byte/work walk remains unchanged apart from delegating depth
admission to this exact model. Boundary tests prove that a Row containing a
leaf fits at depth one, while fraction, square-root and one-cell matrix paths
refuse at depth two and fit at depth three. The direct-break depth boundary is
also retained.

No MTEF bytes, AST values or default profiles changed. Source-only checks were
`rustfmt --edition 2024` and `git diff --check`; no Cargo/native/rustdoc
command, binary, large evidence artifact or upstream checkout path is part of
this return.
