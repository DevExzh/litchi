# Log sections for change 0666

The coordinator can merge these four blocks into the newest position of the
named shared files. This packet does not edit those rollups.

## For `HOTSPOTS.md`

## 0666 — the worksheet MCE proof stops charging declarations cumulatively

Record: [0666-mce-rewrite-residue](../../0666-mce-rewrite-residue.md).
Change 0653's writer now emits namespace bindings at their source boundary,
but its worksheet `MceRewriteEquivalence` still charged every declaration
against every later start tag and rejected all prefixed attributes. Change
0666 removes that stale cumulative charge, passes the active resolver into the
proof and admits bound non-MCE prefixed attributes plus a validated
`mc:Ignorable`; the worksheet parser ignores those attributes. It keeps
prefixed elements, `ProcessContent`, `PreserveElements`, `PreserveAttributes`,
`MustUnderstand`, unknown MCE directives, unbound prefixes, `dyDescent` and
`AlternateContent` on the fallback path. The deterministic real-worksheet
census is 205 parts: 71 borrowed, 30 rewritten candidates, 25 validator-backed
shared completions and 104 admission fallbacks. No wall-clock claim is
registered.

## For `GOAL_AUDIT.md`

## 0666 — a bounded read-side proof removes stale work while preserving the source contract

Record: [0666-mce-rewrite-residue](../../0666-mce-rewrite-residue.md).
The change addresses `docs/GOAL.md`'s no-unnecessary-work clause on the
source-backed worksheet read path. A 288-declaration witness distributed over
96 tags completes the proof even though it exceeds the old cumulative 256
charge; each tag remains below quick-xml's unchanged 256-declaration bound.
The proof allocates no directive set, bounds `mc:Ignorable` tokens at the
existing 4096 MCE limit and falls back on any event-tree-changing directive.
The preservation floor is zero differences over every one of the 205 real
worksheet parts: successful values and exact errors match authoritative
preprocessing. The corpus measurement is scoped evidence, not a registered
timing claim; `dyDescent` still needs its separate x14ac value capture.

## For `REPORT.md`

## 0666 — MCE worksheet admission residue and adjacent consumer audit

Record: [0666-mce-rewrite-residue](../../0666-mce-rewrite-residue.md).
`raw/worksheet/mod.rs` now observes namespace declarations per start tag and
resolves prefixed attributes before deciding whether the MCE proof can replace
the processed stream. A bound non-MCE prefixed attribute is safe because the
worksheet parser does not read prefixed attributes; the value-only validator
still refuses `r:id` relationships inside `sheetData`. `mc:Ignorable` is
validated for NCName syntax, duplicate prefixes, binding, MCE self-reference
and the existing 4096-token bound. Focused adversarial tests cover valid
Ignorable with a bound ignored attribute, unbound/duplicate/MCE prefixes,
`ProcessContent`, `PreserveAttributes` and the distributed declaration chain.
The checked fixtures contain 130 `mc:Ignorable`, 104 `dyDescent`, three other
MCE-directive and zero `AlternateContent` worksheet markers; 25 of 30 rewritten
candidates complete the validator-backed shared path. Query-table extension
lists retain an owned XML tree, auto-filter and data-validation captures retain
bounded detached fragments, and named-sheet-view filter payloads are adapted
into a generated auto-filter root before typed publication, so none of those
four consumers is changed. No release benchmark or allocation measurement was
taken.

## For `ADR_COMPLIANCE.md`

## 0666 — no ADR amendment, limit movement or publication-byte movement

Record: [0666-mce-rewrite-residue](../../0666-mce-rewrite-residue.md).
Authority is decision 1 of change 0652; no accepted ADR is amended. ADR 0005's
mandatory validation and bounded-resource rules remain in force: the resolver
lookup and directive-token scan are bounded, `quick_xml`'s per-start
declaration limit is unchanged, and every proof refusal repeats the
authoritative validation and preprocessing path. ADR 0006's preservation
default remains the source editor's contract: the shared path only reads source
bytes, and the valid `mc:Ignorable` plus bound prefixed-attribute witness opens
and commits a source-backed no-op without changing source bytes. No public API,
error type, limit, dependency, unsafe block, global pool or publication writer
changes. The adjacent fragment consumers remain owned or namespace-adapted
because a borrowed slice would drop the standalone namespace context or typed
unknown-node payload they publish. `performance_claim: none`.
