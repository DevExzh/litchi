# Change 0666 — finish the worksheet MCE rewrite proof residue

Change 0653 left two worksheet read-side costs behind its namespace writer: the
`MceRewriteEquivalence` proof still charged namespace declarations
cumulatively, and it rejected every prefixed attribute even after change 0657's
D4 admitted unfamiliar attributes for byte-preserving value edits. Change
0666 carries decision 1 of change 0652 through the remaining safe part of that
surface, under change 0651 rows 4 and 5.

The proof now counts declarations on each source start tag. The writer emits a
binding at its source boundary and can hoist one only when a dropped ancestor
requires it; the worksheet proof still refuses prefixed elements and every
directive that can change the event tree, so cumulative source declarations
are no longer the relevant bound. The proof receives the active namespace
resolver, admits a bound non-MCE prefixed attribute because the worksheet
parser ignores those attributes, and validates `mc:Ignorable` with the same
NCName, duplicate, binding, MCE-recursion and directive-token checks as the
preprocessor. `ProcessContent`, `PreserveElements`, `PreserveAttributes`,
`MustUnderstand`, unknown MCE attributes, unbound prefixes, `dyDescent` and
`AlternateContent` retain the authoritative fallback path.

The focused suite adds a source-preservation no-op with a bound ignored
attribute, a differential set for malformed and semantic-changing MCE
directives, and a 288-declaration chain spread over 96 tags. The last witness
exceeds the old cumulative budget while staying below the per-tag quick-xml
bound. The real worksheet census covers 205 parts: 71 borrowed, 30 rewritten
candidates, 25 completing the value-only shared traversal, and 104 refusing
the admission surface before the rewrite proof. The corpus has 130 literal
`mc:Ignorable` worksheet markers, 104 `dyDescent` markers, three other MCE
directives and no `AlternateContent` worksheet marker. Every admitted and
fallback result remains transparent against the authoritative path.

The four adjacent processed-buffer consumers remain unchanged. Query-table
extension lists own an XML tree whose unknown nodes and namespace attributes
are published through the typed model. Auto-filter and data-validation
captures detach selected fragments from processed input into bounded writers;
their callers need standalone fragments and source-span replacement checks.
Named-sheet-view filter payloads are adapted into a generated auto-filter
root, and `adapt_filter_root` intentionally retains only modeled filter
attributes; its payload is parsed into the typed view and is not published as
the source slice. A borrowed buffer or direct slice in any of these sites
would change ownership or namespace-context contracts without evidence of a
safe reduction, so no consumer migration is claimed here.

No wall-clock or allocation claim is registered. The retained measurement is
the deterministic admission census and its zero-difference preservation
floor; the packet records the focused test command and the exact corpus
counts. The existing byte-preserving source editor remains the publication
authority, and no public API, accepted ADR, parser limit or error contract is
changed.

Base: `5fa92d7ce`. Branch: `perf/0666-mce-rewrite-residue`. Authority:
decision 1 of [0652](0652-owner-decisions-for-the-third-wave.md), with the
remaining MCE worksheet rows in [0651](0651-queue-refresh-after-the-second-wave.md).
Evidence packet: [results/change-0666](results/change-0666/README.md).

`performance_claim: none`.
