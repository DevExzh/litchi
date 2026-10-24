# 0362ul5 — Final inventory semantic review returns

The frozen `49c25ef67f` capture is diagnostic evidence, not proof that the
remaining heuristic rows can be waived. Source review found further required
engineering after the initial limits, shared-commit and alias migrations.
The following decisions preserve 0351 and the owner's requirement that only
the comprehensive redaction audit in item 22 is deferred.

## Required source corrections

* DrawingML's proposed detached-family migration is returned before merge.
  Ink actions still publish through `Edit::finish` and `prepare`, while
  `commit` is an alias. Its checked owner lacks canonical edit/replay entry
  points. Chart, transform and theme patches also expose an alternate
  `Patch::apply(&Package) -> Package` publication route. The follow-up must
  use the owner entry points and shared Commit, remove compatibility methods,
  place fresh-authoring drafts under `builder`, and migrate every consumer.
  Equal-content replay, exact no-op bytes, explicit limits and O(1) edit
  entry remain required. Passing the 22 heuristic rows alone is insufficient.
* The shared spreadsheet Chart is an authoring draft, not a parser AST or
  immutable projection. Its exports through XLSX/XLSB chart modules must
  move to coherent builder paths. This requires consumer migration, not a
  new non-debt disposition for `Chart::add_series`.
* DOC/PPT/XLS raw Reader mutability is explicitly permitted by 0351 §2.1.6.
  That permission does not excuse implicit raw limits, foreign error results,
  or an unexamined output receiver. A dedicated lane reviews those concrete
  signatures. Canonical shared value types remain allowed: their presence
  alone does not require redundant format wrappers. Error ownership and
  typed causal sources are distinct from value sharing.
* The coordinator now assigns the CFB and ZIP follow-up lanes scheduled by
  0362rd: seven CFB and eight ZIP limits entries receive canonical names and
  required raw limits, with every caller migrated in the same changes. Their
  scheduling boundary is not used as an accepted API exception. Algorithms,
  default choices at callers, streaming behavior and bounded memory must be
  preserved.

The external semantic audit has 62 source-review rows (44 raw Reader, three
owner-name heuristics and 15 transport limits entries) under
`/home/zhuhe/.cache/litchi-api-contract/partial-diagnostic-49c25ef67f62a8bb943e7c6a0bf464e49ae87d07/semantic-review-raw-ast-substrate-2026-10-24.tsv`.
It is a review queue, not validator input or an approved disposition. In
particular, its foreign-value observations require the narrower distinction
above; they are not blanket wrapper requirements.

## Tooling and dispositions

The result-contract resolver is also returned for fail-closed corrections.
Foreign Commit aliases must not disappear when their exact defining type is
absent or unauditable. Duplicate defining graphs and a foreign indexed item
masquerading as a local canonical path must not establish coverage. Alias
identity must resolve transitively, with cycles and ambiguous targets rejected.
These checks supplement the facade-specific accessor contract and the replay
registrations; they do not relax them.

The two Formula setters mutate an arena AST value. The six IWA codec request
types and thirteen generated protobuf types listed in 0362ul2 are also values,
not package edits. Exact semantic disposition categories may describe these
facts, with source evidence and strict row matching. They are neither builder
drafts nor deferred work. No raw-reader blanket category is approved while
its limits/error review is pending.

The all-68 inventory, glossary, complete gate, OfficeCLI corpus replay and
native qualification must bind the resulting assembled source. Earlier
capture and test receipts retain their original source and dependency locks.
No final inventory closure or performance acceptance is claimed here.
