# 0667: the XLSX value-only editor retains and resolves shared strings

Status: retained and implemented in `litchi-xlsx`.
`performance_claim: none`; this record adds no claim-registry entry. The base
is `5fa92d7ce`, the landed 0657 implementation of the D4 dependency rule, and
the branch is `perf/0667-xlsx-value-editor-shared-strings`.

This is the explicit follow-on for the shared-string gap that 0657 left open.
Change 0657 widened the value-only editor's copied-element admission under
decision 7 of 0652, but deliberately kept 0602's D2 shared-string refusal. The
0667 assignment asks for that remaining D2 work under the queue and safety
trade-offs recorded by 0651 and 0652. The change models the one shared-string
relationship and its table for reads, retains the source part byte-for-byte,
and refuses an edit that would change a shared-string index. It does not author,
renumber, remove, or append shared-string entries.

## Authority and format rule

The implementing requirements are 0602 D2 and its gates: capture the table
lazily, supply it to every worksheet parse and reduced readback, preserve index
identity, charge the retained part before materialization, and refuse actions
that add, remove, or renumber an entry. The D4 dependency rule accepted by
0652 decision 7 supplies the admission boundary; 0657 explicitly recorded D2
as deferred, so this record does not treat the earlier record as having already
authorized shared-string modeling.

The local format reference is [ECMA-376 Part 1, 5th edition (December 2016),
§12.3.15, §§18.4 and 18.4.9, and §18.18.11](../../3rdparty/specs/ECMA-376/ECMA-376-1_5th_edition_december_2016.zip).
Section 12.3.15 specifies one internal Shared String Table part per package;
§18.4 describes the table as the workbook-wide indexed list; §18.4.9 defines
`count` as total text references and `uniqueCount` as table entries, including
formatting distinctions; and §18.18.11 identifies `t="s"` as the shared-string
cell type. Section 18.3.1.96 says the cell `<v>` is the index into that table.
These rules are why the relationship has a cardinality check, why an index is
resolved only when a worksheet references one, and why the source table cannot
be changed during a value-only edit.

## Implementation

`SourceCatalogCapture`, `OwnedCatalogCapture`, and `SourceState` now retain one
`SharedStringsState` containing the relationship identity, part URI/content
type, source bytes, and an `Arc<OnceLock<Box<[Text]>>>`. Capturing the part does
not parse it. The parser closure is passed through all source-backed, owned,
stream-observed, complete-fallback, rewritten, reduced-readback, and
multi-sheet paths. The first worksheet containing `t="s"` initializes the
table; a workbook with a shared-string relationship but no referenced cell
leaves the table untouched.

The retained part's bytes are charged to the existing
`MAX_MULTI_WORKSHEET_BYTES` aggregate (64 MiB) before a selected multi-sheet
snapshot is accepted, and the single-sheet routes apply the same bound to the
captured part. The table parser keeps its existing UTF-8, XML, count, text
length, rich-text, and allocation checks. A large generated witness with
66,935 entries (1,863,181 XML bytes) exercises this path within the bound.

Workbook validation permits at most one transitional or Strict shared-string
relationship. The relationship must be internal, target a part with the
shared-string content type, have no outbound relationships, and remain byte
and relationship identical when the source is checked again. Duplicate
relationships, an external relationship, a wrong target content type, or
outbound shared-part relationships fail before a value result is published.

The stored-cell support check now permits an existing `Stored.shared_string`
entry. The reduced readback therefore merges unchanged shared-string neighbors
with parsed edited cells while preserving their original physical indexes.
`Set`, `SetFormula`, `SetSharedFormula`, `Clear`, and `Remove` all call the same
staging guard. A target that already contains a shared-string cell returns
`invalid XLSX structure: value-only edits cannot add, remove, or renumber
shared strings`; edits to numeric or inline-string cells beside such a target
remain available.

No public API item or `litchi-opc` API changed. The shared-string state is an
internal XLSX source-closure detail. There is no consumer/API dependency on
0661's OPC lazy work; any merge conflict in shared traversal tests is a test
combination point, not an API contract overlap.

## Correctness evidence

The admission witnesses cover the positive and negative edges:

* `a_shared_string_part_is_admitted_and_preserved` edits a numeric neighbor,
  reads the shared text back, and verifies that the published
  `xl/sharedStrings.xml` member remains byte-identical.
* `a_shared_string_cell_cannot_be_mutated_or_removed` covers set, clear, and
  remove at a shared-string target and checks the common typed refusal.
* `a_shared_string_table_is_lazy_when_no_cell_references_it` supplies a
  malformed table with only numeric worksheet cells; the edit succeeds because
  the table is never materialized.
* `shared_string_cells_require_a_valid_table_and_index` covers a missing
  relationship, an out-of-range index, and a truncated table.
* `multiple_shared_string_relationships_are_refused_before_parsing` freezes
  the ECMA-376 cardinality rule.
* `a_large_shared_string_table_stays_within_the_retained_part_bound` resolves
  the 66,935-entry table and reads its final index after an adjacent edit.

The source-backed integration witness
`mce_relationships_and_signed_changes_are_refused` now edits a package carrying
the shared part and checks the published member. The existing facts oracle and
shared traversal tests pass with the table closure supplied on every parser
route.

## Measurement

The release producer harness was built with `CARGO_BUILD_JOBS=2` into
`/home/zhuhe/code/litchi-target-0667` and run with three warmups and 15 samples
per case. Its retained outputs are
[`bench/producer.json`](results/change-0667/bench/producer.json) and
[`bench/producer-evidence.json`](results/change-0667/bench/producer-evidence.json).
The generated medium and dense shared-string sheets each contain 40% shared
cells and 64 unique table entries. The evidence census records all five
previously refused producer facts, including the shared-string role, as
admitted on the read variant.

The timing run is a descriptive after-only observation, not a before/after
claim. The p50/p95 nanosecond pairs were:

| case | p50 | p95 | p95/p50 |
| --- | ---: | ---: | ---: |
| medium source open | 61,551 | 64,031 | 1.040 |
| medium source selected cell | 1,333,873 | 1,354,113 | 1.015 |
| medium source planning | 789,482 | 797,492 | 1.010 |
| medium source one edit/save | 1,769,955 | 1,800,285 | 1.017 |
| medium control selected cell | 538,471 | 540,462 | 1.004 |
| medium control planning | 693,422 | 699,342 | 1.009 |
| dense source open | 55,790 | 64,330 | 1.153 |
| dense source selected cell | 18,968,419 | 19,070,961 | 1.005 |
| dense source planning | 11,365,660 | 11,411,290 | 1.004 |
| dense source one edit/save | 23,568,962 | 23,831,772 | 1.011 |
| dense control selected cell | 7,917,801 | 8,135,322 | 1.027 |
| dense control planning | 10,107,727 | 10,358,417 | 1.025 |

The declared per-leg cleanliness rule is p95/p50 ≤ 1.05. The dense open case
is retained as an observation but fails that descriptive floor; the other eleven
cases meet it. No A/A pair, baseline checkout, cold-cache claim, allocation
claim, or speedup is registered. The large-table focused test passed in 0.34 s
with a 67,728 KiB maximum process RSS on the already-built debug test binary;
that RSS is a process observation rather than a package-wide memory claim.

## Limits and retained costs

The table is read-only. A cell using `t="s"` can be selected and read, but any
mutation at that address is refused. New shared-string authoring, table entry
deletion, table renumbering, shared-string rewriting, and edits requiring a
new shared index remain outside this editor. Pivot caches, tables, query tables,
and other value-dependent relationships remain refused. Only one shared-string
part is modeled, and a malformed or unsupported shared part still fails closed.

The producer harness uses a 64-entry table; the separate 66,935-entry witness
checks the retained-byte envelope and parser behavior but does not establish a
large-table throughput claim. The packet has no 0602-style real-fixture census
or paired before/after timing, because the existing real corpus and the
producer-shape publication route need separate follow-on work. The retained
shared-string XML is charged and preserved, while parsed `Text` entries remain
subject to the raw parser's per-string and allocation limits.

The coordinator should cherry-pick the commit named in the handoff and retain
the four packet log sections without adding this work to shared rollup files.
