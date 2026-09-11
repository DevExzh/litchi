# 0515: isolate XLSX changed-output work

`performance_claim: none`

`claim_authorized: false`

This evidence-only batch separates changed-output XLSX parsing from the source
`Worksheet::store` parse and measures changed worksheet compaction on its own.
Production and harness source are unchanged at base revision
`b698732363250547f05d235b8c8c090f2dd17c59`. The source-manifest SHA-256 is
`2482a2f009ab9213d663a6f8a62607398a72006c13f451a56da8dc78ba80ee5c`, and the
ordinary release binary SHA-256 is
`4d5193b285544315399425800903a669700f7bfba5b4f51645df289525bc3106`.

The result identifies the next XLSX work boundary. The full changed-output
`raw::worksheet::parse` call is about 25.8% of the selected commit profile,
but that row includes semantic materialization, XML namespace/MCE handling,
limits and other validation. The source review's conservative reader-sharing
ceiling is only about 6.9% of commit instructions. It is a design ceiling,
not a measured speedup or a promise that all of that work can be removed.
Changed worksheet compaction is about 20.4% of the commit profile. The four
profile lanes are diagnostic simulated-instruction evidence; they do not
authorize a production change.

## Existing output path and proposed seam

For an effective worksheet edit, the transaction obtains the source Store,
rewrites the worksheet, compacts the rewritten XML, parses the exact compacted
bytes for changed-output validation, finishes the web proof, validates styles,
checks every requested change, and clones/reopens the package before
publication. Metadata-only edits skip grid Store verification where the
existing transaction permits it, but still retain output web validation and
package reopen. Exact no-ops retain source/patch identity and skip the
changed-output path.

The next design to test is an output-aware parser feed at the compactor's
writer boundary, started only after the edit is known to be effective. Each
event actually emitted by the writer would feed an equivalent normalized event
to the existing worksheet parser; dropped formatting whitespace would produce
no parser event. The parser result would remain provisional until compaction
reaches EOF. Reuse would require all of the following:

- the feed observes the writer's actual normalized names, attributes, escaping
  and omission decisions, rather than copying source events;
- output UTF-8 and XML well-formedness checks remain in the path;
- x14ac capture and MCE processing run against the output and retain their
  current ordering and limits, with an acceptable borrowed MCE result;
- worksheet depth, value, cell, web and other existing limits remain owned by
  their current boundaries; and
- the provisional result is discarded and the current parser reads the exact
  compacted bytes whenever any proof gate fails.

The exact-byte parser remains authoritative fallback. The single-quote
attribute case demonstrates why: compaction can rebuild a start tag and the
writer can normalize quote delimiters, so an observer of the source event is
not proof of the bytes that will be published. The first safe slice may be
limited to ordinary MCE-free and x14ac-free output with proven-safe
attributes; all other output falls back. Style validation, requested-change
checks, web validation, package reopen, the bounded Store handoff and current
error ordering remain required. In particular, the candidate must preserve
`compact > grid > web` failure precedence, x14ac input ownership, XML
`xml:space`/CDATA/entity semantics, namespaces and MCE, unknown content,
styles, formulas, metadata, and exact no-op behavior.

The detailed source and ADR review is retained in
[`output-semantics-review.md`](../results/change-0515/output-semantics-review.md)
and [`adr-review.md`](../results/change-0515/adr-review.md). The review
describes the seam and proof requirements; it does not describe an implemented
candidate.

## Frozen measurement scope

The normal lane uses the existing four cases
(`one_cell_commit`, `one_percent_commit`, `one_cell_commit_save` and
`one_percent_commit_save`) across `tiny`, `medium` and `dense-wide` shapes.
Each of the two serial CPU-2 repeats has 30 samples after three warmups, one
worker and one child process per capture, retaining 720 durations in total.
Each iteration reopens and stages a fresh workbook outside the timed interval. The
preflight has one sample and no warmup. The normal timer boundaries and sink
behavior are unchanged: setup, staging, expected-output generation, final
readback, post-clock oracles and drops stay outside the timer; commit/save
does not include sink reservation or expected-output generation. Internal
temporary work and destruction inside the operation remain included.

The commit profiles use the dense-wide one-percent commit case, three samples
and no warmup, `--collect-atstart=no`, an exact commit toggle, a reset at
`run_xlsx_update_commit`, and `--separate-callers=3`. Each has three direct
`Edit::commit` calls, six direct changed-output parse calls and six direct
source Store parse calls. The compaction profiles use the same case and reset,
toggle only `changed_worksheet`, and contain six direct compaction calls: two
changed worksheets across three commits. Fixture construction, open/staging,
final runner readback and returned-Commit drops are outside the selected
collection windows. Publication checks inside commit remain included in the
commit profiles; they are outside the compaction-only interval.

The direct parser contexts are disjoint at this scope:

```text
run_xlsx_update_commit -> Edit::commit -> raw::worksheet::parse
run_xlsx_update_commit -> Edit::commit -> Worksheet::store -> raw::worksheet::parse
```

The first is the parse of changed compacted output. The second is the source
Store parse before projection. Nested `Parser::parse`, quick-xml, MCE and
semantic rows overlap their parent and are diagnostic descendants. The
Callgrind reset can also leave synthetic active-ancestor rows in the exclusive
rendering; those rows are accounting scaffolding and must not be added to the
selected edges. Collection-off descendant call counts are metadata, not
measured stores or parses.

## Normal repeat rows

The table reports the exact p50 timer values from the two ordinary release
lanes and their same-build repeat change. All rows have no adverse
p50/mean/p95/p99 drift flag. The frozen thresholds are 5%, 5%, 10% and 15%
respectively.

| Shape | Case | R1 p50 (ms) | R2 p50 (ms) | p50 repeat change | Flags |
| --- | --- | ---: | ---: | ---: | --- |
| tiny | `xlsx_one_cell_commit` | 0.163245 | 0.161960 | -0.7872% | none |
| tiny | `xlsx_one_percent_commit` | 0.302766 | 0.299716 | -1.0074% | none |
| tiny | `xlsx_one_cell_commit_save` | 0.203606 | 0.204371 | +0.3757% | none |
| tiny | `xlsx_one_percent_commit_save` | 0.367766 | 0.367196 | -0.1550% | none |
| medium | `xlsx_one_cell_commit` | 1.768922 | 1.759206 | -0.5493% | none |
| medium | `xlsx_one_percent_commit` | 7.071792 | 7.041130 | -0.4336% | none |
| medium | `xlsx_one_cell_commit_save` | 2.200053 | 2.224253 | +1.1000% | none |
| medium | `xlsx_one_percent_commit_save` | 8.758018 | 8.813127 | +0.6292% | none |
| dense-wide | `xlsx_one_cell_commit` | 109.401480 | 108.489507 | -0.8336% | none |
| dense-wide | `xlsx_one_percent_commit` | 221.568615 | 220.135346 | -0.6469% | none |
| dense-wide | `xlsx_one_cell_commit_save` | 153.434571 | 154.002499 | +0.3701% | none |
| dense-wide | `xlsx_one_percent_commit_save` | 307.731827 | 310.295771 | +0.8332% | none |

Across the same 12 rows, repeat changes are −1.0074% to +1.1000% for p50,
−0.8766% to +1.0695% for mean, −1.1580% to +2.4798% for p95, and −2.6009%
to +3.1528% for p99. None crosses its frozen review threshold. Whole-child
GNU-time maximum RSS is **138,480 KiB** in R1 and **135,732 KiB** in R2
(−1.9844%). This is repeatability context covering the entire child, including
setup and oracles; it is not operation-local RSS, document-memory evidence or
an optimization result.

## Profile attribution

The retained Callgrind summaries and direct-edge analysis are:

| Lane | Selected collection | Direct calls | Collected IR | Normalized share |
| --- | --- | ---: | ---: | ---: |
| commit-r1 | `Edit::commit` | 3 | 11,908,886,205 | 100% of commit lane |
| commit-r2 | `Edit::commit` | 3 | 11,907,930,477 | 100% of commit lane |
| compact-r1 | `changed_worksheet` | 6 | 2,426,404,049 | 20.37% of commit-r1 |
| compact-r2 | `changed_worksheet` | 6 | 2,426,404,144 | 20.38% of commit-r2 |

The commit profile totals differ by **−0.0080%** between repeats. The
compaction totals differ by less than **0.0001%**. Within the separated
commit profiles, the two direct parser roles are:

| Commit lane | Changed-output parse, direct `Edit::commit -> raw::worksheet::parse` | Source Store parse, direct `Worksheet::store -> raw::worksheet::parse` |
| --- | ---: | ---: |
| commit-r1 | 3,069,117,648 IR (25.77%) | 3,070,712,898 IR (25.79%) |
| commit-r2 | 3,069,401,244 IR (25.78%) | 3,069,288,836 IR (25.78%) |

The 25.77–25.78% changed-output row is the complete direct parser call and
includes its semantic/materialization and validation descendants. It is not a
removable-work estimate. Sharing only the reader portion has an approximate
6.9% commit-instruction ceiling under this profile; the exact writer-feed
proof and fallback requirements above determine whether any subset is safe.
Do not add the overlapping reader, namespace, MCE or parser-descendant rows to
the direct edge, and do not use the `changed_worksheet`-only denominator as a
commit speedup denominator.

The changed-output eager parser context further separates these direct children.
Each share below uses the whole commit profile as denominator; the rows are
siblings within the output parser and therefore disjoint:

| Output parser child | R1 Ir | R2 Ir | Commit share | Reader-sharing implication |
| --- | ---: | ---: | ---: | --- |
| `Reader::read_event_impl` | 579,317,679 | 579,317,340 | 4.86% | Duplicate XML event reading is the main target |
| `NsReader::process_event` | 226,209,740 | 226,209,490 | 1.90% | Reuse the compactor's namespace reader state |
| `NamespaceResolver::set_level` | 16,547,784 | 16,547,784 | 0.14% | Reader namespace-level maintenance belongs to that shared pass |
| `NamespaceResolver::resolve_event` | 227,733,072 | 227,733,072 | 1.91% | Still required for semantic event interpretation |
| `Parser::start` | 954,662,998 | 954,665,928 | 8.02% | Cell, row and attribute semantics remain |
| `semantic::materialize` | 244,645,456 | 244,651,489 | 2.05% | Typed value construction remains |

The first three children total 822,075,203/822,074,614 Ir, about 6.90% of
commit instructions. This is the measured work boundary for a shared reader,
not a prediction that a replacement will remove all of it. Additional event
normalization, eligibility and fallback work must be measured in a candidate.

The independent compaction-only profiles show why compaction cannot disappear
with the second reader: `write_start` takes 25.02% of compaction instructions,
the web proof 19.18%, direct attribute iteration 7.57%, and direct writer work
3.75%. Its event reader takes 23.88% and namespace processing 9.32%. These are
disjoint direct compactor children; their denominator is compaction, not commit.
The output writer and web proof still need to run. Small tag shortcuts are not
a substitute for measuring the larger duplicated reader work.

All four Callgrind profile logs report a `brk segment overflow` warning and exit
successfully. The warning, simulated `Ir` definition, and the three-sample
repeat boundary are retained with the raw profiles. They provide attribution
only; profiler elapsed time and profiler RSS are excluded from the native
evidence.

## Evidence status and limitations

Build, preflight, normal, commit and compaction receipts bind the same source
manifest, binary, corpus catalogs, logs and profile artifacts. The machine
readable reports are retained under
[`results/change-0515`](../results/change-0515/), including
[`plan.json`](../results/change-0515/plan.json),
[`scope-review.md`](../results/change-0515/scope-review.md),
[`normal-r1-report.json`](../results/change-0515/normal-r1-report.json),
[`normal-r2-report.json`](../results/change-0515/normal-r2-report.json),
[`commit-r1-inclusive.txt`](../results/change-0515/commit-r1-inclusive.txt),
[`commit-r2-inclusive.txt`](../results/change-0515/commit-r2-inclusive.txt),
[`compact-r1-inclusive.txt`](../results/change-0515/compact-r1-inclusive.txt),
[`compact-r2-inclusive.txt`](../results/change-0515/compact-r2-inclusive.txt),
the retained [`analyze.py`](../results/change-0515/analyze.py), and the
[`annotation-replay.json`](../results/change-0515/annotation-replay.json)
replay receipt. The independent replay found matching complete function
blocks and incoming/outgoing edges for all eight annotation renderings;
equal-cost display ordering may differ.

Normal allocator vectors are unavailable and no values are inferred. No
operation-local allocation, phase-memory, hardware-counter, cache-locality,
physical-I/O, throughput or worker-scaling claim is made. The normal rows are
same-build descriptive repeats below the registered 500-sample latency lane.
Callgrind is simulated instruction attribution, not cycles or wall-clock
timing. The final [verifier](../results/change-0515/verification.json) passes; its
[post-cleanup replay](../results/change-0515/replay-after-cleanup.json) exactly
matches the retained summary without the executable.
[Cleanup](../results/change-0515/cleanup.json) removed only the owned scratch:
1,768 files and 1,040,297,984 allocated bytes across unique inodes, with no
process references before removal. Generated bytecode caches were also removed.
The raw artifacts, replay tools and checksums remain committed.

The release build, workspace formatting, crate-boundary audit, all ten existing
strict registered claims and REPORT classification pass. This unchanged-Rust
batch uses fresh benchmark correctness oracles and replay/negative-vector
checks; it does not claim a new owner unit suite, fuzz campaign or native
Office validation run.

The broader OLE2/OOXML optimization goal remains active. Further ODF work is
deferred until that goal is complete, and iWork remains excluded.
