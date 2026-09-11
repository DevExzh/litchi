# 0515 bounded profile final review

This review closes the read-only scope check for the 0515 XLSX attribution
bundle. Production and benchmark sources remain unchanged at
`b698732363250547f05d235b8c8c090f2dd17c59`. The bundle records diagnostic
instruction attribution only. It authorizes no speedup, latency, memory,
allocation, cache, scaling, or hardware-counter claim.

## Captured scope

The two profile lanes use the existing
`xlsx_one_percent_commit` case with the `dense-wide` corpus, three samples,
and no warmups. The corpus has two 256-by-256 worksheets, and each commit
updates both worksheets. The commit selector collects the three timed
`Edit::commit` bodies after the runner reset. Its direct edges show six
`Worksheet::store` calls, six rewrites, six changed-worksheet compactions, and
six direct post-compaction worksheet parses. The compaction selector collects
only the six `changed_worksheet` bodies. Both selectors use
`--collect-atstart=no` and reset at
`litchi_perf_baseline::run_xlsx_update_commit`. Fixture construction,
open/staging, final runner readback, and returned-`Commit` drops remain outside
the selected profile interval. Publication checks executed inside
`Edit::commit` remain collected by the commit selector; they are outside the
compaction-only interval.

Global `--separate-callers=3` is used only by the commit lanes. It preserves
the two parser roles in the direct call graph:

```text
run_xlsx_update_commit -> Edit::commit -> raw::worksheet::parse
run_xlsx_update_commit -> Edit::commit -> Worksheet::store -> raw::worksheet::parse
```

The first edge is the parser over the changed compacted output. The second is
the source Store parser used before the edit projection. The compaction lanes
have no parser context split because their toggle ends at the compactor.

## Raw profile and repeat checks

The independent analyzer replayed the selected positive direct edges from the
raw Callgrind files and matched their inclusive annotation edges. The raw files
also retain ten zero-call records per lane; these are valid Callgrind metadata
left by collection reset and are excluded from invocation counts. The selected
direct edges are all positive and have these exact counts:

| lane | raw `Ir` summary | runner → commit | changed-output parse | source Store parse | compaction |
| --- | ---: | ---: | ---: | ---: | ---: |
| `commit-r1` | 11,908,886,205 | 3 | 6 | 6 | — |
| `commit-r2` | 11,907,930,477 | 3 | 6 | 6 | — |
| `compact-r1` | 2,426,404,049 | 3 | — | — | 6 |
| `compact-r2` | 2,426,404,144 | 3 | — | — | 6 |

The retained raw SHA-256 values are `2e7abe22e431fd71d2ebc25edbf4e580f3589c7eced37ce4b2757658706c64c0`,
`39d99ba46efd2eb2f5217a59a345ac55373d7993fb11d0a3284ba3842707073e`,
`a1ca644d3ff934a15210dfff5148ee080a02c6860c0da3bb0006950b2f8fd2f6`, and
`c66a7b60c9640f78522220592052eac30693ce9729e7d6f5428f42ee6e0fc951` in that
order. Each capture receipt binds its raw file, report, catalog, and log;
each annotation receipt binds both rendered files to its raw profile.

The direct commit-child rows in the first inclusive annotation are:

| direct child of `Edit::commit` | `commit-r1` | `commit-r2` |
| --- | ---: | ---: |
| worksheet rewrite | 3,236,498,389 (27.18%) | 3,236,508,556 (27.18%) |
| `Worksheet::store` | 3,073,085,526 (25.80%) | 3,071,661,470 (25.80%) |
| changed-output `raw::worksheet::parse` | 3,069,117,648 (25.77%) | 3,069,401,244 (25.78%) |
| `changed_worksheet` | 2,426,404,081 (20.37%) | 2,426,409,489 (20.38%) |

These four direct regions account for about 99.13% of each commit profile.
The rows are disjoint direct children. Descendant rows such as
`Parser::parse`, `quick_xml`, MCE, scanner, and semantic materialization are
inside those parent costs and must not be added again. In the compaction-only
lane, `changed_worksheet` is 100% of that lane's collected `Ir`; this does not
make it 100% of a complete commit.

The largest repeat difference among those four direct rows is below 0.05%.
The compaction summaries differ by 95 `Ir`. This is repeatability evidence for
the selected boundaries; it is not a candidate-versus-control comparison.

## Decision-relevant descendants

In the changed-output parser context of `commit-r1`,
`Reader::read_event_impl` (4.86%), `NsReader::process_event` (1.90%), and
`NamespaceResolver::set_level` (0.14%) total about 6.9% of the commit
summary. Sharing that reader layer is a plausible next design target, subject
to an emitted-byte equivalence proof. `Parser::start` (8.02%) and semantic
`materialize` (2.05%) remain parser work that the attribution does not justify
removing. `NamespaceResolver::resolve_event` (1.91%) remains required unless
separate correctness evidence covers its namespace-resolution behavior.

The compaction-only direct-child rows use the compaction-lane denominator and
therefore do not describe commit cost:

| Compaction child | R1 Ir | R2 Ir | Share of compaction lane |
| --- | ---: | ---: | ---: |
| `compact::write_start` | 607,008,921 | 607,009,111 | 25.02% |
| `Reader::read_event_impl` | 579,317,197 | 579,316,373 | 23.88% |
| `web::Probe::observe` | 465,489,606 | 465,489,606 | 19.18% |
| `NsReader::process_event` | 226,209,164 | 226,209,362 | 9.32% |
| direct attribute iteration | 183,798,702 | 183,798,702 | 7.57% |
| `Writer::write_event` | 90,967,146 | 90,967,146 | 3.75% |

These are sibling direct children of `changed_worksheet`; nested rows remain
inside their parent costs. The event reader and namespace processing in this
table explain why compaction still has substantial work after any parser-reader
sharing seam.

The roughly 25.77% changed-output parse row is the complete selected parser
boundary. It must not be treated as wholly removable. Style validation,
requested-change checks, web checks, MCE/x14ac handling, exact-output fallback,
package reopen, and publication checks remain semantic obligations. A future
writer-fed parser candidate needs exact compacted-byte and parsed-state
differential checks, with fallback to the current exact-byte parser whenever
serialization or preservation proof is incomplete.

No operation-local allocation or phase allocation measurement was collected.
The normal reports correctly publish allocation vectors as unavailable, so no
allocation conclusion is inferred from Callgrind instruction counts.

## Identity and profiler limits

All seven capture receipts report exit status 0, the same release binary
SHA-256 `4d5193b285544315399425800903a669700f7bfba5b4f51645df289525bc3106`,
and the same source-manifest SHA-256
`2482a2f009ab9213d663a6f8a62607398a72006c13f451a56da8dc78ba80ee5c`.
The current 7,196-file source manifest matches the 0514 before manifest, the
current source tree, and the declared base revision. Normal-r2 also completed:
the existing 12-row native control used 30 samples and three warmups per row,
one worker and one child process per capture (720 durations across the two
captures); iterations execute in the harness's in-process loop;
the dense-wide one-percent commit repeat changed p50 by -0.647%, mean by
-0.855%, p95 by -1.158%, and p99 by -1.451%.

Every Callgrind log retains the `brk segment overflow` warning. `Ir` is
simulated instruction attribution, not cycles. Profiler elapsed time and RSS
are excluded from native comparisons, and the normal 30-sample control is
descriptive. Global caller separation can enlarge profile and annotation
files; apostrophe-separated context suffixes must be normalized only for
symbol matching while retained for role classification. Reset-active ancestor
rows are accounting scaffolding rather than extra phase work.

The independent annotation replay in `annotation-replay.json` completed for
all eight raw-to-rendered pairs: fresh `callgrind_annotate` output had the same
function blocks and incoming/outgoing edges. Fresh bytes can differ because
equal-cost rows may be reordered by Perl hash iteration. The direct-edge
analyzer replay currently succeeds as well. Its final internal-child
extension adds the exact parser and compactor child tables recorded in the
0515 change report, and the refreshed `attribution-summary.json` matches that
extension. `verifier-sources.json` binds the current analyzer and verifier
hashes.

## Sealing status and custody

Evidence custody and replay are complete through scratch cleanup. The retained
`verification.json` verifier run exited 0 before cleanup, with summary SHA-256
`fc3f0e95b1cacd7940308b885bede3f53025507d559778a4756f8663f4b3931a` and
verifier SHA-256
`f6d6288b814fded7d0581844c880c7209931c91a30a613ad644ee930b0f44313`.
Its `owned_scratch_present: true` records the state at that run. `cleanup.json`
then removed only `/tmp/litchi-goal-0515` after finding no process references
(1,768 files; 1,040,297,984 unique-inode allocated bytes). The retained
`replay-after-cleanup.json` exits 0, confirms the same summary, and records
`owned_scratch_absent: true`.

Raw profiles, annotations, reports, catalogs, logs, receipts, and review
documents belong to the retained evidence bundle and remain in the repository.
At the independent review checkpoint, only checksum sealing remained. The
coordinator subsequently sealed all 84 evidence artifacts and verified their
exact inventory, filesystem hashes and staged Git bytes. `SHA256SUMS` excludes
only itself.

`cleanup.py` was limited to `/tmp/litchi-goal-0515` after its process-reference
check. It removed the scratch executable and target files as part of that
owned tree; the raw profiles, annotations, reports, receipts, source hashes,
and replay tools stay retained. No unrelated worktree or target directory is
within this cleanup scope.

This bundle therefore supports the next OLE2/OOXML design investigation. It
does not complete the optimization goal, and ODF remains deferred until that
goal is complete.
