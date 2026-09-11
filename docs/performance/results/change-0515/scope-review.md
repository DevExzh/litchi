# 0515 XLSX phase attribution scope review

This read-only review confirms that the frozen 0515 diagnostic plan can
separate the two largest changed-output boundaries without changing the
production or benchmark source. It uses the existing
`xlsx_one_percent_commit`/`dense-wide` runner and the normal release binary
from one source epoch. The result remains instruction attribution evidence;
it is not an optimization, latency, memory, or full-goal claim.

## Frozen workload and ownership

The selected command is equivalent to:

```text
taskset -c 2 valgrind --tool=callgrind --collect-atstart=no \
  --toggle-collect='*litchi_xlsx::workbook::edit::semantic::transaction::Edit*::commit' \
  --zero-before='*litchi_perf_baseline::run_xlsx_update_commit' \
  --separate-callers=3 ... \
  litchi-perf-baseline --case xlsx_one_percent_commit \
  --xlsx-shape dense-wide --samples 3 --warmup 0 ...
```

The commit profile collects exactly the three timed `Edit::commit` calls after
the runner reset. In the dense-wide corpus there are two 256-by-256 sheets,
and the deterministic one-percent update set touches both sheets on each
commit. The direct commit children are consequently six `Worksheet::store`
calls, six changed worksheet rewrites, six `changed_worksheet` calls, and six
direct post-write worksheet parses. Those counts are used only where the raw
profile records the direct edge; descendant call metadata is not an invocation
count because Callgrind can retain calls made while collection is disabled.

The second profile uses the same command and workload but toggles
`*litchi_xlsx::raw::compact::changed_worksheet` instead of `Edit::commit`:

```text
taskset -c 2 valgrind --tool=callgrind --collect-atstart=no \
  --toggle-collect='*litchi_xlsx::raw::compact::changed_worksheet' \
  --zero-before='*litchi_perf_baseline::run_xlsx_update_commit' ...
```

It therefore has six measured compaction bodies (two changed worksheets over
three commits). The compaction interval includes the changed worksheet XML
reader/writer and its web-binding observation. It ends before the transaction
parses the compacted bytes and before style validation, publication, caller
oracles, and teardown.

## Parser context split

The commit profile's global `--separate-callers=3` is sufficient for the two
relevant call chains:

```text
run_xlsx_update_commit -> Edit::commit -> raw::worksheet::parse
run_xlsx_update_commit -> Edit::commit -> Worksheet::store -> raw::worksheet::parse
```

The first is the changed-output validation parse over the compacted worksheet
bytes. The second is the source Store parse needed before projecting the edit.
Callgrind appends caller context to separated function names, so an analyzer
must normalize the base symbol and then classify the direct caller chain. It
must not look for one exact unsuffixed `raw::worksheet::parse` name or sum the
context rows with the unsplit function aggregate. The direct commit edge and
the direct parser-context rows are disjoint at the declared scope; nested
`Parser::parse`, `quick_xml`, MCE, and semantic rows overlap their parent and
are diagnostic descendants only.

Global separation is intentionally limited to three callers. It may enlarge
the Callgrind file and annotation substantially because every function can
receive context variants. This affects profiler storage and collection time,
not the native rows. A profile analyzer should use the direct `>` edges in the
inclusive annotation (or the corresponding `cfn` records in the raw file),
retain the context suffixes, and require one runner-to-commit edge with three
calls plus the expected six direct changed-output parser calls. The source
Store parser's function block can show extra collection-off calls; its
instruction total must be attributed from the collected context rather than
from that metadata alone.

The reset can leave synthetic active-ancestor rows in Callgrind's exclusive
annotation. Those rows are accounting scaffolding for the zeroed collection
window, not additional work; they must not be summed with the selected direct
edges or used as phase totals.

## Exact phase boundaries

In `workbook/edit/semantic/transaction.rs`, the source Store is obtained by
`Worksheet::store()` before the edit's effective-action projection. The
changed XML is produced by the worksheet rewrite, then
`raw::compact::changed_worksheet` emits compact bytes. Only after that call
does the transaction invoke `raw::worksheet::parse(compacted.bytes(), ...)`
for changed-output validation and style validation. This ordering means the
parser context split measures the post-compaction validator, while the
compaction-only profile cannot accidentally include that validator.

The selected commit-only case has no expected-output commit/save operation.
Fixture construction happens before `run_xlsx_update_commit`; the exact
`--zero-before` reset removes its costs. Workbook opening, edit staging,
post-clock semantic checks, final readback, and `Commit` drop are outside the
selected collection window. `collect-atstart=no` and the exact end-anchored
toggle prevent setup and caller work from becoming collected instructions.

## Identity, reproducibility, and limits

Both profiles must use the same normal release executable, source manifest,
Rust/toolchain settings, dense-wide corpus catalog, CPU-2 affinity and one
worker. Each receipt should bind the command, binary SHA-256, source-manifest
SHA-256, report/catalog, raw Callgrind output, log, and the two
`callgrind_annotate` outputs. The normal binary remains the ordinary system
allocator build; no allocator vector is inferred from this profile. A future
candidate needs a separate source epoch and executable, matched semantic and
exact-output checks, and its own operation-local allocation evidence.

Callgrind `Ir` is simulated instruction attribution. The `brk segment
overflow` warning observed in prior profiles must remain attached if it recurs;
successful exit and matching direct edges do not turn it into hardware or
wall-clock evidence. Profiler elapsed time and RSS are excluded from native
comparisons. The two 3-sample repeats establish boundary repeatability only;
they do not authorize a speedup claim. Raw profiles, annotations, reports and
receipts are retained committed evidence. After their hashes, source custody
and replay checks are sealed, only the owned 0515 build scratch, executables
and target files may be removed; no repository evidence, source or unrelated
worktree/target may be cleaned as part of that step.

The changed-output parser and compaction are adjacent but separate evidence
regions. Their instruction totals may guide the next XLSX design, but they
must not be added to nested parser/writer rows or generalized to other OOXML,
OLE2, ODF, or iWork formats.
