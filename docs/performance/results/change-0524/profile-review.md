# Change 0524 profile scope review

This review defines the matched baseline and candidate Callgrind attribution
evidence for the CFB/OLE2 visited-bit lookup/set experiment. It is a
mechanism diagnostic and does not decide native optimization admission. The
companion numerical analyzer owns native timing, allocation, correctness, and
quality gates.

## Scope proof

Each stage contains eight profile children: one `xls-owned` child and three
CFB shape children (`tiny`, `many-small`, and `few-large`) for each of two
repeats. Every child uses five measured constructor calls with
`--collect-atstart=no`, `--toggle-collect`, `--zero-before`, and
`--dump-after` bound to the exact selected owner. The XLS owner is
`SourceBackedWorkbook::from_read_at_with_limits`; its positive owner edge may
come through the `SourceBackedWorkbook::from_read_at` wrapper. The analyzer
records bounded positive ancestry from the benchmark runner and uses the
owner edge's positive `calls=1` record for constructor scope.

Each CFB child retains the corpus-generation setup call as its first numbered
dump. Its positive setup context is accepted only before the timed dumps.
Subsequent numbered dumps must have a positive benchmark-runner path and one
selected constructor call each. The final unnumbered dump must be a
`Program termination` dump with `summary: 0`; it is retained as process
cleanup evidence and is excluded from constructor attribution.

The parser requires exactly one positive incoming edge into the selected
constructor, matching the dump summary and the selected function's raw
self-plus-direct accounting. It rejects a dump whose owner caller is neither
the timed runner (or bounded runner ancestry) nor an allowed CFB setup context.
Child-call metadata remains diagnostic because collection is disabled outside
the exact owner toggle.

## Stage binding

Every build and child receipt carries both the artifact-stage source manifest
and an `execution_stage` plus `execution_manifest_sha256`. The binary path is
stage-local (`/tmp/litchi-goal-0524/<stage>/normal`), so profile reports cannot
silently use the other stage's executable. Baseline native repeat two is the
intentional exception: it retains the baseline binary and report directory
while running under the frozen candidate source workspace. The analyzer checks
that distinction explicitly instead of treating the execution tree as binary
provenance.

The profile report and its native-r1 counterpart must have identical normalized
corpus, source, sink, output, and result identities within each stage. Native
r1 is used for both profile repeats because the second native block occurs
later in the serial A1/B1/B2/A2 schedule. The matched comparison also requires
native identity parity between baseline and candidate before reporting owner-Ir
deltas.

## Attribution comparison

The optional `--compare` report aggregates the five timed constructor dumps
for every matched repeat and workload. It reports baseline/candidate owner
inclusive, self, and direct `Ir`, together with the requested descendant
functions: `collect_exact`, `claim_sector`, `validate_stream_allocations`,
`validate_physical_sector_layout`, and `load_fat`. These values identify where
the candidate changes the selected constructor's Callgrind cost. They do not
establish native latency, hardware instructions or cycles, allocation counts,
RSS, physical I/O, cold-cache behavior, scaling, or native Office-producer
behavior.

The native gate remains authoritative for admission. A lower owner `Ir` alone
does not satisfy the plan's matched p50, allocation, correctness, or quality
requirements. ODF remains deferred while the OLE2/OOXML optimization goal is
active.
