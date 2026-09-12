# Change 0523 profile scope review

This review records the fresh baseline Callgrind attribution for the CFB/OLE2
constructor experiment. It does not compare a candidate, select an
optimization, or make a speedup claim. The executable, source manifest, and
capture receipts are bound by the hashes checked in
[`analyze_profiles.py`](analyze_profiles.py); the normalized corpus and output
identities are also checked against the native reports.

## Scope proof

The profile matrix has eight children: one `xls-owned` child and three CFB
shape children (`tiny`, `many-small`, and `few-large`) for each of two
repeats. Every child uses five samples, `--collect-atstart=no`, and the exact
`--toggle-collect`, `--zero-before`, and `--dump-after` owner. The XLS owner is
`SourceBackedWorkbook::from_read_at_with_limits`; its positive owner edge comes
from the `SourceBackedWorkbook::from_read_at` wrapper. The `run_xls_owned_source_case` runner reaches the wrapper
through a positive raw instruction-cost edge whose Callgrind
call count is zero because the surrounding path was outside collection. The
analyzer records that bounded ancestry and uses the owner edge's positive
`calls=1` record for the constructor scope.

Each CFB child retains six numbered dumps. Part 1 is the corpus-generation
setup call; its positive owner caller is the inlined
`litchi_perf_baseline::run::{{closure}}` context. Parts 2 through 6 have a
positive `run_cfb_open -> OleFile<R>::open` edge and one owner call each, so
only those five parts enter operation attribution. The final unnumbered dump
is retained and has `Trigger: Program termination` with `summary: 0`. The XLS
children have five numbered timed dumps and no setup owner call.

The parser requires exactly one positive incoming edge into the selected
constructor, matching the dump summary and the selected function's raw
self-plus-direct accounting. It rejects a dump whose owner caller is neither
the timed runner (or bounded runner ancestry) nor an allowed setup context.
Child-call metadata remains a diagnostic: collection is disabled outside the
owner toggle, so context records with zero call counts do not establish a
timed invocation by themselves.

## Exclusive constructor descendants

The table reports raw exclusive `Ir` (`self_ir`) summed over the ten timed
constructor dumps for each workload (five per repeat). Percentages use the
selected constructor's summed inclusive `Ir` as the denominator. These are
mechanism diagnostics for this fixed baseline workload.

| Workload | Constructor Ir | `collect_exact` | `claim_sector` | `validate_stream_allocations` | Physical reconciliation | `load_fat` |
| --- | ---: | ---: | ---: | ---: | ---: | ---: |
| XLS owned source | 27,947,884 | 11,202,280 (40.08%) | 4,979,100 (17.82%) | 3,958,820 (14.17%) | 3,983,420 (14.25%) | 501,640 (1.79%) |
| CFB tiny | 490,284 | 10,400 (2.12%) | 900 (0.18%) | 8,570 (1.75%) | 860 (0.18%) | 3,490 (0.71%) |
| CFB many-small | 28,061,254 | 1,528,970 (5.45%) | 92,100 (0.33%) | 1,074,490 (3.83%) | 73,820 (0.26%) | 11,130 (0.04%) |
| CFB few-large | 26,145,238 | 11,143,890 (42.62%) | 4,954,650 (18.95%) | 3,935,640 (15.05%) | 3,963,860 (15.16%) | 533,630 (2.04%) |

The named functions are summarized per actual selected constructor and retain
their positive incoming edges, self cost, direct cost, call count, and caller
records in `profile-analysis.json`. The XLS Callgrind stderr retains the
reported `brk segment overflow` diagnostic; it is carried as a warning and is
not converted into a performance claim.

## Limits

Callgrind `Ir` and annotation rows do not establish native latency, hardware
instructions or cycles, allocation counts, RSS, physical I/O, cold-cache
behavior, scaling, or native Office-producer behavior. ODF remains deferred
while the OLE2/OOXML optimization goal is active.
