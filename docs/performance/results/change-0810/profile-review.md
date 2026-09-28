# 0810 Callgrind mechanism review

The four retained publications pass the exact owner and raw-format checks.
Each `namespace_uri_probe::capture_region_0793` publication has one incoming
owner call, owner self Ir of `10`, and an immediate-child inclusive partition
that reconstructs the owner summary. The two before summaries are
`549292102` and `549348645` Ir; the two after summaries are `537148235` and
`537183147` Ir. These are guest-instruction attribution values for the scoped
wrapper.

In both ordered repeats, the baseline scanner has one direct
`scan_processed_xml` → `NsReader<R>::process_event` edge with `282612` calls
and inclusive Ir of `46148677` and `46196379`. The candidate scanner has no
such direct edge. The scanner still calls `Reader<R>::read_event_impl` 282612
times in every publication; its inclusive Ir is `104447735` and `104463101`
before, then `104433533` and `104486156` after. This records the parser work
restructuring without treating the disappearance of one edge as a global
symbol deletion.

The global `NsReader<R>::process_event` symbol remains in the retained graphs:
the before publications have five positive incoming edges and 283665 total
incoming calls, while the after publications have four positive incoming edges
and 1053 total incoming calls. Those other callers are diagnostic evidence;
global zero-call absence is not required. `NamespaceResolver::push` remains
present with self Ir `13019166` in every before and after publication, so its
cost is retained for comparison.

The raw parser is the hash-bound `parse_raw` implementation from sealed 0784.
The four JSON reports use the frozen 0806 public-workflow schema and match the
sealed large-capture source, fixture, output identity, complete semantic
verification, and capture metrics. The numbered positive publication and empty
termination publication are both checked for every process.

Callgrind Ir is a guest-instruction attribution diagnostic. It makes no claim
about native latency, RSS, native cycles, phase fractions, or production
speedup, and Ir differences add no adoption threshold. Nested inclusive rows
are retained as diagnostics and are not added to the disjoint owner partition.
