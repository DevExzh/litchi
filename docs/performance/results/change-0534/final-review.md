# 0534 final disposition and custody review

The candidate is rejected by the frozen native gate: all eight primary XLS
p50 comparisons are slower, by 1.0596–10.6386%. Constructor and physical self-Ir
gates pass, as does the allocator guard, but those results do not override
native admission. All 86 matched adverse and 67 same-build flags are retained
and individually reviewed; no sample was discarded or cause inferred.

The final runtime is exactly the committed 0533 baseline. Only two independent
physical-layout contract tests remain, adding 116 lines in
`crates/litchi-cfb/src/file.rs`. Baseline plus `tests.patch` reconstructs the
final source; baseline plus `candidate.patch` and `tests.patch` reconstructs
the measured rejected source. The final source manifest is
`9d6a53738f299107582b71e5ddab03922b9646771df4d77b639b3fdc7fd6ff75`, and the
final Rust file hash is
`72fb5295dfbba5f3ed2192a9c24ecbf7d8aa2759f418b77563a4e2ff1042d03b`.

Both candidate and restored source passed all 14 quality gates and 4,382 test
executions per source state: CFB all-features 310, CFB no-default 310, XLS
1,345, DOC 1,187, and PPT 1,230. Across the two quality runs there are 8,764
successful test executions. The five in-memory verifier corruption probes,
physical gate controls, and distinct-stage quality accounting also pass;
these probes are separate from the Rust execution counts.

The campaign retains 48,000 native samples, 1,440 allocator samples, 16
constructor profile children with 80 timed and 12 separate setup dumps, and
four whole-child hardware captures. The verifier binds exact source, binary,
commands, scripts, corpora, output identities, raw vectors and derived reports.
It replays all evidence after cleanup and validates all 118 successful
receipts in serial order: baseline 45, candidate 59 and restored final 14.

Both owned paths, `/tmp/litchi-goal-0534` and
`/home/zhuhe/litchi-goal-0534-target`, were removed after checking accessible
process references. No owned Python cache remains. Full post-cleanup replay
passed, and the final SHA256SUMS inventory includes this review and the
verification receipt. No production performance improvement is retained.

OLE2 and OOXML remain active, with the collector identified for the next
bounded source/code-shape investigation. The rejected paired-prefix loop,
visited-bit fusion and freshness-session mechanisms remain rejected. ODF is
deferred until the OLE2/OOXML optimization goal completes; iWork is excluded.
This completed batch does not complete the broader goal.
