# Independent review and retention

Retain the final bound-first ASCII candidate. The independent source reviewer
and profiler both recommend scoped retention after the final captures; root
agrees. `performance_claim: none` remains, with no registered or general
CFB/XLS performance claim.

Correctness review confirms unchanged empty/NUL/first-forbidden/UTF-16-length
refusal precedence, exact ASCII keys, unchanged Unicode fallback, safe bounded
array indexing and no new heap allocation, cache, source read or public API.
Four independent differential tests and the existing Unicode exhaustive check
cover these invariants. The six quality gates and DOC/PPT consumer tests pass:
4,392 passed, zero failed, 27 existing ignored.

The initial ASCII-first candidate's 11–16% 54016 open regression was not
accepted. Its source, complete checks and captures are archived. The final
bound-first candidate avoids scanning overlong input merely to classify ASCII,
shrinks helper code from the initial 2,672 to 2,440 bytes, and restores measured
54016 opening to approximately baseline. Cold-profile and layout observations
are diagnostic; the exact cause of the initial regression is not isolated.

Final eighth-query stored medians improve about 13–22% on several owned
routes, with independent long-loop and counter confirmation. Simple workflows
improve about 16.5% owned and 8% file; formula-refusal workflows improve about
4–5%. Large 54016/default open-plus-eight remains approximately flat. Missing
indexed queries show no useful gain. These distinctions define the claim.

Retained costs are explicit: 45365-2 build-query medians regress 5.12–9.19%,
and its first owned open-plus-eight window rises 5.39%/3.36%. Other measured
45365-2 windows rise 1.05–4.42%. Build mean/tail flags and the noisy Simple/file
open p99 flag remain in the tables. Later-query improvements do not erase
these short-workflow costs; the 45365-2 construction path remains follow-up
work. No query-cache admission or allocation policy changed in this patch.

All 96 query allocation groups and 12 counted routes match exactly. Nevertheless,
helper stack reservation grows 128 bytes and code grows 807 bytes versus
baseline. Native-child RSS flags reach +8.72% for Simple/file, with additional
Plan1/file, Simple/owned and late-54016/file flags. These measured process
costs are accepted for scoped retention; no uniform memory reduction or RSS
bound is claimed. Their cause is not isolated by query allocator gauges.

The evidence uses warm file caches and one shared host. Physical cold storage,
remote/range latency, concurrency and cross-platform gains remain unproven.
Cargo build durations reflect incremental-target reuse/invalidation and are
not comparable clean-build performance. No native Office or new fuzz-campaign
claim follows this batch. Chain traversal remains the dominant warm profile
component, and the broader GOAL remains active.

Independent evidence review confirmed the final numerical tables, test and
corpus counts, regressions, RSS/profile/assembly costs, archive separation and
reproduction scope. It found no material discrepancy or remaining blocker.
