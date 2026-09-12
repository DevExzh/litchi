# 0528 verification review

The independent verifier implementation checks the unchanged source inventory,
accepted ADR and dependency bindings, exact prior seals, source replay, build
command and binary identity, disk-backed scratch placement, four profile
receipts and their raw reports, and the unique measured publication dump.
All numbered dumps remain included; lifecycle and measured callers are
verified separately. The final process dumps must contain zero Ir.

Publication annotations and raw incoming/direct edges are checked independently
of the report's claimed totals. The canonical publication and retained-scanner
analyses replay exactly. Native phase context remains distinct from fresh
instruction-only publication attribution; no latency gain is inferred.

Root completed the final verifier review. The CLI now writes only when given
an explicit output path and refuses output inside a sealed bundle. Cleanup
and inventory are independently verified instead of requiring a circular seal
hash inside the sealed decision. Placeholder assertion-failure probes were
removed because they did not exercise actual validators. No tamper-test pass
is claimed for those placeholders.

The added quality component validates exact source equality with 0527 final,
all 12 prior successful quality receipts and their logs, and the 1,297 recomputed
test executions. It separately binds and reruns the two actual annotation-name
regression tests. This is source-matched reuse for a source-unchanged profile
batch, not a new broad Rust test execution.

The source/pipeline review identified that both streaming and slice auditors
currently create fresh duplicate-key iterators. A future zero/one-attribute
probe can preserve error precedence by replaying the checked path on every
error or second result, with no state mutation during probing. This is a
reviewed candidate design, not approved production code or a speedup result.

Before cleanup, all source/build/profile/analysis/quality/decision components
must pass with live binary custody. After cleanup, the complete bundle must
pass again with absent owned paths and an exact SHA256SUMS inventory. The
verifier reports incomplete or failed evidence separately; it never authorizes
a production performance claim for this attribution-only plan.

Commands:

```sh
python3 -B docs/performance/results/change-0528/verify.py --component precleanup --strict
python3 -B docs/performance/results/change-0528/verify.py --component all --strict
```

Post-cleanup build replay validates the complete cleanup receipt before accepting
absent temporary binaries. Binary identity remains bound to its original lexical
scratch path; resolving a removed symlink is not used as a storage identity check.
Live binaries must still resolve to the exact owned disk-backed directory.
