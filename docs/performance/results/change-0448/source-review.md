# Source and ADR review

TransferDelayPolicy has two explicit values. Existing constructors and CLI default
to SeparateSleeps. MinimumService starts a monotonic Instant immediately before
the existing fixed wait; fixed-delay accounting and the wrapped read keep their
prior order. After a valid nonempty return, nominal transfer targets are computed
with the unchanged checked integer formula. The remaining wait is the saturating
difference between checked fixed-plus-transfer duration and elapsed service time.
An already satisfied target skips the second sleep. No target is shortened.

The deadline helper checks combined Duration overflow and exact below/equal/above
boundaries. Per-read target counters and byte/short-read counters are identical
between policies. Empty/zero-cap/EOF/error paths retain prior behavior. The new
policy adds no shared queue or aggregate link state; it remains per-call behavior.
Elapsed source work counts toward the floor, so minimum service differs from an
additive fixed-plus-transfer wait and must be selected explicitly.

The CLI allows the new policy only with an explicit transfer rate, whose existing
range-provider and bounded-rate checks still apply. Reports name the actual policy.
The nominal transfer counter is a target component, not actual sleep duration.
Existing publication, source-version, cache, budget and exact-output gates are
unchanged; the real PPTX equivalence test includes both policies.

Only two standalone harness modules change. No production Rust, public document
API, dependency, unsafe-code policy, hidden executor or ambient network behavior
changes. The new monotonic clock stays in opt-in tooling around caller ReadAt.
Accepted ADR contracts (0002/0003/0005/0006/0008/0023/0024) remain intact. The
accepted tree is c950b6c8be822561b498d7bbe87c460873dcbf49, unchanged from its prior
complete read. Native, cold I/O, bounded append and scaling gaps remain outside
this calibration and are not marked complete.

Final review: all owned tests, builds, captures, profiles and derived replay
processes are terminal. Both current Rust files equal the exact retained tested
after-copies. Production Rust, accepted ADRs and sealed 0447 evidence are unchanged.
The user-owned GOAL.md retains its pinned SHA-256 and remains unstaged.
The minimum-service timer starts before fixed waiting, preserves delegation/error
order, and never treats nominal counters as actual sleeping time. Default policy
remains separate sleeps. Three focused tests and real PPTX equivalence pass.
