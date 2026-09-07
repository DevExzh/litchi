# 0463 integration and source custody

Baseline revision is `c079b1c59`. The only initial untracked file is user-owned
`docs/GOAL.md`, whose bytes remain unchanged and unstaged. All accepted ADR and
README hashes match the prior complete review. The previous goal batch made
progress by rejecting a measured shape index rather than retaining below-gate
complexity. No production source change from 0462 remains.

Baseline release build, binary binding, frozen primary and supplementary
protocols, baseline R1 (six lanes, 180 samples) and assembly all pass before
the implementation ACK. Every primary capture uses `check.py` and retains its
per-lane immutable oracle, binary, source, protocol and raw-artifact custody.
The practical gate remains a 3% normal medium/large p50 improvement in both
repeats with negative bootstrap upper bounds, plus individual adverse-metric
and memory review. No timing rerun or after-the-fact gate change is planned.

The candidate reuses the writer's strictly authored-audited XML result only
when a private proof binds the exact source and candidate archive owners and
conservatively proves all final per-part and aggregate audit limits. The proof
is returned with one serialization, never cached on the mutable draft. A miss
must use the existing validator. Candidate reopen and semantic readback,
source-reference precheck, media/domain checks and no-op/patch behavior remain
required. See `proof-design.md` and the final source review for the argument.

The prior commit profile is retained by hash and extracted rows in
`prior-profile-binding.json`; its inclusive FP/DWARF sample shares overlap and
include warmups. The ordinary baseline includes the accepted 0460 staging
fusion, which affects transaction setup. These diagnostic shares motivate the
candidate but are not independently removable costs or current ordinary-build
time attribution.

Primary lifecycle metrics, separate public-API phase clocks and whole-process
counters retain distinct scopes. Public phase clocks exclude setup, checks and
reporting, and warmup rows are discarded. Whole-process perf counters and RSS
include all lifetime work. No cold/range/native-GUI/worker-scaling claim follows
from these measurements. iWork remains outside the implementation scope.

A user-requested fresh fetch and rebase ran during implementation. The remote
was already an ancestor (83 commits ahead / zero behind), so HEAD did not
change. Autostash restored the two in-progress Rust files; all 72 then-present
source/evidence/GOAL files, the tracked diff and working-tree status were
verified unchanged. Baseline source and binary custody therefore remain valid.

Implementation review narrowed proof creation to an existing exact source
archive, matching the ordinary transaction scope. Source-less serialization
remains valid but receives no proof. The design table initially reversed the
final token/text limits; source inspection corrected it to 16 MiB token and
128 MiB text. Both remain no tighter than the writer defaults. No production
audit policy or frozen measurement protocol changed.

The first all-target Clippy gate failed with unchanged source custody: newly
imported `Arc` and `audit` made older fully qualified uses redundant under the
warning-denied policy, and six test limit constructors used an unsupported
`ConfigError` conversion. The failed receipt and log remain retained. The
corrected source must pass a separately tagged retry before measurement.

The first owner test run reached 194 passing unit tests and one failing new
RDF fallback diagnostic. Its fixture used a predicate IRI without the RDF
serializer's required namespace boundary. The predicate was corrected to
`https://example.test/predicate`; no production behavior changed. The failed
run remains archived and the complete owner suite is rerun under a new tag.

The final owner-test retry passes all 381 tests with no ignored or failed
tests. Warning-denied all-target Clippy passes on the same corrected source
epoch. The owner source manifest contains 7,030 files and is bound by hash in
the immutable receipts; benchmark builds and captures follow only after these
checks, with no Rust edits permitted during or after candidate binding.

The candidate passes the frozen normal medium/large latency gate in both
repeats with no adverse elapsed/RSS or allocation-increase flags and is
retained. All 387 harness tests pass (one ignored), as do warning-denied
rustdoc, scoped formatting and crate boundaries. Live precleanup and fresh-copy
portable replay pass; a resealed one-nanosecond summary mutation is rejected.
Cleanup inventories and removes four executable copies totaling 233,058,712
bytes. The owned staging directory and isolated verification copies are absent.
The user-owned GOAL file remains unchanged and unstaged; the full goal remains
open and this batch is progress through a measured production improvement.
