# 0689 review and disposition

Retain the candidate. Root coordinated implementation, tests, source review,
performance profiling and an independent evidence audit. Root alone ran the
Cargo and measurement lane; diagnostic instrumentation was restored before
final source checks. No iWork work or registered claim is included.

- Source review (`chain_next_design`): no production blocker. Existing CFB
  identity and offset checks protect restoration; publication remains complete
  scan only. Resolver-local state does not add a bytes/value/error cache.
- Implementation/test review (`sst_checkpoint_coder`, `sst_checkpoint_tests`,
  root): six added tests exercise SST offsets, ownership, concurrent clones,
  refusal/error ordering and accounting. Root corrected the backward-offset
  case to start from a nonzero SST index and fixed a test-only held cache owner.
  The initial failed lifecycle test remains archived.
- Performance review (`cfb_name_profiler`): retain scoped SST-heavy gains.
  Traces remove 478/97/24 repeated SST links; worksheet work and reads match.
  The 32-byte logical charge applies to every admitted index and stack
  reservation grows 48 bytes. Near-limit allocation geometry, opening/build
  costs and all original and supplementary tail flags remain disclosed.
- Evidence review (`evidence_review`): primary and supplementary numerical
  evidence passes. All 24,000 follow-up owners and 192,000 queries agree.
  Root strengthened trace assertions to bind both the retained starting point
  and final point to the baseline final sector. Wording now distinguishes
  indexed replay walks from preparatory walks, and ordinary-capacity allocation
  deltas from near-limit exceptions. No numerical/evidence blocker remains.

Final validation is 4,400 passed, zero failed, 27 existing ignored, with corpus,
source/hash, counted-I/O and repository evidence checks passing. The main
report preserves missing-target baseline drift and every tail trigger. Larger
controls do not reproduce two original anomalies; synthetic opening p99 costs
remain descriptive. There is no universal tail, RSS, cold-device, remote,
cross-platform or concurrency performance claim. Broader GOAL work remains.
