# Detached ink-action CRUD allocator/runtime evidence

This directory contains an isolated, freeze-gated profile harness for the
detached `litchi_drawingml::ink::actions::edit` API. It measures the public
`Draft`, `Edit`, `Prepared`, `Commit`, and `Patch` workflows without adding a
production benchmark dependency. A process-local counting allocator records
requested allocation bytes and incremental peak live bytes; `/usr/bin/time -v`
records whole-process RSS.

The exploratory matrix uses deterministic direct-action sources at 8, 128,
and 1,024 actions. It covers detached draft creation, one-operation scalar
edits, exact source-backed no-ops, structural add/remove/move, and caller
output-cap refusals at each baseline count. Scaled and near-limit batch lanes
queue many distinct scalar edits, repeated writes to one scalar, root
insertions, ID-bearing removals, clear operations, and disjoint ID-bearing
moves. A detached opaque namespace payload lane and every
source-backed fixture retain comments, XML identifiers, and opaque vendor
payloads. Commits are checked after timing through their source-checked patch
and inverse, while no-ops also require source byte identity and source
allocation sharing.

The repeated runner is intentionally gated and must be run from an isolated
detached clean Git checkout at the approved committed head; the shared
worktree is not a valid build source. Before building, it resolves Cargo
metadata and rejects every path-backed package file and retained extra that is
modified or untracked relative to that checkout's `HEAD`. Registry packages
remain external dependencies. The isolated harness's committed
`harness/Cargo.lock` is authoritative; the repository root `Cargo.lock` and
`docs/GOAL.md` are not build inputs.

The frozen owner and semantic tests have already passed their repository gates:

```sh
PROFILE_FROZEN=1 \
  CARGO_TARGET_DIR=/var/tmp/litchi-ink-action-edit-profile-target \
  bash docs/report/spec-gap-validation-evidence/ink-action-edit-performance/run_profile.sh
```

The bounded correctness smoke can be reproduced without starting the repeated
profile:

```sh
PROFILE_FROZEN=1 \
  bash docs/report/spec-gap-validation-evidence/ink-action-edit-performance/smoke.sh
```

The final receipt will use three fresh processes, twenty measured samples per
lane, and two warm-ups. It will retain raw JSON, allocator counters, elapsed
samples, incremental peak-live deltas, `/usr/bin/time -v` receipts, source
manifests and hashes, host/toolchain data, exact commands, and a report
recomputed from the raw samples. A separate freeze-gated smoke runner uses one
warm-up and one sample to exercise the batch correctness lanes without
standing in for the final profile. The profile makes scoped absolute
observations; it makes no before/after speedup, native-application, or
package-wide performance claim.

The opaque payload lane measures the detached API's bounded copy of a complete
namespace-bearing payload. It does not claim a full InkML model or action
execution semantics. The source-backed lanes measure only the detached action
part; host package relationships and publication remain outside this owner.
The repeated-write lane checks only its final value and cost. The public API
does not expose an internal coalescing diagnostic, so the evidence makes no
coalescing claim.

Caller-cap refusal checks capture source bytes and parsed action state before
the timed edit, retain the input profile through the failed finish, and verify
the post-failure source with a fresh public no-op readback. Raw refusal
receipts expose matching source hashes and state gates; they do not imply a
failed commit patch exists.
