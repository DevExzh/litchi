# 0808 frozen-driver readiness review

This is a read-only review of the frozen 0808 driver and manifest set. It did
not run Cargo, rustfmt, a probe, a capture, Callgrind, or any owned driver.
The review covers `build.py`, `capture.py`, `profile.py`, `quality.py`,
`probe_quality.py`, `apply_candidate.py`, `restore_candidate.py`,
`analysis.py`, the custody helper, and the packet schemas.

## Resolved qualification reader finding

The qualification lane deliberately runs the **allocation** binary:
`capture.py` sets `binary_kind = "allocation"` for `LANE == "qualification"`
and records that binary in each receipt. The reader now matches that protocol
at `analysis.py:509`:

```python
kind = "allocation" if lane in {"allocation", "qualification"} else "native"
```

This makes the before-only 18-report qualification use the allocation binary
and allocator fields while keeping qualification elapsed samples outside the
timing comparison. The qualification audit was accepted before application,
so the earlier classification finding is resolved and is not an open replay
blocker.

## Intentional effective allocation schedule

`plan.json` defines the six native block orders. The allocation lane
intentionally inherits the first two native orders (`before/after`,
`after/before`) in `capture.py`, `analysis.py`, and the independent audit.
That gives allocation the fixed effective AB/BA schedule while preserving the
single schedule source and the paired resource guard. The workflow has
already accepted this effective allocation protocol; no separate
`allocation.orders` amendment is required for this archived run.

## Checks that are otherwise aligned

The archive/manifest chain is coherent: the before source is tied to the
current base, the after source is tied to the one-file candidate, the patch is
`git apply`-compatible, and the application requires an independently accepted
18-report qualification audit before changing production. The after build
requires the before target and revalidates all three before binary identities;
the capture and profile drivers validate both build source manifests while the
candidate source is live. The quality driver is scoped to `litchi-pptx` and
its six package checks, and the probe-quality driver expects the repaired
36-test PPTX probe.

The native/allocation capture order is explicit for native and deterministic
for the current allocation prefix. Each receipt carries a source-bound binary
descriptor and the driver checks the live source after every process. The
profile driver keeps the owner-scoped Callgrind publication separate from
latency and RSS claims.

Two lower-priority hardening points remain for the final replay reader. It
should compare each lane's recorded `complete.plan_sha256` with the current
plan hash instead of ignoring that field, and should check the recorded source
manifest revision/base and candidate-manifest descriptor against the build
source census. The build/application drivers already perform the stronger
checks during execution, but independent replay should bind those duplicate
schema fields as well.

## Readiness disposition

The driver and archive packet were ready for the baseline, qualification, and
candidate quality sequence. That sequence stopped at the after-quality
Clippy gate because of three pre-existing diagnostics in `opened/tests.rs`,
before after-build, native/allocation capture, or profiling began. The reader
classification and effective allocation AB/BA schedule have no open blocker
for the archived protocol. No performance or adoption conclusion follows
from this review.
