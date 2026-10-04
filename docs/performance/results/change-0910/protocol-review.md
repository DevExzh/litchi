# 0910 native owning-input versus materialized-ReadAt protocol review

This packet measures two existing PPTX ingress routes on the exact media-rich
corpus qualified by change 0909. It is a descriptive native comparison and
does not authorize a production bridge or an optimization claim.

## Route definition

The owning leg must move one fresh byte clone into
`litchi_pptx::Package::from_vec`. The materialized leg must move one fresh byte
clone into `litchi_core::OwnedSource`, construct the caller `Arc<dyn ReadAt>`,
open through `SourceBackedPackage::from_read_at`, consume
`into_opc_package`, and adopt the result through
`litchi_pptx::Package::from_opc_package`. Both legs then perform the same
opened-presentation transaction, replace all eight text boxes on each of the
eight selected slides, commit, publish, and serialize through `to_bytes`.

The byte clone and materialized source adapter are intentionally part of the
primary owning-byte ingress comparison. Corpus generation, source hashing,
expected-output construction, and the independent preservation oracle are
outside the measured operation. A separate phase view may report the route
cost after input preparation; it must not replace the primary total silently.

The primary timer starts immediately before the per-iteration byte clone and
stops after publication and the route-local package, transaction, commit,
source owner, adapter, and temporary output allocations have been dropped.
The caller sink is prepared and its capacity reserved before the timer and is
retained for post-timer checks. The same bounded sequential sink and maximum
logical write size must be used for both routes. The sink's later destruction
is outside both route timers.

Phase clocks are non-overlapping:

* `clone` covers the fresh `Vec` copy;
* `ingress` covers route construction after the copy, including complete
  source-backed materialization and `Package::from_opc_package`;
* `snapshot` covers `opened_presentation_transaction`;
* `edit_commit` covers all 64 replacements and `commit`;
* `publish` covers commit publication, `to_bytes`, and the identical sink
  writes.

Route-specific inverse, forward, or stale-patch checks do not belong in the
primary clock. They must be run outside the clock with equivalent correctness
gates, or the exact same sequence must be applied to both routes.

The primary run must not wrap only the materialized leg in an instrumented
`ReadAt` adapter: counter and atomic overhead would then be a benchmark-only
route difference. Read counters, if needed, belong to a separate diagnostic
capture. A logical `ReadAt` count is not a disk, network, or range-I/O count.

## Sampling

Each route runs in a fresh child process. Twelve process pairs alternate
`owned, materialized` and `materialized, owned` order, use CPU 12, and run
serially with no concurrent workload. Each child uses 30 untimed warmups and
20 retained samples. The binary, source archive, lockfiles, compiler, flags,
CPU placement, and environment are bound in every receipt.

Samples within a child are correlated through allocator and cache state. The
independent units for route comparison are therefore the 12 paired
process-level summaries, preferably each child's median. Report per-process
statistics, order-specific summaries, and paired median ratios or differences.
Any bootstrap interval must resample the 12 paired blocks. Pooled p95/p99
values may be shown as descriptive tails, but must not be treated as 240
independent observations or as a precise optimization confidence interval.

## Required gates

Every measured sample must produce the expected output length and stable
digest, exactly 64 changed text boxes, and the same bounded sink result. The
corpus gate binds 200 slides, eight text boxes per slide, eight deterministic
2 MiB media parts, 229 logical OPC parts, 445 ZIP members, 17,017,139 source
bytes, and source SHA-256
`61b2b99083ca27ebd37955db600955e3f41289b93dba71951983164239eff757`.

The semantic oracle must verify slide order and topology, selected and
unselected text, relationships, content types, untouched logical Part
payloads, and every media payload. The source must remain unchanged. The
independent raw census must inspect representative outputs from both routes
and bind the 445-member order and archive comment, exactly the eight selected
slide changes, and the 437 untouched members. Equal output bytes do not prove
that the writer copied rather than decoded, recompressed, or recopied a
member.

The owning route's source-preservation capability and the materialized route's
loss of physical provenance are preflight observations, not timed gates. A
different output digest between the routes must be retained and investigated;
the 0909 output identity cannot be assumed merely because it was observed
there.

## Expected mechanism and limits

`Package::from_vec` retains the owned source and preservation provenance.
After an edit revokes exact-source publication, the preservation writer can
still copy unchanged source members and regenerate only the changed closure.
The source-backed-to-materialized route fully decodes the catalog into an
ordinary `OpcPackage`; `Package::from_opc_package` does not carry physical
source provenance, so ordinary publication work is expected. This mechanism
is a hypothesis to measure. No copy-versus-regenerate count may be inferred
from output equality or from the 0908 Callgrind fractions.

The result is limited to one generated in-memory fixture and the 0910 native
build. It does not cover filesystem or remote/range input, cold cache,
producer variation, signed/encrypted/malformed archives, ZIP64 or adversarial
descriptor cases, cancellation, allocation failure, concurrency, or a
production provenance bridge. `/usr/bin/time` maximum RSS is a whole-child
envelope including setup, warmups, samples, verification, and artifact work;
it is not an operation-level memory measurement. The 0910 build must not be
compared numerically with the differently configured 0907 binary.
