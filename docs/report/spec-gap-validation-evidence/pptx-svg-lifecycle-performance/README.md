# PPTX source-backed SVG lifecycle performance evidence

This directory retains bounded allocator, wall-time, and RSS evidence for the
source-backed PPTX SVG attach/detach lifecycle. The harness is intentionally
kept outside the production workspace dependency graph. It uses synthetic PPTX
packages with a direct `p:pic`, a PNG compatibility fallback, and the native
`asvg:svgBlip` extension when the fixture selects an attached SVG.

The final run is sealed only after the production lifecycle owner and its
focused correctness tests are frozen. Each lane runs in three fresh processes
with twenty measured samples after warm-up. The process-local allocator reports
requested allocation bytes, reallocations, deallocations, live-byte balance,
and incremental peak live bytes. `/usr/bin/time -v` supplies whole-process RSS.

The report separates isolated capture/clone/no-op operations from the named
end-to-end attach and detach sequences. End-to-end timing includes only the
steps named by its lane and records the resulting package reopen checks. The
attach and detach lanes use the ordinary fallible `commit()` path, so their
timed closure includes dependency-closure validation before publication.
Fixture construction, expected-result derivation, and post-run assertions are
outside the timed closure unless the lane explicitly says `end_to_end`.

The evidence makes no native Office, rendering, or general performance claim.
The SVG and raster payloads are generated test inputs, and the bytes supplied to
`SourceSvgAttachmentReplacement` or the detached attachment value are caller-owned input
buffers whose construction is outside the library allocation measurement.

The matrix also includes a namespace-heavy capture lane with 252 inherited
bindings and 1,024 opaque descendants. It deliberately exercises
active-binding lookup through the scoped prefix index without claiming linear
namespace handling. A paired
refusal lane reaches the owner resolver's 16,384 active-binding limit from a
syntactically complete source; raw receipts retain the exact refusal reason so
an earlier package/XML admission refusal cannot be mistaken for the namespace
limit result.

Two inventory lanes contain 256 and 1,024 direct raster pictures in one slide,
all sharing the same raster relationship. They call the public full-slide
`images()` inventory and check every descriptor, making repeated per-picture
owner resolution visible while keeping the media payload modest. These are
absolute workload observations; they do not claim linear scaling or isolate a
native Office algorithm.

Paired inventory lanes put a distinct unused local namespace declaration on
each of those 256 and 1,024 pictures. They retain the same raster and shape
workload while exposing persistent local namespace-context capture and lookup.

When sealed, use:

```sh
PROFILE_FROZEN=1 \
  CARGO_TARGET_DIR=/var/tmp/litchi-pptx-svg-lifecycle-profile-target \
  sh docs/report/spec-gap-validation-evidence/pptx-svg-lifecycle-performance/run_profile.sh
```

The verifier checks the source manifest, raw receipts, allocator equations,
expected success/refusal status, semantic readback, output hashes, and the
recomputed report. It also checks that all source inputs are byte-identical
before and after the build and run. Each raw receipt also records and verifies a
FNV-1a-64 hash of its complete synthetic package input across fresh processes.
The runner removes a newly created isolated Cargo target on exit while
retaining the raw receipts and report under this directory.
