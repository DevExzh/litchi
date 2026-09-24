# Source-backed DOCX SVG lifecycle measurement contract

This directory defines the bounded evidence contract for commit
`892441d95db29da4351390716ef5c65b4c7c97de` (`OPC57680dc86`). It is an
absolute profile of the committed public source-backed DOCX API. It does not
compare against an older implementation and does not claim a speedup,
rendering result, or native Office result.

The acceptance run uses three fresh processes and at least twenty measured
samples per lane after warm-up. The initial smoke run intentionally uses one
process and one sample per selected lane; it validates wiring and semantic
receipts only. The full run is gated by `PROFILE_FROZEN=1` so that costly
measurement cannot be mistaken for the scaffold phase.

Each sample reports these quantities separately:

- wall time for the lane's named scope, with capture, staging, commit,
  publication, reopen, inverse reopen, inverse, selected payload read,
  validation, and failed-source readback phase timings where applicable;
- direct allocation bytes and reallocation old/new bytes from a process-local
  `GlobalAlloc` observer;
- incremental peak live allocation bytes, computed from live allocator bytes;
- allocator calls, deallocations, failure/underflow flags, and the checked live
  byte equation;
- output bytes and semantic, opaque-content, lazy-media, and exact-inverse
  gates;
- process RSS from `/usr/bin/time -v`, recorded in the sidecar timing receipt.

Peak live allocation bytes are not RSS. RSS includes the runtime, code,
allocator arenas, page cache effects, and other process mappings. The runner
retains both receipts and never substitutes one for the other.

The named scenario matrix is:

- native SVG and floating-picture source capture from copied, hashed native
  submodule fixtures;
- lazy metadata inventory at one and 64 owners, proving media members stay
  cold until an explicit `data()` request;
- single attach and single detach at 1, 16, and 64 owners;
- same-story batch attach and batch detach at 1, 16, and 64 owners;
- shared SVG cleanup across 64 owners, including exact raster and opaque
  payload checks;
- exact physical inverse publication for a single attach, plus the 64-owner
  batch topology refusal as a recorded bounded refusal;
- a large unchanged media member under a two-MiB managed execution cap;
- exact no-op detach publication at 64 owners.

The 64-owner batch attach and exact-inverse batch lanes are expected typed
refusals because the committed OPC topology overlay has a 64-Part operation
bound. Their receipts are successful only when the public API returns
`litchi_docx::Error::Opc(litchi_opc::OpcError::SourceBackedOverlayUnavailable
{ .. })`, the source package can be republished byte-for-byte after the
failure, and reopened metadata matches the pre-operation snapshot. The
message text is retained for diagnostics after the typed match; it is not the
classification gate. Silently dropping work or emitting a partial package
fails verification.

Synthetic fixtures are deterministic and include the opaque document marker
`docx-svg-profile-opaque-v1`, unchanged media bytes, one/16/64 owner counts,
512-byte SVG, 128-KiB SVG, 1-KiB raster, and an 8-MiB unchanged media member.
Fixture construction and caller-owned payload construction are outside the
library allocation boundary. The timed operation includes only the steps
listed by its lane and the public semantic readback needed to validate the
result.

The retained `baseline-source/` tree freezes the relevant committed
DOCX/DrawingML/OPC source files. The source manifest compares those retained
copies with the clean committed checkout before the build, then hashes the
stable evidence paths before and after the run. Every raw receipt carries both
SHA-256 and FNV-1a-64 input identity. Any missing receipt, source drift, hash
mismatch, allocator failure, semantic failure, opaque mismatch, or process
failure causes the verifier to fail closed.

Opaque preservation validation uses the ordinary `NormalPackage` reader and
is reported in `validation_ns`; it is not hidden inside an unlabelled phase.
The named phase durations are disjoint and their sum is bounded by the lane's
elapsed time. Explicit selected-media reads use `payload_ns`.
