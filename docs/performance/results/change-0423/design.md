# Matched public lifecycle boundaries

The two new opt-in source-backed selectors use exactly the deterministic
source/destination archives, selected slide positions and collision pattern of
the owned lifecycle selectors. Plain copies one slide; media-rich copies one
slide and eight 2 MiB image leaves. The three-slide source and two-slide
destination become a three-slide destination after insertion. Layout reuse is
required. This is an admitted direct-image fixture, not arbitrary closure
support.

For both roles, archive Vec clones and the bounded output sink reservation are
outside the operation. The common sink maximum is twice the owned baseline
output length plus 65,536 bytes; writes are bounded to 65,536 bytes. The
source-backed adapter is also constructed outside. The owned timer covers both
Package::from_vec constructors, opened snapshots, plan, atomic application and
OPC sequential publication. The source-backed timer covers both public
from_read_at constructors, plan and public publication. Reopen, independent
semantic/package/raw-record checks and object destruction follow each timer.
The source-backed report has a separate lifecycle key; the original plain
phase selector retains its schema and boundaries.

The two writers may select different relationship IDs and serialize different
ZIP bytes. Each role must produce its own stable output hash. The media oracle
must bind copied image references to the correct source payload, verify content
types and the complete added-part/member sets, and retain untouched destination
raw metadata and ordering. The source-backed oracle cannot require the owned
writer's identical copied slide XML. For this deterministic fixture, it
normalizes only the ordered double-quoted `r:embed` values, then requires all
other inserted slide XML bytes to match, including geometry, style and unknown
markup. Other lexical forms fail closed. The presentation must add exactly
one relationship. Source byte/version and stale/foreign
refusal checks remain untimed hard gates.

The frozen protocol uses normal 100-sample/10-warmup and allocator
30-sample/3-warmup lanes. Each corpus has owned R1, source R1, source R2, owned
R2 fresh processes, serialized on CPU 2 with one worker. This is a same-revision
baseline, not a candidate/control experiment. Only within-role repeat drift is
tested. Shared-host background activity is uncontrolled. Normal latency and
allocator latency are never compared.

## ADR constraints

| Accepted constraint | This batch |
| --- | --- |
| 0001 / 0002 / 0024 API layers and crate ownership | Benchmark-only selectors call existing public APIs; no production dependency or ownership change. |
| 0003 source-checked edits and atomic publication | Existing owned commit/patch checks and source-backed stale/foreign refusal checks remain gates. |
| 0005 explicit I/O, finite resources and evidence | In-memory caller-supplied sources, bounded sequential sinks, separate instrumentation lanes, frozen protocol and retained source identities. |
| 0006 lossless preservation and fail-closed behavior | Independent closure/media/type checks and untouched raw ZIP records gate capture; no guessed closure or partial edit is admitted. |
| 0010 / 0011 physical package ownership | ZIP inspection stays in the isolated benchmark's existing oracle helpers; format APIs and archive boundaries are unchanged. |

No proposed ADR or new unsafe implementation is required. This enabler provides
matched current API observations; it does not relax the broader program's
production optimization, native evidence or strict lint requirements.
