# 0494 methods and scope

This is a baseline for one opened DOCX paragraph replacement and sequential
save. It reuses the existing deterministic 0188 corpus and publication helper.
It makes no before/after production-speedup claim and does not replace the
read-only managed 0493 experiment.

## Providers and clocks

The warm matrix has six arms and two executables. Each formal process retains
30 samples after three warmups; two repeats reverse role and provider order.
The pilot retains three samples per arm and executable. The owned provider
has no source counter. Instrumented bytes and the recently staged file expose
logical reads. Short-read and delayed adapters additionally expose their
underlying source reads. The short-read cap is 4 KiB. Both transport controls
use a 64 KiB cap and a 100 MiB/s transfer model: `range-zero` has zero fixed
request delay; `delayed` adds 1 ms per request. The former is still bandwidth
paced. Neither is a network measurement.

Corpus generation, provider construction, trace and sink reservation, and
preflight patch oracles are outside the warm operation clock. Open, paragraph
edit, commit, sequential publication, commit diagnostics, source/candidate XML
comparison, and commit/document/package teardown are inside. Output remains
live after the clock for exact-byte, semantic, unselected-part, relationship,
and eight-media-member checks. Warm source-version fences and provider drop
are outside the operation clock. Raw reports state these boundaries.

The cold lane starts a fresh parent and measured child for each observation.
Every child performs all preflight reads and reservations before fsync,
DONTNEED, and the final fincore probe. Eligibility requires zero resident,
dirty, and writeback bytes plus positive measured process storage reads and
positive logical source reads. FileSource construction and source-version
fences are timed. The source ZIP tail is padded to page alignment. Cold and
warm file observations therefore have different source/clock boundaries and
are not a controlled cache-residency comparison. Formal cold observations
number 30 per role/repeat, 120 in total; the pilot contains six fresh children.

## Correctness and accounting

Every admitted row proves one materialized main Part, one changed operation,
unchanged source version, and exact expected output. Patch replay, inverse,
stale-target, and foreign-source checks are untimed preflight gates; timed
commit XML is compared with that preflight source and candidate. These are
source-XML patch checks, not a claim of whole-artifact patch identity.

The constructor uses actual default ReadLimits and an 8 MiB/128-entry source
cache. It is unmanaged: no ExecutionContext reservation or managed budget
return-to-zero claim is made. Ordinary managed edits retain their typed refusal.

Raw ranges permit independent reconstruction of counts, byte totals, request
sizes, and compressed-media overlaps. Logical requests and underlying adapter
requests are separate quantities. Empty reads reach the immediately wrapped
provider; the range provider handles them without a physical request.
Unavailable counters remain unavailable. Sink evidence records total bytes,
write count, and largest write; it does not provide a full write histogram.
Decompression, recompression, arbitrary memory-copy bytes, and lock-wait time
are not measured by these counters.

Allocator runs use a separate system-allocator binary. The operation region
includes a process live-byte baseline. The peak increment subtracts that
baseline; it is distinct from absolute allocator peaks and RSS. Warm GNU-time
RSS is one whole-child observation per process. Cold /proc RSS and VmHWM refer
to the measured child; its GNU-time receipt covers the parent and waited child.
Neither is operation-local peak RSS. Retained report traces and post-clock
oracles contribute to process memory.

## Evidence and ADRs

The exact final gates, retained binaries, source manifest, reproduction patch,
frozen protocols, helper hashes, raw captures, deterministic analysis, and
verification receipts form the evidence chain. The helpers retain selected
environment variables, tool versions, argv, process termination state, and
private-file cleanup. They do not promise a hermetic environment. The shared
host and advisory CPU lock permit unrelated workload interference.

| ADRs | Binding for this change |
| --- | --- |
| 0001, 0004 | Provider and instrumentation controls remain in the benchmark tool. |
| 0002, 0010, 0011, 0024 | Existing format/OPC owners perform the operation; no archive dependency is added to a facade. |
| 0003 | Existing isolated edit, commit, source-checked replay/inverse and conflict semantics are exercised. |
| 0005 | Explicit providers, bounded observer/sink storage, measured scope and typed unavailable budgets. |
| 0006 | Untouched parts/media and exact changed output are checked; failures refuse an evidence row. |
| 0008 | This extends measured scenario evidence and makes no broad format-support claim. |

The full goal remains incomplete: managed ordinary edits, genuine borrowing,
atomic filesystem save, independent producers, concurrent scaling, and broader
CRUD coverage require further implementation and measurements.

Analysis tables summarize selected allocator quantities (calls, allocated and
deallocated bytes, region peak and peak increment). Every raw row retains the
additional failed-allocation count, live-byte endpoints, and historical peak
fields. Those raw fields are not silently replaced by zero or promoted to
independent per-operation RSS observations.
