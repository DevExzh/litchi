# 0816 protocol review

This protocol freezes the bounded source experiment at production revision
`c8ff2b9f65`. It is a read-only harness study. It does not change a production
crate, add a public format operation, or claim an end-to-end CRUD result. This
review was prepared from the frozen packet plan and the reusable
`tools/perf-execution` entry point; it did not run Cargo, a binary, a workload,
or a profiler.

## Question and evidence gap

The question is whether a caller-visible bounded source (a maximum return
size per `ReadAt` call, with an optional fixed delay per non-empty successful
call) changes finite-budget scaling across the public CFB bulk-read and
source-backed OPC ordered-Part routes. The comparison keeps the immutable
corpus, task floor, route APIs, requested widths, and cache-state protocol
fixed.

The gap is deliberately narrower than a first short-read demonstration.
0786 measured CFB and ordered Parts under finite execution budgets with an
in-memory source, but that source returned the requested range immediately.
0787 and 0788 added cache-state and scheduling controls without a delayed
caller source. 0498 is useful historical context because it measured
source-backed Part batches with short-read and fixed-delay controls, but it is
not an equivalent control: it is an earlier Parts-only batch study with
different corpora and capture protocol, and it has no matched CFB route or
this 0786 large/mixed fresh and large primed finite-budget cross-product.
Those results remain separate and are not pooled with 0816.

This batch therefore supplies descriptive category-15 low-level read evidence.
It cannot close physical-cold storage, a network service, cross-session
contention, native-format CRUD, publication, or a cross-format speed claim.
iWork is excluded and ODF remains deferred.

## Public routes and corpus

The harness must call the existing public APIs through these routes:

* CFB uses `SharedOleFile::open_with_limits`, `bulk_read`, and
  `SharedOleBulkRead::read_streams` over an immutable `ReadAt` source.
* Parts uses `SourceBackedPackage::from_read_at_with_execution_context` and
  `read_parts_ordered` over the same kind of immutable `ReadAt` source.

The generated corpus has 32 payload members. A large shape has 32 members of
256 KiB. A mixed shape has 31 such members followed by one 4 KiB member. The
CFB and OPC container bytes are generated once per sample configuration and
their payloads have deterministic per-member SHA-256 values. The CFB stream
names and ordered OPC URIs identify the same 32 logical payloads.

The 4 KiB CFB stream is at the regular-FAT cutoff and must retain the existing
container classification. The protocol does not use a MiniFAT claim to
explain the result. Structural OPC members are container metadata; the Parts
route verifies the 32 selected payload parts in request order.

## Provider arms

Each CFB or Parts sample creates a new immutable in-memory source with the
following caller-controlled settings:

| Arm | `source_max_read_bytes` | `source_delay_us` | Meaning |
| --- | ---: | ---: | --- |
| local | 0 | 0 | Uncapped, no-delay reference |
| capped | 65,536 | 0 | Each successful non-empty read returns at most 64 KiB |
| delayed | 65,536 | 250 | The capped source also waits 250 µs per non-empty successful call |

The cap applies to the provider's returned byte count after the normal bounds
check: it cannot exceed the caller buffer, remaining source bytes, or the
configured cap. A cap of zero means uncapped. A positive short read is a
successful prefix; the production reader must issue subsequent reads to obtain
the requested bytes. The provider must not manufacture progress at EOF or
return bytes beyond `len()`. For an offset representable by the host `usize`,
`offset >= len()` returns zero. An offset that cannot be represented by
`usize` retains the harness's explicit invalid-input error on narrow hosts;
that conversion error is not a successful read and does not acquire delay.

The delay belongs inside the caller's `ReadAt::read_at` operation, after the
call is known to request a non-empty in-range read and before it returns its
successful bounded prefix. It therefore contributes to the route wall-clock
interval. `len()`, `version()`, empty-output calls, out-of-range/EOF calls, and
the cap calculation must not acquire that delay. This is a requested sleep in
an in-memory simulation; it is not network RTT, disk latency, or a physical
range source.

The source's normal path must remain allocation-free apart from the output
buffer supplied by the caller. The diagnostic observer may count logical
calls, requested and returned bytes, short reads, request-size buckets, and
active/max simultaneous reads, but those counters must be compiled only under
the separate `source-metrics` feature. The ordinary binary must retain no
observer atomics or source-counter work.

The CLI defaults both new settings to zero when their flags are omitted. The
bounded domains are `0..=1,048,576` bytes and `0..=100,000` microseconds.
Source settings are valid for CFB and Parts; OPC is rejected when either is
nonzero because `OpenSession::from_bytes` does not use the external `ReadAt`
source. The report must serialize the selected source settings rather than
silently treating an omitted value as a different arm.

## Frozen 72-case matrix

The packet's `litchi.performance.0816.plan.v1` lists exactly 72 cases. Every
case uses `task_floor = 65,536` and one requested width from `1, 2, 4, 8`.
The 18 four-width groups are:

* CFB, large, fresh and primed, for each of the three provider arms;
* CFB, mixed, fresh, for each of the three provider arms;
* Parts, large, fresh and primed, for each of the three provider arms; and
* Parts, mixed, fresh, for each of the three provider arms.

No small-shape, OPC, zero-floor, width-32, iWork, ODF, or native CRUD case is
part of this packet. The width is a caller execution-context limit; it is not
a guarantee that the route admits that many simultaneous reads.

## Timing and state boundaries

Corpus construction, container serialization, source/provider construction,
route metadata setup, priming, output verification, observer snapshot, result
drop, and JSON serialization are outside the per-sample route wall interval.
Each sample owns a new finite root budget and execution context. The measured
closure is the public CFB `read_streams` or Parts `read_parts_ordered` call.
The source delay is inside that closure and is consequently measured. Process
CPU time brackets the operation with the existing process-clock scope and may
be slightly wider than wall timing. Whole-child RSS, if retained by the
capture driver, includes setup and teardown and is not operation-local.

`fresh` means a new route object after metadata setup over a warm immutable
in-memory source. It is not physical cold cache. `primed` performs one verified
preload on the same route object before the clock, drops the preload result,
resets source observers, and then measures the next public operation. The
preload consumes the sample's finite budget but is outside the wall interval.
Primed Parts are an explicit cache-hit control: their timed operation is
expected to avoid payload source calls when the existing cache semantics hold.
Primed CFB measures the existing retained session/cache state; it must not be
described as a physical warm file or a remote cache result.

The observer reset occurs after all priming reads and immediately before the
timed operation. The observer snapshot occurs after the timed call has
returned and before source teardown. Observer timing is a diagnostic lane and
must not be pooled with native timing. Verification happens after timing and
must fail the run if any returned member is missing, reordered, changed, or
has a mismatched SHA-256 digest.

## Finite budget and conservation checks

The context retains the reusable harness limits: finite memory, input/output,
objects, depth, work, worker and I/O reservations, CPU-task ceiling, maximum
in-flight tasks/bytes, aggregate parallel threshold, and the selected 64 KiB
task floor. The source cap and delay change provider behavior only; they must
not change the declared budget limits or bypass the public route's admission
and cancellation checks.

The source observer's conservation rules are:

* every logical call increments `logical_calls` once, including a short
  successful call; empty or EOF calls are counted only if the production route
  actually invokes the provider;
* `requested_bytes` is the sum of caller buffer lengths and
  `returned_bytes` is the sum of valid prefixes, so returned bytes never
  exceed requested bytes;
* a call returning fewer bytes than requested increments `short_reads`, while
  a non-EOF positive short read must allow the route to make progress;
* `active_reads_after_operation` is zero after the timed call, and
  `max_simultaneous_reads` is no greater than the requested worker width; and
* the request-size histogram sums to `logical_calls` and uses stable bucket
  boundaries.

The route-level oracle is stronger than counters: every case must return all
32 ordered payloads and exact member hashes. After the returned batch/package,
session, and context are dropped, worker and I/O reservations must be zero and
the cumulative CPU-task usage must remain within its ceiling. The report's
resource receipt must not claim conservation of memory, input, output,
objects, depth, work, or cache bytes unless those dimensions are actually
snapshotted. A finite configured limit is not itself a measured usage value.

## Capture and analysis contract

The qualification lane has one sample and no warmup per case. Native timing
uses six counterbalanced blocks with orders
`forward, reverse, forward, reverse, reverse, forward`, 30 measured samples,
and three warmups per report. The observer lane uses two blocks,
`forward, reverse`, with two samples and no warmups. The packet therefore
expects 120 qualification reports, 720 native reports, and 240 observer
reports: 1,080 reports and 22,200 measured samples across all lanes. Native
and observer outputs remain separate.

Width scaling pairs each width with width one within the same route, shape,
state, provider arm, task floor, and native block. Provider comparisons pair
the capped arm with local and the delayed arm with capped at the same width
and block. The primary summary uses nearest-rank process quantiles and the
median of six paired block ratios. Bootstrap uses 10,000 resamples, seed
816816, 95% endpoints at ranks 250 and 9,749. These are descriptive controls;
the plan has no adoption threshold or production optimization claim.

The report schema is conditional by source arm: the local `0/0` arm uses
`litchi.execution-baseline.v1`, while either nonzero source setting uses
`litchi.execution-range-baseline.v1`. This is an intentional additive schema
extension for the simulated source scope; both schemas carry route, source
settings, state, budget floor, corpus identity, per-sample wall/CPU values,
verification, resource receipt, and source metrics. The packet reader must
accept exactly the schema selected by the recorded source settings, rather
than treating the two arms as byte-identical reports. The plan schema is
`litchi.performance.0816.plan.v1`; changing either report schema or the case
cardinality requires reopening this review before capture. Historical reports
and the source-observer lane are never silently merged into native statistics.

## Review gates before build

The source handoff must pass a static review against this protocol before root
builds or captures:

1. CLI parsing has explicit defaults, inclusive bounds, duplicate/unknown flag
   rejection, and route validation for source settings.
2. The provider's cap, delay placement, EOF behavior, overflow handling, and
   zero-length behavior match the provider arms above, including the
   platform-portable treatment of an unrepresentable offset.
3. Native builds have no source observer atomics; all observer state is behind
   `source-metrics`, and observer reset/snapshot boundaries preserve timed
   scope.
4. CFB and Parts construct the configured provider for every sample and use
   existing public APIs; priming does not leak into the timed interval.
5. Budget creation remains per sample, the 64 KiB floor is applied, and
   worker/I/O/CPU checks are retained through teardown.
6. Output serialization records the actual arm, selects baseline-v1 for
   local `0/0` and range-v1 for nonzero source settings, and keeps OPC's
   source metrics explicitly not applicable. The offline reader must enforce
   this conditional schema contract rather than silently accepting either
   schema for every arm.

Any failed gate is a source change request to the harness owner. No build or
capture should proceed on an implementation that merely parses the new flags
while continuing to use the uncapped, zero-delay source.
