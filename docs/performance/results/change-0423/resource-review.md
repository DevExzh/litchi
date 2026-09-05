# Resource boundaries and remaining work

The source-backed public planner retains copied media and XML buffers. Its
preparation loop reads each selected image and clones its bytes into staged
storage. Public publication prepares again and checks the staged payloads,
then builds the topology publication. The retained first plan and second
preparation can overlap. These source-code observations explain which resource
boundaries need measurement; they are not profiled hotspot or zero-copy claims.
Relevant owner code is crates/litchi-pptx/src/presentation/source_cross_copy.rs,
particularly prepare staging, publish_cross_slide_copy_to_stream and the
candidate reread reservation.

Both entry points use default OPC ReadLimits. At this revision these include a
512 MiB input limit, 512 MiB per-part and total-part limits, 100,000 parts, and
XML depth 256. Each source-backed package also uses the default finite payload
cache (8 MiB, 128 entries); those cache bounds exclude explicit staged plan
buffers and pinned mandatory metadata. Source cross-copy uses checked staging and reread reservations
against the total-part limit, with explicit execution budgets when supplied.
This fixture does not approach those limits and establishes no near-limit
retention bound. Native producers, borrowed ingress, range latency, physical
filesystem reads and parallel execution are outside this capture matrix.

Every iteration creates fresh package/cache state. Warmups exercise the code
and process allocator without retaining a selected-part cache into the next
sample. The common filesystem-cache flag is part of the harness configuration;
these selectors perform no filesystem input inside the operation.

InstrumentedSource copies requested byte ranges from an immutable Vec and
updates SeqCst counters. Its overhead is included inside source-backed API
calls. The read counters measure logical requests and returned bytes, including
repeated reads. They do not measure storage traffic or remote range behavior.

The allocator lane uses serialized_region_peak_v3. Region peak is an absolute
process live-request-byte maximum in callback observation order, including
bytes already live at entry and any background-thread callbacks. The source-backed publisher consumes its editor inside the interval; the
source view, plan, publication snapshot, input adapters and output sink remain
live at exit. The owned path retains its packages and opened artifacts through
exit. These are different retained-artifact boundaries. The peak is not object-owned memory, physical realloc copy overlap,
RSS or post-drop retention. The observer mutex affects scheduling, so its
latency cannot substitute for the normal lane.

Whole-process time-v RSS includes corpus creation, correctness/refusal oracles,
iterations and teardown. The two roles use distinct corpus setup oracles;
whole-process RSS therefore does not isolate the timed operation. No cross-role
memory reduction is claimed. Future retention work needs explicit snapshots
at plan, publication and concrete object-drop boundaries, plus near-limit
failure evidence. Existing phase-only heaptrack profiles remain historical;
this batch adds no profiler trace or CPU hotspot attribution.

## Retained observations

Mean operation-region peaks in both allocator repeats were 1,359,983 bytes
(plain owned), 1,059,110 bytes (plain source), 272,736,283 bytes (media owned)
and 229,300,257 bytes (media source). Individual vectors vary as the process
retains report bookkeeping across iterations; these means are not identical
per-sample values. The result table and summary retain entry/exit live values,
region peaks, lifetime peaks and separate whole-process RSS.

Every source-backed plain sample recorded 123 source reads / 14,105 bytes and
423 destination reads / 58,432 bytes. Media samples recorded 1,794 source
reads / 50,368,324 bytes and 975 destination reads / 16,845,847 bytes. Repeated
logical reads are included. The source media sink accepted 33,599,843 bytes in
1,031 writes, each at most 65,536 bytes; the plain sink accepted 31,514 bytes
in 183 writes. These are public lifecycle observations, not physical I/O or
unique-byte counts. Phase timing is retained, but phase-specific allocation
and read accounting still need explicit instrumentation.
