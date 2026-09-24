# Profile acceptance contract

The final receipt is admissible only when the shared `formatcode2` owner and
its semantic tests are frozen for the run. The profile must then satisfy all
of these gates:

- the source manifest is byte-identical before and after the build and all
  source hashes still match;
- every lane has exactly three fresh-process JSON receipts with twenty
  measured samples;
- allocator counters pass the allocation/reallocation/deallocation balance,
  no allocation failure is observed, and the invalid-live-byte flag remains
  clear;
- element and attribute readback, source-backed `read_shared` pointer sharing,
  no-op byte identity, scalar-edit semantic readback, clone source-sharing,
  and `write_to` sink byte-count/hash checks pass outside the timed interval;
- both malformed lanes for both owners reject as expected;
- the report is recomputed from the raw samples and contains no comparison or
  speedup claim.

The timer excludes fixture construction, source hashing, output comparisons,
reopens used for semantic gates, and report generation. Requested allocation
bytes are direct allocation sizes plus successful reallocations' new sizes;
old realloc sizes are retained separately and are not double-counted. Peak
live bytes are incremental above the timed closure's live-before baseline.
RSS is a whole-process value from `/usr/bin/time -v`, not per-operation RSS.
