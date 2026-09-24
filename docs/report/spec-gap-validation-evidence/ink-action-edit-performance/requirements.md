# Ink-action edit profile acceptance contract

The final receipt is admissible only after the detached `actions_edit` owner
and its semantic tests are frozen. The run must satisfy all of these gates:

- the source manifest is byte-identical before and after the build and all
  retained source hashes still match;
- the runner executes from an isolated clean committed Git checkout and
  rejects every local path-backed package or extra input that is modified,
  untracked, or outside that checkout without a retained source snapshot;
- the isolated harness `Cargo.lock` is the authoritative lockfile; the root
  `Cargo.lock` and `docs/GOAL.md` are not build inputs;
- the selected `PROFILE_ARM` is one of the pinned baseline or candidate arms;
  its source pin is an ancestor of the captured head and its five exact
  production source hashes match `profile_pins.py`;
- a matched candidate capture uses the isolated child of the clean baseline
  pin with only the reviewed InkAction source change; both arms use the same
  Rust/Cargo 1.95.0 toolchain, harness lockfile, fixtures, allocator, flags,
  lanes, process count, warm-ups, and sample count;
- all 34 lanes have exactly three fresh-process JSON receipts with twenty measured
  samples after two warm-ups;
- allocator requested-byte accounting, reallocations, deallocations, live
  bytes, peak-live deltas, allocation failures, and invalid-live flags pass;
- detached drafts read back with the requested action count and opaque payload
  bytes;
- scalar, distinct-scalar-batch, repeated-scalar-write, add,
  insertion-batch, remove, removal-batch, clear-batch, and move commits read
  back with the expected semantic action count/order and preserve opaque source
  bytes;
- batch lanes queue scaled operation counts against the retained source, use
  canonical direct selectors, and validate XML identifiers after insertion,
  removal, and disjoint moves;
- every successful edit patch applies to the exact source and its inverse
  restores the original source; exact no-ops also retain source identity;
- every caller-cap lane refuses with the expected output limit and leaves the
  independently captured source bytes and parsed action state unchanged; the
  receipt retains matching pre/post source hashes and a fresh public no-op
  readback proof;
- draft and refusal lanes mark source/inverse/output checks as N/A where those
  checks do not apply; refusal receipts retain the actual resource and limit;
- the report is recomputed from raw samples and makes no comparison or
  speedup claim.

Fixture construction, source hashing, output comparisons, patch application,
inverse checks, and report generation stay outside the timed operation. The
timer covers the named public workflow, including its own bounded validation
and output allocation.

The repeated-scalar-write lane checks only the final value and measured cost.
The public API exposes no internal coalescing diagnostic, so the profile makes
no coalescing claim.
