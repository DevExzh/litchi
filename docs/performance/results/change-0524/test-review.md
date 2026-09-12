# 0524 CFB differential test review

Status: test artifacts prepared against the frozen 0524 base
`477281a2f83c3256bf3cc06fbc5c57724c9b6bbc`.

This review covers only private CFB validation tests. The production
`CheckedBitSet` candidate has not been applied, and this agent performed no
production edit, build, test execution, benchmark, profile, or capture.

## Artifacts

[`tests.patch`](tests.patch) adds three candidate-dependent tests and the
independent collector differential to the `litchi-cfb::file::tests` module.
The candidate-dependent tests are:

- `checked_bit_set_test_and_set_matches_insert_at_word_boundaries` walks
  every bit for bounded lengths 0, 1, 63, 64, 65, 127, 128, 129, and 130,
  then revisits them in reverse order. The existing `contains` plus
  `insert` behavior is the oracle for the fused return value and final
  words, so the test covers both sides of every 64-bit word boundary and
  repeated visits without copying the proposed implementation.
- `checked_bit_set_test_and_set_preserves_state_on_bounds_error` compares
  the fused out-of-range error with the existing insertion error after bits
  0 and 64 are set. It snapshots the words and bit length and verifies that
  the failed fused operation leaves them unchanged.
- `checked_bit_set_test_and_set_preserves_state_on_missing_backing_word`
  truncates the private word vector after setting bit 0, then compares the
  missing-word error with the existing insertion helper. It checks that the
  failed operation does not alter the truncated state and that the previously
  set bit remains observable. The truncation reaches a checked structural
  error; it does not claim allocator-failure behavior.

[`collector-tests.patch`](collector-tests.patch) is a standalone extraction
of the independent hunk in `tests.patch`, for retaining that coverage if the
candidate is rejected. It adds
`scratch_differential_matches_owned_chain_helper_and_resets`, which compares
`SectorChainScratch::collect_exact` with the existing
`collect_sector_chain_exact` over a bounded table of valid, empty, cyclic,
invalid-start, invalid-index, invalid-marker, too-short, early-end, late-end,
and empty-chain marker cases. It compares successful sector vectors and exact
displayed error text, runs the same cases twice through one scratch object,
checks that errors clear the output and visited length, and checks retained
sector/visited buffer identity across successful reuse. The existing helper
remains the behavioral oracle; the test does not reimplement chain walking.

## Semantic and ADR boundaries

The tests stay inside the private CFB owner and add no API, dependency,
archive type, provider, unsafe code, resource limit, or execution policy.
They preserve the existing checked bounds and backing-word error authority,
the collector's validation order, and scratch reset-on-error behavior.

The design was reviewed against the accepted ADR set recorded by
[`change-0524/adr-manifest.json`](adr-manifest.json), with particular
attention to ADR 0001/0002/0024 private ownership and crate topology, ADR
0005's measured-evidence and bounded-resource requirements, ADR 0006's
non-mutating validation and typed failure behavior, ADR 0008's verification
custody, ADR 0010/0011 archive ownership, and ADR 0026's CFB directory-owner
boundary. OLE2/OOXML remains the active priority; ODF remains deferred.

No test asserts allocation counts, byte counts, peak memory, or allocator
failure. Those claims require the separate allocator evidence lane.

## Verification status

Both patch files pass `git apply --check` against the frozen working tree.
Apply `tests.patch` once to install the complete test set. Apply
`collector-tests.patch` alone when only the candidate-independent hunk is
being retained; it must not be applied after `tests.patch`, because that would
apply the same collector hunk twice.
The candidate-dependent patch is intentionally not compilable until the
planned private `test_and_set` method exists. No build or test command was
run here, so compilation and focused execution remain root-owned follow-up
work after the production candidate and test patch are frozen.
