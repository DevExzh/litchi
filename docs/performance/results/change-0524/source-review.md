# 0524 CFB candidate source review

This review binds the frozen candidate, which was subsequently rejected. The final retained test-only source is reviewed separately in `final-review.md` and `temp-test-review.md`.
`scope: independent read-only review of the candidate and test patches`

`performance_claim: none`

I reviewed `crates/litchi-cfb/src/file.rs` after the candidate and final test
patches were applied, [`candidate.patch`](candidate.patch),
[`tests.patch`](tests.patch), the standalone [`collector-tests.patch`](collector-tests.patch),
the design and profile reviews, and the applicable accepted ADRs. This review
did not edit production or test source and did not run a build, test, profile,
capture, or commit.

## Source and patch custody

The frozen baseline source has SHA-256
`5f70ab560dd2d954e487fabeaac19d01d6b6158d74a6576834e01940d2cf4e72` and Git
blob SHA-1 `1da9ad4f2c0b6989e3d5e7828cb7fc964b5702ec`. After applying the
candidate and final test patches, the current source has SHA-256
`ec44f1209fe7b895bb3e2d55316eb3cf9e890017a74d6925fe9027e88370e737` and Git
blob SHA-1 `85fb33500f01658d057f4c9a0d05e8593d4afef4b`. That current SHA-256
matches `preflight-source-manifest.json`.

The candidate patch SHA-256 is
`70ee4c2f1e6fb945e30d7c569d8e96b43b2337db35f00e4cf4786739c5f2369c`, the
final combined tests patch SHA-256 is
`24dad1628652fd060ac0fd5d30f97f8d1441738ad048e6ac0fb83895f66cbb3e`, and the
standalone collector-test extraction SHA-256 is
`e77e8b6656550d2f1a10e2dbc40aabf19d20635aa74a47c99b6779b8b35241a1`. The
candidate and final combined tests patches each passed `git apply --check`
before application. `collector-tests.patch` is a standalone rejection
extraction of the differential test; it is not applied in addition to the
combined tests patch.

The three test-only formatting corrections are recorded in
`preparation-adjustment.json` (SHA-256
`b6e19ae0ed1f9e764a329e0153ea2f914026fee9f7a8c7da905a631887d571de`); they
wrap one local binding and add configured match-arm commas. No production
logic was changed by that preparation step, and no source edit was made by
this review.

The applicable accepted constraints are ADR 0001 (typed safety), ADR 0002
(CFB ownership boundary), ADR 0003 (fallible and non-mutating state), ADR 0005
(bounded resources and measured admission), ADR 0006 (validation and hostile
input behavior), ADR 0008 (verification), ADR 0010/0011 (container ownership),
ADR 0024 (current topology), and ADR 0026 (OLE directory ownership). The
candidate stays within those constraints and keeps OLE2/OOXML ahead of
deferred ODF work.

## Candidate semantics

The patch adds private `CheckedBitSet::test_and_set` beside the existing
`contains` and `insert` methods and changes only the reusable
`SectorChainScratch::collect_exact` call site. It does not alter the owned
`collect_sector_chain_exact` helper or any other bit-set user, including
directory, FAT, MiniFAT, and physical ownership paths. It adds no public API,
dependency, unsafe code, source I/O, or allocation.

At [`file.rs#L66`](../../../../crates/litchi-cfb/src/file.rs#L66), `insert`
checks the bit length, computes the word and mask, checks the backing word, and
then ORs the mask. The candidate performs those same checks in the same order,
reads the prior mask, ORs it, and returns the prior state. A clear bit returns
`false` and receives the old set state; a set bit returns `true` and remains
set. The two existing checked diagnostics are preserved exactly:

```text
bit index {bit} exceeds checked bit-set capacity {bit_len}
bit index {bit} has no backing word
```

The missing-word branch returns before mutation. An out-of-range bit also
returns before word lookup or mutation. The operation performs no fallible or
infallible allocation.

## Collector order, state, and resource proof

At [`collect_exact`](../../../../crates/litchi-cfb/src/file.rs#L2796), the
candidate leaves the following sequence unchanged:

1. Scratch reset, empty-chain handling, invalid start-marker rejection, and
   the declared-count versus allocation-table check.
2. Fallible reservation of `sectors` under `"sector-chain entries"`.
3. Fallible growth/preparation of the visited map under `"sector-chain map"`,
   including the same bit length and zeroing.
4. Per iteration, checked sector conversion and allocation-table bounds.
5. Cycle detection, now fused with the same checked bit set operation.
6. The already-reserved `sectors.push`.
7. Allocation-table lookup and the early/late `ENDOFCHAIN` and invalid-marker
   checks.
8. Scratch reset on every error.

The old path did `contains(slot)` followed by `insert(slot)` immediately before
the push. The new path does `test_and_set(slot)?` at that same position, so a
cycle still returns before the push or next-marker read, and the bit is still
set before the push on a successful visit. A missing backing word or bit-set
capacity error still occurs before the push. Since `prepare_visited` sets
`bit_len` to `allocation_table.len()` and the preceding slot check uses that
same length, those private bit-set errors are defensive checks rather than new
reachable collector outcomes.

The exact vector reservation, resource labels, capacity assumption, and
allocation-failure boundary remain untouched. On cycle, marker, table, or
other collector failure, the existing enclosing result handler clears both
the sector buffer and visited bit length. `claim_sector`, `claim_chain`,
`validate_stream_allocations`, and `validate_physical_sector_layout` are not
changed, so ownership and physical-sector validation remain after chain
collection with their existing error order.

## Test coverage

The current [`tests.patch`](tests.patch) adds four same-module tests:

- word lengths 0, 1, 63, 64, 65, 127, 128, 129, and 130 compare the fused
  return/state against `contains` plus `insert` on first and repeated visits;
- an out-of-range bit compares the exact display error and checks words,
  `bit_len`, and existing set bits remain unchanged; and
- a deliberately truncated backing vector compares the exact missing-word
  display error and checks state remains unchanged before a valid repeat.
- a two-pass differential matrix compares `SectorChainScratch::collect_exact`
  with the unchanged `collect_sector_chain_exact` for valid, empty, cyclic,
  invalid-start, invalid-index, early/late terminal, invalid-marker,
  overlong-table, and empty-non-terminal cases. It compares complete sector
  results or exact display error text, checks scratch cleanup after every
  error, and checks retained buffer pointers/capacities are reused on a later
  successful collection.

These tests directly exercise the new operation at word boundaries and both
defensive error branches. The unchanged scratch tests already cover successful
reuse, growth before walking, empty chains, cycle errors, early terminal
markers, late terminal markers, and reset after errors. The unchanged general
helper remains available as a behavioral reference, while existing CFB owner
and consumer tests cover ownership, overlap, FAT/MiniFAT markers, physical
range, and malformed-count boundaries.

The differential coverage requested by the design review is therefore now
present in the combined patch. The standalone collector extraction carries the
same test for rejection or isolated review and does not add a second production
source change.

The retained initial preflight receipt has `exit_code: 101` after 273 passed
and five failed tests: three explicit `E122` `QuotaExceeded` errors and two
filesystem failure-state assertions in shared `/tmp`. All four new tests are
listed as passing in that retained output. The
[`preflight-environment-adjustment.json`](preflight-environment-adjustment.json)
(SHA-256
`cd3b28e1fc0cbe8d615a7b7e8dc0e3fd857b8bafb3f3170ac79c6a72a7e90acf`) records
removal of only the five proven 15-byte artifacts from PID 1370116, with
absence checks, before repeating the same source under the owned disk-backed
`TMPDIR=/home/zhuhe/litchi-goal-0524-target/test-tmp`. The final
[`preflight.receipt.json`](preflight.receipt.json) (SHA-256
`fc0f34f0b253a017608a839aec484eadfbf9c2d54608624827d24a00a1e086e2`) reports
`exit_code: 0`, 278 passed and 0 failed, and `source_unchanged: true` for
source manifest SHA-256
`7720b9dc1a135e411def3b9164f60b37695fbd73a46c69fb455c3c145d4f84ce`.
Therefore the current source SHA-256 above is the authoritative candidate
hash. The initial failures remain recorded as an environment and custody
event rather than a candidate semantic failure.

## Review disposition

The applied candidate source change is a narrow, safe work-elimination
experiment: it fuses one checked lookup/set pair without removing a validation,
changing state lifetime, moving a fallible boundary, or touching ownership and
physical reconciliation. The combined test patch is applicable and now
contains both direct bit-set checks and the collector differential. The
candidate preflight is now clean under the owned temporary directory.
Production adoption still requires the declared candidate build, matched native
and allocation evidence, and all quality gates; this source review makes no
performance or adoption decision.
