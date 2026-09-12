# 0532: attribute the CFB claim-sector success path

The 0532 batch is a current-head OLE2 attribution baseline. It measures
revision `70e04d90181847ffe24c2ec8b8d5e1cd3b47981f` with source manifest
`c2fbb838229f7fcb40d51d7d8bf90721d7ff2e9658601db60e28342ccd6bd514` and
adopts no runtime source, dependency, unsafe-code, validation-policy or
resource-policy change. It follows the rejected 0531 OOXML binary and keeps
OLE2 ahead of deferred ODF; iWork is outside the scope.

## Evidence binding

The [campaign README](../results/change-0532/README.md) binds the matrix,
source and binary receipts, corpus identities, commands and measurement
boundaries. The [native and allocation analysis](../results/change-0532/analysis.json)
passes with 24,000 native samples and 720 separate allocation samples. Its
same-build review has zero variations above 5%; allocation elapsed time is
excluded from native latency, and the report makes no before/after speedup,
physical-I/O, cold/range, native-producer or scaling claim.

The four current CFB native rows have these p50 baselines (repeat 1 / repeat
2):

| Workflow | Native p50 | Allocation calls / reallocations | Allocated bytes | Incremental peak bytes |
| --- | ---: | ---: | ---: | ---: |
| CFB tiny | 2.410 / 2.400 µs | 32 / 1 | 4,456 | 3,384 |
| CFB many-small | 140.751 / 140.670 µs | 544 / 7 | 214,186 | 193,026 |
| CFB few-large | 91.441 / 90.980 µs | 29 / 1 | 213,571 | 205,129 |
| `xls_owned_source_open_one_cell` | 129.775 / 131.910 µs | 126 / 25 | 223,774 | 191,084 |

The [profile analysis](../results/change-0532/profile-analysis.json) passes
eight parent-constructor children with 40 timed constructor dumps and six
separate setup dumps. Positive incoming edges classify setup from timed calls;
all final process dumps terminate at zero Ir. The [quality summary](../results/change-0532/quality-summary.json)
records five fresh checks: CFB all-features tests (306 executions), CFB
no-default-features tests (306), XLS tests (1,345), Clippy and rustdoc, for
1,957 fresh test executions. The [quality-reuse binding](../results/change-0532/quality-reuse.json)
confirms that the exact 0531 restored source and receipts supply eight prior
quality gates and 4,757 successful test executions; those reused executions
are not relabeled as fresh 0532 tests.

## What the profiles establish

The [mechanism analysis](../results/change-0532/mechanism-analysis.json)
records one raw `claim_sector` call per physical sector in every one of the 40
timed dumps. Each call has 15 self Ir. The [assembly index](../results/change-0532/assembly-index.json)
contains six `claim_sector` variants; every variant is 363 bytes, reserves
112 stack bytes, and has a 15-instruction recorded success path. The entry
shape includes the prologue, the `usize` bounds branch and ownership branch;
the success path stores the role, writes the `Ok` discriminant and returns.
Formatted error paths remain in the function body. These are static
x86-64/Callgrind observations, not latency or removable-work estimates.

Exclusive owner shares are disjoint at the reported owner level. They must
not be added to inclusive parent costs or treated as independently removable:

| Profile shape | `claim_sector` | physical reconciliation | stream allocation validation |
| --- | ---: | ---: | ---: |
| XLS owned one-cell, repeat 1 | 17.8164% | 14.2536% | 14.1656% |
| XLS owned one-cell, repeat 2 | 17.8209% | 14.2572% | 14.1692% |
| CFB few-large, both repeats | 18.9505% | 15.1609% | 15.0530% |
| CFB many-small, both repeats | 0.3282% | 0.2631% | 3.8291% |
| CFB tiny, both repeats | 0.1836% | 0.1754% | 1.7480% |

The claim owner is therefore material in the XLS and few-large profiles, but
small in tiny and many-small CFB. Stream validation is the larger adjacent
owner in the latter two shapes. The [mechanism rows](../results/change-0532/mechanism-analysis.json)
also show that the three owners are measured as separate exclusive categories;
the full physical reconciliation remains a separate pass.

## Next measurement to evaluate

The highest-justified next candidate is a narrow compiler-layout experiment
around `claim_sector`, not a change to the CFB ownership model:

1. Move the three formatted `OleError::CorruptedFile` constructions into
   private `#[cold] #[inline(never)]` helpers, retaining their exact text and
   branch order.
2. Add only an ordinary `#[inline]` hint to `claim_sector` so the compiler can
   expose the success path at callers. Do not force `#[inline(always)]`; the
   six existing variants and their code size must be measured again.
3. Keep the checked body and its fallibility unchanged: the
   `u32`-to-`usize` conversion, `get_mut` bounds check, duplicate ownership
   check, role publication, `Result` discriminant and all error precedence
   stay in the same logical order. The source is
   [`claim_sector`](../../../crates/litchi-cfb/src/file.rs#L1025), and its
   per-sector caller is [`claim_chain`](../../../crates/litchi-cfb/src/file.rs#L1051).

The candidate must leave [`validate_stream_allocations`](../../../crates/litchi-cfb/src/file.rs#L1058)
in its existing collect-then-claim order and retain its mini-sector bounds,
fallible map operations and duplicate-stream errors. It must also leave the
full [`validate_physical_sector_layout`](../../../crates/litchi-cfb/src/file.rs#L1157)
scan in place, including the FAT lookup and unclaimed-marker error. The
collector's checked reservation, visited-bit check, marker checks and
failure reset in [`collect_exact`](../../../crates/litchi-cfb/src/file.rs#L2778)
remain part of the contract. No collection/claim fusion, bounds bypass,
physical-pass elimination, error rewrite or per-sector dump campaign is
justified by this evidence.

The assembly was captured from the standalone performance workspace. That
workspace has no explicit LTO profile, so root-workspace release LTO must not
be inferred as an explanation for the current prologue or as a property of a
future candidate. A fresh serial ABBA comparison is required after any source
change. A 3% primary-native p50 improvement in both repeats is the proposed
follow-up rule, but it is not frozen by this attribution batch and cannot be
used as an admission result here.

## Candidates to keep rejected

The 0524 private visited-bit fusion remains rejected after failing the native
admission rule; the 0279 freshness-session proposal is not reinstated. The
current measurement supplies no evidence to revive either one, and it does
not authorize aggregating the three owner shares into a single optimization.
The 0531 OOXML MCE binary remains rejected by its final native comparison and
is unrelated to this CFB attribution.

No runtime speedup or production change is retained from 0532. Both owned scratch paths have been removed. The
[post-cleanup verifier](../results/change-0532/verification.json) passes and
the [SHA-256 inventory](../results/change-0532/SHA256SUMS) seals the evidence. The [temporary-storage probe](../results/change-0532/temporary-storage-probe.json)
records an environment `errno 122` disk-quota failure; it is a custody note,
not a performance or correctness result.
