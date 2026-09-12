# 0535 collector attribution review

This is a read-only review of the baseline collector in the sealed 0534
Callgrind evidence. It separates the exclusive self Ir of
`SectorChainScratch::collect_exact` from the Ir of its visible direct callees.
It does not add a runtime patch, a candidate, a capture, or an admission
claim. The 0534 paired role/FAT-prefix loop remains rejected, as do the
visited-bit fusion and freshness-session directions.

The inputs are the [0534 baseline profile analysis](../change-0534/baseline/profile-analysis.json)
(`status: pass`, `performance_claim: diagnostic-only`, SHA-256
`ca375510ab289c2381ca268ea647e3519e1d7f83484de979f4f842d81f11b217`), its
retained positive timed annotations such as the [XLS-owned dump](../change-0534/baseline/profile-r1-xls-owned.part-2.inclusive.txt),
and the [many-small CFB dump](../change-0534/baseline/profile-r1-cfb-many-small.part-2.inclusive.txt).
The profile stage is bound to plan SHA-256
`4e1a13bdfa7f7c0268065d52def107f085b07cc0587d39685d9587de181d40dc`, baseline
source manifest SHA-256
`ce5a705508299624c2611d525940ce1502b3b9f20196814cf449e9286f8f11d6`,
and baseline normal binary SHA-256
`44c80a425a719441e0182d4a7932fcbe6309a84a816a75d75c764e90a0a13e0f`.

The evidence contains eight baseline profiles, 40 positive timed constructor
dumps, and six separate CFB setup dumps. XLS parts 1–5 are timed. CFB part 1
is setup and parts 2–6 are timed; setup is excluded from every number below.
The selected XLS constructor is
`SourceBackedWorkbook::from_read_at_with_limits`, reached through its positive
`from_read_at` edge. The selected CFB constructor is `OleFile::open`, reached
through the positive benchmark-runner edge. Each selected constructor edge has
one call. These incoming edges establish scope; child-call metadata is not a
timing counter.

## Parent and owner boundaries

The parent constructor total remains its own denominator. The existing
exclusive owner rows remain separate rows. The following values are sums over
the five timed dumps in each profile and are shown only to preserve those
boundaries; they must not be added together.

| Profile | Parent constructor inclusive Ir | `collect_exact` self Ir | `validate_stream_allocations` self Ir | Physical reconciliation self Ir |
| --- | ---: | ---: | ---: | ---: |
| XLS-owned, repeat 1 | 11,317,204 | 5,601,140 | 1,814,485 | 1,991,710 |
| XLS-owned, repeat 2 | 11,319,136 | 5,601,140 | 1,814,485 | 1,991,710 |
| CFB few-large, either repeat | 10,264,094 | 5,571,945 | 1,640,190 | 1,981,930 |
| CFB many-small, either repeat | 13,979,277 | 764,485 | 533,435 | 36,910 |
| CFB tiny, either repeat | 244,607 | 5,200 | 4,285 | 430 |

`collect_exact` is a positive out-of-line target in every selected timed dump.
The root mini-stream call uses the separate owned-result
`collect_sector_chain_exact` helper. The scratch calls reviewed here are the
MiniFAT and regular-FAT calls at the two `validate_stream_allocations` call
sites in [`file.rs`](../../../../crates/litchi-cfb/src/file.rs#L1079).
`validate_physical_sector_layout` remains a required positive out-of-line
owner. An absent or inlined `claim_sector` row remains an explicit allowed
outcome and does not represent free work.

## Collector self versus direct callees

For each dump, the analyzer's collector `self_ir` is the exclusive self cost
of the selected `collect_exact` function ID. Its `direct_ir` is the sum of the
visible direct-edge costs, and the selected `inclusive_ir` satisfies
`inclusive_ir = self_ir + direct_ir`. For example, the first XLS repeat has
`1,163,298 = 1,120,228 + 43,070` Ir.

| Shape and repeat | `validate_stream_allocations` → `collect_exact` edge metadata | Collector self Ir / timed dump | Visible direct children / timed dump | Collector inclusive Ir / timed dump |
| --- | ---: | ---: | ---: | ---: |
| XLS-owned, repeat 1 | 1 + 9 = 10 | 1,120,228 | 43,070 = 41,806 `memset` + 1,264 `finish_grow` | 1,163,298 |
| XLS-owned, repeat 2 | 1 + 9 = 10 | 1,120,228 | 43,151 = 41,806 `memset` + 1,345 `finish_grow` | 1,163,379 |
| CFB few-large, either repeat | 4 | 1,114,389 | 21,638 = 20,807 `memset` + 831 `finish_grow` | 1,136,027 |
| CFB many-small, either repeat | 256 | 152,897 | 13,218 / 13,105 / 13,012 | 166,115 / 166,002 / 165,909 |
| CFB tiny, either repeat | 3 | 1,040 | 208 = 46 `memset` + 162 `finish_grow` | 1,248 |

The three many-small values correspond to timed parts 2–3, 4, and 5–6. In
those same parts the `memset` edge is 12,850 Ir and the `finish_grow` edge is
368, 255, and 162 Ir respectively. The repeated profiles have the same
values. The aggregate collector self rows above therefore exclude the
`memset` and `finish_grow` children; they do not include child cost by virtue
of being called “self”.

The raw function graph has only two positive direct callee names under the
collector function in these dumps: `__memset_avx2_unaligned_erms` and
`alloc::raw_vec::RawVecInner<A>::finish_grow`. The `memset` edge is consistent
with [`prepare_visited`](../../../../crates/litchi-cfb/src/file.rs#L2783)
clearing `visited.words` with `fill(0)`. `finish_grow` is the lower-level
allocation-growth path reached by a fallible vector reservation; the edge
alone cannot say whether a particular growth came from `sectors` or from the
visited words. The collector has no visible direct edge to `Vec::push`,
`Vec::try_reserve_exact`, `CheckedBitSet::contains`, or
`CheckedBitSet::insert`.

The call labels printed beside child edges must be handled separately from
their Ir. In the first XLS timed dump of each repeat, the raw annotation shows
`memset (126x)` and `finish_grow (56x)`; in the other four XLS dumps it shows
the same Ir with `memset (12x)` and `finish_grow (5x)`. CFB shows collector
edge labels of `memset (5x)` and `finish_grow (2x)` for few-large,
`memset (257x)` and `finish_grow (2x)` for many-small, and `memset (4x)` and
`finish_grow (2x)` for tiny. The collector parent edge in those CFB rows is
4, 256, and 3 respectively.

These discrepancies are the expected collection-off/setup caveat in the
0534 analyzer: Callgrind can retain child-call context from work performed
outside the selected collection toggle. The labels identify positive graph
edges and help locate code, but they are not operation-local dynamic call
counts. They cannot be converted into allocation counts, per-stream call
counts, or a claim that `finish_grow` performed that many allocations.

## Bitset insertion attribution

The baseline does contain a separate out-of-line
`litchi_cfb::file::CheckedBitSet::insert` function record. That answers the
binary-wide question, but it does not show that the collector calls that body.
No positive `collect_exact` → `CheckedBitSet::insert` edge appears in the raw
function edges for any of the 40 timed dumps. The direct insertion edges that
are visible are siblings of the collector under its parent:

- CFB many-small has `validate_stream_allocations` → `CheckedBitSet::insert`
  with 81,920 inclusive Ir and a displayed 4,096-call edge.
- CFB tiny has the same parent edge with 480 inclusive Ir and a displayed
  24-call edge.
- XLS-owned has a 240-Ir `CheckedBitSet::insert` body reached from directory
  loading, outside `validate_stream_allocations` and outside the collector.

The CFB many-small and tiny sibling edges are the
`claimed_mini_sectors.insert` operation after a complete scratch collection,
at [`file.rs`](../../../../crates/litchi-cfb/src/file.rs#L1147), not the
`self.visited.insert(slot)` operation inside the collector at
[`file.rs`](../../../../crates/litchi-cfb/src/file.rs#L2853). Charging those
parent edges to `collect_exact` would double count ownership publication.

`CheckedBitSet::contains` is explicitly marked `#[inline]`; `insert` is not.
The absence of a collector edge is therefore consistent with the successful
collector insertion being inlined or otherwise optimized into its self Ir,
while the separate body remains available to other callers. Callgrind alone
cannot distinguish that shape from a low-cost or collection-off omission. A
fresh 0535 binary disassembly of both `collect_exact` and `CheckedBitSet::insert`
is required before making a stronger inlining claim. The same inspection
should map the two reserve sites, the `fill(0)` sequence, and the exact push
store in the bound binary.

The source control flow that contributes to collector self Ir is retained in
[`SectorChainScratch::collect_exact`](../../../../crates/litchi-cfb/src/file.rs#L2799):
entry reset, empty/start/table checks, exact sector reservation, visited-map
preparation, checked index and cycle checks, bitset insertion, exact vector
push, FAT/MiniFAT lookup, end-marker checks, and failure reset. A local self-Ir
reduction would still need to preserve the separate MiniFAT/FAT scratch
namespaces, collection-before-claim order, exact error precedence, fallible
resource labels, and post-error reset state.

No runtime change is accepted from this review. The next bounded step is the
0535 final-binary assembly review requested by root; only after that review
could a narrow cold-error/layout experiment be considered. It must preserve
the rejected visited fusion and freshness directions and use the complete
native and allocation guards before any adoption decision. OLE2 and OOXML
remain the active optimization priority; ODF is deferred until that goal
completes, and iWork remains outside scope.

No Rust build, test, profiler capture, or existing evidence rewrite was run
for this review.
