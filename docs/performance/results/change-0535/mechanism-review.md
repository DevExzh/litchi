# 0535 mechanism review: the CFB chain collector

This review keeps the 0535 runtime at the retained baseline. The bound
binary shows that the ordinary `SectorChainScratch::collect_exact` path
already has the successful visited-bit operation and sector-vector append in
the caller. The separate `CheckedBitSet::insert` symbol is present for other
callers, but the collector has no hot call to it. That removes a proposed
call-level explanation without providing a timing or speedup claim.

The 0534 paired role/FAT-prefix loop remains rejected. The 0524 visited-bit
fusion and 0279 provider freshness-session directions remain rejected as
well. Physical reconciliation and the existing ownership boundaries remain
required work.

## Evidence and scope

The review is bound to the frozen [0535 plan](plan.json), the [native
diagnostic](analysis.json), the [host receipt](host.json), and the [assembly
index](baseline/assembly-index.json). The index binds the assembly to plan
SHA-256
`6f1fa017c41d2e9e48043795509479dc28e566508c4eb7545c847fece4611d34`, source
manifest SHA-256
`9d6a53738f299107582b71e5ddab03922b9646771df4d77b639b3fdc7fd6ff75`, and
normal binary SHA-256
`caa376f8313f4c2342e20ca4d3ee5126b06805206581fe276dbf5a76ab94028a`.

The [collector disassembly](baseline/assembly-2.stdout) is the 1,436-byte
`SectorChainScratch::collect_exact` symbol at `0x2f292c0`. The [bitset
disassembly](baseline/assembly-1.stdout) is the 223-byte
`CheckedBitSet::insert` symbol at `0x2f1b1c0`. Both are from that same normal
binary. The raw [XLS instruction dump](baseline/profile-r1-xls-owned.callgrind.1)
and [CFB setup dump](baseline/profile-r1-cfb-few-large.callgrind.1) plus the
first CFB timed dump (`profile-r1-cfb-few-large.callgrind.2`) use
`positions: instr`, `events: Ir`, and absolute positions.

The validated [instruction analyzer](instruction_analysis.py) and its retained
[instruction report](instruction-analysis.json) report schema
`litchi-ole2-change-0535-instruction-analysis-v1` and no performance claim.
The report's validation gives:

| Check | Result |
| --- | ---: |
| raw dumps | 22 |
| positive timed dumps | 20 |
| separately classified CFB setup dumps | 2 |
| collector rows in the assembly index | 1 |
| `CheckedBitSet::insert` rows in the assembly index | 1 |
| collector instructions in the bound symbol | 313 |
| relocation bias for every profile | `0x0` |
| collector instruction Ir equals collector function self Ir | true |
| positive parent owner attribution | true |
| direct-call and jump metadata kept separate | true |

XLS parts 1–5 are timed. For each CFB repeat, part 1 is the setup path
through the benchmark closure and parts 2–6 are timed `run_cfb_open` paths.
The positive incoming constructor edges establish those roles; setup is not
added to timed owner totals. A parent constructor's inclusive Ir, an
exclusive owner row, and the collector's exclusive self Ir remain separate
denominators. The [earlier collector attribution review](collector-attribution-review.md)
records the parent and child boundaries in detail.

The analyzer's per-workload aggregates keep the constructor workloads
separate. Across the ten timed XLS dumps, collector self is 11,202,280 Ir;
addresses from `0x2f29480` through `0x2f2952e` account for 11,194,720 Ir,
99.9325 percent. Across the ten timed CFB dumps, collector self is 11,143,890
Ir and that range accounts for 11,140,800 Ir, 99.9723 percent. These are
instruction-cost location results; they are not operation-local elapsed-time
results and are not combined into a mixed-workload denominator.

The raw function graph also has positive direct edges for `memset` at the two
visited-map clear sites and `RawVecInner::finish_grow` at the growth sites.
The `cfn`/`calls` labels are retained graph metadata. For example, the first
XLS dump displays 126 `memset` calls and 56 `finish_grow` calls in the
collector context, but collection-off context can contribute those labels.
They cannot be read as operation-local dynamic call counts, allocation
counts, or a per-stream cost. The positive Ir owner path and the mapped
`fn` self-cost are the attribution evidence used here.

## Generated code shape

The collector begins by clearing the retained result length at `0x2f292e4`
and the retained visited-map logical length at `0x2f292ec`. The empty-chain,
start-marker, table-length, and fallible-capacity checks precede the ordinary
loop. The sector-vector growth path reaches `finish_grow` at `0x2f29775`; the
visited-word growth path reaches `do_reserve_and_handle` at `0x2f29834`. The two mapped
`memset` sites are `0x2f2940a` for newly grown words and `0x2f29451` for the
retained-map clear. These calls and their ordering are part of the current
allocation and freshness behavior.

The valid-chain loop is the contiguous block beginning at `0x2f29480`:

1. `0x2f29480` compares the current sector with the visited map's logical
   length. `0x2f29489`–`0x2f29494` derive and check the visited word index.
2. `0x2f2949a`–`0x2f294aa` form the bit mask and load the selected word.
   `0x2f294ae` performs the duplicate-bit test, and `0x2f294b2` branches to
   the cycle diagnostic when the bit is already set.
3. On the success path, `0x2f294bd` ORs the mask into the word. There is no
   call from this loop to `CheckedBitSet::insert`; the checked word access,
   test, and set are generated in `collect_exact` itself.
4. `0x2f294c1`–`0x2f294c8` check the retained sector-vector capacity. The
   normal append at `0x2f294ee` stores the sector directly, `0x2f294f5`
   increments the length, and `0x2f294f8` writes the new length. There is no
   hot `Vec::push` call in this body.
5. `0x2f294fc` loads the next FAT/MiniFAT marker. `0x2f2950d` and
   `0x2f29516` retain the early-end and invalid-marker checks. The next
   sector is carried at `0x2f29523`, and `0x2f29529` returns to `0x2f29480`.

The stack state in this block must not be described as removable latency.
The initial sector state is stored at `0x2f2945c`; the current slot used by
the cycle/error paths is stored at `0x2f294b8`; the loaded next marker is
stored at `0x2f29500`; and the next sector state is stored at `0x2f2951f`.
The stores feed final-marker, error-formatting, and next-iteration control
paths. Their presence in an Ir map does not establish that deleting any one
would be safe or faster.

The collector's rare branches occupy the remainder of the symbol. Bounds
errors branch at `0x2f2959a` and `0x2f295e4`; cycle formatting begins at
`0x2f29690`; final-marker handling begins at `0x2f296cb`; early end at
`0x2f296f9`; and invalid-marker formatting at `0x2f2971e`. They converge on
the formatter path around `0x2f29565`, which restores the error result and
resets both retained lengths before returning. The map and vector growth
failure paths similarly preserve the existing resource and reset behavior.

The out-of-line `CheckedBitSet::insert` body has a different shape. Its
success path at `0x2f1b1cc`–`0x2f1b1f1` checks `bit_len`, checks the word
length, loads the words pointer, ORs the mask, and returns success. Its
out-of-range and missing-word diagnostics begin at `0x2f1b201` and
`0x2f1b249`, respectively, and each calls the formatter. This body is real
binary code, but the validated collector records contain no positive
`collect_exact` → `CheckedBitSet::insert` edge. Changing this out-of-line
body alone therefore has no demonstrated effect on the collector's ordinary
loop.

The compiler has placed the collector's valid loop before its diagnostic and
growth tails, and the instruction map puts almost all timed collector self Ir
in that valid loop. That supports a narrowly scoped code-layout question
around cold diagnostics. It does not show that the current error layout is a
native timing cause, and it does not justify a static instruction-latency
calculation.

## Native context and disposition

The native diagnostic used 4,000 samples and excluded 20 profiler-timing
samples. It has no candidate and no speedup claim. The same-build p50
observations are retained exactly as measured:

| Case | Repeat 1 p50 | Repeat 2 p50 | Same-build p50 change |
| --- | ---: | ---: | ---: |
| XLS owned source | 95,615 ns | 95,590 ns | −0.0261% |
| CFB few-large | 68,840 ns | 75,790 ns | +10.0959% |

There are 13 same-build variation flags in `analysis.json`, including the CFB
p50 row. They are diagnostic observations; neither their cause nor a
before/after gain is inferred.

No runtime change is accepted from this review. If a future bounded
experiment is staged, it must isolate rare error formatting while retaining
the current bitset checks, reset points, fallible reservations, resource
labels, FAT/MiniFAT namespace separation, marker-check order, and
collect-before-claim order. It must leave `claim_chain` and physical
reconciliation untouched. The next-candidate document records that narrow
question and its admission gates.

This review ran no Rust build, test, profiler capture, or evidence rewrite.
OLE2 and OOXML remain the active optimization priority; ODF is deferred and
iWork is outside scope.
