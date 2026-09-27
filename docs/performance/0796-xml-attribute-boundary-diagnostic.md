# 0796 — direct XML attribute construction and boundary diagnostics

Diagnostic only. The rejected 0794 candidate remains rejected; production source
is unchanged from `2b53da0c98`.

## Protocol

The [packet](results/change-0796/README.md) compiles the exact baseline OPC
helper and archived rejected 0794 OPC helper as separate modules of one release
binary. The standalone manifest and dependency lock are frozen before the first
build. A separately generated [literal input catalog](results/change-0796/fixtures.json)
is compared byte-for-byte with the binary's case listing. All 35 inherited
architecture-input hashes and the production source census remain unchanged.

Thirty-one inputs cover distinct attribute counts 0, 1, 4, 5, 8, 9, 16, 17, 32,
33, and 64; valid duplicates after 1, 4, 5, 32, and 33 attributes; 4,096-byte
quoted, unterminated, and unquoted duplicate tails after 1 and 33 attributes;
and three syntax-error tails after 0, 4, and 33 attributes. Each has construction
and consumption modes. Construction creates, black-boxes a reference to, and
drops the iterator without advancing. Consumption includes construction, iteration
through first error or natural exhaustion, checksum bookkeeping, and destruction.

Root ran builds and captures serially. Native timing uses CPU 12, six source-leg
pairs with the frozen alternating orders, 30 samples after three warmups, and
4,096 repetitions of the same hot input per sample. Each case/mode/leg/block
runs in a fresh process. Process p50 is the fifteenth ordered sample; the
reported ratio is the median of six paired p50 ratios. The interval uses 10,000
bootstrap draws, seed 796079, and zero-based ranks 250 and 9749. A ratio above
1.05 with interval lower bound above 1 is a diagnostic flag, not a workflow veto.

Callgrind uses two opposite leg orders, one sample with one iteration and no
warmup, and five events: guest instructions, conditional branches and simulated
misses, indirect branches and simulated misses. Collection starts disabled and
zeros, toggles, and dumps at the exact named non-inlined mode/leg function.
The counter lane and repeated native timing lane have different iteration counts;
simulated counters are not native hardware measurements.

## Validation and limits

Formatting, locked release build, locked release check, and Clippy with warnings
denied pass. The independent binary's catalog and semantic self-check pass.
The copied helper modules retain external `cfg(test)` declarations; this probe
does not compile their unit-test modules or run `cargo test`. Their production
test evidence is inherited from 0794, with exact source hashes. Fresh semantic
checks compare accepted borrowed key/value sequences, first error variants and
positions, two terminal `None` calls, and clones at six transition boundaries
against a fail-fast adapter over quick-xml 0.41.0.

Every timed error is at the input tail. No separate valid attribute following an
error exercises reference-adapter recovery suppression; that coverage limitation
is retained explicitly. Opaque iterator construction forces materialization and
can differ from optimized inlined production callers. The two local modules may
also compile differently from separately built production crates. Native samples
reuse hot input, and include timer/function-pointer and checksum overhead. These
results cannot establish public-workflow latency, allocation counts, or production
adoption. The previous workflow rejection remains authoritative.

## Results

All 992 process reports preserve their semantic and result checks, covering
22,568 measured samples (22,320 native and 248 counter samples). An independent
reader reconstructs all literal inputs, accepted sequences, error positions,
and every sample checksum. A separate scalar raw-cost reader confirms five-event
conservation and positive/empty termination policy across all 496 counter dumps.
All 248 exact owners qualify with one incoming call and conserved self-plus-child
partitions. All 62 independently computed paired native rows match the full analyzer.

The table selects boundary and error cases; the full paired results remain in
the packet. Ratios are candidate/baseline, so values above one are slower.
Guest instruction counts below are repeat 0, one iteration.

| Case | Mode | Native ratio | Bootstrap interval | Before Ir | Candidate Ir |
|---|---|---:|---:|---:|---:|
| distinct-0 | construct | 8.678752 | 8.248358–11.804714 | 70 | 140 |
| distinct-0 | consume | 1.070628 | 1.069980–1.077228 | 170 | 225 |
| distinct-1 | consume | 0.964448 | 0.920096–0.999746 | 574 | 568 |
| distinct-4 | consume | 1.001081 | 0.988482–1.068491 | 1,702 | 1,822 |
| distinct-5 | consume | 0.872759 | 0.824471–1.313560 | 2,645 | 2,404 |
| distinct-8 | consume | 0.923773 | 0.915049–0.938279 | 4,327 | 4,111 |
| distinct-9 | consume | 0.944431 | 0.940364–0.959241 | 5,520 | 5,369 |
| distinct-16 | consume | 1.022298 | 0.995007–1.031786 | 9,272 | 9,278 |
| distinct-17 | consume | 0.982798 | 0.960817–1.021349 | 10,383 | 10,414 |
| distinct-32 | consume | 0.996332 | 0.983188–0.999079 | 25,332 | 24,091 |
| distinct-33 | consume | 0.883829 | 0.866771–0.888214 | 49,736 | 41,248 |
| distinct-64 | consume | 0.825757 | 0.808593–0.838914 | 80,677 | 71,909 |
| duplicate-valid-after-32 | consume | 0.532182 | 0.521305–0.537795 | 49,438 | 24,546 |
| duplicate-long-quoted-after-1 | consume | 24.384922 | 23.758672–24.789628 | 754 | 25,412 |
| duplicate-long-unterminated-after-1 | consume | 24.514611 | 23.497580–24.729153 | 754 | 25,400 |
| duplicate-long-quoted-after-33 | consume | 0.922036 | 0.907887–0.935156 | 74,555 | 66,115 |
| syntax-flag-after-0 | consume | 1.309300 | 1.261347–1.370248 | 253 | 331 |

Between-process p50 spread exceeds 5% in 66 of 124 case/mode/leg groups.
The complete spreads and outliers are retained; no samples or blocks were dropped.
For example, distinct-0 construction after-leg p50 spans 48,680–86,010 ns per
4,096-iteration sample. The bootstrap intervals describe these six observed
block pairs and do not remove host or code-generation uncertainty.

All 31 opaque construction rows trigger the diagnostic flag: their native
ratios range from 7.890472 to 9.094537. The iterator grows from 120 to 192 bytes,
and the one-iteration distinct-0 construction region doubles from 70 to 140
guest instructions. The ratio must not be extrapolated to production
construction: the opaque reference prevents normal elimination and inlining,
and this very short repeated operation is sensitive to generated loop code.

Eight of 31 consumption rows trigger the diagnostic flag. Long quoted and
unterminated duplicate values after one accepted attribute cost about 24 times
as much natively, with guest instructions rising from 754 to 25,412 and 25,400.
The candidate's lexical iterator reaches the value before its own duplicate
check, while the baseline rejects the duplicate key first. Both return the same
error and positions; their amount of fail-fast work differs. At 33 accepted
attributes, other savings dominate these cases and the aggregate native ratios
are below one. The nonmonotonic boundary results prohibit treating one direct
helper ratio as representative of a document workload.

The next bounded candidate should preserve duplicate detection before value
scanning and reduce state initialization/bookkeeping for short tags. It must
then pass fresh paired public-workflow and resource gates. This packet provides
diagnostic evidence only; it neither reverses 0794 nor changes any CRUD coverage
or optimized baseline. OLE2/OOXML remain in scope, ODF deferred, iWork excluded.

The owned temporary target was removed after verifying the captured binary's
identity (112,130,069 logical bytes removed). Post-cleanup replay uses the exact
binary identity in the retained cleanup witness. The packet seals raw captures,
source and lock identities, analysis, independent audits, and review records;
unrelated files and existing worktrees are preserved.
