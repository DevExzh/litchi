# 0795 results review — bounded interpretation of the rejected iterator

This is an offline review of the retained 0795 raw captures and their two
independent readers. It does not rerun a build, profiler, or native command.
The 0794 disposition remains controlling: the five-file iterator candidate is
rejected, the exact baseline remains production, and these diagnostics cannot
authorize adoption.

## Evidence boundary

The packet compares the exact 0794 baseline with the archived candidate under
the same `namespace_uri_probe::capture_region_0793` owner. The Callgrind leg
has twelve positive dumps (tiny, medium, and large; two order-reversed
repetitions) and twelve zero termination dumps. It records guest `Ir`, `Bc`,
`Bcm`, `Bi`, and `Bim`; these are simulated counters, not hardware cycles or
retired instructions. The native leg has four large captures at `cycles:u` with
frame-pointer stacks. Its samples include warmups and whole-process work, so
they are stack observations rather than latency measurements.

`callgrind-analysis.json` reports twelve qualified owner partitions with no
qualification failures and conservation of all five counters. The independent
`root-cg-totals.json` reader agrees for every positive dump and finds every
termination dump zero. `native-analysis.json` reports 4,097 owner-qualified
samples and a public descendant in each one; `root-native-counts.json` agrees.
The two baseline native processes each retain one unresolved owner sample, so
the packet correctly leaves native phase fractions uncomputed. The stale
candidate build was rejected before capture and contributes no profile data.

## Paired Callgrind result

The percentages below are `(candidate - baseline) / baseline` for each paired
shape and event. `r0/r1` are the two order-reversed repetitions; they are
paired drift controls, not a confidence interval.

| Shape | Ir change r0/r1 | Bc change r0/r1 | Bcm change r0/r1 | Bi change r0/r1 | Bim change r0/r1 |
| --- | ---: | ---: | ---: | ---: | ---: |
| tiny | -0.157% / -0.220% | -2.072% / -2.339% | -7.702% / -7.839% | -1.451% / -1.543% | +25.917% / +26.357% |
| medium | -0.131% / -0.183% | -1.559% / -1.726% | -6.319% / -5.968% | -1.388% / -1.306% | +25.201% / +25.379% |
| large | -0.393% / -0.387% | -2.086% / -2.073% | -6.880% / -7.032% | -1.627% / -1.630% | +35.076% / +34.970% |

The large owner region therefore has a small, repeat-stable simulated
instruction decrease while conditional branches and simulated conditional
misses decrease, indirect branches decrease slightly, and simulated indirect
misses increase materially. The `Bim` result is a guest-simulator observation;
it does not say that the host predictor incurred the same misses. The pattern
also cannot explain the 0794 workflow p50 result by itself: Callgrind has a
different execution model and this packet has no fresh native timing control.

The large selected-symbol rows show where the binary attributes some of this
guest work, with the usual parent/child overlap:

| Selected row, large repeat 0 | Baseline self Ir | Candidate self Ir | Change |
| --- | ---: | ---: | ---: |
| `inspect_element` | 29,165,311 | 29,774,878 | +609,567 (+2.09%) |
| `CheckedAttributes::next` | 8,660,171 | 8,936,993 | +276,822 (+3.20%) |
| quick-xml iterator state | 42,805,826 | 38,818,090 | -3,987,736 (-9.32%) |

The same self rows recur in the second large repetition. The checked-iterator
row's simulated `Bim` self count changes from 3 to 40,470, which is a useful
line/function target for the next diagnostic, not a hardware branch result.
The `allocation_named_functions` group falls sharply in self Ir, but it is a
lexical function census; it is not an allocator-call count and must not be used
as an allocation or latency explanation. Inclusive rows and selected groups
overlap, so they cannot be summed into savings or shares.

## Native sampled-stack observations

| Repeat | Leg | Whole-process samples | Owner samples | Inspector samples | Checked-helper samples | Unresolved owner samples |
| ---: | --- | ---: | ---: | ---: | ---: | ---: |
| 0 | before | 3,128 | 1,016 | 208 | 99 | 1 |
| 0 | after | 3,369 | 1,023 | 239 | 156 | 0 |
| 1 | after | 3,386 | 1,032 | 247 | 151 | 0 |
| 1 | before | 3,128 | 1,026 | 207 | 95 | 1 |

The owner-filtered helper observations increase from 99/95 to 156/151 and
the inspector observations are 208/207 versus 239/247. These are overlapping
sample and frame counts, with different owner totals and warmup/process work;
they are not proportions of cycles or proof of more executed calls. The two
unresolved baseline owner frames forbid phase fractions even though every
qualified sample has the public descendant.

The symbol census has the expected source visibility for the two designs. The
baseline shows `quick_xml::events::attributes::IterState::check_for_duplicates`
51 times in each run and no candidate `OwnCheck`/`SeenNames` symbols. The
candidate shows `OwnCheck::check` 28/30 times and `SeenNames::check` 12/8
times, while quick-xml duplicate-check observations fall to 21/10. The
`CheckedAttributes::next` symbol is visible in 96/92 baseline samples and
146/144 candidate samples. These are sampled stack rows, not call counts.

The source predicts that candidate `SeenNames::check`, `position`,
`offset_in`, and `end_of` may inline into `CheckedAttributes::next` or
`OwnCheck::check`. A missing helper frame therefore does not prove no work;
conversely, a present frame proves only that a sampled stack contained it.
Generic quick-xml iterator frames may remain for parsing work even when its
duplicate-check path is disabled. The baseline's cold, non-inlined takeover
symbols are expected only if the public corpus reaches the 33-name boundary.

## What the comparison supports

The safe reading is a work shift in the generated code: the candidate's local
checked-iterator symbol carries more self Ir and simulated indirect-miss
accounting while quick-xml iterator-state self Ir falls. The native stacks make
the local helper more observable in these owner-filtered samples. This is
consistent with the source-level possibility that candidate bookkeeping and
state dispatch replace quick-xml duplicate checking, but the captures do not
establish that relationship as a cause of the 0794 timing regression.

The concrete paths worth isolating are:

* construction and the zero-attribute `None`/`Done` path, where the candidate's
  inline `OwnCheck` changes the live iterator layout even without an
  allocation;
* one through four accepted names, where candidate `SeenNames::Inline`, raw
  byte comparison, `offset_in`, and `end_of` replace the baseline quick-xml
  duplicate check;
* the 5, 9, 17, and 33-name transitions, where vector growth, the linear to
  ordered-map transition, and baseline's cold takeover/reparse are distinct;
  and
* early valid and malformed duplicates, where baseline checks the key before
  value parsing but the candidate may scan a complete quoted or unterminated
  value before duplicate recovery.

The known iterator layout difference, 120 bytes in the baseline versus 192
bytes in the candidate, should be recorded as a separate layout observation.
It can affect construction, moves, and ABI/stack traffic without exposing a
stable symbol. It should not be inferred from a prologue or from a sampled
stack count.

## Next bounded measurement

Before considering another implementation, run a separately labelled direct
attribute diagnostic with the exact baseline and archived candidate over
`0, 1, 4, 5, 8, 9, 16, 17, 32, 33, 64` distinct names. Include an early
duplicate, a long quoted or unterminated duplicate, and representative syntax
errors. Keep this lane separate from the frozen public tiny/medium/large
workflow. Use the same owner-scoped Callgrind events and native stack settings
when root schedules the work, then compare each threshold pair with the same
source identity, build settings, and collection boundary.

The direct report should attribute, where symbols or source lines permit,
`IterState::check_for_duplicates` against `SeenNames::check`,
`offset_in`/`end_of`, the inline/vector growth path, the 33-name takeover or
ordered-map transition, and duplicate-value recovery. It should report
`size_of::<CheckedAttributes>()` separately from counters. Inline and unresolved
paths must remain explicitly unqualified; nested rows must remain non-additive.

This measurement can decide whether a later candidate should test one bounded
bookkeeping or layout change at a time. It cannot relax the exact semantic
error-position contract or the 0794 public timing/resource gates. Until a
narrow path is isolated and a complete workflow requalifies it, retain the
baseline and do not infer a causal optimization from the 0795 diagnostics.
