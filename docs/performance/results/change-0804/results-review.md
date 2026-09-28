# 0804 results review

Result: the completed results packet is internally consistent and remains
diagnostic-only. I found no results blocker. This review is independent of the
offline analyzer's generated summary and does not authorize production use.

## Native raw review

The native receipt set contains 936 successful children: six alternating block
pairs for each of 39 cases and two modes. Every child has 30 samples, and the
nearest-rank p50 is sample index 14 after sorting. All reports have matching
semantic oracles, eight clone advances, and measured iterator sizes of 128
bytes for both helper layouts. The independent checksum and native audit
matches all 28,080 native samples and the 78 paired case/mode rows.

Selected raw process p50s agree with the report's ratios:

```text
distinct-32 consume before 5442548 5407058 5472827 5396898 5431188 5373717
                   after  3949320 3919600 4118331 4120901 4078831 4084511
                   median after/before 0.751753
distinct-33 consume before 5703169 5687860 5690489 5679919 5700409 5688999
                   after  4319862 4389393 4309473 4307612 4136511 4308652
                   median after/before 0.757407
distinct-64 consume before 14856796 15493290 14945886 14832885 15013867 14843726
                   after  13712920 13289158 13625629 13629160 13477659 13430659
                   median after/before 0.908234
duplicate-valid-after-1 construct before 6890 5930 6960 5930 5930 5930
                              after  6560 6820 5940 6760 6880 6850
                              median after/before 1.145025
syntax-equals-value-after-4 consume median after/before 1.024252
```

The full analysis has zero diagnostic regression flags. The 5% process-p50
spread set contains 75 of 156 case/mode/leg groups, split into 63 construct
and 12 consume groups; that count matches an independent scan of all raw p50
vectors. The four construction medians above 1.05 have bootstrap intervals
that include 1, so they do not meet the packet's regression rule. The report
correctly keeps these variability flags visible and does not attribute the
construction variation to different constructor work: both construction legs
resolve to the same surviving symbol in the binary.

The selected Callgrind table also matches raw positive-dump totals on both
repeats. For example, `distinct-32/consume` is 35,978 Ir before and 26,130
after, while `distinct-33/consume` is 38,142 before and 27,487 after. These
are guest-counter diagnostics; they do not establish native latency or a
unique instruction-level cause.

## Profile amendment and ownership

The failed original profile attempt is retained under `profiles-failed-0` with
child exit 0 and driver exit 1. It produced only the zero termination dump;
there was no positive dump to count. The `nm -C` witness contains
`after_construct`, `before_consume`, and `after_consume`, but no
`before_construct`, so mapping both construction legs to `after_construct` is
the necessary explicit scope repair.

The hash-bound resolution manifest ties the repair to the same binary, build
inputs, original plan, and original capture driver, and marks both the native
lane and probe/binary as unchanged. The resolved 312-child profile lane uses
the following owners:

```text
before/construct -> attribute_boundary_probe::after_construct
after/construct  -> attribute_boundary_probe::after_construct
before/consume   -> attribute_boundary_probe::before_consume
after/consume    -> attribute_boundary_probe::after_consume
```

All 312 positive dumps and all 312 empty termination dumps parse correctly.
Exact owner qualification succeeds for 312/312 profiles: one exact owner, one
positive incoming call, one call, matching incoming cost, and a valid
self-plus-direct-child partition. All five counters are conserved, and no
termination dump contributes counters. The profile receipts contain the
expected four artifacts per child and no `.callgrind.2` files.

## Quality and claim bounds

Each mirror leg has five crates with 20 tests each, for 100 passing tests per
leg, plus warnings-denied Clippy. The release probe's format, build, check,
Clippy, fixture, and semantic gates all pass. The setup metadata omission and
the unguarded continuation into quality are documented in
`execution-note.txt`; the quality source archive was unchanged while the
metadata was completed, and no build or native child was retried.

The main report accurately describes the comparison as 0803 after-control
versus direct byte equality in the bounded linear scan. Its tables and counts
match `summary.md`, `analysis.json`, the independent native audit, and the
raw profile totals. The wording keeps workflow advancement, production
adoption, production-baseline comparison, public workflow speed, and
cross-format claims false. The profile counters are described as guest
diagnostics, and the construction symbol repair is disclosed rather than
presented as a before/after constructor implementation difference.
