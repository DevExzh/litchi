# 0913 final review

Review status: **PASS**. The packet is complete and internally consistent. The
candidate outcome is **rejected** and `adopted` is false; this review makes no
new speedup claim.

The rejection follows both frozen owned-route admission conditions. The paired
owned total ratio is 1.005714080656532, below the 1.05 target, and its 95%
paired bootstrap interval is 0.9733079247831775–1.0174857562044162, whose lower
bound is not above 1. The materialized route is also neutral to slower at
0.984089747530512 (95% interval 0.9730520628620976–0.9901747761629278). The
nine latency flags are retained and treated as observations requiring review;
none is silently discarded. Whole-process RSS has no >5% route regression, and
the packet makes no allocation-count or peak-live-byte claim.

The statistical custody checks are sufficient for this decision: all 48
children passed, covering 960 measured samples and 2,400 guarded operations;
the independent reducer matches all 40 compared metric groups; and the four
ZIP censuses agree on 445 physical members, 437 unchanged members and the
expected eight changed members. The analysis, independent statistics,
comparison receipt and decision receipt are bound below by bytes and SHA-256.

The span evidence is correctly scoped. The 64 KiB admission cap leaves all
eight 2,097,797-byte compressed media members on the existing baseline
passthrough path. The generated archive has 437 eligible file members and the
static one-read bound for those members is 211,315 bytes. That is an
archive-layout bound, not a dynamic read, allocation or copied-byte count; the
16,777,216 decoded media bytes remain outside that bound. This evidence cannot
rescue the failed native gate.

The candidate production files were restored to the recorded parent bytes and
all three candidate-only tests were removed from the production tree. The five
candidate files remain available through the retained `candidate-source/`
copies bound by `candidate-retention.json`; this review does not bind mutable
root candidate paths. Source reconstruction and cleanup checks pass; the
before/after replay results are byte-identical. The iWork resolution is
independently bound to
commit `1b12cef9a97bbe3597039f7e3caadc2c45c905ce`, with its replay and cleanup
receipts; it is separate from the rejected performance result.

`seal.py` and `integrate.py` were reviewed statically. The seal helper builds a
private-index performance commit from the recorded parent and verifies every
packet byte before publishing the performance ref. The integration helper
constructs a merge tree in a temporary index, checks existing worktree bytes,
uses Git's normal index lock for compare-and-swap publication, and leaves
foreign untracked files alone. No build, test, benchmark, profiler or workload
was run for this review.
