# 0806 — fresh public workflow and resource qualification

**Measured disposition: rejected; no production change retained.** The main
workflow analysis and independent raw audit found no eligible capture or
lifecycle benefit, so the six-file production change was rejected and restored
exactly to the recorded base, `e3ff267ee3454e71d66f177f54f3cd05e0d9cce5`.
Cleanup and independent cross-format replay are complete. iWork remains
outside the scope, and the broader performance goal remains open.

The machine-readable packet is [results/change-0806/README.md](results/change-0806/README.md).
The packet's [execution notes](results/change-0806/execution-notes.md),
[source review](results/change-0806/source-review.md), and
[qualification review](results/change-0806/qualification-review.md) are the
authoritative custody and review records.

## Scope and frozen decision boundary

The candidate applies the bounded XML attribute iterator change to five helper
copies and the shared OPC tests:

```text
crates/litchi-ole-common/src/xml_attributes.rs
crates/litchi-opc/src/xml_attributes.rs
crates/litchi-opc/src/xml_attributes/tests.rs
crates/litchi-sign/src/xml_attributes.rs
crates/litchi-xldm/src/xml_attributes.rs
crates/xml-minifier/src/xml_attributes.rs
```

The implementation disables quick-xml's duplicate check, performs a bounded
borrowed-name scan, and falls back to the existing ordered map at the declared
boundary. The iterator algorithm preserves duplicate precedence,
malformed-value behavior, iterator fusion, and byte positions according to the
source and semantic reviews. The original candidate archive also contained an
OLE visibility narrowing; the exact two-token compatibility repair is recorded
below and was required for the measured candidate source. The inherited source comments
still carry older 0802 wording; the packet records that mismatch rather than
treating a comment as performance evidence.

The frozen main matrix has six shapes—`tiny`, `medium`, `large`, `vendor`,
`unicode-vendor`, and `valid-4attr`—with `capture`, `commit`, and `lifecycle`
rows. Native timing is six alternating paired blocks, 30 measured samples and
three warmups on CPU 12: 216 reports and 6,480 samples. The separate allocation
lane has two paired blocks and three samples without warmup: 72 reports and
216 samples. The before-only qualification has 18 reports and 18 samples.
Four large-capture heaptrack children and eight offline decodes are a separate
resource-accounting gate; they are not latency or RSS claims.

The main policy requires at least one capture or lifecycle improvement of 3%,
with the paired-ratio bootstrap upper endpoint below one. Any row whose median
ratio exceeds 1.05 with a bootstrap lower endpoint above one rejects the
candidate. Each paired allocation block must also avoid increases in median
allocation calls, allocated bytes, net live bytes, and peak bytes above entry. The cross-format
control is a separate DOCX/XLSX veto lane: a row rejects when its median ratio
exceeds 1.05 and its bootstrap lower endpoint exceeds one. The main and cross
seeds are 806080 and 806081. Main process p50 uses nearest rank; cross-format
p50 uses the integer midpoint of the two middle samples. Both use 10,000
median bootstrap resamples with zero-based interval endpoints 250 and 9749. The full policy is in
[adoption-policy.json](results/change-0806/adoption-policy.json), and the
matrix is in [plan.json](results/change-0806/plan.json).

The dedicated `valid-4attr` fixture has 12 slides and 96 text tags. Each text
tag has four distinct, quoted, namespaced attributes and each slide has four
root namespace declarations. The four qualified names are
`lx1:probeOne`, `lx2:probeTwo`, `lx3:probeThree`, and `lx4:probeFour`; each
sample verifies 96 tags, four attributes per tag, 384 attribute occurrences,
384 value occurrences, and exact namespace URIs and values. The fixture is an
independent extension-preservation case, not a replacement for the ordinary
and vendor controls.

## Evidence completed before workflow trials

The before source remained bound to the base commit throughout setup. Three
probe setup versions are retained. Setup 0 exposed process-global allocator
interference with synthetic test counters; setup 1 then exposed dead-code,
documentation, and test-module placement issues; setup 2 passed the three
probe quality gates but its new reports used `valid4-attr` instead of the
frozen `valid-4attr` identity. Those 18 setup-2 reports are retained as
unaccepted evidence and are never pooled. The final probe quality record passed
formatting, 36 tests, and warnings-denied Clippy. The setup lineage and reader
repairs are described in [execution-notes.md](results/change-0806/execution-notes.md).

The final before qualification completed all 18 rows. Root replay and
independent review verified the source, package, full-text, readback, and
preservation identities. The 15 inherited rows match their sealed 0792
qualification fields; the new extension fields are explicitly null on those
rows. The six shapes' retained package identities are:

| Shape | Source and capture | Commit/lifecycle |
| --- | --- | --- |
| `tiny` | 31,530 bytes, `26b1487517882224f820f9f2384c6428dcc40fa2d12c306d74df2809f4003f49` | 31,549 bytes, `10d6819120dc881e89ae4d191a85b550d6bb11efc20418565fac605b79adc07a` |
| `medium` | 40,788 bytes, `50ad2f81099ee29d4768d7080b7fc51ea2b5ca2aadd531031efcab65e8d5409e` | 40,806 bytes, `eee54f80d033e9e423625157dff22943f146e5989f814e61adc9fbbbda2ddcf4` |
| `large` | 215,220 bytes, `9c46542b763fc4bef63dfe4336cadd2bfba2b7e7b3f18a376c3924eb5643b3e9` | 215,240 bytes, `af76af57c33a47c69cf921d41fd1b961a566557c372ca0e546e2e4085157ee15` |
| `vendor` | 42,433 bytes, `e8b519bbbcf4140fe3d57f0cedbeb0f266fa0bbc3d1509d4b392151f323f2be8` | 42,450 bytes, `4ceca02399d8e993511be03046ed0970a23ad0d942ba95925c769eed24fdb4b4` |
| `unicode-vendor` | 42,370 bytes, `1b18f9837140109e64aba0f28ab23e69e86e52df71360d9773d5b8fd4aae426a` | 42,386 bytes, `e6b516c103604d8f1de75861eb1bd36bbf2c39f055d2a11c75acf3f1b1231ac8` |
| `valid-4attr` | 42,304 bytes, `a15d498e5201902f749b9be08e3d698bf1c1ebc8a73aba467558105c382519f3` | 42,323 bytes, `7af3fecea9679c1f7e713a7b0c5aa4b0e261e5a30b2f1ff6947c637fdc419133` |

The before source manifest is 1,182,921 bytes with SHA-256
`fbbead35ddcc01aae68676d9f47b8f7bf1d94108ffa56825979e8340fe25fc0d`.
The cross-format before qualification completed one report containing eight
rows and eight samples, bound to the same before source. The after builds and
all frozen capture children completed successfully; the resulting main
analysis contains 306 reports and 6,714 samples.

## Supplemental protected micro-preflight

The original candidate failed the first production quality attempt because the
constructor stopped using the retained private `unchecked_attributes` helper.
The exact constructor-reuse amendment was mirror-tested with 70 baseline tests,
100 amended tests, and warnings-denied Clippy, and both native probe builds
completed. The fresh protected preflight then ran 936 reports and 28,080
samples. It compares direct `construct` and `consume` micro-inputs, not public
Office workflows. Its [analysis](results/change-0806/amendment-preflight/analysis.json),
[independent audit](results/change-0806/amendment-preflight/root-native-audit.json),
and [decision](results/change-0806/amendment-preflight/decision.json) all retain
the raw report custody and numerical replay.

The required one- and two-attribute consume benefits pass the supplemental
policy:

| Case | After/before p50 | Median change | Bootstrap interval |
| --- | ---: | ---: | ---: |
| `distinct-1` | 0.6219275223 | -37.807% | 0.6079273343–0.6252781494 |
| `distinct-2` | 0.8109240519 | -18.908% | 0.7941903222–0.8114593622 |

All protected consume classes pass their preflight veto. Five non-protected
consume rows remain diagnostic regressions and are preserved explicitly:

| Case | After/before p50 | Median change | Bootstrap interval |
| --- | ---: | ---: | ---: |
| `distinct-4` | 1.2255830582 | +22.558% | 1.2098363613–1.2385104074 |
| `distinct-32` | 1.0516035528 | +5.160% | 1.0454343867–1.0552642513 |
| `duplicate-valid-after-4` | 1.1140148348 | +11.401% | 1.1009783977–1.1239792296 |
| `syntax-flag-after-4` | 1.2755048745 | +27.550% | 1.2587487288–1.2864896402 |
| `syntax-equals-value-after-4` | 1.2837389554 | +28.374% | 1.2714482525–1.2965331890 |

The decision therefore records `advance_to_workflow_trials: true` and
`production_adoption: false`. Its claim is limited to protected native
micro-input timing. It records no allocator, profile, Callgrind, resource, or
public-workflow speedup claim. The supplemental rows are not combined with
the 18 workflow qualification rows, the cross-format rows, or any later
workflow samples.

## Source-quality repairs and current gate

The immutable original candidate application remains in
[application.json](results/change-0806/application.json). The constructor
repair has its own [amendment application witness](results/change-0806/quality-amendment-application.json).
The first production quality run passed formatting but failed the locked
all-features/all-targets check because `litchi-sign`'s private
`unchecked_attributes` method became unused. The repair reuses the existing
inline helper and does not suppress that lint.

The next quality run passed formatting and then failed compilation because the
candidate had narrowed the OLE helper's `BytesStartExt` trait and
`CheckedAttributes` type from `pub` to `pub(crate)`. Downstream
`litchi-crypto` imports that trait and calls `checked_attributes()`. The exact
visibility amendment changes only those two declaration tokens, restoring the
baseline API; its [review](results/change-0806/visibility-amendment-review.md)
and [application witness](results/change-0806/visibility-amendment-application.json)
are retained.

`quality-2` passed all six gates: formatting, all-features/all-targets
compilation, tests, warning-denied Clippy, warning-denied rustdoc, and dependency
boundaries. The 500 test suites report 12,885 passed, zero failed and 89 ignored.
The retained [quality summary](results/change-0806/quality-summary.json) records
those counts from the successful production gate. Root froze the after-build
handoff against the visibility-restored source. All three main after-builds,
all three after-probe-quality gates (36 tests), and the cross-format after-build
completed successfully. The main native/allocation captures, cross-native
captures, and four heaptrack children also completed successfully.

## Main workflow result: rejected

The frozen main analysis uses the paired process p50 for each row and a
10,000-resample median bootstrap with seed 806080. The table reports the
after/before ratio, its median change, and the 95% interval. A benefit is
eligible only for `capture` or `lifecycle`, requires at least a 3% improvement,
and requires the interval's high endpoint below one.

| Case | After/before p50 | Median change | 95% interval |
| --- | ---: | ---: | ---: |
| `tiny/capture` | 1.001528 | +0.153% | 0.996748–1.004903 |
| `tiny/commit` | 1.005779 | +0.578% | 0.999207–1.008328 |
| `tiny/lifecycle` | 1.011651 | +1.165% | 1.006953–1.014534 |
| `medium/capture` | 1.001236 | +0.124% | 0.997932–1.004487 |
| `medium/commit` | 1.003197 | +0.320% | 0.997917–1.008333 |
| `medium/lifecycle` | 1.010959 | +1.096% | 1.004373–1.014342 |
| `large/capture` | 1.018526 | +1.853% | 1.012106–1.020932 |
| `large/commit` | 0.989316 | -1.068% | 0.985353–0.992373 |
| `large/lifecycle` | 1.031050 | +3.105% | 1.024719–1.033025 |
| `vendor/capture` | 1.019683 | +1.968% | 1.009681–1.022140 |
| `vendor/commit` | 1.003400 | +0.340% | 0.985528–1.014469 |
| `vendor/lifecycle` | 1.015042 | +1.504% | 1.012985–1.020089 |
| `unicode-vendor/capture` | 1.019829 | +1.983% | 1.013545–1.025324 |
| `unicode-vendor/commit` | 1.004961 | +0.496% | 1.003113–1.007479 |
| `unicode-vendor/lifecycle` | 1.014539 | +1.454% | 1.012834–1.016230 |
| `valid-4attr/capture` | 1.021855 | +2.186% | 1.004563–1.031418 |
| `valid-4attr/commit` | 1.008884 | +0.888% | 1.001891–1.017104 |
| `valid-4attr/lifecycle` | 1.018966 | +1.897% | 1.017324–1.020225 |

No capture or lifecycle row satisfies the benefit gate. The largest eligible
change is `large/lifecycle`, which is 3.105% slower and has an interval wholly
above one. The only faster row is `large/commit`, an ineligible mode; it is
1.068% faster and does not satisfy the required 3% capture/lifecycle benefit.
No row exceeds the 5% latency veto threshold with its interval wholly above
one. The independent [root audit](results/change-0806/root-audit.json) reports
306 reports, 6,714 samples, no latency violations, and
`production_adoption: false`. The machine-readable [analysis](results/change-0806/analysis.json)
records `benefits: []` and `adoption_eligible: false`; the retained
[disposition](results/change-0806/disposition.json) records `status: rejected`.
Root restored every allowlisted production file to the exact before bytes; the
[restored-source witness](results/change-0806/restored-source.json) binds that
state to `e3ff267ee3454e71d66f177f54f3cd05e0d9cce5`.

## Allocation and profile findings

The 72-report, 216-sample allocation lane passes every resource guard. Across
all shapes and modes, neither paired block increases allocation calls, allocated
bytes, net live bytes, or peak bytes above entry. The large capture illustrates
the measured resource change:

| Large capture operation metric | Before | Candidate |
| --- | ---: | ---: |
| Allocation calls | 72,106 | 10,788 |
| Allocated bytes | 4,767,939 | 853,507 |
| Net live bytes | 278,201 | 278,201 |
| Peak above entry bytes | 338,955 | 338,955 |

The allocation reduction is a scoped resource result, not a latency result and
does not override the required workflow benefit. The independent audit reports
`resource_guard: true` with no resource violations.

The profile lane has four qualified heaptrack reports and eight offline
decodes. Every owner and whole-process conservation check passes in both
repeats. The owner allocation calls are 72,106 before and 10,788 after; whole
process calls are 608,290 and 241,630. The nested quick-xml duplicate-check
count is 61,342 before and zero after, while nested notes-inspector counts are
61,274 and 456. These nested stacks overlap and are mechanism diagnostics; the
profile analysis makes no timing or RSS claim. See
[profile-analysis.json](results/change-0806/profile-analysis.json).

## Cross-format control

The cross-format lane completed 12 native reports and 2,880 samples after its
one-report, eight-sample before qualification. The official
[cross analysis](results/change-0806/cross-analysis.json) and
[independent raw audit](results/change-0806/cross-root-audit.json) pass after
binary cleanup. All eight DOCX/XLSX rows have no veto under seed 806081.

## Counts and final validation

The evidence classes remain separate:

| Evidence class | Reports | Samples | Status |
| --- | ---: | ---: | --- |
| Main before qualification | 18 | 18 | complete and accepted |
| Cross-format before qualification | 1 | 8 | complete and accepted |
| Protected amendment preflight | 936 | 28,080 | complete; micro-only advancement |
| Setup-2 qualification archive | 18 | 18 | retained but unaccepted; never pooled |
| Main native after/before timing | 216 | 6,480 | complete; no eligible benefit |
| Main allocation after/before | 72 | 216 | complete; resource guards pass |
| Cross-format native after/before | 12 | 2,880 | complete; independently replayed |
| Heaptrack/profile lane | 4 | 4 | complete; 8 decodes qualified |

Main, cross-format and profile evidence totals 323 reports and 9,606 samples.
With the separate protected amendment preflight, qualified evidence totals
1,259 reports and 37,686 samples. The 18 unaccepted setup reports are never
pooled into those results.

The main analysis, independent raw audit, cross-format analysis and independent
cross audit, profile qualification, supplemental raw audit and quality summary
all replay after cleanup. The owned target directory and all 19 copied binaries
were removed, with separate main, cross, setup and supplemental cleanup witnesses.
No candidate production change remains.

The retained setup failures, quality failures, reader repairs, and amendment
history document the execution chain. Archives, raw reports, logs, source
manifests, rejection disposition, restored-source witness and reviews remain
available for replay. The final packet seal binds this report and all five
performance indexes alongside the retained evidence.
