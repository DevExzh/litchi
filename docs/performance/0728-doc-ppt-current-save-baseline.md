# 0728 — current DOC/PPT length-changing save baseline

The refreshed baseline confirms material common-container staging and finish
cost under the implemented 0663 Reuse policy. On the two DOC fixtures, finish
occupies 28.95–30.67% of the common-editor lifecycle; staging, which renders,
reopens, recaptures and discovers targets, occupies 53.83–57.42%. These are
within-owner phase fractions for the container control, not public DOC phase
fractions or an achievable end-to-end speedup. Production remains unchanged at
`1dc122a2bc`; `performance_claim: none`.

| Fixture / route | Whole p50, μs | Whole mean, μs | Stage median % | Finish median % |
| --- | ---: | ---: | ---: | ---: |
| FloatingPictures.doc, public edit/commit | 984.82–1089.01 | 991.14–1078.85 | — | — |
| FloatingPictures.doc, common Reuse | 243.05–247.64 | 243.66–247.45 | 53.83–54.10 | 30.13–30.67 |
| NoHeadFoot.doc, public edit/commit | 108.01–110.83 | 110.43–112.89 | — | — |
| NoHeadFoot.doc, common Reuse | 43.88–44.80 | 44.60–45.30 | 56.52–57.42 | 28.95–29.29 |
| 45543.ppt, public slide removal/commit | 1128.81–1140.66 | 1115.55–1132.03 | — | — |
| 45543.ppt, common Reuse control | 271.53–278.34 | 278.43–286.68 | 43.58–43.77 | 17.96–18.78 |

Ranges span six independent processes per route. Phase entries are ranges of
per-process medians of sample-level phase/whole ratios. All raw samples and
per-process p95, p99 and maxima remain in the packet; no sample is dropped or
trimmed. These ranges are descriptive, not confidence intervals. The larger
DOC public route has a visibly wider process range and should not be treated
as a stable single-number estimate.

The existing Rewrite control has lower common-container p50 ranges:
166.80–177.67 μs, 32.98–33.68 μs and 210.36–211.63 μs respectively. It also
changes 394, 40 and 51 normalized raw-directory bytes respectively, while all
Reuse and public outputs preserve the source directory image outside the
planner-owned allocation fields. Logical streams and the checked semantic
witnesses agree. These controls do not authorize switching the Reuse default
or treating the policies as preservation-equivalent. This is not a replication
of the old 0663 picture.doc result; the current fixtures are named above.

The source review distinguishes the mechanisms. DOC's batched stream edit
renders and validates a candidate, discards that rendering, then renders again
at finish. PPT's public slide removal uses its own embedded editor and writer;
it does not have that common-editor stage/finish duplication. Replaying PPT's
exact changed streams through the common editor is an alternative control.
Subtracting separate format and container medians would not attribute nested
cost. The old 0617 fractions are not used as current evidence.

Separate instrumented processes give identical allocation regions across all
three repetitions of each route. Public FloatingPictures.doc allocates
15,512,580 bytes in 15,756 calls, NoHeadFoot.doc 1,061,429 bytes in 1,907 calls,
and 45543.ppt 11,776,674 bytes in 5,663 calls. Their region-relative peaks are
3,128,190, 209,515 and 2,686,521 bytes. For common Reuse, DOC finish alone
allocates 1,397,553 bytes/615 calls and 125,206 bytes/249 calls respectively.
The opened/staged editor and returned output have explicit retained lifetimes;
negative finish retained-byte deltas reflect releasing editor ownership while
keeping the output. Separate phase peaks are not additive. These counters are
not RSS or copied-byte counts. [Complete route and allocation results](results/change-0728/analysis.md).

The prospective matrix contains two cycles with reversed case/route order,
three fresh processes per cell/cycle, two warmups and 30 measured owners per
native process: 54 processes, 1,620 measured lifecycles and 108 warmups.
Another 27 allocation processes take one owner each. All 81 processes are
pinned to CPU 12 on AMD EPYC 9R45 with Rust 1.95.0, release builds and warm OS
caches. Fixtures, complete workspace source, accepted constraints, probe,
executables, commands and scripts are bound before collection. Native timing
uses the system allocator; allocation instrumentation runs separately.

The DOC edit replaces paragraph zero with 45 UTF-16 code units, versus source
lengths six and 76. The PPT edit removes the selected second live slide
(slide ID 472, persist ID 10), taking 11 live slides to ten. Both DOC streams
change logical length; PPT appends its updated history and changes the document
stream length. A physical CFB size change alone is not accepted as proof.

Untimed checks validate CFB structure, exact stream membership/content,
source-to-output unchanged streams, allowed changed paths, root/storage CLSIDs,
normalized raw directory preservation, DOC target and survivor paragraphs and
available auxiliary projections, and PPT ordered survivors with direct public
value/live-record comparisons. Digests are report fields rather than substitutes
for those direct comparisons. Unavailable projections are explicitly reported;
PPT dependencies inside the allowed changed PowerPoint Document stream are not
claimed fully verified. The probe rejects all 23 applicable case/control
combinations, repeated 621 times across the matrix. These include stream loss,
untouched same-length mutation, metadata mutation and wrong requested edits.

Final qualification passes formatting, four probe unit tests, warning-denied
all-target Clippy and rustdoc, and 18 route/lane smoke processes. Initial compile,
test-helper, lint and formatting failures, earlier sources and review findings
remain archived. Independent review closes the substantive oracle findings.
The analyzer, separately implemented statistics/custody auditor, and 15 evidence
corruption controls pass. Cleanup removes only the two owned build/binary roots;
post-cleanup checks replay the exact results with executable identity witnesses.
No production code changed, so this batch makes no new full-workspace correctness,
Office, cross-platform, RSS, cold-storage or concurrency claim.

The next bounded investigation is attribution inside the actual public DOC
owner, using its existing commit/CFB observer hooks and checking observer cost.
That can establish whether returning the already validated rendering through a
batched handoff materially helps the whole edit/commit lifecycle. Any candidate
must preserve validation, failure atomicity, exact no-ops, successive edits,
source/policy freshness and bounded retained memory. A generic render cache is
not justified by this baseline. PPT needs its own public-route attribution;
the common-control fractions do not nominate an equivalent PPT optimization.
The existing 0652 placement decision and 0663 policy remain authoritative.

[Evidence and replay index](results/change-0728/README.md). The broader non-iWork
goal remains active; no registered claim or CRUD coverage status is promoted.
