# 0918 — DOC stream-to-Arc publication handoff

The isolated three-line candidate is **rejected** by its prospective latency
gate. It saves requested allocation bytes, but FloatingPictures at two and
eight edits fails the required native confidence bound. Production code stays
unchanged; the complete experiment is retained without exclusions or reruns.

This batch follows [0917's real-DOC repeated-edit study](0917-doc-repeated-edit-retention.md).
The production baseline is unchanged between `bb19ca74bd52` and `f5485b3ca858`.
The retained 0917 probe measures fresh public open/edit, repeated replacements,
commit, output materialization and local owner destruction. Preparation,
reference routes, preservation oracles and JSON remain outside that boundary.

## Diagnosis and candidate

The [source audit](results/change-0918/design-audit.md) traces clone-first DOC
mutation, stream publication, CFB rendering and required validation. It confirms
that the existing retained-render handoff and earlier CFB batching/layout work
are already adopted. The proposed change targets a separate temporary copy:
`RevisionEditor::commit` constructs each immutable stream allocation from a
cloned mutable `Vec`. Constructing the `Arc<[u8]>` directly from the original
slice retains independent immutable bytes while removing the temporary `Vec`
allocation and its extra payload copy.

The isolated candidate changes only those WordDocument, table and changed-Data
conversions. It retains the mutable editor's vectors, the missing-Data add path,
atomic replacement batch, resource/protection/signature checks, retained-render
policy, all required CFB validation and final strict-owner/public-reader checks.
It introduces no cache, borrowed published storage, public API or dependency.

Scoped Callgrind profiles use the retained native probe with zero warmups and
one sample, collection toggled around `measure_one`. `run_case` is invoked once
within that scope. These are guest instruction and memory-reference diagnostics;
copy/SHA dispatch and allocator behavior prevent interpreting their shares as
native time. The zero-collection attempt and whole-process exploration are
retained separately from the accepted scoped profiles.

The accepted original profiles contain 27,097,518 guest instructions for
FloatingPictures at two edits, 224,621,383 at 32 edits, and 24,163,805 for
NoHeadFoot at 32 edits. At 32 FloatingPictures edits, inclusive
`RevisionEditor::commit` accounts for 57.04% and stream publication for
50.18%; these nested shares overlap. The `memcpy` row is 53.28% of guest
instructions, which motivates inspecting copies but does not establish a
native-time saving. Three reproduction profiles preserve scope and exact
output hashes; their instruction totals vary, including +2.57% for the
32-edit FloatingPictures run. All six accepted/reproduction captures pass
the [strict offline verifier](results/change-0918/profile-replay.py).
The retained nonfatal Callgrind `brk segment overflow` warning and the
failed zero-event collection attempt remain visible in the
[profile report](results/change-0918/profile-report.md).

## Frozen comparison and decision

A separate comparison uses fresh matched baseline/candidate native processes,
with an unchanged probe and exact output/semantic/preservation gates. Native
comparison covers both fixtures and all four edit counts under default
retention. Allocation comparison covers all three retention policies. The
frozen comparison protocol defines sample counts, ordering, uncertainty and
admission criteria before any comparative measurements.

The [protocol](results/change-0918/comparison-protocol.json) requires at least
5% fewer median allocated bytes for one 32-edit default-policy fixture, no
allocator cell/repetition above +1% allocated bytes or +5% peak live bytes,
and all eight native bootstrap lower bounds at least `1/1.05 = 0.952381`.
Twelve matched rounds alternate version order and reverse fixture/count order.
Each native process has five warmups and 50 measured samples. The separate
allocator lane runs two repetitions of every fixture/count/policy/version,
with zero warmups and one sample.

All 192 native children (9,600 measured samples, 960 warmups) and 96 allocator
children passed. The capture ran from 18:52:29 to 18:53:36 UTC on 2026-10-05,
pinned to CPU 14. Coordinated workers were idle; the observed one-minute host
load ranged from 3.18 to 5.96 and unrelated host activity was uncontrolled.
The [reduction](results/change-0918/comparison-reduced.json) retains every
child and all adverse flags. Its adoption result is false.

| Fixture | Edits | Pooled p50 baseline → candidate (µs) | Median paired baseline/candidate ratio | Paired 95% CI | Latency gate |
| --- | ---: | ---: | ---: | --- | --- |
| NoHeadFoot | 1 | 118.680 → 117.100 | 1.009751 | [1.003715, 1.024616] | pass |
| NoHeadFoot | 2 | 146.341 → 144.711 | 1.009500 | [1.006593, 1.018833] | pass |
| NoHeadFoot | 8 | 300.011 → 298.391 | 1.005195 | [0.998718, 1.007822] | pass |
| NoHeadFoot | 32 | 1020.935 → 1007.074 | 1.013304 | [1.010230, 1.017089] | pass |
| FloatingPictures | 1 | 1142.545 → 1085.405 | 1.020084 | [0.961492, 1.139740] | pass |
| FloatingPictures | 2 | 1215.905 → 1341.456 | 0.908414 | [0.871292, 1.008271] | **fail** |
| FloatingPictures | 8 | 2561.472 → 2561.471 | 0.970854 | [0.933940, 1.028418] | **fail** |
| FloatingPictures | 32 | 8304.077 → 8081.245 | 1.038432 | [1.001944, 1.048652] | pass |

The bootstrap resamples the twelve matched process ratios (10,000 draws,
seed 918); pooled quantiles are descriptive and are not the gate statistic.
The two failed intervals span 1, so this experiment does not establish a
certain intrinsic slowdown. It does fail to establish the required bounded
regression, which is sufficient to reject adoption. Shared-host noise does
not authorize dropping rows or weakening the frozen threshold.

| Fixture | Edits | Default allocated bytes baseline → candidate | Calls baseline → candidate | Peak live bytes, both |
| --- | ---: | ---: | ---: | ---: |
| NoHeadFoot | 1 | 1,008,777 → 996,873 | 1,768 → 1,766 | 303,110 |
| NoHeadFoot | 2 | 1,282,207 → 1,257,236 | 2,269 → 2,265 | 312,324 |
| NoHeadFoot | 8 | 3,181,346 → 3,053,550 | 5,368 → 5,352 | 370,820 |
| NoHeadFoot | 32 | 16,781,690 → 15,823,914 | 19,024 → 18,960 | 642,194 |
| FloatingPictures | 1 | 14,304,400 → 14,233,642 | 15,498 → 15,496 | 4,284,552 |
| FloatingPictures | 2 | 17,369,691 → 17,217,392 | 18,014 → 18,010 | 4,370,204 |
| FloatingPictures | 8 | 37,737,565 → 36,869,577 | 33,187 → 33,171 | 4,891,053 |
| FloatingPictures | 32 | 164,233,879 → 156,566,655 | 95,057 → 94,993 | 7,635,361 |

Both allocator repetitions match exactly for every allocation metric. All
three policies show fewer allocated bytes/calls and unchanged peak live and
retained output bytes. The 32-edit default savings are 5.7073% for NoHeadFoot
and 4.6685% for FloatingPictures; the allocation-benefit gate therefore passes.
The unchanged peak demonstrates that these temporary handoffs were not the
peak-setting allocation overlap in these workloads.

All per-round p50/p95/p99 and whole-child RSS results remain in the reduction.
Across 96 native pairs there are 17 p50, 19 p95 and 19 p99 increases above 5%,
including two NoHeadFoot two-edit p99 flags and adverse FloatingPictures rows
at every count. No native RSS pair exceeds +5%. One allocator RSS pair does:
FloatingPictures, eight edits, release policy, repetition one rises from
13,768 to 14,604 KiB (+6.07%). No allocator byte or peak-live pair crosses its
adverse threshold. These observations are retained, not used as exclusions.

## Build and correctness scope

The baseline is the retained 0917 build from `bb19ca74bd52`. The candidate
uses `f5485b3ca858` plus the isolated three-line patch; production sources and
build configuration are unchanged between those revisions. Both use the exact
eight-file 0917 probe and its separate, locked dependency graph. The
[source receipt](results/change-0918/candidate-source.json) binds the lean
281,774,080-byte archive, patch, fixture export and complete file manifests.
The [final build receipt](results/change-0918/candidate-build-final.json) is
authoritative: its before/after file rows match, and both executable hashes
match the original candidate build. The earlier build receipt remains a
historical record. [Build comparability](results/change-0918/build-comparability.md)
distinguishes recorded settings from inferred defaults and unavailable
historical compiler arguments.

All 96 premeasurement correctness children passed: both versions and lanes,
both real fixtures, four edit counts and three retention policies, with zero
warmups and one sample. These children do not enter performance statistics.
Checks include exact output identities, public semantic reread, untouched
stream/directory/CLSID preservation, patch inversion, exact no-op sharing,
atomic refusal and the retained-render transitions.

The isolated candidate passed 1,365 DOC tests (two ignored), 14 doctests
(12 ignored), warning-denied all-target/all-feature Clippy, DOC formatting,
and the existing changed-Data picture transaction test. The
[quality receipt](results/change-0918/candidate-quality.json) also retains
failures: whole-workspace formatting finds existing unrelated drift, and
warning-denied public rustdoc fails with identical baseline/candidate
diagnostics for private `parse_recording` links. Private-item rustdoc also
fails on unresolved links. This is not an all-green workspace certification.

The root workspace lock required an offline update for quality checks. Its
original and generated copies are retained externally; the separate probe
lock did not change. The first failed lock/Clippy attempt's stderr was
overwritten by its successful rerun, so that attempt is disclosed without
claiming complete raw-log custody or an unproven cause.

## Reproduction and evidence custody

The [comparison manifest](results/change-0918/comparison-manifest.json)
binds the unchanged 0917 probe, validator, fixtures, protocol, two source/build
lineages and four executables before capture. Native and allocator stdout,
stderr and process metrics have deterministic gzip companions below the
packet; each receipt binds compressed and uncompressed size/SHA-256. Large
Callgrind files, source archives/manifests and all executables remain external.
Their receipts record file names, sizes, hashes and regeneration commands.
The original qualification raw files also remain external.

In-place replay passed, but the independent literal-copy replay exposed an
absolute fixture-path mismatch that the pre-capture path-resolution smoke
did not exercise. A separate hash-bound replay adapter supplies the original
recorded input path through the unchanged validator's existing alias field,
after verifying the copied fixture bytes. The
[dated repair note](results/change-0918/replay-repair.md) records this limitation.
No frozen script, raw report, receipt, statistic or admission result is changed.
The [final independent review](results/change-0918/review.md) recomputed all
statistics, replayed the exact final adapter twice in a fresh copied hierarchy,
and rejected altered path aliases, fixture bytes, gzip bytes and reduced
results. It confirms the scoped rejection; its temporary copied/corrupted
trees were cleaned after retaining commands and proof receipts.

Run the retained verifiers from the checkout:

```sh
python3 -B docs/performance/results/change-0918/profile-replay.py
python3 -B docs/performance/results/change-0918/qualification-replay.py
python3 -B docs/performance/results/change-0918/comparison-analysis.py
```

The qualification/profile commands require the external custody paths in
their receipts. A relocated comparison packet requires the reconstructed
repository hierarchy documented in its
[replay repair](results/change-0918/replay-repair.md), including the copied
0917 validator, protocol, probe and fixtures; it is not an arbitrary standalone folder.
Do not overwrite the retained schedule. A new experiment needs a fresh packet
and empty external root, with identities frozen before its measurements.

To rebuild the candidate, generate/extract the archive using
`candidate-source.json`'s exact `regeneration_argv`, apply the retained
`candidate.patch` at its root, and use the exact locked/offline standalone
release command in `candidate-build-final.json`. The supplementary workspace
lock is only for the separate quality checks. Baseline regeneration uses the
0917 source/build receipts. Profile regeneration uses
`profile-reproduce.sh` with a fresh external destination. The synthetic
admission controls and independent review are retained alongside the packet.
Exact large-source-manifest regeneration commands are recorded in
[candidate reproduction](results/change-0918/candidate-reproduction.md).

The [postcapture custody check](results/change-0918/postcapture-source.json)
verified all 10,898 exported source rows, eight probe files, fixtures and final
executable hashes, and confirmed that root production still matches baseline.
Cleanup removed 22,409,812,460 bytes of temporary exports, superseded files and
build caches. Both final executables remain at their recorded external paths,
along with the source archive, complete manifests, lockfiles, build/quality
logs and retained compiler-fingerprint evidence. Post-cleanup stability passed.

## Scope and remaining work

The workload repeatedly replaces paragraph zero; it is not a bulk edit of
distinct paragraphs. Source bytes and final semantic text are fixed, but
preserving the append-style stream representation grows output over longer
edit sequences. The default retained-render ceiling bounds that retained
buffer, not the whole transaction's transient memory. Native timing includes
fresh in-memory open, all repeated edits, commit, output materialization and
local destruction; it excludes file I/O and post-sample correctness oracles.
Allocator totals are requested bytes and region-relative live bytes, while
whole-child RSS also includes untimed setup and oracles.

The experiment adds no producer, selector, distinct-object coverage,
cold/range-source, concurrency or native Office certification. Other work on
the shared host remains uncontrolled. Coordinated build/profile/test work
must be idle for the timed comparison, but that does not imply an otherwise
idle machine. The broader non-iWork program remains open.
