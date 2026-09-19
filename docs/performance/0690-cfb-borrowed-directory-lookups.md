# 0690 — avoid constructing owned keys for ASCII CFB lookups

Status: final validation passed; scoped retention independently reviewed. `performance_claim: none`.
OLE2/OOXML remain active; iWork is excluded.

## Measured hypothesis

After [0689](0689-xls-sst-chain-checkpoint.md), a fresh baseline at
`01d83efcf` attributes 18.00% self samples to `directory_name_data` in a
2,000,000-query owned 54016 first-cell run (0.789 seconds). The path rebuilds
both original UTF-16 and uppercase comparison SmallVecs for each requested
stream name, although the directory already retains validated comparison keys.
The lookup only needs the requested UTF-16 length and its ordered uppercase
code units. This suggests avoiding owned temporary keys for bounded ASCII
queries, while preserving the existing Unicode fallback and stored metadata.

An exact-spelling shortcut at the first tree node was considered and rejected
before implementation: Workbook is not the root sibling-tree candidate in
54016, Plan1, Simple or 45365-2. A new persistent stream descriptor would add
API/identity/lifetime concerns for work that may be removable inside the CFB
lookup owner. The experiment first measures the smaller lookup-only design.

## Evidence scope

The [packet](results/change-0690/README.md) uses unchanged 0684/0686 probes and
fixtures. Native A/A then A/B/B/A covers 24 case/source groups, 14,400 fresh
owners and 115,200 queries. Nine separate 50,000-query processes per leg provide
longer controls for 12 groups. Allocation, counted source reads and whole-
process hardware counters/RSS are captured separately. All long-loop cases
use the default 2 MiB index ceiling, including `54016-missing-1048576`;
the native matrix uses 1 MiB for that case.

A focused directory-lookup probe is supplemental mechanism and regression
evidence, not an end-to-end Office workflow. Both source modes use warm OS file
caches on a shared host with CPU 12 affinity. No physical cold-device, remote,
concurrent scaling, cross-platform or native Office claim follows.

## Implementation and proof

A private borrowed ASCII key handles only queries of at most 31 bytes.
After this bound and ASCII classification, it performs the existing empty,
NUL and first-forbidden-character checks. Every other query falls through
once to `directory_name_data`, preserving UTF-16 limits and error precedence.
Both the legacy and shared CFB lookup owners use the helper; writers, parsed
name caches and stored metadata remain unchanged.

The comparator orders original UTF-16 length first, then ASCII-uppercase
units against the cached Unicode comparison slice, then comparison-slice
length. Thus ASCII `Stream` still matches Unicode `ſtream`, while `STRASSE`
does not expand `straße`. Each tree node still requires its validated name
cache entry. A generic private traversal helper keeps each owner's bounded
walk single-sourced without a trait object or an enum holding the large key.
No public API, dependency, retained state, source read, source-freshness fence,
chain check or memory reservation changes. Per-node uppercase conversion is
an explicit performance risk checked by the wide-tree supplemental probe.

## Program priority beyond this batch

This is a scoped repeated-query improvement, not a substitute for end-to-end
CRUD work. Independent review identifies the real PPTX edit in change 0653
as a larger next target: 24.19 ms remained versus an 8.3 ms marker-free control,
with six whole-slide markup-compatibility passes per edit. These are historical
measurements, not fresh 0690 results. The next investigation should attribute
current pass identities and test transaction-local reuse with explicit
invalidation and memory pricing before adding a cache. Remaining XLS query
search/read costs need separate attribution before further optimization. The
[bounded next investigation](results/change-0690/next-pptx-investigation.md)
records the reviewed call path, invalidation obligations and measurement plan.

## Correctness validation

Seven new tests and extended existing matrices compare bounded ASCII
validation and comparison against the owned-key reference, including all
one/two-byte ASCII inputs, length boundaries, invalid names, Unicode and
supplementary-plane fallback. A frozen shared-reader traversal oracle covers
wide trees, case folding, the length-first witness, nested/root behavior and
missing/mismatched cached keys. A test-only altered comparison-slice length
checks the comparator's final tie-break. Existing legacy-reader tests remain.

All six CFB/XLS/facade gates and the DOC/PPT consumer test gate pass:
**4,407 passed, zero failed, 27 existing ignored**. This is scoped Rust
correctness evidence, not a new native Office or cross-platform certification.
The existing `parse_cfb` fuzz target was inspected, but this environment has
neither a nightly toolchain nor `cargo-fuzz`; no fuzz campaign is claimed.
Tool availability is retained in the packet.

## Paired query and workflow results

The primary eighth-query medians change as follows (B1/A1 and B2/A2):

| Target | Owned | File, warm OS cache |
|---|---:|---:|
| 54016 first, 2 MiB index | −8.93% / −8.93% | −1.82% / −0.91% |
| Plan1 first | −5.88% / −6.86% | −3.69% / −1.66% |
| Simple first | −6.52% / −5.49% | −1.66% / −0.48% |
| 45365-2 first | −5.77% / −5.77% | −1.60% / −3.69% |
| 54016 numeric late | −1.36% / −0.68% | −1.49% / −2.78% |

Owned `54016-stored-2097152` falls 560 → 510 ns. Longer same-owner controls confirm
−5.80%/−5.85% there, −3.80%/−3.93% for Plan1, −6.17%/−7.83% for Simple,
and −4.33%/−5.34% for 45365 first. File loop means are mostly within roughly
1%, including small regressions; no substantial general file-query gain is
claimed. Missing-target loops are also approximately flat.

These gains do not carry uniformly into complete workflows. `54016-stored-2097152`
open-plus-eight costs +1.24%/+1.62% owned and +1.22%/+1.08% file.
Its open phase is nearly flat; first/build queries carry roughly 1–2% costs.
The generated fixture workflow costs +1.93%/+2.33% owned and +2.57%/+1.90%
file, with first/build costs around 2%. Simple first workflows improve
−2.34%/−1.70% owned and −1.00%/−0.43% file. Formula-refusal workflows are
within 1%, with errors uncached. Index-disabled 54016 workflows cost
+0.58%/+1.17% owned and +1.06%/+0.41% file. These measured costs remain explicit even
though they fall below the initial 5% review trigger.

The primary paired median trigger is Simple missing owned q3: 70 → 80 ns
(+14.29% in both pairs). Its q8 stays 70 ns. There is no directory lookup on
an indexed missing result, so this is not attributed to removed key work.
The primary 45365-late file B1 first/build queries rise +92.20%/+52.27%,
and its workflow rises +54.80%, while B2's workflow is −3.21%. This large
one-leg anomaly is preserved and investigated separately below, not hidden
by a combined average. The 54016 missing owned workflow's mixed +0.79%/−5.15%
also has baseline drift and is not claimed as a gain.

Primary paired descriptive tail flags include Simple first/file warm-mean
p99 +70.54%/+63.55%, 45365 first/file warm-mean p99 +49.67%/+51.92% (A/A
+65.99%), and its open p99 +15.62%/+14.80% (A/A +13.81%). Smaller flags
include 54016 missing q3 p95/p99, Simple missing q3 mean/q8 p95, Simple missing
file q1 mean, and Plan1 late file q1 p99. Every flagged value and all individual
phases remain in `tail-regressions.md` and `native-comparison.json`.

## Mechanism, lookup controls and resource costs

The uninstrumented 2M-query profile changes 0.789 → 0.754 seconds (−4.46%),
with approximate cycles 3.489 → 3.338 billion. Candidate `ascii_lookup_key`
accounts for 1.59% self and `find_entry` 6.76%; the former 18% owned-key
construction is absent from the dominant query path. Total time, longer
controls and hardware counters support the gain, not symbol disappearance
alone. Both profiles report zero lost samples.

Owned 54016 extra-query instructions change 6,744 → 6,436 (−4.57%), and
cycles 1,747 → 1,646 (−5.77%). Simple instructions/cycles fall 5.26%/5.03%,
Plan1 4.10%/3.62%, and numeric-late 54016 only 0.63%/1.25%. Branch counts
fall, but numeric-late branch misses rise from about 1.21 to 2.15 per extra
query. Tiny or negative differenced cache-miss/page-fault values do not support
a locality claim; no uniform branch-prediction improvement is asserted.

Supplemental legacy lookup means improve 31–41% for short ASCII queries
across widths 1/31/257, including mixed case and the Unicode stored name
`ſtream`. Nested lookup improves about 37%, and missing-name lookup about
38%. Unicode-query/fallback controls remain within about 1.2%. Invalid-name
refusal is a real cost: roughly 6.7 → 7.9 ns, +17.77%/+16.60%. This probe
prices depth for generally six-byte ASCII keys; it does not establish
maximum-length key scaling or end-to-end CRUD latency.

The initial supplemental run had an unused-result warning in its warmup code.
Its exact probe, builds and captures remain in `initial-lookup/`. After an
explicit discard was added, both revisions were rebuilt with warnings denied
and all supplemental measurements repeated. The timed loop and production
sources are unchanged. The figures above use the final supplemental captures;
initial observations are separately audited and are not pooled. Main XLS
measurements, profiles and quality results remain on the same frozen sources.

All 96 allocation groups (three identical repeats per binary) match exactly,
including index construction; all 12 complete opening/eight-query counted
I/O routes also match. The removed SmallVecs were inline, so allocation parity
is expected and is not the CPU proof. No retained memory or budget charge is
added. Measured native-child RSS medians rise at most 1.02%; lower observed
values, including a 7.73% decrease in one Plan1 short process, are not a memory
bound or claimed retained-state saving.

Whole `.text` grows 637,379 → 638,227 bytes (+848, about 0.13%). `query_cell`
grows 34,104 → 34,142 bytes; its 2,040-byte stack reservation is unchanged.
Cursor/walk/cold-error and owned directory-key function sizes stay unchanged.
The new ASCII helper is 570 bytes with no explicit stack reservation in the
inspected x86_64 body. This does not establish a peak-stack or clean-build gain.

## Larger controls and remaining tail uncertainty

Four flagged groups receive an additional A/A then A/B/B/A capture with
1,000 fresh owners per leg and ten warmups: 24,000 owners and 192,000 queries.
Frozen binaries/probes and all outcomes match. Original 100-owner captures
remain intact; windows are not pooled and the follow-up does not replace them.
Raw supplementary records use lossless deterministic gzip outside probe timers.

The 45365-late file workflow anomaly does not recur: follow-up medians are
−1.40%/+0.85%, with q1 −0.63%/+0.58% and q2 −1.61%/+1.12%. The 45365-first
file warm-mean p99 flag also does not recur (−35.75%/−2.34%), and its open
p99 is −0.05%/−0.61%. No gain is inferred from the original unstable window.

The Simple missing owned cost **does recur**: q3 stays 70 → 80 ns in both
pairs, with flat A/A and ABBA baseline medians. Its q3–q8 mean p50 rises
71.67 → 76.67/78.33 ns (+6.98%/+9.30%), and p99 rises +6.25%/+8.33%.
Its complete workflow changes +2.37%/+0.01%. This is a retained small absolute
cost, not dismissed as noise or hidden by the stored-query improvement.

The Simple first/file warm-mean p99 flag **also persists**: +46.90%/+10.84%,
2,153 → 3,163 ns and 2,830 → 3,137 ns. A/A tails are 3,348/3,198 ns
(−4.48%), and the ABBA baseline tail rises +31.42%. The control variability
limits causal attribution, but does not erase the candidate tail observation.
Its workflow p50 is −0.30%/−1.24%, and workflow p99 +4.37%/−1.02%.
A new 45365-late warm-mean p99 single-leg flag is +0.78%/+57.88%, with
A/A −32.83% and ABBA baseline −33.39%. These distributions establish no
universal tail improvement. Complete supplementary phase statistics remain
in `followup-comparison.json` and `followup-summary.md`.

## Disposition

Retain the private lookup-only change for its independently reproduced warm
ASCII lookup and owned XLS repeated-query gains. The 1.2 ns invalid-name
cost, 10 ns missing-query cost, first/build workflow costs, persistent Simple
file tail flag, new late-file tail flag and 848-byte code increase remain
explicit. No retained-state charge or allocation/I/O tradeoff is introduced.
The decision does not imply that every workload improves.

The [review record](results/change-0690/review.md) records independent source,
test, priority, performance and evidence review. All 126 real XLS fixtures and
the full generated 70,001-cell visitor/digest comparison match. All repository
evidence gates pass; final audits bind sources, constraints, builds, profiles,
assembly, raw captures and outcomes. No registered claim or CRUD coverage
promotion follows. The broader non-iWork GOAL remains active, with the larger
PPTX edit-path investigation identified above as the next priority.
