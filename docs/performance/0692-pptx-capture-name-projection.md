# 0692 — reuse slide XML for capture root and name projection

Status: retain the measured capture-only optimization.
`performance_claim: none`. Baseline: `e9bd3360c`.

The real 13-slide one-edit phases fall from **22.76–22.88 ms to 17.11 ms**
median, a paired **24.83–25.24% reduction**. Capture and commit each avoid one
redundant MCE processing pass per slide. The change retains the original root,
relationship, identity, name and notes validation order, without a persistent
XML cache or public API change.

The [evidence packet](results/change-0692/README.md) contains frozen builds,
source/constraint/corpus/probe/binary hashes, raw native samples, separate
allocation diagnostics and profiles, temporary trace/restoration receipts,
and independent code/evidence reviews. The broader GOAL remains active;
iWork is excluded, with no coverage promotion or registered performance claim.

## Native scope and individual results

Thirteen workflows have two pre-edit A/A legs and four A/B/B/A comparison
legs: **7,800 native samples**, each from a fresh package, with five warmups
per process. The table preserves baseline a2/a3 and candidate b0/b1 medians;
paired changes are b0/a2 and b1/a3. CPU 12 is pinned on this shared Linux host,
with warm caches and no demonstrated host-wide quiescence. The preceding A/A
total medians range from −1.99% to +3.51%, with total p95/p99 drift below 3.60%.
Small 1–4% notes-case differences are observations close to that drift range,
not strong general speedup claims.

Timers cover capture, working clone, set text, commit and apply. The enclosing
total includes interphase timer/check overhead. File read, generated-corpus
construction, initial open, target selection, save and reopen are excluded;
this is not a complete open/edit/save measurement. Inputs are the real
`slide-section-test`, its marker counterfactual, generated 12-slide × 8-text-box
content, POI `prProps` notes, and LibreOffice `tdf131082` notes. Two-edit cases
select shapes on distinct slides.

| Workflow / input | Baseline medians (ms) | Candidate medians (ms) | Paired change |
| --- | ---: | ---: | ---: |
| one-real | 22.7578 / 22.8808 | 17.1082 / 17.1063 | -24.83% / -25.24% |
| one-control | 7.5416 / 7.5622 | 6.7687 / 6.7593 | -10.25% / -10.62% |
| one-generated | 1.8682 / 1.8725 | 1.7473 / 1.7236 | -6.47% / -7.95% |
| one-notes-poi | 1.0855 / 1.0843 | 1.0569 / 1.0610 | -2.63% / -2.15% |
| one-notes-lo | 1.7549 / 1.7390 | 1.6986 / 1.6921 | -3.21% / -2.69% |
| noop-real | 11.5702 / 11.5704 | 8.5170 / 8.5552 | -26.39% / -26.06% |
| noop-control | 3.8359 / 3.8310 | 3.4752 / 3.4375 | -9.40% / -10.27% |
| noop-generated | 0.9150 / 0.9076 | 0.8355 / 0.8331 | -8.69% / -8.21% |
| noop-notes-poi | 0.4938 / 0.4978 | 0.4851 / 0.4858 | -1.76% / -2.41% |
| noop-notes-lo | 0.6499 / 0.6579 | 0.6275 / 0.6307 | -3.44% / -4.13% |
| two-real | 23.0698 / 23.0332 | 17.2437 / 17.3572 | -25.25% / -24.64% |
| two-control | 7.7174 / 7.6973 | 6.9374 / 6.9184 | -10.11% / -10.12% |
| two-generated | 2.0843 / 2.0823 | 1.9418 / 1.9464 | -6.84% / -6.53% |

| Real one-edit phase | Baseline medians (ms) | Candidate medians (ms) |
| --- | ---: | ---: |
| capture | 10.5732 / 10.6561 | 7.7068 / 7.6974 |
| clone | 0.0168 / 0.0176 | 0.0168 / 0.0166 |
| settext | 1.0971 / 1.0964 | 1.0877 / 1.0823 |
| commit | 10.9574 / 11.0154 | 8.1865 / 8.2081 |
| apply | 0.0966 / 0.0972 | 0.0960 / 0.0959 |

| One-edit allocation diagnostic | Baseline | Candidate |
| --- | ---: | ---: |
| alloc_calls | 350,148 | 270,242 |
| realloc_calls | 9,030 | 6,950 |
| requested_bytes | 26,015,747 | 19,915,826 |
| peak_above_start | 463,159 | 463,159 |
| net_live_change | 195,730 | 195,730 |

Allocation columns above describe the complete timed one-edit interval.
Requested bytes already include the full new size of successful reallocations;
the separate realloc-request counter must not be added again.

All total means and p95 values improve in both comparison pairs; the largest
total p99 increase is 2.02%. There are **21 phase-tail review triggers** above
5%, all at p95/p99, retained in `native-review-triggers.json`. Clone p99 is
especially variable (the maximum increase is 125.30%, or 7.33 µs, on no-op
POI); other triggers affect apply, no-op commit and two-edit set-text. The
real two-edit set-text p99 rises 23.08%, or 0.302 ms, in one pair. No phase median or
mean exceeds the 5% regression trigger. These tails remain visible; there is
no universal tail-latency guarantee and no pooling away unfavorable legs.
Raw nearest-rank tails and seeded within-leg median bootstrap intervals are
in `native-summary.json`; bootstrap intervals are descriptive, not independent
population estimates.

## Mechanism, allocation and memory costs

One capture on the real deck falls from **44 to 31 MCE calls**, and from
**819,319 to 550,141 input bytes**. Owned output bytes fall from 944,179 to
634,121. There are still five presentation passes and two per slide: the
combined root/name projection and the unchanged notes scan. Commit capture
also has 31 calls, processing 549,999 input bytes after the edit. The generated
case falls from 43 to 31 calls; generated and marker-control outputs remain
borrowed. Baseline trace counts reuse 0691's exact production source; fresh
candidate capture/apply traces and restoration hashes are retained separately.

The private helper processes one slide, validates its root, obtains an owned
name result, then drops processed XML before allocating a fallback name.
It retains only the earliest name error, continues later root/relationship
checks, and consumes that error at the original name position. After that
first error it skips later name projections. Notes loading remains after
slide identity/name checks. Public part readers retain their existing behavior.

On this 64-bit ABI a temporary entry grows from **24 to 48 bytes**: +24 bytes
per slide, or 312 bytes for the real deck and 98,304 bytes at the default
4,096-slide ceiling. Successful name allocations move into the final snapshot
without duplication, but now remain live during later root processing.
The vector uses fallible exact reservation after the existing slide-count
check; at most one error is retained. No processed XML, cache or extra state
survives capture. Allocation-failure opportunities can differ because repeated
work is removed and the scratch reservation is now fallible; ordinary typed
validation errors and their precedence are preserved.

Across all 78 case/phase allocation groups, three samples repeat exactly and
**peak-above-start and net-live change match baseline**. For the real one-edit
case, allocation calls fall 22.82%, requested bytes fall 23.45%, and total peak
above start stays 463,159 bytes. This measures the chosen corpus, not a proof
that arbitrarily large names cannot raise transient peaks.

Native open-plus-capture counter slopes (210 minus 10 iterations, divided by
200) fall from 213.45M to 159.05M instructions and 54.00M to 39.96M cycles.
Page-fault slopes rise from 167.2 to 206.7 per operation; this observed
increase remains disclosed. Whole-child peak RSS is 5,688 versus 5,576 KiB in one
run per binary, insufficient for an RSS-bound claim. Candidate sampling still
attributes 18.17% self cost to MCE `start`; residual MCE/notes work remains.
Kernel symbols were unavailable. The native executable's `size` text column
grows by 4,048 bytes; no build-time or code-size improvement is claimed.

## Preservation, verification and disposition

Every native process passes the probe's semantic/revision/save/reopen checks.
The oracle compares sorted inventories, relationships, content types,
non-part metadata and untouched payload hashes; it is order-insensitive.
No-op requires equal before/after serialization and revision, not equality
between serialization and the original ZIP archive. Two-edit checks exclude
only the two edited slide payloads and validate both markers after reopen.
Edited-part unknown markup, physical metadata and ordering remain the domain
of the library preservation tests.

The marker counterfactual changes 43 URI spellings across 103 ZIP members.
Names, timestamps and uncompressed lengths are preserved; flags/external
attributes change and compression is regenerated. It changes semantics and
serves only as a mechanism control.

Five focused tests compare the actual capture with a frozen old two-phase
catalog sequence, checking successful names and typed/Debug error equality.
They cover Transitional/Strict, MCE, explicit/empty/fallback names, earliest
name-error retention, later bad roots or missing relationships, catalog
ID/target uniqueness, and notes/name/tail precedence. The main catalog already
rejects duplicate IDs, relationship IDs and case-folded targets; the capture
loop's later identity/target guards remain defensive for stable graphs.
Initial test-fixture failures were corrected before candidate builds; no
production change was needed to make the tests pass.

Formatting, all-feature/all-target checking, Clippy and warning-denied rustdoc
pass. Default-feature owner tests pass **908** with two existing ignored;
all-feature owner tests pass **922** with the same two ignored; the narrow
facade tests pass **45**. Counts include integrations/docs where applicable
and overlap between feature runs.

Full gate results are retained in `integration/results.json` and
`evidence/results.json`. The facade uses the narrow `pptx` feature with standard
rustc warnings. Three attempts to apply `-D warnings` to existing facade
feature combinations exposed unrelated unused helpers, an unused-mut test
fixture and an unused ODF field; their logs remain under `integration-initial`.
The owner check/Clippy/tests and rustdoc still deny warnings; no facade source
was changed. Existing OPC fuzz entrypoints were inspected, but
this environment has neither nightly nor cargo-fuzz; no fuzz campaign or
native Office application validation is claimed. No physical cold-cache,
remote-source, concurrency, cross-platform or general PPTX result is claimed.

The primary median improvement is substantially larger than measured A/A
drift, is reproduced by both candidate legs, and is supported by reduced
processing and allocation counts. Retain this small capture-local change
with its transient storage and phase-tail costs disclosed. Further work should
profile the remaining notes scan and presentation passes before proposing
additional reuse or a budgeted cache.
