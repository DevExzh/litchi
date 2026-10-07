# 0966 — DOC SPRM reservation census

**Status: completed diagnostic; no production candidate, adoption, or
production change.** The formal capture passed 48/48 fresh children and all
24 cells. Repeated reports are exact, errors are zero, every existing
correctness gate is true, and the independent post-capture audit passed. The
census has no timing, allocator-traffic, removable-work, or speedup claim.

## Question and source boundary

At `crates/litchi-doc/src/sprm.rs:182-186`, the current parser reserves
`max(grpprl.len() / 4, 8)` `Sprm` elements before it knows whether the input has
records. Inputs shorter than 32 bytes request the eight-element minimum. The
census asks how often this lifecycle sees empty or short groups. An empty-input
fast path was the only source-closed initial candidate; the result now points
to a broader bounded review of short non-empty parsing and its result
consumers. No CFB writer, recapture, transaction-boundary, parsed-vector
cache, late CFB deferral, or public-API change is selected.

The five external source reviews under `/home/zhuhe/litchi-performance-0966/`
provided the ranking, scenario limits, parser call boundary, and separation
from the speculative transaction design. Current `HEAD` is
`a499d34f5119a034bc92abde1e1dfcf529159b8d`. The path-scoped diff from
`73559c87c37898b675100b02985181a7764c66d2` is empty for
`crates/litchi-doc`, `crates/litchi-cfb`, and `crates/litchi-ole-common`.
The relevant parser consumers are `parse_chpx`/`parse_papx`
(`tracked_revision/codec.rs:453-525`), the `revisions()` resource check
(`tracked_revision/package.rs:1117-1119`), and the edit guard's revision scan
(`body_text.rs:4383-4388`).

## Evidence boundary

0942 remains the current-source CPU evidence: 18,209 formal samples included
5,866 lifecycle-marker samples, and `parse_sprms` had 5.7% self weight in both
FloatingPictures/32 repetitions. This is sampled CPU ancestry inside the DOC
lifecycle, useful for ranking only.

0919 is a different allocation report. Its `21,213` and `1,168` figures are
allocation-event counts within the `parse_sprms` subtree, not logical parser
invocation counts; operand allocations may be included. It reports
8,263,944 requested bytes for FloatingPictures/32, including 19,781 reported
384-byte minimum-capacity events, and 1,154 such events in the NoHeadFoot/32
pattern. These inclusive allocation observations cannot be numerically
compared with the 0966 logical call counts and do not establish CPU cost,
empty-input prevalence, or savings.

## Method and controls

The fixed matrix used `NoHeadFoot.doc` (26,112 bytes,
SHA-256 `45e5df073f34314da6f39d2dad119fb2ef23470878fd2df67f632864cd92ea48`)
and `FloatingPictures.doc` (335,360 bytes,
SHA-256 `41cc0cd8f2d7390266f844f5bddbcd29d08ac5338abc33832d9c9a6ed5d0f85d`)
at 1/2/8/32 edits and `zero`/`default_8mib`/`release_8mib` retention. It
used two fresh serial children per cell, zero warmups, one sample, and no
retry or replacement. The parser body, grammar, reservation policy, result,
error behavior, and returned bytes were unchanged. The diagnostic patch was
external only; the production sources were restored, so production remains
unchanged. `timing_claim` was false; capacity is a logical `Vec`-capacity
model, not allocator traffic or RSS.

The external build/control receipt reports the final 125,930,408-byte
diagnostic ELF (SHA-256
`d257020c302f92096c34d8bc821f0a88878900ab270f52a189b4e2889eaaf375`), 24
focused parser tests, two census guard/error controls, harness compile-only,
ordinary no-feature check, and 1,130 `litchi-doc` tests passed with 2 ignored
after fixture overlay. The receipt preserves two earlier compile/check
failures and two earlier fixture-overlay failures. This validates the staged
method and source overlay; it is not full-workspace QA.

## Census result

The bounded reduction is
[`analysis/census-summary-v1.json`](results/change-0966/analysis/census-summary-v1.json),
11,652 bytes, SHA-256
`53339c9a9c9e22af0f1773fdd2621351948ef0b207263ac3744a04f63d4121d5`.
The checker receipt passes all 48 reports, 24 cells, conservation, exact
repeat identity, zero errors, and all correctness gates. The raw reports remain
external at `/home/zhuhe/litchi-performance-0966/capture/doc-sprm-census-v1`.
The preferred child terminal/report manifest is
[`analysis/child-terminal-report-manifest-v1.json`](results/change-0966/analysis/child-terminal-report-manifest-v1.json),
35,266 bytes, SHA-256
`b23f755a37fe7d5edb06130647ea8b3d16b91e4cf411a3e5bc02c6ff5a0dd96f`.
The supplemental cell-detail file is external at
`/home/zhuhe/litchi-performance-0966/analysis/census-cell-details-v1.json`,
83,816 bytes, SHA-256
`6b45e9a2ddce1f5edde696619dea9edb8dcd7be772f129ef25179cfef9d56c1b`; its
48-child/24-cell cross-check passed. The independent post-capture audit is
`/home/zhuhe/litchi-performance-0966/postcapture-audit-v1.json`, status
`pass`, SHA-256
`f6fe42f9896e9895e9d9669bd36daae95ba87dc4d59691871357cbec15ffe293`.

Each row below is one representative cell; all three retention routes and both
fresh repeats have the same values.

| Fixture | 1 edit | 2 edits | 8 edits | 32 edits |
| --- | ---: | ---: | ---: | ---: |
| `NoHeadFoot.doc` calls / empty / logical bytes | 170 / 31 / 68,544 | 187 / 34 / 75,072 | 310 / 52 / 122,304 | 1,162 / 124 / 449,472 |
| `FloatingPictures.doc` calls / empty / logical bytes | 4,742 / 220 / 1,985,136 | 5,220 / 240 / 2,180,832 | 8,109 / 360 / 3,363,072 | 20,025 / 840 / 8,230,272 |

At 32 edits, the phase counts (`calls / empty`) are:

| Phase | NoHeadFoot | FloatingPictures |
| --- | ---: | ---: |
| `snapshot_open` | 63 / 11 | 1,657 / 80 |
| `edit_open` | 28 / 6 | 950 / 40 |
| `staged_replacement` | 944 / 96 | 15,696 / 640 |
| `final_commit` | 127 / 11 | 1,722 / 80 |
| `output_copy` / `other` | 0 / 0 | 0 / 0 |

Empty inputs are a minority in every cell. Short non-empty inputs dominate the
bounded parser census: at 32 edits the 4–7-byte bucket contains 974
NoHeadFoot and 12,203 FloatingPictures calls (the separate zero-byte bucket
contains the empty calls). These are logical call and modeled-capacity
observations, not CPU or allocator traffic.

## Disposition

The empty-only candidate is not adopted. The next useful source review is a
bounded short-input parse/result-consumer study that preserves whole-input
consumption, malformed-input ordering, PAPX/CHPX semantics, and existing
resource limits. The speculative late-CFB-publication design is not revived.
No new timing baseline follows from this diagnostic; the production source
remains unchanged. The supplement and independent audit now complete the
result custody record.
