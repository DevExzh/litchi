# Change 0188: DOCX/PPTX ordinary-root lifecycle evidence

Date: 2026-08-18

## Decision

Retain eight opt-in filesystem selectors that time a fresh high-level root open
plus one representative query. The existing source-backed DOCX and PPTX path
ingress remains unchanged; this tranche closes the missing end-to-end timing
boundary between the previously separate open-only and prepared-query cases.

| eager byte-owner control | source-backed filesystem candidate |
|---|---|
| `pptx_file_eager_open_slide_count_lifecycle` | `pptx_file_source_open_slide_count_lifecycle` |
| `pptx_file_eager_open_selected_slide_lifecycle` | `pptx_file_source_open_selected_slide_lifecycle` |
| `docx_file_eager_open_paragraph_count_lifecycle` | `docx_file_source_open_paragraph_count_lifecycle` |
| `docx_file_eager_open_full_text_lifecycle` | `docx_file_source_open_full_text_lifecycle` |

The timer covers `fs::read` plus `from_bytes` and the named query for the eager
control, or `Presentation::open`/`Document::open` plus the same query for the
source candidate. Root construction is not prepared outside the lifecycle
timer. Corpus construction, full semantic/archive comparison, source hashing,
and independent positional replay stay outside timing. The default matrix
remains 36 cases / 198 records; the selectable matrix rises from 324 to 332.

## Correctness and source-read gates

The fixed PPTX corpus has 200 slides, eight text boxes per slide, eight 2 MiB
incompressible media members, 445 archive members, and archive SHA-256
`61b2b99083ca27ebd37955db600955e3f41289b93dba71951983164239eff757`.
The fixed DOCX corpus has 200 paragraphs, eight 2 MiB incompressible media
members, 20 archive members, and archive SHA-256
`a4a2e4921235a6da6b38e31d26ddcca1301909885e37330ab4f83ecc0c4e04f4`.

Each child still performs the established untimed eager/source semantic and
archive gates. Across both source legs, all 60 PPTX count samples were
catalog-only with zero slide/media payload overlap; all 60 selected-slide
samples covered only slide 100 with no unselected-slide or media overlap; and
all 120 DOCX samples covered the main-document compressed range during owner
preparation with zero main/media/unselected/core overlap during the query.
Classification failure aborts the run.

Focused selector/parser/scope tests, the integrated eight-case smoke, strict
all-target harness Clippy with warnings and deprecations denied, scoped
formatting, and diff checks pass.

## Measurement protocol

A single exact release binary (SHA-256
`dba03ad4b0a0ea0d726661edf5d8bf7028ae63ee913d67fc045aa7b41a20720f`)
ran on CPU 2 in A1-eager/B1-source/B2-source/A2-eager order. Every leg used
three warmups and 30 retained warm-cache samples per case; every sample ran in
a fresh child process. The raw reports identify base revision
`749559b4b562d5b86e31ae06d6e8ff2b63afca1b` and deliberately record a dirty
worktree because the additive lifecycle selectors were not yet committed.
The release binary was rebuilt after final source review and before A1; later
documentation-only edits do not affect it.

Predeclared same-implementation drift ceilings are 5% for p50/mean, 10% for
p95, and 15% for p99. A statistic is accepted only when the source-backed
candidate is lower in both paired directions and both eager/source drifts pass.
Because this is the minimum 30-sample evidence tranche, every p95/p99 cell is
conservatively withheld even when one individual tail cell passes its ceiling.

## Result

| Lifecycle | A1 -> B1 p50 reduction | A2 -> B2 p50 reduction | Accepted |
|---|---:|---:|---|
| PPTX open + slide count | 83.57% | 81.78% | none |
| PPTX open + selected slide | 83.63% | 81.92% | none |
| DOCX open + paragraph count | 96.45% | 96.28% | none |
| DOCX open + full text | 95.88% | 96.17% | none |

All directions are descriptively lower, but no latency statistic is accepted.
PPTX source-backed p50 drift is 10.54% for count and 9.65% for selected slide;
DOCX source-backed paragraph-count p50 drift is 5.97%; and DOCX eager full-text
p50 drift is 5.03%. Each exceeds the predeclared 5% ceiling. Mean drift fails
the same workloads, and the tranche-wide minimum-sample policy withholds every
p95/p99 cell. The selectors and raw reports remain useful correctness and
attribution evidence; the directional deltas are not a speedup claim.

This accepts only warm fresh-process elapsed time for the named generated
corpora and lifecycle scopes. It does not establish physical-I/O,
cold-filesystem, decompression, allocation/RSS, throughput, scaling,
edit/save, real-producer, broad OOXML, or iWork behavior.

Artifacts:

- [machine-readable summary](../results/ooxml-root-lifecycle-0196-summary.json)
- [artifact manifest](../results/ooxml-root-lifecycle-0196-manifest.json)
- compressed A1/B1/B2/A2 raw reports listed in the manifest
