# Change 0120: PPTX ordinary-root filesystem source-path evidence

Date: 2026-08-15
Status: correctness and logical-read evidence; no release performance claim

## Scope

The standalone performance harness adds eight opt-in filesystem selectors over
the fixed media-rich PPTX corpus used by the existing source-backed PPTX
publication controls:

| eager control | source-path candidate |
|---|---|
| `pptx_file_eager_open` | `pptx_file_source_open` |
| `pptx_file_eager_list_slides` | `pptx_file_source_list_slides` |
| `pptx_file_eager_slide_count` | `pptx_file_source_slide_count` |
| `pptx_file_eager_selected_slide` | `pptx_file_source_selected_slide` |

The corpus has 200 slides, eight deterministic text boxes per slide, and eight
incompressible 2 MiB PNG media Parts. The default matrix remains 36 cases and
198 records; the eight selectors are opt-in and bring the selectable matrix to
227 names.

## Timing contract

Every filesystem sample runs in a fresh child process with the existing warm
and cold-requested cache states and process metrics. The source-path candidate
uses the public `litchi::Presentation::open(path)` root path. The eager open
control times `fs::read` plus `Presentation::from_bytes`. Query controls prepare
their eager byte root or source-path root before timing and time only the named
query:

- `slide_count` counts only; it does not enumerate slide payloads.
- `list_slides` calls the public `Presentation::slides()` and materializes the
  complete owned `Vec<Slide>`; it never substitutes a lazy iterator.
- `selected_slide` calls `Presentation::slide(100)` and does not use
  `slides().nth(...)`.

Full corpus hash, archive length, slide count, slide text/name hashes, metadata,
slide size, and eager/source parity are checked outside the timed interval.
The eager controls have no `ReadAt` implementation or source replay; their
structured sample scope is explicitly `not_applicable_eager_pptx`. Existing
generic filesystem counters remain present for schema compatibility and are
not interpreted as eager PPTX source counters.

## Untimed source replay

Each measured source-path sample performs one separate untimed replay through
`litchi_pptx::SourceBackedPresentation` over an instrumented positional source.
The replay records total requests and exact overlap with compressed ZIP payload
ranges for slides and media:

- open and slide-count must have zero slide/media payload overlap;
- selected slide must overlap only target slide 100, with no unselected slide or
  media overlap;
- list-slides must overlap all slide payloads and no media payload.

The report stores the replay source hash, request-size distribution, classified
payload counters, union-covered bytes, full-range coverage counts, semantic
hash, and classification string. Selected-slide classification requires the
complete target compressed range; list classification requires every one of the
200 slide ranges. These observations are logical range evidence for this
generated corpus. They are not claims about
physical disk I/O, decompression volume, allocation, RSS, latency, throughput,
or cold-cache behavior.

The final integrated one-sample debug replay is retained as a compact,
content-bound [correctness summary](../results/pptx-root-source-path-smoke-0120-summary.json).
It intentionally omits elapsed samples because this change makes no performance
claim; the record preserves the environment, corpus identity, exact build-input
manifest hash, and the source/eager classification gates.

## Verification and limits

The child verifies the source hash and complete semantic parity after timing;
the parent also guards the source hash between every child invocation. A source
classification failure is fatal rather than being recorded as a success. The
case names, operation scopes, overlap arithmetic, and zero/nonzero replay
accounting have focused unit coverage.

No latency, tail-latency, allocation, peak-memory/RSS, CPU-counter,
decompression, physical-I/O, or release-ABBA claim is made by this change.
Those claims require a frozen production tree, release builds, matched eager
and source controls, CPU/storage protocol, retained raw samples, and the
existing correctness gates.
