# Change 0368: ODT source-backed document catalog selectors

Date: 2026-09-02

Status: Accepted bounded diagnostic evidence; no production candidate

Performance claim: none

## Decision

The performance harness adds three opt-in selectors for the existing
source-backed ODT document catalog:

`odt_source_backed_catalog_open`, `odt_source_backed_catalog_list`, and
`odt_source_backed_catalog_query`.

The open selector times a fresh
`SourceBackedDocumentCatalog::from_read_at`. The list selector prepares the
source-backed owner outside the timed interval and times only `catalog()`. The
query selector prepares the owner and resolves the selected index outside the
timed interval, then times only `block_at(5000)`. Semantic ordering,
catalog/selected-block digests, archive topology, and media identity checks
remain outside the timers.

## Corpus and protocol

The selectors reuse the fixed large ODT corpus with 10,008 entries, 13 ZIP
members, and eight deterministic 2 MiB `Pictures/*` members. The archive is
16,811,815 bytes with SHA-256
`d63726138d0a50c8ff7e150af4a86385df1a34d886bb5f61f985c78ac79b0220`.
The control was collected on CPU 2 with 30 warmups and 500 retained samples
per selector.

Every sample also performs an untimed instrumented-source replay. The replay
records logical reads and bytes, range overlap, `content.xml` overlap,
untouched and Pictures reads, payload reads, and source-version observations.
Open includes the required `mimetype` and manifest reads. After preparation is
reset, list requires zero source, content, untouched, ordinary-payload, and
Pictures reads; query permits content reads only and requires zero untouched,
ordinary-payload, and Pictures reads. No phase reads a Pictures payload.

## Descriptive control result

The retained [machine-readable control report](../results/odt-source-catalog-0368-control.json)
was produced by binary SHA-256
`cc6f5a148f0788210814254f521c681238cf77ce9eba29ff3b29b5486d6c6ae8`, from
revision `14884ced9d8b29b7d2155134025986e9315ac771`, with `dirty: true`.
The observed elapsed-time statistics are descriptive only:

| Selector | Min ns | P50 ns | Mean ns | P95 ns | P99 ns | Max ns |
| --- | ---: | ---: | ---: | ---: | ---: | ---: |
| `odt_source_backed_catalog_open` | 1,949,506 | 2,032,560 | 2,060,901.762 | 2,228,040 | 2,436,085 | 3,162,246 |
| `odt_source_backed_catalog_list` | 20 | 125 | 130.77 | 291 | 410 | 611 |
| `odt_source_backed_catalog_query` | 1,599,544 | 1,697,603 | 1,718,240.906 | 1,870,487 | 2,016,478 | 2,431,429 |

This is a dirty current-control report, not a clean A/B comparison. It does
not authorize a speedup or establish a production baseline.

## Validation and claim boundary

The focused catalog oracle passed `1/1`, and the selectable-case-count test
passed `1/1` after the count assertion was corrected. The selector smoke also
passed; its JSON is intentionally not retained. Scoped Clippy passed with
explicit allowances for unrelated existing lint classes.

The initial standalone harness suite recorded `233 passed, 7 failed, 1
ignored`. The count assertion was then fixed and its focused test re-passed;
the remaining failures are unrelated to Change 0368:

- `docx_source_edit_is_deterministic_and_emits_complete_evidence`
- `media_rich_odp_scalar_and_batch_text_box_replacements_are_matched`
- `media_rich_odt_scalar_and_batch_resource_replacements_are_matched`
- `xlsx_cell_values_matched_controls_are_deterministic_and_bounded`
- `xlsx_row_visibility_matched_controls_cover_single_and_bounded_batch`
- `xlsx_vendor_extension_cell_crud_shape_is_opt_in_and_preserved`

The selectable registry is `404`; the default matrix remains `36 cases / 198
rows`. This change is harness-only evidence coverage. `performance_claim` is
`none`: no latency, throughput, physical-I/O, allocation, decompression,
cold-cache, fixed-memory, RSS, OOM-prevention, or broad ODF claim follows.

