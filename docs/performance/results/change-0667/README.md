# Evidence: change 0667, XLSX value-editor shared-string admission

Change record: [`0667-xlsx-value-editor-shared-strings`](../../0667-xlsx-value-editor-shared-strings.md).

Disposition: retained and implemented in `litchi-xlsx` and the producer-shape
harness. `performance_claim: none`. This packet closes 0602 D2 after 0657's
D4 admission, with a lazy shared-string table, preserved indexes, a bounded
retained part, and a typed mutation refusal.

## Contents

| Path | Purpose |
| --- | --- |
| `bench/producer.json` | Release producer-shape timing, 12 cases, three warmups and 15 samples per case. Descriptive after-only observations. |
| `bench/producer-evidence.json` | Producer corpus facts and verdict census. Medium and dense read variants have 40% shared cells and 64 unique entries; the shared-string fact is admitted. |
| `bench/large-table-run.txt` | Re-run of the 66,935-entry admission witness after compilation, including elapsed test time and process RSS. |
| `decision.json` | Change decision and evidence disposition. |
| `cleanup.json` | Retained artifacts and scratch cleanup. |
| `gates.txt` | Final formatting, lint, tests, docs, and harness gate tails. |
| `log-sections.md` | Four coordinator-ready sections; no shared rollup files are changed here. |

## Provenance

Base: `5fa92d7ced5a78f3c2c84a6afe9d0ba404253717` (0657). Branch:
`perf/0667-xlsx-value-editor-shared-strings`. The release harness binary was
built with `CARGO_BUILD_JOBS=2`, `CARGO_TARGET_DIR=/home/zhuhe/code/litchi-target-0667`,
`cargo build --release --locked --manifest-path tools/perf-baseline/Cargo.toml`.
Its SHA-256 is `444d4ff970cf5c59cf656f097e8035daa7b81e9f421656dcb2f2766985540ad2`.

The local format citation used by the record is [ECMA-376 Part 1, 5th edition,
§12.3.15 and §§18.3.1.96, 18.4, 18.4.9, 18.18.11](../../../../3rdparty/specs/ECMA-376/ECMA-376-1_5th_edition_december_2016.zip).
It specifies one internal shared-string part, workbook-wide indexing, the
`count`/`uniqueCount` meanings, and `t="s"` cell indexing.

The producer timing run is intentionally not a registered performance result:
it has no base binary or A/A pair. Its p95/p50 floor is applied per case for
descriptive stability; the dense source-open case is excluded from that floor
because it measures 1.153. The other eleven cases are at or below 1.05.

## Correctness scope

The tests cover lazy materialization, valid text readback, index bounds,
truncated tables, duplicate relationships, exact source-part preservation,
shared-target mutation refusal, and a 66,935-entry table. The full
`litchi-xlsx` library suite and the source-backed integration target are listed
in `gates.txt`. No public API or OPC crate code changed.
