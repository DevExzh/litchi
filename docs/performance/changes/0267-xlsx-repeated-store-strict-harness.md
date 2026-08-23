# Change 0267: XLSX repeated-store strict schema and harness

## Status

Landed in `ebd0c83ea4ea49c05f693d6af5c93296cf812ca9` and
`1b3b5b094b8146d83df1c49a852efe4afdfeeaba`. This change adds neutral,
opt-in XLSX repeated-store evidence selectors and the strict schema needed to
validate their reports. It does not establish a latency result or change a
production library path.

## Selectors and pinned corpora

The four opt-in selectors are:

- `xlsx_source_repeated_store_medium`
- `xlsx_source_repeated_store_oversized`
- `xlsx_source_repeated_store_medium_reacquisition_control`
- `xlsx_source_repeated_store_oversized_reacquisition_control`

They add four `Case` entries, taking the current selectable matrix to **389**
names while leaving the default **36 cases / 198 records** unchanged. Both
scenarios use the pinned generator
`litchi-xlsx-source-repeated-store-corpus-v1`, the selected member
`xl/worksheets/sheet1.xml`, four 48-by-48 worksheets, and 9,216 scalar entries.
The medium archive is 4,226,429 bytes with source SHA-256
`dfff7ec0c749d9e404091776f15a8fb690985af7f58efdfe659dbeaed7145036`; its
selected worksheet is 63,294 uncompressed bytes. The oversized archive is
4,236,114 bytes with source SHA-256
`3cf797e44ef51189a4b62d040cf39ff2af670ebd909c6e806f387b51e72ecfec`; its
selected worksheet is 8,389,041 uncompressed bytes.

Each timed operation runs the four semantic queries `cell`, `cells`, `visit`,
and `stored_extent` eight times. The cache limits are 8 MiB with two entries
for the medium scenario and 8 MiB with 128 entries for the oversized scenario.
Every report records the exact timing scope
`semantic_query_only; explicit PartData reacquisition excluded`, managed
Budget diagnostics, source/cache counters, and the typed semantic projection.
Samples are warm-only and run in fresh child processes.

The two primary selectors use the source-backed cached store and are a possible
future ABBA comparison path only when the same selector is compared across
revisions. The two `reacquisition_control` selectors explicitly reacquire
`PartData` on every query: the medium control proves cache eviction, while the
oversized control proves the bypass path. Their claim scope is structural
cache/read control only; their elapsed/query vectors are excluded from the
candidate latency comparison and cannot masquerade as primary evidence.

## Strict evidence boundary

`tools/perf_abba_summary.py` now validates the exact case-local corpus
manifest, source and full semantic identities, distinct repeated-query
projection identity, query order/count, timing arithmetic, cache/read/Budget
counter arithmetic, warm/fresh-child configuration, and globally unique
positive child-process IDs. It rejects selector renames, arbitrary or
partially rewritten corpora, mixed primary/structural scopes, forged result
channels, and schema mutations. Allocator samples are accepted only in the
strict measured/unavailable/overflow forms owned by the existing allocator
vocabulary. Structural controls are validated but omitted from the primary
elapsed/source/sink summary.

The focused strict-summary suite covers the repeated-store contract. The
selectors and verifier are correctness and evidence-boundary infrastructure;
this change registers no latency, allocation, RSS, physical-I/O, throughput,
decompression, producer, or production-performance claim.
