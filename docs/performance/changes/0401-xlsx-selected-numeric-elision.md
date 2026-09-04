# Change 0401: XLSX selected numeric ownership elision

Date: 2026-09-04

Status: accepted scoped production optimization. The retained latency claim is
limited to the listed mean, p95, and p99 statistics for the fixed filesystem
XLSX selected-cell operation and protocol below. The p50 result is adverse in
both paired directions and is explicitly rejected; no median speedup claim is
made.

`performance_claim: scoped`

`claim_authorized: true`

## Production change and boundary

The control is production commit
`0859063be5a67bd2aafb3531f2126020b2b5000d`. Production commit
`87f26d5ee02a1903e668bf7f60fa3ef954a0c3fb` adds a private borrowed lexical
validator and uses it to elide ownership for unselected values. The borrowed
`Number::validate_lexical(&str)` check shares the validation behavior of the
owned `Number::new` constructor, including rejection of invalid and
non-finite worksheet numbers, without taking a `Box<str>` for a value that the
query will not return.

The selected worksheet scanner applies the elision only when a cell is
unselected, non-formula, and non-inline, and its kind is numeric or untyped.
It still validates a nonempty numeric lexical form before discarding it, then
recycles the existing scanner scratch state. Selected numbers retain their
exact lexical spelling and ownership. Formula cells retain their cached
numeric value and formula ownership, and inline strings—including unselected
inline strings—continue through their normal decode and validation path.
Other cell kinds and the existing eager fallback are unchanged. No public
type, API, dependency, executor, or format-neutral behavior is introduced.

## Corpus and fixed selected-cell oracle

The opt-in `xlsx_file_selected_cell` selector remains the sole timed case. It
uses the pinned medium
`litchi-xlsx-cell-values-source-edit-media-multi-sheet-v1` corpus: four
48-by-48 worksheets (9,216 numeric cells), 17 ZIP members, 4,226,429 archive
bytes, and source archive SHA-256
`dfff7ec0c749d9e404091776f15a8fb690985af7f58efdfe659dbeaed7145036`.
The selectable registry remains **420** and the default matrix remains
**36 cases / 198 rows**.

Each fresh child prepares `litchi::Workbook::open(path)` and the mixed-case
selector `bEnCh01` before timing. It selects canonical `Bench01` (zero-based
worksheet position `1`) and reads `M29` (zero-based row `28`, column `12`).
The typed oracle is a stored Number with exact lexical value `1028012`, and
the selected-cell evidence digest is
`36e53d9002ae8c433ad918b400196fb886fa675f850076808ac51327d1f42ac1`.
The selected `Bench01` sheet has 2,304 numeric cells: one selected cell and
2,303 unselected numeric cells in this fixed query shape.

## Normal ABBA evidence

The normal release run used stable Rust/Cargo/Rustdoc 1.98.1, CPU affinity 2,
one execution worker, warm filesystem state, and a fresh child per sample.
Each leg used 20 warmups and 500 retained samples in strict
`A1(control) / B1(candidate) / B2(candidate) / A2(control)` order. The timer
contains only case-insensitive worksheet selection followed by the exact
`M29` read; workbook open and query preparation are outside the timer.

Positive values mean the candidate is faster. The accepted statistics are
listed first; p50 is retained as an adverse result and rejected:

| Statistic | A1 control → B1 candidate | A2 control → B2 candidate | Verdict |
| --- | ---: | ---: | --- |
| mean | `+0.099577940251%` | `+0.026562239637%` | accepted |
| p95 | `+0.625379111895%` | `+0.198122423529%` | accepted |
| p99 | `+1.170167332729%` | `+0.045344544337%` | accepted |
| p50 | `-0.012690677428%` | `-0.035254218167%` | rejected; adverse in both directions |

The accepted result is therefore a narrow mean/tail observation for this
selector. The adverse p50 values are not converted into a median claim or
rounded away.

## Allocator evidence

A separate warm allocator ABBA used the same CPU-2, one-worker, fresh-child
shape with three warmups and 30 retained samples per leg, in the same
`A1 / B1 / B2 / A2` order. Each implementation produced a constant vector on
both of its legs:

| Metric | Control | Candidate | Candidate − control |
| --- | ---: | ---: | ---: |
| Allocation calls | `84,221` | `81,918` | `-2,303` |
| Deallocation calls | `84,206` | `81,903` | `-2,303` |
| Reallocation calls | `12` | `12` | `0` |
| Failed allocation calls | `0` | `0` | `0` |
| Allocated bytes | `10,706,565` | `10,690,444` | `-16,121` |
| Deallocated bytes | `10,705,182` | `10,689,061` | `-16,121` |

The exact call/byte deltas correspond to the 2,303 unselected numeric cells
in this fixed `Bench01!M29` oracle. The allocator observation accounts for a
7-byte lexical allocation per such cell on this Rust 1.98.1 run; that fixed
7-byte lexeme size is oracle-specific and is not a general cell-size or
memory claim. Allocator elapsed time is observational only. The live and
peak values, including process snapshots, are not claim metrics.

## Correctness and reproducibility

The production commit adds exact regression coverage for the ownership and
validation boundary:

- `borrowed_number_validation_matches_owned_constructor_errors`
- `streaming_0401_numeric_elision_preserves_selected_lexemes_and_unselected_validation`
- `streaming_0401_numeric_elision_keeps_selected_and_unselected_errors_exact`
- `streaming_0401_numeric_elision_keeps_formula_cache_ownership_and_parity`
- `streaming_0401_numeric_elision_does_not_skip_unselected_inline_validation`

Together these tests cover valid signed-zero, whitespace, and exponent
spellings; malformed and non-finite numeric errors; selected lexical
ownership; formula cached-number parity; and malformed, oversized, and
surrogate-bearing unselected inline strings. The existing XLSX selected
scanner, eager fallback, source/facade, and full-library validation remain
the correctness gates for the surrounding path. The expected raw normal and
allocator evidence is linked from the
[Change 0401 evidence bundle](../results/change-0401/).

## Authorized claim and exclusions

`performance_claim: scoped`; `claim_authorized: true`. The authorized claim
is limited to the accepted mean, p95, and p99 normal release-binary elapsed
statistics for case-insensitive selection of `bEnCh01` followed by the exact
stored-number read at `M29`, over the fixed medium filesystem XLSX corpus and
the stated CPU-2, one-worker, warm fresh-child ABBA protocol.

No median/p50 speedup claim follows. No claim follows for allocator elapsed
time; allocator live/peak values or RSS; physical I/O; cold or cold-cache
behavior; throughput; other cell types, queries, ranges, or corpora; or
general XLSX behavior.
