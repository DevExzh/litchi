# 0302 — Residual single-op and batch performance misses

- Status: implemented on `uo2/perf-residual`, cut from `feat/unified-ops-api`
  at `928b19b4db`, trunk `cfef34af91` merged (`7f41bba583`).
- Date: 2026-10-16
- Inputs: the open misses against `docs/UNIFIED_OPS_GOAL.md`'s PERFORMANCE BAR
  listed in records 0286 F1 and 0291 (DOCX single `set`), 0295 F7 (ODT find),
  0297 (Numbers open + get and 1k batch), 0299 (Pages 10k batch) and 0291
  (RTF fan-out against the precompiled `Plan`).
- Scope: `litchi-core` selector parsing; `litchi-ops` `target`, `find`,
  `batch/execute` (field admission only), `format/docx`, `format/odt`, and
  separate small hunks in `format/iwork{,/pages,/numbers/coalesce}`,
  `write.rs`, `transaction.rs` and `Cargo.toml`.
- `performance_claim`: paired measurements in §3. `profiling` builds (release
  with debug info, stripped before timing) pinned with `taskset`, processes in
  B A A B order, each harness an in-process ABBA. Instruction counts are
  callgrind (`perf` is blocked, `perf_event_paranoid=4`). Evidence:
  [results/0302/](results/0302/), driver `results/0302/measure.sh`, summary
  `results/0302/summary.txt`. Host: shared 32-core EPYC, load 4–12.

## 1. Summary

| Item | Before | After | Status |
|---|---:|---:|---|
| DOCX single paragraph `set` (`docx_fanout_bench`, 1 ¶) | 1.110–1.125 | 1.105–1.111 | **improved, still a miss** (§4.1) |
| DOCX `/body > p`, 100 ¶ | 1.057–1.065 | 1.059–1.062 | unchanged |
| ODT find, 1 match in 10,000 ¶ | 1.104–1.108 | 1.016–1.017 | fixed |
| ODT find, every ¶ of 10,000 | 1.029–1.032 | 0.991–1.006 | improved |
| ODT merge, 10,000 ¶ | 1.069–1.071 | 1.066–1.089 | unchanged within noise |
| Pages batch of 1 on the 10,000-¶ document | 1.629 | 1.017–1.019 | fixed |
| Pages 1k batch | 1.728–1.736 | 1.098–1.101 | fixed (budget 2×) |
| Pages 10k batch | 2.418–2.474 | 1.608–1.615 | improved (10k only has to scale linearly) |
| Numbers 1k batch | 1.443–1.488 | 1.320–1.336 | improved (budget 2×, met before and after) |
| Numbers open + first `get` | 1.249–1.259 | 1.236–1.256 | **unchanged, miss** (§4.3) |
| RTF fan-out vs `Plan` route | 3–9 points | 3–9 points | **unchanged** (§4.4) |

Every output stays byte-identical: each harness checks both legs' bytes (or
texts, for DOCX) before timing; the query-corpus differential, the corpus
fidelity test, the checklist goldens, `path_allocations`, the capability and
registry goldens (both feature sets) and the OfficeCLI corpus pass unchanged,
nothing re-blessed.

## 2. Changes

### 2.1 `litchi-core`: plain paths in one pass, smaller steps

- `Path::parse` first tries `parse_plain`: `/` and plain segment names, each
  with at most one `[N]`. It fills the path in place (a parsed path is several
  hundred bytes; returning it through an `Option` copied it once more). Any
  other input — quotes, other predicates, whitespace, an index of zero or one
  that overflows, a step or byte limit, a malformed path — goes to the
  general parser unchanged, which builds the path or reports the error. Test
  `plain_fast_path_matches_the_general_parser`: every string of up to five
  bytes over `/aB[]10 "'\t*` (the fast path's result must equal the general
  parser's, `Display` and predicates included), real paths, and refusals.
- `OneOrMany` (the predicate list of a step, a query's alternatives) is an
  enum `Zero | One(T) | Many(Vec<T>)` instead of an `Option` beside a `Vec`:
  `Step` 120 → 96 bytes, `Path` 760 → 616, `Query` 248 → 176. No public type
  or method changed.

### 2.2 `litchi-ops` shared code

- `has_query_syntax` returns `false` at once when the text holds none of the
  bytes that can make it true (`"'`, `,>*:()` and whitespace). The exhaustive
  classifier test of record 0291 covers it.
- Batch field admission (`batch::execute::fields`) tests each field in wire
  order and names one only when it is present and refused; it built a
  19-entry name table per item (0.5 K instructions). Same errors, same order.
- `find::Matcher` keeps a `memchr::memmem::Finder` for a case-sensitive
  pattern: `str::find` built its searcher on every call. A byte match of a
  UTF-8 pattern in UTF-8 text lies on character boundaries, so the offsets
  are those `str::find` reports (test
  `prebuilt_search_reports_the_offsets_of_str_find`). `memchr`, already a
  `markdown` dependency, is now a plain dependency of `litchi-ops` (no new
  crate in any lockfile).

### 2.3 ODT find and merge prefilter (0295 F7)

`format/odt/write/find::surface_of` decodes each paragraph's run texts once
into a reusable list and lists the paragraph's runs only when its joined text
passes a filter: `Matcher::may_match` for find (exact when it says no; every
match lies inside one group, so inside its paragraph's joined text), and
`contains('{')` for merge (`surface::merge` already skipped groups without a
brace). Rejected paragraphs skip the per-run scope, rewritability and id
bookkeeping. Text decode errors still surface for every paragraph.

### 2.4 DOCX adapter (0286 F1/F2, 0291 §5)

- `format/docx/few.rs`: `Few<T>` keeps one element inline, more in a `Vec`.
  The resolved nodes, admitted targets and admitted assignments of a `set` use
  it: three allocations fewer for the one-target, one-property call.
- `Docx::set`/`set_command` borrow their `TargetArg` (one fewer 656-byte
  copy); `list_request` edits the assignment list in place.
- `DocxOwner::Source` and the cached metadata are boxed: `DocxState` was
  1,840 bytes, copied twice and allocated for every committed handle.

### 2.5 Pages coalesced paragraph sets (0299)

- `pages::plan` maps a direct `/body/p[N]` (exactly the step names `body`
  and `p`, no predicate on `body`, one positional index) to `(section, line)`
  from each section's paragraph start, cached per handle
  (`Iwork::paragraph_starts`, computed once by counting each section text's
  `\n`). The object tree is not built. Any other spelling, or a document with
  fewer paragraphs, takes the tree route as before. Test
  `direct_paragraphs_match_the_tree` compares both on every Pages fixture for
  every N up to past the end.
- `Pending` stages `(line, text)` by appending; `flush` sorts each section's
  list stably and keeps the last text of a repeated line, which is what the
  per-line map held.

### 2.6 Numbers coalesced cells (0297)

The run's planned-cell set was a `BTreeSet<Address>` probed and filled per
item; it is a `HashSet` of packed coordinates with a multiply-rotate hasher
(`CellHasher`, rotated on finish so buckets see every input bit).

## 3. Measurements

Medians per process run, B A A B. Binaries: `results/0302/binaries.sha256`
(base `928b19b4db`; after `7f41bba583`, `m4_formats_bench` from
`b046c2a64c`). The after binaries include trunk `cfef34af91`, which changed no
leaf on these routes (`litchi-odt`, `-docx`, `-rtf`, `-numbers` untouched;
`litchi-pages` only its body-table merge transaction).

### 3.1 DOCX single paragraph `set`

| run | leaf µs | unified µs | ratio |
|---|---:|---:|---:|
| base 1 / 2 | 7.55 / 7.72 | 8.50 / 8.57 | 1.125 / 1.110 |
| after 1 / 2 | 7.45 / 7.58 | 8.28 / 8.37 | 1.111 / 1.105 |

Callgrind, per call ((Ir@3000 − Ir@1000)/2000): unified − leaf 7.78 K → 6.64 K
instructions in the full-feature build (5.86 K in a `docx`-only build). Of the
removed 1.1–1.9 K: 3 allocations and frees, 0.8 K of struct copies
(`memcpy` from the handle state, `TargetArg` and `list_request` moves),
0.25 K in `Path::parse`, 0.17 K in `has_query_syntax`.

### 3.2 ODT find and merge, 10,000 paragraphs (`m4_formats_bench`)

| scenario | base 1 / 2 | after 1 / 2 |
|---|---:|---:|
| find, 1 match | 1.104 / 1.108 | 1.016 / 1.017 |
| find, every paragraph | 1.032 / 1.029 | 1.006 / 0.991 |
| merge | 1.069 / 1.071 | 1.066 / 1.089 |

A first version without the prebuilt searcher fixed the 1-match case (1.08)
but slowed the every-paragraph case to 1.075: matching paragraphs were
searched twice with a fresh `str::find` searcher each time.

### 3.3 Pages batches on the 10,000-paragraph document (`pages_bench`)

| scenario | base 1 / 2 | after 1 / 2 |
|---|---:|---:|
| batch 1 | 1.629 / 1.629 | 1.017 / 1.019 |
| batch 1,000 | 1.736 / 1.728 | 1.101 / 1.098 |
| batch 10,000 | 2.474 / 2.418 | 1.608 / 1.615 |

Callgrind at base, per 10k batch: tree build 32 M and its drop 7.5 M
instructions, resolution 12.7 M, parsing 15.6 M, staging 5.4 M, over a leaf
leg of about 97 M. After §2.5 neither the tree nor its drop runs.

### 3.4 Numbers (`numbers_unified_bench`, 7 ABBA rounds)

| scenario | base 1 / 2 | after 1 / 2 |
|---|---:|---:|
| open + `get` | 1.259 / 1.249 | 1.236 / 1.256 |
| `set` | 1.002 / 1.009 | 1.014 / 1.026 |
| batch 1,000 | 1.443 / 1.488 | 1.320 / 1.336 |

Callgrind (`--toggle-collect` on the legs): unified − leaf per item 7.2 K →
4.6 K instructions (path text 2.2 K → 1.25 K, planned-cell set 1.5 K →
0.2 K, field admission 0.5 K → 0.1 K).

### 3.5 RTF (`rtf_destination_bench`)

Fan-out / `Plan` route / edit-closure batch ratios are unchanged within
noise (`summary.txt`; e.g. table-cell text ×1: 1.140–1.148 / 1.108–1.115
before, 1.155–1.156 / 1.114–1.117 after).

## 4. Misses left, with measured explanations

### 4.1 DOCX single paragraph `set`: 1.105–1.111

Instructions are 5.3 % over the leaf (6.64 K of 124.6 K), the time 10.5–11 %.
A callgrind cache simulation (`--cache-sim=yes`, the host's 32 KiB L1s) shows
why: per call the unified leg has 3.25 K L1 instruction misses against the
leaf's 2.63 K, and 135 D1 read plus 208 D1 write misses against 20 + 26. The
leaf's working set nearly fills L1; the adapter, resolver and registry code
and the larger frames push it out, and the misses land in the leaf's own
commit (`owned::Edit::commit` copies its package into the boxed editor's and
the committed handle's heap state). The remaining instruction excess: shared
resolution 1.2 K (two `Nav::child` steps 0.6 K), `Path::parse` +
classification 0.9 K, registry admission 0.6 K, handle adoption 0.6 K (two
`Arc`s, kept: the state is shared across signature dispositions), and frame
moves. Follow-up: keep the leaf editor and committed package unboxed on the
single-op path, which needs an inline `DocxEditor` variant and is a wider
adapter change than this record.

### 4.2 Pages 10k batch: 1.61

0.18 µs per coalesced set over the leaf's `Edit::set`: parsing each item's
path text (0.95 K Ir), the owned text of each set (0.5 K) and the transaction
frame. The batch scales linearly; the 2× target is set at 1k (met: 1.10).

### 4.3 Numbers open + first `get`: 1.24–1.26

The first read builds the workbook's object tree (0.24 M Ir) and dropping the
handle drops it (0.17 M). A `get` that skips the tree needs a lazily built
per-table tree in the iWork tree module, which the iWork structure record
(0301) is changing; follow-up there.

### 4.4 RTF fan-out vs `Plan`

A plan-cache hit costs about 1.0 K Ir (query hash 0.25 K, full comparison
0.2 K, thread-local lookup, `Arc`). The measured gap to the `Plan` route
(about 0.2 µs) is mostly the harness: the fan-out leg passes
`query.clone()` of an owned `Query<'static>`, 2.0 K Ir and two allocations
per call, which the `Plan` leg does not pay. Hashing less would save at most
0.15 K. No change made; follow-up: a harness route that passes `&Query`
under a new name.

## 5. Gates

Build environment per the rules (`CARGO_BUILD_JOBS=6`, `RUST_TEST_THREADS=8`,
one cargo command at a time, target pruned between suites). Every exit status
was checked; all ran on the merged tree (`b046c2a64c`).

| Command | Exit |
|---|---:|
| `cargo fmt --all --check` | 0 |
| `cargo clippy -p litchi-core -p litchi-ops -p litchi -p litchi-py --all-targets --all-features -- -D warnings` | 0 |
| `cargo clippy -p litchi-ops --no-default-features [--features F] --all-targets -- -D warnings`, F = none, docx, odt, ods, odp, oth, odm, pages, numbers, keynote, rtf, doc, ppt, xls, xlsx, xlsb, pptx, markdown | 0 each |
| `cargo test -p litchi-core --all-features` (315 passed) | 0 |
| `cargo test -p litchi-ops --all-features` (2,298 passed; includes the query-corpus differential, corpus fidelity, checklist goldens, `path_allocations`) | 0 |
| `cargo test -p litchi --all-features` (260 passed) | 0 |
| `cargo test -p litchi-ops --all-features --test capabilities --test registry` (not re-blessed) | 0 |
| the same with `--no-default-features --features` every litchi-ops feature but `encryption` (not re-blessed) | 0 |
| `RUSTDOCFLAGS=-Dwarnings cargo doc --no-deps --all-features -p litchi-core -p litchi-ops` | 0 |
| `python3 tools/check_crate_boundaries.py` | 0 |
| `python3 tools/non_iwork_gate.py verify` | 0 |
| `RUSTC_BOOTSTRAP=1 python3 tools/check_api_glossary.py --target-dir …` (37 profiles, 1,800 bans, 133 families) | 0 |
| `crates/pyo3-litchi/tests/run_tests.sh` (275 tests) | 0 |
| `cd tools/perf-baseline && cargo check --locked --all-targets` (after `cargo update --offline -w`: its lockfile lists `memchr` under `litchi-ops`) | 0 |
| `cargo test -p litchi-cli --all-features --test officecli_corpus` (not re-blessed) | 0 |

No capability, registry or OfficeCLI-visible behaviour changed, so the delta
table and the OfficeCLI corpus were not regenerated. No public name was added
or changed; no glossary row changes.

## 6. Not claimed

- That the DOCX single `set` meets the 10 % budget.
- Any change of a frozen perf-baseline route; only `m4_formats_bench` gained
  a filter so `M4_ONLY=odt_find_1` runs one scenario.
