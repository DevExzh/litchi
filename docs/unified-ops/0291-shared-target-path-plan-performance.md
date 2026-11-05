# 0291 — Shared target, path and plan performance

- Status: implemented on `uo2/perf-path`, cut from `feat/unified-ops-api`
  at `99cc3138eb`.
- Date: 2026-10-15
- Inputs:
  - record [0286](0286-ooxml-adapter-performance.md) F1 (DOCX single
    paragraph `set` at 1.12–1.13×, about 4.6 K instructions of it in the
    shared target and path code);
  - record [0287](0287-remaining-query-performance.md) gaps: RTF selector
    compile per call, the generic walker step, `query.clone()` in mutation
    resolvers, nested XLS source checks;
  - record [0288](0288-odf-follow-ups.md): ODB single `set` at 1.35× on a
    1,000-query database.
- Scope: `litchi-ops` `target`, `path/{arg,resolve,vocabulary}`,
  `query/{arg,cache,walk}`, `format/{xls,odf_tree}`, and one-line call-site
  changes in the mutation resolvers of `format/{doc,docx,odp,ppt,pptx,xls,
  xlsb,xlsx,odf_tree}`.
- `performance_claim`: paired measurements in §3. Release code (`profiling`
  profile: release with debug info) pinned with `taskset -c 31` on the
  shared 32-core host at load 7–15. Processes ran in B A A B order; each
  harness is itself an in-process ABBA. Instruction counts are callgrind
  (`perf` is unavailable, `perf_event_paranoid=4`). Evidence:
  [results/0291/](results/0291/), with the driver
  `results/0291/measure.sh`.

## 1. What changed

1. **Vocabulary lookups indexed at compile time.** Every path step looked
   up its name in the family's alias table, then the element rule of the
   kind it found. Each was a binary search with a case-folding comparison,
   about 400 and 300 instructions.
   - Each `Vocabulary` now carries an `AliasIndex`, built by a `const fn`
     when the static is defined. It holds first-byte buckets over the
     sorted aliases, each alias's rule position, and the root's rule
     position.
   - A name is found by a short scan of its bucket, and the same lookup
     gives the element rule (`Vocabulary::canonical_rule`). The resolver's
     `classify` uses that pair instead of a second search when the adapter
     names its kind canonically, which every adapter does; otherwise it
     looks the rule up as before.
   - `Vocabulary::rule` answers the root directly and other canonical names
     through the alias index. It falls back to the old search for anything
     else, which misses as before.
   - The test `indexed_lookups_match_the_sorted_table_searches` compares
     both lookups with the replaced binary searches. It covers every
     vocabulary, alias, element name and upper-case variant, plus near
     misses.
2. **Target classification in one pass.** `has_query_syntax` decided whether
   a `(`, `)` or whitespace byte was query syntax by first finding the last
   top-level `/` in a separate pass. It now keeps a flag that a later
   top-level slash clears. The test
   `one_pass_query_syntax_matches_the_two_pass_classifier` compares it with
   the old classifier on all 813,616 strings of up to five bytes over the
   bytes it treats specially, plus a set of real targets.
3. **`&TargetArg` implements `IntoTarget`.** A target parsed once with
   `TargetArg::parse` can be passed by reference to every later call. The
   path or selector is then borrowed: it is not parsed, classified or
   copied again. This is the "pre-parsed target" of 0286 F1. Callers that
   pass text still parse on every call. No public name was added: this is
   a trait impl on existing types.
4. **Per-thread plan cache** (`query/cache.rs`). A selector passed as text
   or as a parsed `Query` was compiled for the document's format on every
   call, about 3.5 K instructions for a short RTF selector.
   - **Key.** The parsed query, compared in full after a hash; the
     adapter's format and type; and `raw`. Compilation reads nothing else,
     so a cached plan is the plan a fresh compile builds.
   - **Admission.** A query is cached the second time its thread sees it
     among the last 64 distinct queries. A selector run once pays only for
     its hash, not for an owned copy of itself.
   - **Bounds.** At most 32 plans per thread, replaced oldest first.
     Failures are never cached.
   - **Sharing.** A hit shares the plan by `Arc` (new `PlanRef::Shared`),
     so walks stay `Send`, and nothing is shared between threads. On
     0287's objection to hidden mutable state: the cache is not observable
     except in time. A cached plan equals a fresh compile, which the tests
     check through the plans' `Debug` forms.
   - Tests: `a_repeated_query_is_compiled_once_and_shared` and
     `failures_are_not_cached_and_the_cache_is_bounded`.
5. **Kind-only walk visits.** When a query has one alternative that tests
   only the element kind and has no positional predicate (`cell`, `p`,
   `shape`), `Engine::visit` answers from the kind bit. It skips the
   per-alternative loop and the out-of-line shortcut call. The answer is
   the one the loop gives: that shortcut matches every node of its kinds,
   and with no position it yields.
6. **Borrowed selectors in mutation resolvers.** Twenty-eight
   `crate::query::run(nav, query.clone())` calls now pass `query`, a
   `&QueryArg`. Since record 0287 a borrowed `QueryArg` is an `IntoQuery`
   that compiles from the reference.
7. **ODB trees deferred after a commit.** `OdfTree::from_commit` built a
   database's object tree eagerly, because the catalog parse can fail. A
   committed package's catalog cannot fail it:
   - the ODB leaf's commit parses the new catalog with the same default
     limits before it publishes (`snapshot.database.catalog()?` in
     `Edit::commit_core`);
   - a metadata commit leaves `content.xml` as it was.

   The database tree is therefore built on the first read, as the drawing,
   chart, image and formula trees already were. The test
   `committed_database_trees_are_built_on_first_read` checks two things.
   After a `set`, the tree is not built. The tree built later equals the
   one a fresh open of the saved bytes builds. (`add` still builds the
   tree, to report the new node's path.)
8. **Nested XLS source checks share the outer one.**
   - `Xls::finish` now runs its read in a fence scope for the handle and
     skips its own check inside such a scope.
   - The XLS navigator now has a fence key (the handle's shared owner).
     Resolver and walker steps therefore cover the adapter's nested fenced
     reads too, as they already did for XLSX, XLSB and PPTX (record 0262).
   - The unified per-match check stays, as the owner decided on
     2026-10-14. So does the check of every public read.
   - The test `nested_fenced_reads_share_the_outer_check` counts source
     observations through a counting `ReadAt`. A sheet's `View::text`,
     which reads the name through a nested fenced reader, now observes the
     source once instead of twice. A `get` observes it once, and a query
     observes it at least once per match.
   - `shared_nested_checks_still_refuse_a_changed_source` appends to the
     file after the reads. A node read before the change, and a new `get`,
     then both report `SourceChanged`.

## 2. Breakdown before optimizing (callgrind, per call)

**DOCX single paragraph `set`.** The unified call took 134.5 K
instructions, the leaf 125.4 K: 9.1 K more.

| part | base | after |
|---|---:|---:|
| `<&str as IntoTarget>::into_target` | 1,674 | 1,424 |
| — `has_query_syntax` | 489 | 238 |
| — `PathArg::parse` (`litchi-core` `Path::parse`) | 984 | 984 |
| `path::resolve::get` | 3,316 | 2,147 |
| — `Vocabulary::canonical`, two steps | 790 | 204 |
| — `resolve::rule`, root and steps | 592 | 116 |
| — the adapter's `Nav::child`, two steps | 770 | 770 |
| `Docx::resolve_targets` | 3,428 | 2,259 |
| `Commit::into_file` (DOCX adapter: dropping the leaf patch) | 1,584 | 1,588 |

The shared target and path cost fell from 4.99 K to 3.57 K. The unified
call is now 132.7 K against the leaf's 125.0 K, 7.7 K more.

**RTF fan-out.** The selector compile cost about 3.5 K instructions per
call. A cache hit costs about 0.8 K: hashing the query, comparing it with
the entry, and the `Arc`. Over the whole `rtf_destination_bench` run,
selector compiling fell from 1.05 M to 0.19 M instructions. The remaining
0.08 M are first sightings and admissions.

**Walker.** On ODS warm cell queries the visit took about 70 instructions
per cell. The per-alternative loop was 25 of them, and the out-of-line
shortcut call 24.

## 3. Measurements

Binaries are not committed; `results/0291/binaries.sha256` lists them:

| set | built from |
|---|---|
| base | `99cc3138eb` |
| after | `b2187dd6c1`, this record's code |
| after-v1 | `odf_gaps_bench` only: the working tree before the final `canonical_rule` and `Box` changes, which do not touch the ODB route |

All three were built with:

```text
cargo build --profile profiling -p litchi-ops \
  --features docx,pptx,xlsx,xlsb,xls,doc,ppt,rtf,odt,ods,odp,odg,odc,odi,odb,oth,odm,odf-formula \
  --example <name>
```

### 3.1 DOCX single paragraph `set` (`docx_fanout_bench`)

| run | leaf µs | unified µs | ratio |
|---|---:|---:|---:|
| base 1 / 2 | 7.64 / 7.61 | 8.51 / 8.53 | 1.115 / 1.121 |
| after 1 / 2 | 7.42 / 7.42 | 8.24 / 8.29 | 1.111 / 1.118 |
| `/body > p`, 100 paragraphs, base → after | | | 1.056–1.070 → 1.056–1.059 |

- In instructions, the overhead over the leaf fell from 9.1 K (7.3%) to
  7.7 K (6.2%).
- In time, the unified leg is 0.27 µs faster (−3.1%). The leaf leg also
  reads 0.2 µs faster in the after binary, although its code did not
  change: code layout.
- **The ratio stays at 1.11–1.12: still a miss against the 10% budget.**
  The remainder is listed in §5.

### 3.2 RTF destination writes (`rtf_destination_bench`, eight runs each)

Medians of eight runs per binary; every sample is byte-identical to the
native oracle. The `Plan` route is unchanged.

| family | targets | fan-out base | fan-out after | `Plan` route | edit-closure batch base → after |
|---|---:|---:|---:|---:|---:|
| table_cell_text | 1 | 1.19 | 1.16 | 1.12 | 1.06–1.07 → 1.06–1.07 |
| footnote_text | 1 | 1.18 | 1.14 | 1.10 | 1.08–1.09 → 1.08–1.09 |
| comment_text | 1 | 1.17 | 1.13 | 1.09 | 1.05–1.06 → 1.05 |
| picture_removal | 1 | 1.33 | 1.28 | 1.17 | — |
| picture_removal | 64 | 1.36 | 1.34 | 1.33 | — |

- The fan-out improved by 3–5 points. It is now 3–4 points from the
  precompiled `Plan` route, and 11 points on picture removal.
- The edit-closure single set (direct paths) is unchanged at 1.05–1.09×.
  Its paths arrive as parsed `Path`s, so neither target change applies,
  and its resolution is a small part of it.
- **Unexplained:** `header_paragraph_text` batches read 1–2 points slower
  (100 targets: 1.24 → 1.25; 1,000: 1.21 → 1.23, eight runs each). Whole
  run instruction counts fell by 68 M, and this route makes no plan,
  query or vocabulary call that changed beyond the faster lookups. It is
  listed in §5 with 0287's other layout readings. The full table is
  `results/0291/rtf-ratios-8-runs.txt`.

### 3.3 ODB single `set` on a 1,000-query database (`odf_gaps_bench`)

| scenario | base (2 runs) | after (2 runs) |
|---|---:|---:|
| batch_1 | 1.351 / 1.345 | 1.001 / 0.999 |
| batch_10 | 1.087 / 1.079 | 0.999 / 1.005 |
| batch_100 | 1.019 / 0.999 | 0.997 / 1.000 |
| batch_1000 | 0.993 / 0.979 | 0.979 / 0.983 |

- Fixed: 2.40 ms per call against the leaf's 2.39 ms, down from 3.17 ms.
- The other ODF object scenarios are unchanged within ±1%, and every
  output is byte-identical to the leaf's.

### 3.4 Queries

| harness, route | base | after | change |
|---|---:|---:|---:|
| ODS warm `cell`, 1,000 × 10, µs | 369–391 | 339–341 | −10% |
| ODS cold `cell`, µs | 436–439 | 396–427 | −5% |
| ODS warm, callgrind Ir (2 samples) | 33.7 M | 27.1 M | −19.7% |
| ODT `query("p")`, 10,000 paragraphs, µs | 1,119–1,134 | 1,055–1,099 | −4% |
| ODT query iteration, callgrind Ir | 207.0 M | 175.8 M | −15.1% |
| XLS eager positional `get` + text, 1,000 cells, µs | 236–240 | 204–206 | −14% |
| XLS eager get, callgrind Ir | 1,155.6 M | 885.4 M | −23.4% |
| XLS eager compiled `cell` scan + text, callgrind Ir | 555.8 M | 525.9 M | −5.4% |
| XLS source-backed get, 1 / 1,000 cells, ratio | 1.43–1.44 / 1.66–1.68 | 1.39 / 1.59–1.60 | |

- The XLS `get` gains come from the vocabulary index, since workbook paths
  have two steps.
- The ratios against the borrowed leaf loops remain large:
  - ODS warm at 16–23×;
  - ODT `query_all` at 3.1–3.9×.

  The remaining cost is the adapter's child step and the moves of each
  match.
- The ODS and ODT write scenarios and opens are unchanged within noise.

### 3.5 Source checks per XLS read

The source-backed `get` plus `view().text()` route already made two
unified checks at base: one in the resolver and one in `View::text`.
Callgrind counts the same 742,202 `View::text` checks for 742,202 texts
in both builds. 0287's count of four predates that. A fenced read nested
in another fenced read does share the outer check now. A sheet's text
observes the source once instead of twice (§1.8).

## 4. Differential and parity evidence

The query-corpus differential (`tests/query_corpus_differential.rs`)
passes with `--all-features`, and its golden was not touched. The
following are unchanged and pass without re-blessing:

- the corpus fidelity test;
- the checklist goldens;
- `query_early_stop_parity`, `node_view_allocations` and
  `path_allocations`;
- the capability and registry goldens;
- the OfficeCLI corpus.

The new tests are listed with each change in §1.

## 5. Gaps and follow-ups

- **DOCX single paragraph `set` is still 1.11–1.12×.** It is 7.7 K
  instructions over the leaf:
  - **`Path::parse` in `litchi-core`, about 1.0 K.** A target passed as
    text is parsed every call. Callers that repeat a target can now parse
    it once and pass `&TargetArg`. A faster parser belongs to the selector
    grammar's owner.
  - **The DOCX adapter's `Nav::child`, about 0.8 K.**
  - **Handle bookkeeping, about 2 K.** The four allocations of a new
    immutable handle are made in the DOCX adapter: the `Arc` state and
    handle, the boxed editor and the boxed committed package. So is
    `Commit::into_file`'s 1.6 K, which drops the leaf patch. Neither is
    shared-layer code, and this record was told to keep out of
    `format/docx`. Follow-up for the OOXML owner, with 0286 F2.
  - **Registry and target admission and the remaining resolver frame,
    about 3 K.**
- **RTF fan-out** is 1.13–1.16× for single targets and 1.28× for picture
  removal. Most of the remaining gap to the `Plan` route is hashing and
  comparing the query on each hit. A cache keyed by selector text would
  skip the parse as well, but the measured route passes a parsed `Query`.
- **Walker.** The kind-only path removes the per-alternative cost. The
  per-match moves and the adapter's child steps remain: about 610 → 540
  instructions per ODS cell. XLSX was not re-measured here; its walk takes
  the same path.
- **Unexplained wall-clock reading:** RTF `header_paragraph_text` batches,
  +1–2 points (§3.2).
- **Unchanged here:** RTF `header_paragraph_text` batches (1.24×), ODT
  `batch_bold_1000` (1.32–1.38×), ODS `batch_10000` (1.31×).

## 6. Gates

The build environment is per the rules file: `CARGO_BUILD_JOBS=6`, one
cargo command at a time. Each exit status was checked.

| Command | Exit |
|---|---:|
| `cargo fmt --all --check` | 0 |
| `cargo clippy -p litchi-ops -p litchi -p litchi-py --all-targets --all-features -- -D warnings` | 0 |
| `cargo clippy -p litchi-ops --no-default-features [--features F] --all-targets -- -D warnings`, F = none, docx, xlsx, pptx, xlsb, xls, doc, ppt, rtf, odt, ods, odp, odg, odc, odi, odb, odf-formula, oth, odm | 0 each |
| `cargo test -p litchi-ops --all-features` (includes the corpus differential) | 0 |
| `cargo test -p litchi --all-features` | 0 |
| `cargo test -p litchi-ops --all-features --test capabilities --test registry` (not re-blessed) | 0 |
| the same with `--no-default-features --features doc,docx,ppt,pptx,xls,xlsx,xlsb,rtf,odt,ods,odp,oth,odm,odg,odc,odi,odb,odf-formula,odf-objects,pages,keynote,numbers,opc,markdown,formula,yaml,eval,test-support` (every feature but `encryption`) | 0 |
| `RUSTDOCFLAGS=-Dwarnings cargo doc --no-deps --all-features -p litchi-ops` | 0 |
| `python3 tools/check_crate_boundaries.py` | 0 |
| `python3 tools/non_iwork_gate.py verify` | 0 |
| `RUSTC_BOOTSTRAP=1 python3 tools/check_api_glossary.py --target-dir …` (33 profiles, 1,420 bans, 55 families) | 0 |
| `crates/pyo3-litchi/tests/run_tests.sh` (267 tests) | 0 |
| `cd tools/perf-baseline && cargo check --locked --all-targets` | 0 |
| `cargo test -p litchi-cli --all-features --test officecli_corpus` (not re-blessed) | 0 |

No capability, registry or OfficeCLI-visible behaviour changed. The delta
table and the OfficeCLI corpus were therefore not regenerated. No
dependency was added.
