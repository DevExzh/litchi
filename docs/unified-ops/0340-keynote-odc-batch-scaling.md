# 0340 — Keynote paragraph-set batches and ODC series batches scale linearly

- Status: implemented on `uo2/keynote-batch`, cut from `feat/unified-ops-api`
  at `2083926091` (trunk merged four times: `c31198a5af`, `4282e69000` (record 0338), `ed832736c3` (record 0336), `7ea8acee2a` (record 0341)).
- Date: 2026-10-17
- Input: record [0337](0337-perf-closure.md) follow-ups **F-0337-KEYNOTE**
  (the only 1k-batch budget miss: 21.9–22.4× at 1k, 387–390× at 10k,
  quadratic) and **F-0337-ODC** (added cost 0.75 → 3.68 µs an operation
  between 1k and 10k); record [0314](0314-iwork-batch-equivalence.md)
  (a batch equals its items applied one at a time); the goal's
  PERFORMANCE BAR (a 1k-op batch within 2× of one hand-written leaf edit,
  scaling linearly).
- Scope: `crates/litchi-ops/src/format/iwork/keynote/{write,coalesce}.rs`,
  `crates/litchi-ops/src/format/odf_tree/{nav,tree}.rs`, a test in
  `crates/litchi-ops/tests/keynote_unified.rs`, `tools/unified_bench_matrix.py`
  and `docs/unified-ops/results/0340/`. No leaf crate, no public API, no
  registry or capability changed.
- `performance_claim`: §4. Record 0337's harness
  (`tools/perf-baseline/unified-batch-closure`, unchanged: same route names),
  release with `debug = 1`, features `pages,keynote,numbers,odc`, one pinned
  CPU (`taskset -c 30`) of the shared 32-core host at load 6.4–8.8, three
  trials (processes) per binary with the trunk (`before`) and branch
  (`after`) binaries alternating trial by trial, each case an in-process
  ABBA of 5 rounds (10 samples a leg) after an untimed exact-bytes check.
  Instruction counts are callgrind (`perf` is blocked). Evidence:
  [results/0340/](results/0340/) (`build.sh`, `measure.sh`,
  `binaries.sha256`, `scaling.py` → `scaling.md`, `abba/`, 88 KB).

## 1. Summary

| item | before (trunk, this session) | after |
|---|---|---|
| Keynote 1,000 paragraph sets (budget 2.0) | 20.40–21.32 | **1.093–1.110** |
| Keynote 10,000 paragraph sets (budget 2.0) | 378.3–382.9 | **1.722–1.766** |
| Keynote added µs/op, 100→1k → 1k→10k | 114.4 → 242.8 (quadratic) | **0.52 → 0.44** (linear) |
| ODC 10,000 series label sets | 1.589–1.629 | **1.067–1.080** |
| ODC added µs/op, 100→1k → 1k→10k | 0.54 → 3.61 (superlinear) | **0.55 → 0.33** (linear) |
| Pages, Numbers (no code change) | 1.08 / 1.33–1.39 at 1k | unchanged: 1.07–1.08 / 1.35–1.38 |
| bench-matrix open misses | 103 | **101** (both Keynote batch rows superseded by meets) |

Every iWork and ODC batch route is within 2× at 1k and at 10k and scales
linearly. Results are byte-identical: every case's batch output equals the
leaf's exact bytes, the batch-equivalence corpus passes at 60 sequences per
fixture, and the query-corpus differential, corpus fidelity, CRUD checklist
goldens and capability/registry goldens pass unchanged (nothing re-blessed,
§5).

## 2. Keynote (F-0337-KEYNOTE)

Record 0337's callgrind put 94 % of a 1,000-set batch in
`Edit::coalesce_set`: every paragraph set split the base and the staged text
on `\n` (cancel check, `touch`), scanned the base for the hull, opened a
leaf slide-text edit for `admits_combined_splice`, and re-split and re-joined
the whole staged text. Each is O(text) per item.

Changes (`format/iwork/keynote/write.rs`):

- **Paragraph staging.** A text's group (`TextGroup`, record 0314) now holds
  its base's paragraph byte offsets and UTF-16 offsets (one scan when the
  group opens), the replaced paragraphs (`None` keeps the base's) and a
  count of replaced paragraphs that differ from the base. A staged text is
  `Staged::Text` (a whole text) or `Staged::Paragraphs` (the group's base
  with its replacements). A paragraph set costs its paragraph; the text is
  joined once, when the run publishes (`flush`).
- **The record 0314 checks, per paragraph.** "Emptied" is one paragraph that
  is now empty; "cancels" is: the set writes the base's paragraph, the
  paragraph differs now, and it is the only one that differs (the count);
  the hull is two offset lookups. A whole text staged first keeps the
  previous split-based checks (`cancels_text`), so mixed whole/paragraph
  sequences behave as before.
- **Leaf admission once per group.** The leaf documents its answer for every
  sequence of splices within the hull, so it holds for any region within an
  admitted one (the run already cached it so). `widest_admitted` asks for
  the whole text first (one leaf call, the common case); if that is refused
  and the touched region is admitted, it widens the region paragraph by
  paragraph by binary search on each end (O(log n) leaf calls), so later
  sets inside it make no leaf call. A refused region still publishes the
  run first, as before.

Changes (`format/iwork/keynote/coalesce.rs`), after the staging fix left
the per-item path work dominant (callgrind of three 10,000-set batches:
path comparisons in the run's memo 25 %, resolution 18 %):

- the memo of resolved paths (16 entries, record 0314) compares the last
  step's index first, then the steps from the last;
- the run memoizes up to four parents (a path's steps but the last) with
  their tree item, resolves only the last step from the parent
  (`resolve::get_from`), and, once a positional step of a name has resolved
  under a parent to the paragraph that the parent's N-th paragraph child is,
  takes later `name[N]` steps of that parent directly (the iWork
  navigator's own contiguous-children check). Any other step, or a miss,
  takes the walk; a parent miss takes the full resolution.

Per coalesced item, `coalesce_set` went from O(text) to a flat 7.5 k
instructions at 1,000 sets and 7.2 k at 10,000 (callgrind, three ABBA
rounds, 8 batches); of that, the item's path parse (`TargetArg::parse`,
outside the adapter) is 1.5 k and the paragraph value coercion, done twice
(`joins` and staging), 0.7 k.

Equivalence (record 0314): `iwork_batch_equivalence` passes with
`LITCHI_BATCH_EQUIVALENCE_CASES=60` (Pages, Keynote, Numbers; 698 s,
debug). New tests: `long_paragraph_runs_equal_their_items_one_at_a_time`
(120-paragraph body; sets of empty paragraphs, sets back to the base text
and again, repeated targets, the `paragraph[N]` spelling; batch bytes equal
one-at-a-time bytes; a run without cancelling sets is one commit) and the
unit tests of `TextGroup` (offsets with multi-unit characters, empty
paragraphs, the differ count) and `cancels_text`.

## 3. ODC (F-0337-ODC)

Callgrind, `--toggle-collect` on `File::batch_with_limits`, three ABBA
rounds (8 batches), trunk binary:

| count | `coalesce_set` instructions per item, before | after |
|---:|---:|---:|
| 1,000 | 10.6 k | 4.4 k |
| 10,000 | 73.5 k | 4.4 k |

The cost was the set plan's resolution of `/series[i]`: the object
navigator (`odf_tree/nav.rs`) selected a positional step with
`filter(kind).nth(i)` over the root's children, O(i) per item, O(N²) per
batch. Fix:

- the navigator tries the slot `position` past the first child of the kind
  and accepts it when the item there has the kind and its stored ordinal is
  `position` (the iWork navigator's check, record 0326), else counts as
  before;
- tree items store that ordinal (`Item::ordinal`), and `TreeBuilder::push`
  counts per (parent, kind) in a map instead of rescanning the parent's
  children for every push (the same O(N²), paid when the tree is built).

The rest of the ODC batch is the leaf's own commit (`Edit::commit` 78 % of
the unified leg at 10,000).

## 4. Measurements

Min–max of the three trials' ratios; µs are medians of the trials'
medians. Fixture of 10,000 targets for every count.

| format | ops | trunk leaf µs | trunk unified µs | trunk ratio | after leaf µs | after unified µs | after ratio |
|---|---:|---:|---:|---|---:|---:|---|
| keynote | 100 | 5,534 | 15,727 | 2.813–2.848 | 5,688 | 5,837 | 1.026–1.029 |
| keynote | 1,000 | 5,601 | 118,720 | 20.40–21.32 | 5,625 | 6,244 | **1.093–1.110** |
| keynote | 10,000 | 6,046 | 2,303,948 | 378.3–382.9 | 6,055 | 10,673 | **1.722–1.766** |
| odc | 100 | 42,562 | 43,176 | 1.010–1.014 | 43,317 | 43,779 | 1.011–1.014 |
| odc | 1,000 | 44,277 | 45,372 | 1.018–1.027 | 44,723 | 45,684 | 1.021 |
| odc | 10,000 | 57,067 | 90,674 | 1.589–1.629 | 58,408 | 62,325 | **1.067–1.080** |
| pages | 100 | 2,177 | 2,195 | 1.008–1.014 | 2,326 | 2,337 | 1.005–1.013 |
| pages | 1,000 | 2,263 | 2,444 | 1.077–1.083 | 2,420 | 2,609 | 1.072–1.080 |
| pages | 10,000 | 3,153 | 4,987 | 1.565–1.582 | 3,432 | 5,287 | 1.509–1.544 |
| numbers | 100 | 303 | 338 | 1.115–1.131 | 300 | 344 | 1.116–1.146 |
| numbers | 1,000 | 820 | 1,106 | 1.332–1.389 | 814 | 1,111 | 1.351–1.376 |
| numbers | 10,000 | 6,218 | 8,969 | 1.442–1.538 | 6,101 | 9,216 | 1.489–1.530 |

Every case: exact bytes, one unified and one leaf commit. Linear scaling
(`scaling.md`, record 0337's rule: the unified layer's added µs an
operation 1k→10k within 1.5× of 100→1k, or under 0.5 µs): after, all four
are linear — keynote 0.52 → 0.44, odc 0.55 → 0.33, pages 0.20 → 0.19,
numbers 0.28 → 0.31; trunk keynote and odc are not (114 → 243,
0.54 → 3.61).

Keynote at 10,000 (1.72–1.77) is the closest to the budget: the leaf leg is
one `replace` (6 ms), and the unified layer adds 0.44 µs an item, of which
the path parse and the per-item value coercion are the largest parts (§2).

The bench matrix (`results/0289/bench-matrix.md`) lists the after runs as
record 0340's rows on record 0337's series; the 0337 Keynote batch rows
(1k, 10k) are "miss (superseded by 0340)"; open misses 103 → 101. The
before runs are kept as evidence and not listed (the matrix does not repeat
trunk states).

## 5. Gates

Tree: the full set below ran on this branch with trunk `4282e69000` (record 0338) merged. Trunk then gained record 0336 (`ed832736c3`: the raw layer and `File::dump`, which touch no file of this record); after that merge the focused set at the end of this section ran again. Build environment per the
rules (`CARGO_BUILD_JOBS=6`, `RUST_TEST_THREADS=8`, the memory-capped
cargo shim, one cargo command at a time); every exit status checked.

| Command | Exit |
|---|---:|
| `cargo fmt --all --check` | 0 |
| `cargo clippy -p litchi-ops -p litchi -p litchi-py --all-targets --all-features -- -D warnings` | 0 |
| `cargo clippy -p litchi-ops --no-default-features [--features keynote / odc / odg / odi / odb / odf-formula / pages / numbers] --all-targets -- -D warnings` | 0 each |
| `RUSTDOCFLAGS=-Dwarnings cargo doc --no-deps --all-features -p litchi-ops` | 0 |
| `cargo test -p litchi-ops --all-features`: `--lib` (935), `--doc` (210), the test targets in five groups (1,417; includes the query-corpus differential, corpus fidelity, the CRUD checklist goldens, `capabilities` and `registry`, not re-blessed) | 0 each |
| `cargo test -p litchi-ops --no-default-features --features` every litchi-ops feature but `encryption` `--test capabilities --test registry` (not re-blessed) | 0 |
| `LITCHI_BATCH_EQUIVALENCE_CASES=60 cargo test -p litchi-ops --features keynote,pages,numbers,odc,odg,odi,odb,odf-formula --test iwork_batch_equivalence --test keynote_unified --test odf_objects_unified --test iwork_unified --lib` (before the first trunk merge), and `LITCHI_BATCH_EQUIVALENCE_CASES=60 cargo test -p litchi-ops --features keynote,pages,numbers --test iwork_batch_equivalence` on the merged tree (710 s) | 0 each |
| `cargo test -p litchi --all-features` (234) | 0 |
| `python3 tools/check_crate_boundaries.py`; `python3 tools/non_iwork_gate.py verify` | 0 each |
| `RUSTC_BOOTSTRAP=1 python3 tools/check_api_glossary.py --target-dir target/glossary` (40 profiles, 2,313 bans, 253 families) | 0 |
| `crates/pyo3-litchi/tests/run_tests.sh` (286 tests) | 0 |
| `cd tools/perf-baseline && cargo check --locked --all-targets` | 0 |
| `cargo check --locked --offline --manifest-path tools/perf-baseline/unified-batch-closure/Cargo.toml --all-targets` | 0 |
| `python3 -m unittest tools.test_officecli_corpus tools.test_officecli_capability_delta tools.test_no_3rdparty_references tools.test_check_crate_boundaries tools.test_non_iwork_gate` | 0 |
| `python3 tools/officecli_capability_delta.py --check`; `python3 tools/officecli_corpus.py --check` | 0 each |
| `python3 tools/unified_bench_matrix.py --check`; `python3 tools/check_feature_matrix.py` | 0 each |

After the record 0336 merge (`ed832736c3`), the focused set, exit 0 each:
`cargo fmt --all --check`; the all-features clippy above and the
`--no-default-features` one; `cargo doc` of litchi-ops; `cargo test -p
litchi-ops --all-features --lib` and `--test iwork_batch_equivalence --test
keynote_unified --test odf_objects_unified --test iwork_unified --test
capabilities --test registry --test charts`; `cargo test -p litchi
--all-features`; the crate-boundary, non-iWork, tool unit-test,
capability-delta, corpus, bench-matrix and feature-matrix checks.

After the record 0341 merge (`7ea8acee2a`, merged at `eea11baa22`; detection
and open code only, the bench matrix regenerated), each exit 0:
`cargo fmt --all --check`; the all-features clippy; `--no-default-features`
clippy with no feature and with `keynote`, `odc`, `pages`, `numbers`; `cargo
test -p litchi-ops --all-features --test keynote_unified --test iwork_unified
--test iwork_table_cells --test odf_objects_unified --test capabilities
--test registry`; `LITCHI_BATCH_EQUIVALENCE_CASES=60 cargo test -p litchi-ops
--features keynote,pages,numbers --test iwork_batch_equivalence`; `python3
tools/unified_bench_matrix.py --check`.

No capability, registry, public name or OfficeCLI-visible behaviour
changed, so the goldens were not re-blessed, the delta table and the
OfficeCLI corpus were not regenerated, and no glossary row changes. No
leaf crate changed (the Keynote fix needed no leaf hook).

## 6. Gaps and follow-ups

- **F-0340-KEYNOTE-ITEM**: Keynote 10,000 paragraph sets stay at 1.72–1.77
  (within budget, the tightest iWork row). The remaining per-item cost is
  the path parse, a double value coercion (`joins` and staging each coerce
  the paragraph text) and, for steps other than a positional paragraph,
  the relative walk. Passing the coerced value from the barrier check to
  staging would save about 4 % of the item.
- The widening search (§2) is exercised only when the leaf refuses the
  whole text; the corpus fixtures and the harness admit it, so that branch
  is covered by the equivalence corpus only where a fixture's storage has
  an attribute boundary. Its correctness rests on the leaf's documented
  monotonicity, which the run already relied on.
- Record 0337's other follow-ups (F-0337-STRUCT, -DOCX-PROP, -ODT-QUERY,
  -RTF-SINGLE, -XLSX-HEADER, -XLSB-NOTE, -PATH) are untouched.

## 7. Not claimed

- Any change for formats other than Keynote and the ODF object formats;
  Pages and Numbers were re-measured only.
- ODG, ODI, ODB and ODF formula batch speed after the navigator change
  (they share it; not re-measured).
