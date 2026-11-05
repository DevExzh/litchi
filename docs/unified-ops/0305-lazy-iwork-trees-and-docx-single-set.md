# 0305 — Lazy iWork trees, the DOCX single-set cost, the RTF `&Query` route

- Status: implemented on `uo2/perf-lazy`, cut from `feat/unified-ops-api`
  at `f24868dfb5`, trunk `e7fe30fb24` merged (`109639b9e7`; the merge
  changed no source on the measured routes), then trunk `6c8d517eff`
  (`191da4bf26`), then trunk `bcdcadadf0` (record 0307, `604fcdf18b`) and
  `9940d9bfe0` (`7adf1cce79`).
- Date: 2026-10-16
- Inputs: the open misses against `docs/UNIFIED_OPS_GOAL.md`'s PERFORMANCE
  BAR from records 0298 (Keynote open + first `get`), 0297/0302 (Numbers
  open + first `get`), 0286/0291/0302 (DOCX single paragraph `set`) and
  0302 §4.4 (RTF fan-out harness); and, added by the coordinator during
  the work, record 0307's 1,000-paragraph DOCX text set (≈2.5×) and
  paragraph add (≈2.1×).
- Scope: `litchi-ops` `format/iwork{,/*}` (new `format/iwork/part.rs`),
  `format/docx{,/write,/publish,/commit_tests}`, one hunk in `write.rs`
  (the DOCX arm of `File::set`), `examples/rtf_destination_bench.rs` (new
  route names only), the new `examples/docx_large_bench.rs`, and one leaf
  hook in `litchi-docx` (`source_backed.rs`: a snapshot memo, §2.4).
- `performance_claim`: paired measurements in §3. `profiling` builds
  (release with debug info, stripped before timing) of
  `--features docx,pages,numbers,keynote,rtf`, pinned with `taskset -c 30`,
  processes in B A A B order, each harness an in-process ABBA. Instruction
  and cache counts are callgrind (`perf` is blocked,
  `perf_event_paranoid=4`), the cache simulation with the host's L1 shape
  (`--I1=32768,8,64 --D1=32768,8,64`). Evidence:
  [results/0305/](results/0305/): driver `measure.sh`, `summary.txt`, raw
  harness outputs, `binaries.sha256`, `docx-cachesim.txt` and its driver
  `docx_cg.sh`. Host: shared 32-core EPYC, load 5–6.

## 1. Summary

| Item | Before | After | Status |
|---|---:|---:|---|
| Keynote open + first `get` (`keynote_bench open_get`) | 1.098–1.105 | 1.023–1.034 | **fixed** (≤ 5 %) |
| Numbers open + first `get` (`numbers_unified_bench open_get`) | 1.211–1.214 | 1.060–1.088 | improved; 6–9 µs over, **met** by the +20 µs arm (§4.2) |
| DOCX single paragraph `set` (`docx_fanout_bench`, 1 ¶) | 1.077–1.095 | 1.053–1.058 | **fixed** (≤ 10 %), cache misses reduced (§3.3) |
| DOCX `/body > p`, 100 ¶ | 1.061–1.063 | 1.053–1.060 | unchanged within noise |
| RTF fan-out, true plan-cache-hit cost (`&Query` vs `Plan`) | hidden by the harness | 0.99–1.02 (×1: +0.4–2.2 %) | measured; not material, no change (§4.3) |
| DOCX 1,000 ¶ text set, default (source-backed) owner | 2.51 (1,878–1,884 µs) | 1.56–1.57 (1,168–1,175 µs) | improved, **miss** (§4.5) |
| DOCX 1,000 ¶ paragraph add, default owner | 2.11 | 1.46 | improved, **miss** (§4.5) |
| DOCX 1,000 ¶ text set / add, eager owner | 1.00 / 0.38 | 1.01 / 0.38 | met (no size-dependent adapter cost) |

Every output stays byte-identical. The query-corpus differential, the
corpus fidelity test, the checklist goldens, `path_allocations`, the
capability and registry goldens (both feature sets) and the OfficeCLI
corpus pass unchanged; nothing was re-blessed. In addition a scratch dump
(not committed) of every Pages, Keynote and Numbers fixture under
`test-data/` (125 packages: all four views as text and JSON, a full child
walk with kind, paths, text, properties and child counts, 7,244 fresh-handle
`get`s including positional cell, merge and A1 probes, and 16
queries) printed byte-identical output (2,046,134 bytes) at base and after.

## 2. Changes

### 2.1 Lazily built iWork tree parts (`format/iwork/part.rs`)

The first read used to build the whole object tree: every Keynote slide's
placeholders, notes, drawable inventory, tables and the theme's layout
catalog; every Numbers cell. The eager tree now holds the root and the
containers every read needs (Numbers sheets and tables with their merge
regions, the Keynote root), and *parts* hold the rest:

- **Keynote slides** are one part whose items are materialized one by one;
  the first slide item reads the layout catalog (once, shared by all
  slides). The eager build still decodes the show projection
  (`Package::slides`), so a show that fails to decode still fails every
  read, as before.
- **A Keynote slide's children** are one part, built whole the first time
  a read reaches them, and rebuilt with the slide's character runs appended
  when a read reaches a paragraph's children (the former whole-tree
  `full` rebuild, now per slide). Every item keeps its index.
- **A Numbers table's cells and merged ranges** are one part of slots, each
  materialized on first touch. A cell step (`A3`) binary-searches the
  leaf's sorted stored-cell list (the leaf's own lookups binary-search the
  same list) and builds that one item; an unstored in-grid cell is a miss
  without materializing anything.

Items are addressed by `At { part, index }` (`Node` stays 16 bytes). A
step's kind and ordinal are read without materializing a slot
(`kind_at`/`ordinal_at`), so `slide[2]`, `cell[N]`, `cell` queries by kind
and the `Last` step touch only the item they return. `Children`, `Nav`,
views and the Keynote/Numbers write helpers read children through
`Iwork::kids`. The views list items in the document order the former
single arena had (each container followed by its part, then every part's
deferred character runs, part by part), which the dump above checks.

Tests: `format::iwork::part::tests` asserts what each read builds
(`/slide[2]/placeholder[title]` builds slide 2's part only, no slide item
and no catalog; a slide's format materializes that slide and reads the
catalog once; a paragraph's children add that slide's runs; a view
completes every part; one Numbers `get` materializes one of 110 cells, a
kind query none, reading the matches all).

### 2.2 DOCX single-operation `set`

- `File::set` calls `Docx::set_file`, which adopts the committed package
  where the leaf commit returns and drops the leaf patch there; the general
  `Docx::set` kept returning the whole `Commit` (424 bytes) through four
  frames only for `into_file` to discard the patch. `set_inner` is generic
  over what it publishes (`Publish` for `Commit` and `Docx`); transactions
  still get the commit with its patch and diagnostics.
- A `text`-only set on a direct `/body/p[N]` path (exactly the step names
  `body` and `p`, no predicate on `body`, one positional index) of an eager
  owner reads the paragraph's transaction position from the cached body
  index — the block the resolver would return — instead of running the
  generic resolver, node construction and target admission. It then runs
  the same steps as the general route: the editor first (so the same
  security refusals), the same registry admission, the same
  `apply_with_policy` and leaf replacement. Any other spelling, owner,
  property list, a missing paragraph or a read failure takes the general
  route, which reports it. Test
  `direct_paragraph_text_matches_the_general_route` compares both routes'
  bytes and errors on every DOCX fixture (eager and source-backed owners,
  first paragraphs, the last, one past the last, a text, a number and a
  `null` value, three spellings).
- `DocxOwner::Eager` holds its package unboxed (56 bytes): one allocation
  and free fewer per committed handle.

Tried and rejected: keeping the leaf editor unboxed (`DocxEditor::Edit` and
`WorkEdit::Eager` holding the 448-byte `Edit` by value) moved it by value
four to six times per call; it cost +1.2 K instructions and +30 D1 write
misses per call (build B in `docx-cachesim.txt`) and was reverted.

### 2.3 RTF fan-out harness route

`rtf_destination_bench` gains route G, `unified_fanout_ref_*`: the fan-out
passes `&Query`, so the per-thread plan cache is hit without the harness
cloning an owned `Query` per call, and `fanout_ref_over_plan` compares it
with the precompiled-`Plan` route F. Existing route names and fields are
unchanged; the order string is now `ABCDEFGGFEDCBA`.

### 2.4 Source-backed DOCX transaction snapshot memo (`litchi-docx`)

Record 0307's harness opens its document with `open_bytes`, whose default
`SourceMode::Auto` keeps a source-backed owner; its leaf leg edits an eager
package. On the eager owner the unified set and add of a 1,000-paragraph
document cost what the leaf costs (§3.6), so nothing in the adapter scales
with document size. On the source-backed owner, callgrind of 20 sets
(43.4 M instructions per call against 17.5 M on the eager owner) showed
three whole-document passes the eager owner does not pay:

| pass | per call |
|---|---:|
| `source_backed::Package::edit` → main-document transaction snapshot (`scan_document_with_context`) | 7.2 M |
| the commit's `apply` → the same snapshot of the same revision, again | 7.2 M |
| `PreparedTopology::apply_source_xml` → `audit_source_part_replacement_xml` (litchi-opc) | 9.5 M |

`source_backed::Package` now memoizes the last unmanaged transaction
snapshot it built (`main_snapshot`, shared by its clones beside the
existing `main_admission` memo). It is served only when its source
identity (lineage, source revision, part name) matches and its bytes equal
the current payload (pointer or byte equality), so a revision is scanned
once, not once per edit and again per commit. The managed branch already
reused the edit's snapshot through its `before` hint; it is unchanged.
Test `unmanaged_edits_share_one_transaction_snapshot` checks the memo
answers for the same revision and bytes, never for other bytes, and that
an edit from the memo publishes the bytes of an edit from a fresh scan.

## 3. Measurements

Ratios are unified/leaf medians per process run, runs in the order base-1,
after-1, after-2, base-2 (`summary.txt`).

### 3.1 Keynote (`keynote_bench`, 8 rounds × 30)

| scenario | base 1 / 2 | after 1 / 2 |
|---|---:|---:|
| open | 0.982 / 0.983 | 0.982 / 0.998 |
| **open + get** | 1.105 / 1.098 | **1.023 / 1.034** |
| set title | 1.036 / 1.038 | 1.008 / 1.006 |
| set notes | 1.075 / 1.080 | 1.047 / 1.024 |
| set frame | 1.057 / 1.063 | 1.023 / 1.015 |
| 1k batch | 1.078 / 1.080 | 1.074 / 1.071 |
| duplicate | 1.063 / 1.063 | 1.029 / 1.029 |
| add slide | 1.013 / 1.010 | 1.003 / 1.005 |
| add text box | 1.056 / 1.056 | 1.030 / 1.026 |
| remove slide | 1.027 / 1.028 | 1.011 / 1.014 |

Callgrind (`--only open_get --rounds 1 --iterations 5`, ten unified first
reads): program 559.7 M → 541.4 M instructions, −1.83 M per first read of
the base's 2.56 M tree build. What a first read still builds: the show
projection (0.45 M, kept eager for its error semantics and the slide
count), slide 2's part (0.31 M). Writes resolve their targets through the
same parts, so every write scenario got cheaper too.

### 3.2 Numbers (`numbers_unified_bench`, 7 samples)

| scenario | base 1 / 2 | after 1 / 2 |
|---|---:|---:|
| **open + get** | 1.211 / 1.214 | **1.060 / 1.088** |
| set | 1.006 / 0.998 | 1.000 / 1.026 |
| batch 1,000 | 1.354 / 1.297 | 1.316 / 1.315 |
| merge add, sheet/table lifecycle, rename, moves | 0.97–1.03 | 0.97–1.04 |

Open + get: the base built and dropped 110 cell items (about 0.21 M
instructions to build, 0.17 M to drop); after, one cell. Program total
357.8 M → 355.4 M for the run's ten tree builds.

### 3.3 DOCX single paragraph `set` (`docx_fanout_bench`)

| run | leaf µs | unified µs | ratio |
|---|---:|---:|---:|
| base 1 / 2 | 7.64 / 7.56 | 8.23 / 8.29 | 1.077 / 1.095 |
| after 1 / 2 | 7.54 / 7.55 | 7.94 / 7.98 | 1.053 / 1.058 |
| 100 ¶ base 1 / 2 | 344.5 / 344.7 | 365.5 / 366.3 | 1.061 / 1.063 |
| 100 ¶ after 1 / 2 | 347.0 / 345.8 | 365.6 / 366.4 | 1.053 / 1.060 |

Cache simulation, per call ((3000 loops − 1000 loops) / 2000):

| build | leg | Ir | I1 misses | D1 read misses | D1 write misses |
|---|---|---:|---:|---:|---:|
| base | leaf | 125.1 K | 2,618 | 116 | 120 |
| base | unified | 131.7 K | 3,209 | 327 | 469 |
| after | leaf | 124.7 K | 2,659 | 144 | 108 |
| after | unified | 128.7 K | 2,979 | 227 | 299 |

Unified − leaf: instructions +6.6 K → +4.0 K (5.3 % → 3.2 %), L1i misses
+591 → +320, D1 read misses +211 → +82, D1 write misses +349 → +191. The
patch-free commit alone (build CD) removed 146 write misses (they were
the moves of the 424-byte commit into cold stack lines) but barely moved
the time; the direct route removed the resolver's instructions and code
footprint (its cold lines are the L1i misses), which is what moved it.

### 3.4 RTF (`rtf_destination_bench`)

Route G (fan-out with `&Query`) against route F (precompiled `Plan`),
after builds: ×1 targets 1.004–1.022, ×100 0.999–1.006, ×1,000
0.993–1.002. The fan-out with the owned clone (route D) was 1.02–1.07 over
F at ×1, so the clone (record 0302: 2.0 K instructions, two allocations)
was most of the gap that record saw; a plan-cache hit itself costs about
0.04–0.08 µs on a 7–8 µs single-target fan-out.

### 3.5 Pages (`pages_bench`, regression check)

batch 1 / 1,000 / 10,000: base 1.014–1.017 / 1.084–1.088 / 1.529–1.535,
after 1.018–1.019 / 1.087–1.098 / 1.568–1.573. The coalesced Pages route
builds no tree; callgrind of the 10,000-item scenario counts 1,519.64 M
instructions at base and 1,519.76 M after (+0.007 %), so the 2.5 % spread
at 10k is layout and host noise, not added work.

### 3.6 DOCX on a 1,000-paragraph document (`docx_large_bench`)

The middle paragraph of record 0307's fixture shape; µs per call, base
`191da4bf26` (this branch before §2.4) and after `5cda11f624`, B A A B
(`docx-large-*.txt`):

| leg | base 1 / 2 | after 1 / 2 |
|---|---:|---:|
| leaf text set | 751.4 / 749.0 | 746.2 / 747.9 |
| unified set, default owner | 1,884.3 / 1,878.5 | 1,174.9 / 1,167.9 |
| unified set, eager owner | 752.2 / 748.1 | 752.7 / 751.7 |
| leaf `insert_paragraph` | 1,097.0 / 1,093.7 | 1,088.0 / 1,094.6 |
| unified add, default owner | 2,310.0 / 2,307.9 | 1,589.0 / 1,599.2 |
| unified add, eager owner | 418.4 / 418.9 | 419.1 / 413.4 |

(The unified add on the eager owner is faster than the leaf leg because
the leaf `insert_paragraph` the harness times is a different, heavier
leaf call than the one the unified `add` stages; the comparison that
matters is the owner: default against eager.)

## 4. Misses and residuals

### 4.1 Keynote open + first `get`: met

1.023–1.034. The residual is the eager show projection (needed so a show
that fails to decode fails the first read, as before) and slide 2's own
part.

### 4.2 Numbers open + first `get`: 1.060–1.088

6–9 µs over a 100 µs leaf leg: within the open budget's +20 µs arm, above
its 5 % arm. What remains is the eager per-table merge read (about 35 K
instructions; kept so a table whose merge storage fails still fails the
first read), the sheets and tables, one cell, and the harness's owned copy
of the 11 KB input (`open_bytes` takes a `Vec`; the leaf leg borrows).

### 4.3 RTF plan-cache hit: not material

Hashing and comparing the query on a hit is about 1 K instructions (record
0302); a cheaper key would save at most a fraction of that on a 7 µs call.
No change made.

### 4.4 DOCX single `set`: met, residual 3.2 % instructions

The rest is shared code outside this record's scope: path parsing and the
target classification (`litchi-core`, `target.rs`), registry admission and
value coercion, and the handle's two `Arc`s.

### 4.5 DOCX large-document set and add on a source-backed owner: 1.46–1.57

After §2.4 the source-backed set costs about 420 µs over the leaf, the add
about 500 µs: the remaining size-dependent pass is litchi-opc's audit of
the replacement main part when the source overlay publishes it
(`PreparedTopology::apply_source_xml` → `check_for_replacement` →
`audit_source_part_replacement_xml`, 9.5 M instructions on 1,000
paragraphs). The transaction commit has already verified those bytes
(`verify_with_policy`, paid by the eager leaf too); the eager route skips
the publisher's second audit through the commit's `VerifiedSource` proof
(change 0754), the source overlay has no way to receive it. Follow-up
(litchi-opc, outside this record's scope): let `apply_source_xml` accept
the transaction's publication proof for the exact target bytes. A handle
committed from a source-backed owner edits through its effective topology,
whose snapshot path (`effective_document_snapshot`) is not memoized; a
chain of sets on successive handles still scans each revision once per
edit — part of the same follow-up.

## 5. Gates

Build environment per the rules (`CARGO_BUILD_JOBS=6`, `RUST_TEST_THREADS=8`,
one cargo command at a time, test executables pruned between chunks). Every
exit status was checked. The whole list ran on `109639b9e7` + this record
and again after each trunk merge and §2.4; the exits and counts below are
the final merged tree's (`7adf1cce79`), run after the 07:46 session loss
under the memory-capped cargo scopes.

| Command | Exit |
|---|---:|
| `cargo fmt --all --check` | 0 |
| `cargo clippy -p litchi-docx -p litchi-ops -p litchi -p litchi-py --all-targets --all-features -- -D warnings` | 0 |
| `cargo clippy -p litchi-ops --no-default-features [--features F] --all-targets -- -D warnings`, F = none, docx, pages, keynote, numbers, rtf | 0 each |
| `cargo test -p litchi-docx --all-features` (2,399) | 0 |
| `cargo test -p litchi-ops --all-features` in chunks: `--lib` (918), `--doc` (210), the test targets in five groups (1,226) | 0 each |
| `cargo test -p litchi-ops --no-default-features --features` every litchi-ops feature but `encryption` `--test capabilities --test registry` (not re-blessed) | 0 |
| `cargo test -p litchi --all-features` (230; trunk removed the iWork facade tests) | 0 |
| `RUSTDOCFLAGS=-Dwarnings cargo doc --no-deps --all-features -p litchi-ops -p litchi-docx` | 0 |
| `python3 tools/check_crate_boundaries.py` | 0 |
| `python3 tools/non_iwork_gate.py verify` | 0 |
| `RUSTC_BOOTSTRAP=1 python3 tools/check_api_glossary.py --target-dir target/glossary` (39 profiles, 2,305 bans, 221 families) | 0 |
| `crates/pyo3-litchi/tests/run_tests.sh` (278 tests) | 0 |
| `cd tools/perf-baseline && cargo check --locked --all-targets` | 0 |
| `cargo test -p litchi-cli --all-features --test officecli_corpus` (not re-blessed) | 0 |

No capability, registry or OfficeCLI-visible behaviour changed, so the
delta table and the OfficeCLI corpus were not regenerated. No public name
was added or changed; no glossary row changes.

## 6. Not claimed

- That Numbers open + first `get` meets the 5 % arm of the open budget.
- That a DOCX set or add on a large source-backed document meets the 10 %
  single-op budget (§4.5).
- That the unboxed DOCX editor helps (it did not; §2.2).
- Any change of a frozen route name; `rtf_destination_bench` only gained
  route G's fields.
