# 0286 — OOXML adapter performance misses

Status: implemented on `uo2/perf-ooxml` (cut from `feat/unified-ops-api` at
9fe45629ab); awaiting coordinator review and merge.

Inputs: record 0281 F1 (DOCX single paragraph `set`, 1.19×) and F2 (PPTX
open + first metadata `set`, 1.31×), record 0283 (XLSX chart `set`,
1.12–1.13× at 10 charts, 1.35× at 100), record 0281's borderline DOCX cell
writes, and record 0280 G4 (PPTX `strict_scene` clones an `Arc` per shape)
and G6 (the DOCX single set's commit wrapping). Budget: the goal's
PERFORMANCE BAR, a single-op write within 10% of the hand-written leaf edit.

## 1. Summary

| Item | Before | After | Status |
|---|---:|---:|---|
| PPTX open + first `/metadata` set (0281 F2) | 1.311 | 1.053–1.058 | fixed |
| XLSX chart `set`, 10 / 100 charts (0283) | 1.125–1.132 / 1.350–1.352 | 1.014 / 1.010–1.016 | fixed |
| DOCX single paragraph `text` set (0281 F1) | 1.203–1.212 | 1.117–1.134 | **improved, still a miss (§4.1)** |
| DOCX cell `align` / `text` (0281, borderline) | 1.086–1.093 / 0.997–1.085 | median 1.085 / 1.077 (spread 1.00–1.16) | inside at the median; no change claimed (§4.2) |
| PPTX `strict_scene` per-shape `Arc` clone (0280 G4) | 937–947 ns per warm `get` (500 × 40 deck) | 437–448 ns | fixed |

Every output stays byte-identical: the chart-set and metadata harnesses
compare both legs' bytes before timing, the capability and registry goldens
pass unchanged (nothing re-blessed), and the `litchi-ops` suite (including
the corpus and query-digest differentials) passes.

## 2. Changes

### 2.1 `litchi-opc`: the main-document partname is remembered

`OpcPackage::main_document_part` re-normalized the `officeDocument` target
(`PackURI::from_rel_ref`, about 4.4 K instructions) on every call; a DOCX
leaf edit and commit call it five times, the unified set six. The package
now keeps the resolved partname with a copy of the relationship it came
from (`OnceLock<Arc<MainPartname>>`, shared by clones) and reuses it only
while the current main-document relationship has the same identifier,
target, base, source and mode. Nothing has to invalidate it: a retarget, a
removal or a new relationship simply fails the comparison and resolves as
before. The error cases are unchanged (never cached). Test
`remembered_main_document_part_follows_relationship_changes`.

### 2.2 `litchi-opc`: the preserving conversion defers part payloads

`SourceBackedPackage::to_opc_package_preserving_source` (the PPTX metadata
owner, PPTX/DOCX promotion, XLSX source adoption) inflated every member
again and re-read every relationship member through `get_part`. It always
holds the source archive in memory, so it now builds each part over a
deferred payload of that archive — the same `DeferredPartSource` an owning
`OpcPackage::from_shared_vec_with_limits` open uses — after admitting every
catalog part's declared size against `PartBytes` and `TotalPartBytes` (the
admission that open makes before deferring). The provenance step reads part
relationships through `part_metadata`, which does not decode. Parts inflate
on first `get_part`; an untouched package still publishes its exact source.
The non-preserving `to_opc_package*` conversions are unchanged. Test
`preserving_materialization_defers_part_payloads_until_read` (no part
inflated by the conversion; one inflated on read; exact-source output).

Semantic delta: a member whose deflate stream or CRC is corrupt is now
reported when that part is first read (or published), not by the
conversion. The source-backed handle never decoded it at open either, and
an owning deferred open already behaves this way.

### 2.3 `litchi-xlsx`: chart inventory memo per workbook lineage

A unified `set` on `/Sheet1/chart[N]` scanned the worksheet drawing's chart
inventory in path resolution (`Package::chart_xml`), and the native edit
scanned it twice more (`Edit::chart`, `Edit::set`). The workbook's shared
state gains `chart_inventories`, a single-entry memo modelled on the format
table cache: owner-held and shared by the lineage, keyed by the drawing
payload allocation through a `Weak` (a changed drawing never reuses an
inventory; the memo never keeps a payload alive), failures never cached,
and inventories of more than `MAX_CHARTS` frames returned but not retained.
`Package::{chart_count, chart_xml, chart_anchor}` and `Edit::frames` use it;
`Edit::add` keeps its own scan of the staged drawing. Both legs now scan a
drawing once per snapshot.

### 2.4 `litchi-pptx`: borrowed shape views

`shape::owned::Scene::views()` borrows each shape's `Shape<'_>` view in
pre-order (what `Handle::view` returns for each item of `iter()`), without
cloning the scene per shape. The ops `strict_scene` admission, which runs on
every source-owner slide read, uses it.

### 2.5 `litchi-ops` DOCX adapter

- The eager owner's body block index is cached per handle
  (`DocxState::body`); `/body` steps no longer clone the main part to build
  a `Document` view on every resolution. Source-backed owners are not cached
  (their reads are fenced against the source).
- A committed handle is built in one step (`Opened::adopted`, a small
  `docx`-gated constructor in `format/signature.rs`) instead of
  `with_owner().with_origin().with_signature()`, which allocated two
  handles only to drop them.
- `list_request` returns the admitted assignments as they are when no
  `listStyle`/`start` key is present, instead of copying them.

Bench harness additions (no route change): `DOCX_FANOUT_LOOPS` sets the
profile loop count of `docx_fanout_bench`; `OOXML_METADATA_PROFILE=native|unified`
(with `OOXML_METADATA_LOOPS`) loops one leg of the PPTX open + first set
pair in `ooxml_metadata_bench`.

## 3. Measurements

Release builds; `taskset -c 13`; AMD EPYC (32 cores) with load average
7.4–17.6 from other agents; each harness's own in-process ABBA; the
"before" binaries built from 9fe45629ab, the "after" binaries from this
branch, the two run alternately (before, after, before, after) in one
session. Raw outputs (all under 10 KB):
[`results/0286/`](results/0286/), binary SHA-256s (stripped copies) in
`results/0286/binaries.sha256`. Rebuild: `cargo build --release -p
litchi-ops --features docx,pptx,xlsx,xlsb --example <name>`;
`CARGO_NET_OFFLINE=true cargo build --release --manifest-path
tools/perf-baseline/{unified-chart-add,unified-query}/Cargo.toml`.

| Harness | Scenario | Leaf before → after | Unified before → after | Ratio before → after |
|---|---|---:|---:|---:|
| `ooxml_metadata_bench` | PPTX open + first `/metadata` set | 377 → 394 µs | 494 → 416 µs | 1.311 → 1.053 (t2 1.057) |
| | PPTX / XLSX / XLSB repeated sets | | | 1.007–1.058 → 1.007–1.053 (unchanged) |
| `unified-chart-add --compare set` | XLSX, 1 chart | 178 → 171 µs | 185 → 173 µs | 1.039 → 1.015–1.016 |
| | XLSX, 10 charts | 228 → 176 µs | 257 → 179 µs | 1.125–1.132 → 1.014 |
| | XLSX, 100 charts | 727–732 → 211 µs | 981–990 → 213 µs | 1.350–1.352 → 1.010–1.016 |
| `docx_fanout_bench` | `set /body/p[1] text` | 8.27–8.38 → 7.68 µs | 10.02–10.08 → 8.63–8.68 µs | 1.203–1.212 → 1.124–1.130 |
| | `/body > p`, 100 paragraphs | | | 1.073–1.075 → 1.063–1.064 |
| `docx_element_set_bench` (5 trials each) | cell paragraph `align` | 160–162 µs | | 1.086–1.093 → 1.076–1.158 (median 1.085) |
| | cell `text` | | | 0.997–1.085 → 1.000–1.147 (median 1.077) |
| `unified-query` | `pptx_large_bytes_get_shape_v1` | 99 ns | 937–947 → 437–448 ns | 9.5 → 4.4–4.5 |
| | `pptx_large_bytes_all_shape_v1` | 280–314 µs | 2.31–2.99 → 2.08–2.15 ms | 8.3–9.5 → 7.3–7.7 |

Earlier interleaved runs in the same session agree: DOCX single set
1.192–1.195 → 1.117–1.134 (three trials); PPTX first set 1.057–1.058.

Instruction counts (callgrind, `profiling` profile, per call):

- DOCX single set: leaf 144.5 K → 125.6 K, unified 159.9 K → 134.1 K
  (unified − leaf 15.4 K → 8.5 K).
- PPTX open + first set: leaf 7.04 M; unified 9.69 M before (open 3.55 M,
  materialization 4.11 M, of which part inflation 1.88 M and relationship
  re-reads through decoding `get_part` the rest; edit and commit 1.31 M).

## 4. Misses left, with measured explanations

### 4.1 DOCX single paragraph `text` set: 1.12–1.13×

Of the remaining 8.5 K instructions per call over the leaf's 125.6 K:

- about 4.6 K in the shared target and path code: `TargetArg::parse`
  (`has_query_syntax`, `PathArg::parse`, 1.7 K) and `path::resolve::get`
  (step canonicalization in `Vocabulary::canonical`, `resolve::rule`, the
  generic frame; 2.9 K). These belong to the shared query/path owner
  (`crates/litchi-ops/src/{target,path}`), which this record was told not
  to change.
- about 2 K of handle bookkeeping inherent to immutable handles: the new
  `Docx` (`Arc` state and `Arc` handle), the boxed editor and the boxed
  committed package — four allocations the leaf does not make.
- about 1.9 K in the DOCX adapter itself (registry admission 0.8 K, the
  `/body` and `p[N]` steps 0.8 K, target admission 0.3 K).

The time ratio (1.12–1.13) follows the instruction ratio (1.07) plus the
unified leg's extra allocations; malloc/free are the largest single
remaining cost (132 vs 124 calls per op). Follow-ups: **F1** pre-classified,
pre-canonicalized parsed targets for the shared path owner (the `&str`
target is parsed and its steps canonicalized on every call; a reused
`Path` value avoids only the first); **F2** inline storage for the
one-node/one-target/one-assignment case in the DOCX adapter (three
allocations, about 1% — not worth a new dependency in `litchi-ops`).

### 4.2 DOCX cell writes

The cell paths were not changed beyond the shared memo of §2.1. In five
alternating trials per binary the medians are 1.085 (`align`) and 1.077
(`text`) after versus 1.088 and 1.076 before; single trials spread from
1.00 to 1.16 at load average 8–13, so the harness's 9 ABBA rounds are too
few to separate them. Inside the budget at the median; no improvement is
claimed. A quiet-host rerun with more rounds would settle it (**F3**).

## 5. Gaps and follow-ups

- **F1** shared target/path parsing and canonicalization cost per call
  (path owner).
- **F2** DOCX adapter inline storage for single targets (optional, ≈1%).
- **F3** DOCX cell writes on a quiet host with more rounds.
- Record 0262 F3 (a source-backed PPTX metadata revision without any
  materialization) is no longer needed for the budget; it would still save
  the remaining materialization and relationship re-reads (§3).

## 6. Gates

Build environment per the rules file; every exit status checked.

| Gate | Command | Exit |
|---|---|---:|
| format | `cargo fmt --all --check` | 0 |
| clippy | `cargo clippy -p litchi-opc -p litchi-xlsx -p litchi-pptx -p litchi-ops -p litchi -p litchi-py --all-targets --all-features -- -D warnings` | 0 |
| clippy per feature | `cargo clippy -p litchi-ops --no-default-features [--features docx\|pptx\|xlsx] --all-targets -- -D warnings` | 0 each |
| tests | `cargo test -p litchi-opc --all-features`, `-p litchi-xlsx`, `-p litchi-pptx`, `-p litchi-docx` (2,389 passed), `-p litchi-ops`, `-p litchi` | 0 each |
| goldens | `cargo test -p litchi-ops --all-features --test capabilities --test registry`, and with every feature but `encryption` (not re-blessed; no diff) | 0 |
| rustdoc | `RUSTDOCFLAGS=-Dwarnings cargo doc --no-deps --all-features -p litchi-opc -p litchi-xlsx -p litchi-pptx -p litchi-ops` | 0 |
| crate boundaries | `python3 tools/check_crate_boundaries.py` | 0 |
| non-iWork | `python3 tools/non_iwork_gate.py verify` | 0 |
| glossary | `RUSTC_BOOTSTRAP=1 python3 tools/check_api_glossary.py --target-dir …` (33 profiles, 1,410 bans, 46 families) | 0 |
| perf-baseline | `cd tools/perf-baseline && cargo check --locked --all-targets` | 0 |
| Python | `crates/pyo3-litchi/tests/run_tests.sh` (265 tests) | 0 |

No capability, registry or OfficeCLI-visible behaviour changed, so the
delta table and the OfficeCLI corpus were not regenerated. New public API:
`litchi_pptx::shape::owned::Scene::views` (borrowed counterpart of `iter`,
named after `Handle::view`; no glossary row changes).
