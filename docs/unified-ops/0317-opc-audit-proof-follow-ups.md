# 0317 — Audit-proof follow-ups: retained proofs for DOCX commits, derived indexes for chains, depth-3 slide indexes

- Status: implemented on `uo2/opc-proofs`, cut from `feat/unified-ops-api`
  at `30eed270b5`. Commits: `16ccc38cac` and `176d9db30d` (xml-minifier),
  `111b0d3c4d` (litchi-opc), `c426f44ce4` (litchi-docx), `e9cefdaf2a`
  (`pptx_large_bench`), and this record. Trunk `cbbdcd1fde` (record 0314,
  iWork only; no file this record changes) is merged last (§6.1).
- Date: 2026-10-16
- Input: record 0313 §4.1–§4.3, from the follow-ups the coordinator listed:
  - the DOCX default-owner set still scans the original once;
  - chained edits on a committed handle fall back to a large run or to the
    complete audit;
  - PPTX shape edits sit deeper than the depth-2 boundary index.
- Scope:
  - `xml-minifier`: `audit.rs`, `audit/successor.rs`, one debug helper in
    `audit/namespaces.rs`, the tests `tests/successor.rs` (updated) and
    `tests/successor_chain.rs` (new).
  - `litchi-opc`: `source_backed.rs`, `source_backed/{revision,view}.rs`, and
    the new test `tests/source_xml_proofs.rs`.
  - `litchi-docx`: `document/transaction.rs`, `source_backed.rs`,
    `source_backed/owner.rs` and `source_backed/document_memo_tests.rs`.
  - `litchi-ops`: one new example, `examples/pptx_large_bench.rs`.
  - No `litchi-pptx` source and no file under `litchi-ops/src/` changed. In
    particular, nothing on the DOCX adapter's read path (`Children::next`,
    record 0315) changed.
- `performance_claim`: §4. `profiling` builds (release with debug info,
  stripped for timing) of
  `cargo build --profile profiling -p litchi-ops --features docx,doc,odp,xls,rtf,odt,pptx,ppt --example {docx,pptx}_large_bench`.
  - Each binary was linked on its own; linking two at once hit the 26 GB
    shim cap.
  - Runs were pinned with `taskset -c 30`, processes B A A B.
  - The base build is this branch with
    `crates/{xml-minifier,litchi-opc,litchi-docx}/src` at `30eed270b5`, so
    it has the new PPTX bench.
  - Instruction and cache counts come from callgrind with
    `--cache-sim=yes --I1=32768,8,64 --D1=32768,8,64`, because `perf` is
    blocked.
  - Evidence is in [results/0317/](results/0317/): `measure.sh`, `cg.sh`,
    `gates.sh`, the raw harness outputs, `docx-pptx-cachesim.txt`,
    `binaries.sha256` and `host.txt`.
  - Host: a shared 32-core EPYC.

## 1. Summary

| Item | Before (0313) | After | Status |
|---|---:|---:|---|
| DOCX 1,000 ¶ text set, default owner (µs; Ir per call) | 407–411; 10.00 M | **23.3–23.4; 0.52 M** | **fixed** (§2.3) |
| the same set / leaf `replace_paragraph_text` | 0.534–0.536 | **0.031** | met |
| set on a handle committed from the default owner | 412–416; 10.22 M | **25.1–25.4; 0.71 M** | **fixed** (§2.2) |
| add on a handle committed from the default owner | 267–270; 5.89 M | **81.8–82.3; 1.37 M** | **fixed** (§2.2) |
| the same committed add / the eager owner's add | 0.64 | **0.195** | |
| add on the default owner (µs; Ir) | 73.1–73.6; 0.907 M | 74.2–74.8; 0.950 M | **+1.0–1.6 % wall, +4.8 % Ir** (§5.1) |
| PPTX 1,000-shape slide: shape text set + write (µs) | 20,335–20,474 | 19,227–19,385 | −5.2 %; −5.4 % Ir |
| PPTX shape add + write (µs) | 29,294–29,338 | 27,819–27,849 | −5.0 %; −8.2 % Ir |
| PPTX unified / leaf, set + write and add + write | 0.995–1.005 | 0.991–1.006 | unchanged; the leaf gains equally |
| every other 0311 / 0313 DOCX leg | | within 1 % wall, 0.4 % Ir | unchanged |

All output stays byte-identical:
- Each bench checks that its paired legs write the same bytes before
  timing.
- The query-corpus differential, corpus fidelity, the checklist goldens, the
  capability and registry goldens and the OfficeCLI corpus pass without
  re-blessing (§6).

## 2. Design

### 2.1 `xml-minifier`: derived indexes (`16ccc38cac`, `176d9db30d`)

Record 0313's successor proof returned a proof without boundaries. So the
next edit of the successor's bytes was proved from the *original's* index,
and its run had to cover both edits.

`verify_successor` now gives the successor an index of its own, at the base
proof's depth and within the same bounds:

- **A splice** derives the replacement's index from three parts:
  - the base's boundaries before the run, which the replacement repeats
    byte for byte together with the scan that reaches them;
  - the window's boundaries, which a `BoundaryRecorder` resumed at the run's
    start records while the window is audited. It works from the run
    start's offset and counters and knows the base's parents. By the
    invariant at `audit::scan`, the window scan agrees with the complete
    scan on every token of the window, its depth and its bindings;
  - the base's boundaries from the run's end on, moved by one signed shift
    of offset and counters. The window is balanced and closes every binding
    it opens, so the state after it is the original's with re-totalled
    counters.

  When the window opens a parent, or drops the last boundary of a parent
  other than the run's own, the parents and scopes are compacted, so every
  parent of an index is referenced by one of its boundaries. Otherwise they
  are kept as they are.
- **An identical replacement** shares the base's index.
- **A replacement audited completely** records its own index during that
  same audit, so a chain recovers after a fallback.
- **An element window** (the pair route's `Window` proof) derives the index
  from the boundary run that holds the same edit. This re-audits that run's
  few bytes.

Soundness discipline:
- Debug builds re-derive every splice proof with a recording audit of the
  replacement (`debug_check_successor_proof`). They also compare the derived
  index with the recorded one: offsets, counters, a one-to-one
  correspondence of parents, and equal bindings in each parent's scope.
- A deliberately broken shift (an off-by-one in the events counter) makes
  the new chain tests fail through that check.
- Scope sharing may differ from a fresh recording, which is why parents are
  compared by correspondence, not by number.

`IndexDepth` (`CHILDREN` = 2, the default; `GRANDCHILDREN` = 3) makes the
recording depth explicit:
- `verify_indexed` and `verify_indexed_replacement` take it;
- `VerifiedSource::index_depth` reports it;
- the index records levels `1..=depth` and searches the deepest first.

Namespace snapshots are now behind `Arc`, so derived indexes share them.
The charge is the clone charge plus two words.

**Bounds** are those of 0313 at every depth:
- one boundary per 32 bytes of the indexed payload, plus 64, at 20 bytes
  each;
- 4 bytes per parent, with every parent referenced by a boundary;
- namespace snapshots within the 8 MiB multi-window budget;
- inputs shorter than `u32::MAX`.

A derivation over any bound keeps no index, and the next edit is audited
completely. The test `a_derived_index_over_its_bound_is_dropped` shows this.

### 2.2 `litchi-opc`: the overlay keeps the effective Part's proof (`111b0d3c4d`)

- `PreparedPart::xml_proof` keeps the proof its replacement's audit
  returned:
  - from `prepare_topology_state`, for plan replacements;
  - from `replace_effective_payload`, for revisions.

  It is kept only on unmanaged packages and only when it covers the
  replacement's exact allocation, so it adds only its index. A revision
  keeps at most `RETAINED_EFFECTIVE_XML_PROOFS` (4). Installing a fifth
  strips the first other Part's proof, and that Part's next edit is then
  proved from the physical original's.
- `audit_physical_replacement_xml` takes a `predecessor`, the effective
  revision's retained proof. Once the physical original is remembered as
  passing, the replacement is audited with `predecessor.verify_successor`
  when the predecessor has an index, and from the original's proof
  otherwise.
  - The verdict and error are the replacement's own complete source audit
    whatever bytes the predecessor proves. So the contract of 0311 holds:
    the pair verdict, the original's failure reported every time, and
    managed packages remembering nothing.
- New public accessors, which audit nothing:
  - `PreparedPartView::xml_audit_proof` returns the revision's proof for a
    replaced Part and the remembered original's for an unchanged one;
  - `SourceBackedPackage::xml_audit_proof` returns the remembered
    original's.
- `write_part_overlays_shared_to_stream`, the PPTX owner's write route,
  audits each overlay through `audit_physical_replacement_xml` instead of
  the plain pair audit. Its verdict and error are the same. So the slide is
  remembered with its indexed proof on the first write, and later writes
  scan only the edited run.
- `source_xml_index_depth` gives PresentationML slides, layouts, masters,
  notes slides, notes masters and handout masters `GRANDCHILDREN`
  (`p:sld/p:cSld/p:spTree` content), and every other Part `CHILDREN`.

### 2.3 `litchi-docx`: the commit gets the remembered proof (`c426f44ce4`)

- **Attaching the proof.** The source-backed owner attaches the OPC
  package's retained proof for the transaction base's exact allocation
  (`with_retained_main_proof`), and the snapshot memo then keeps the proven
  snapshot:
  - on the physical route, the remembered original's indexed proof;
  - on a committed revision, the revision's own proof. The snapshot now
    shares that allocation instead of copying the effective bytes.

  `Snapshot::with_publication_proof` still drops a proof for any other
  allocation.
- **Using it in the commit.** `commit_preserving_unmodified` takes the gate's
  verdict on a proven base from the proof, and audits the candidate with
  `verify_successor`. The candidate's verdict and error are its complete
  source audit, the original is not scanned, and the candidate's proof keeps
  a derived index for the next commit. A failing candidate is published
  unproven, so the writer's audit refuses it where it always has, as before.
- **A base without a proof.** A source-identity base without a proof (the
  owner's first commit) takes `verify_indexed_replacement`, so its candidate
  carries an index. Identity-free snapshots (the eager `Package`, the leaf
  legs) keep the plain `verify_replacement`. An index there would cost them
  0.54 M instructions per call (+3.2 % on `leaf_set`, measured on an
  intermediate build), and no writer of theirs keeps the proof.

## 3. Tests

- **`xml-minifier`, `tests/successor_chain.rs` (9 tests).** Every verdict is
  compared with `verify_source`, error and offset included. Every passing
  successor keeps exactly as many boundaries as a fresh recording audit at
  its depth.
  - Chained edits 36 paragraphs apart are each proved by their own
    paragraph, and so are an insertion and an edit inside the inserted
    paragraph.
  - A removal is proved by its neighbour's run. An edit of the document
    element's start tag is audited completely, records its own index, and
    the chain goes on. An identical successor shares the index.
  - Eleven defect kinds on a chained proof are refused with the complete
    audit's errors: unbalanced tags, closing an ancestor, an unbound prefix,
    a duplicate attribute, an unknown entity, `]]>`, a control character, a
    late declaration, `xmlns:xmlns`, an empty prefix binding, and a
    DOCTYPE.
  - A prefix declared inside a window scopes the next edit inside it, and
    is unbound outside it.
  - The event budget is re-totalled along a chain of insertions, and the
    fourth insertion is refused exactly as the complete audit refuses it.
  - Derived indexes over their bound are dropped.
  - A byte-order mark keeps physical offsets along a chain.
  - At depth 3, a shape set and a shape add on a 200-shape slide are
    spliced by their own shape. At depth 2 the same set is audited
    completely.
  - Random chains: 2,000 rounds on a document at depth 2 and 2,000 on a
    slide at depth 3, where every accepted successor becomes the next base.
    Each run yields more than 100 splices, every derived index is compared
    in debug, and every verdict matches the complete audit.
- **`xml-minifier`, `tests/successor.rs`.** Updated for the new signatures.
  A successor of an indexed proof now keeps boundaries; one of a plain
  proof keeps none.
- **`litchi-opc`, `tests/source_xml_proofs.rs` (5 tests).**
  - A committed revision keeps its proof over its own allocation, at depth
    2.
  - A chain of two edits far apart, then an insertion, each publishes what
    a fresh package publishes for the same final bytes.
  - Seven edits of the chained revision, defects included, have a fresh
    package's verdict and error.
  - A revision keeps at most four replacement proofs, and a Part that gave
    its proof up still edits and publishes exactly.
  - A managed package keeps no proof.
  - Written overlays:
    - the slide is remembered at depth 3;
    - later writes (a shape add) publish a fresh package's bytes;
    - four defects are refused with a fresh package's error, twice each;
    - an invalid original is reported on every write and never remembered.
- **`litchi-docx`, `preserving_commits_prove_their_candidates_from_retained_proofs`.**
  - The first commit takes the indexed pair, and its proof has an index.
  - The physical owner's next base carries the remembered proof over its
    own allocation.
  - A committed revision's base shares the revision's allocation and proof.
  - A chain of commits far apart keeps a proof at every step.
  - Every chain writes the bytes of one fresh edit of the same paragraphs.

## 4. Measurements

### 4.1 `docx_large_bench` (µs per call; B A A B)

| leg | base 1 / 2 | after 1 / 2 |
|---|---:|---:|
| leaf text set | 763.8 / 762.4 | 758.8 / 760.4 |
| **unified set, default owner** | 409.4 / 407.2 | **23.4 / 23.3** |
| unified set, eager owner | 761.1 / 756.9 | 756.0 / 756.5 |
| **unified set, committed handle** | 416.0 / 412.0 | **25.4 / 25.1** |
| leaf `insert_paragraph` | 1,106.1 / 1,103.6 | 1,096.6 / 1,097.5 |
| unified add, default owner | 73.6 / 73.1 | 74.8 / 74.2 |
| unified add, eager owner | 424.1 / 421.1 | 420.9 / 421.2 |
| **unified add, committed handle** | 269.6 / 267.0 | **82.3 / 81.8** |
| open + add, default owner | 1,404.9 / 1,450.8 | 1,441.0 / 1,500.5¹ |
| open + add, eager owner | 1,461.7 / 1,518.4 | 1,531.6 / 1,459.9¹ |

¹ These legs open a handle on every call, and their samples vary by ±4 %
on this host. Their instruction counts move by +0.4 % and −0.1 % (§4.2).

### 4.2 Callgrind per call ((15 − 5 loops) / 10)

| leg | base Ir | after Ir | after L1i misses | after D1 read misses |
|---|---:|---:|---:|---:|
| leaf set | 17.08 M | 17.08 M | 2,588 | 23,064 |
| **default-owner set** | 10.00 M | **0.52 M** | 4,566 | 17,991 |
| **committed set** | 10.22 M | **0.71 M** | 3,233 | 25,209 |
| eager set | 17.06 M | 17.08 M | 2,897 | 23,097 |
| leaf add | 24.09 M | 24.09 M | 2,301 | 21,279 |
| eager add | 8.08 M | 8.09 M | 2,895 | 12,880 |
| default-owner add | 0.907 M | 0.950 M | 5,107 | 17,069 |
| **committed add** | 5.89 M | **1.37 M** | 4,012 | 31,675 |
| open + add, default owner | 31.57 M | 31.69 M | 10,291 | 54,173 |
| open + add, eager owner | 30.74 M | 30.71 M | 7,674 | 37,994 |

**Default-owner set.** The 9.1 M scan of the original that remained in 0313
is gone. The remaining 0.52 M is:
- the transaction's own work: projection, paragraph compaction and the
  `/body` resolution;
- the successor proof itself: the prefix and suffix comparisons, the
  audited paragraph and the derived index;
- the overlay's memo hit.

**Committed set.** The 0.19 M it costs over the default-owner set was not
broken down further.

**Committed add.** The insertion is now proved from the revision's own
derived boundaries. Its run is the inserted paragraph, not paragraphs 2–501.

### 4.3 `pptx_large_bench` (new; µs per call, `PPTX_LARGE_LOOPS=5`; B A A B)

| leg | base 1 / 2 | after 1 / 2 |
|---|---:|---:|
| leaf shape text set + write | 20,446 / 20,424 | 19,393 / 19,371 |
| unified shape text set + write | 20,335 / 20,474 | 19,227 / 19,385 |
| leaf shape add + write | 29,251 / 29,139 | 27,734 / 27,649 |
| unified shape add + write | 29,338 / 29,294 | 27,849 / 27,819 |
| unified set + write, committed handle | 16,834 / 16,720 | 15,470 / 15,436 |
| unified add + write, committed handle | 24,994 / 24,954 | 24,388 / 24,144 |
| write only, leaf / unified | 0.7 / 0.8 | 0.7–0.8 / 0.8 |

Callgrind ((3 − 1 loops) / 2):

| leg | base Ir | after Ir |
|---|---:|---:|
| leaf / unified set + write | 441.3 M / 441.7 M | 418.4 M / 418.1 M |
| leaf / unified add + write | 578.2 M / 578.9 M | 531.7 M / 531.5 M |
| committed set + write | 373.2 M | 338.5 M |
| committed add + write | 486.4 M | 463.4 M |

**What the base paid.** Each write paid a pair audit of the slide against
the physical original:
- about 23 M instructions for a set, where the element window proves the
  replacement;
- about 46 M for an add, which no element window covers, so the replacement
  was scanned completely.

**What remains.** After the change the remembered slide and its depth-3
index prove both from the edited shape, and the audit drops out of the
profile. The rest of each leg is litchi-pptx's edit, which this record
did not touch:
- `Scene::read` and `transition::validate_source` read the 1,000-shape scene
  several times per edit, which is most of the remaining ~420 M
  instructions;
- the scene is parsed again for the overlay's limits.

**Committed handles.** The committed legs gain less. A litchi-pptx overlay
is relative to the physical slide, so their two edits (shape 2, then shape
501) form one run of about half the slide (§5.2).

**Unified against leaf.** Unified/leaf stays within 0.991–1.006 before and
after, inside the single-op budget, because both legs write through the
same route.

**Cache counts.** The after builds show more simulated L1i misses on the
PPTX legs (0.61 M against 0.42 M per set call), and none on the DOCX legs.
The wall times fall with the instructions, so this is attributed to code
layout of the new example binary. It was not investigated further.

## 5. Residuals

### 5.1 The default-owner add pays for its derived index

The default-owner add now costs 43.6 K more instructions per call (+4.8 %)
and 1.0–1.6 % more wall time. That is the cost of deriving the inserted
revision's index: about 1,000 boundaries copied and shifted. The second
commit (`176d9db30d`) already cut it from 59 K to 44 K, with a signed shift
per field and no compaction when no parent changes.

The derivation is what makes the committed add 3.2× faster, so it is kept.
The add is still 0.18× the eager owner's and 0.068× the leaf's.

To remove the cost, the derivation could be made lazy, on the first
successor of the proof. That would need interior mutability in an immutable
proof, and is a possible follow-up.

### 5.2 PPTX overlays are relative to the physical slide

litchi-pptx keeps an edited slide as an overlay relative to the physical
part and audits it only when it writes. A handle committed twice therefore
writes one replacement that differs from the original in two places.

The fix is the one §2.2 made for `PreparedTopology`: litchi-pptx would keep
each overlay's proof and hand it to the writer as a predecessor. That needs
a write entry point that takes proofs, and is a follow-up. Overlays on the
topology route (after a structural edit) already go through
`prepare_topology_state` and benefit.

### 5.3 The first commit still scans the original twice

On a fresh default owner, the first commit's indexed pair audit scans the
original, and the overlay's `verify_indexed` scans it again to remember it.
The `fresh_auto_add` leg is unaffected because the add commits no proof.

Handing the pair's base proof to litchi-opc would remove the second scan.
That needs an apply entry point taking the original's proof as well, and is
a follow-up.

### 5.4 The PPTX edit itself dominates PPTX legs

About 95 % of a PPTX edit-and-write is the leaf's scene reads and
validation (§4.3). It is the same in both legs, and outside this record's
audit-proof scope.

## 6. Gates

Every command below ran on `176d9db30d` (the final code), one cargo command at a time.
The build environment and the memory-capped shim of the rules were used,
and test executables were pruned between chunks. Every exit status was
checked (`results/0317/gates.sh`).

| Command | Exit |
|---|---:|
| `cargo fmt --all --check` | 0 (and again after the trunk merge, §6.1) |
| `cargo clippy -p xml-minifier -p litchi-opc -p litchi-docx -p litchi-pptx -p litchi-ops -p litchi -p litchi-py --all-targets --all-features -- -D warnings` | 0 |
| `cargo clippy -p litchi-ops --no-default-features [--features F] --all-targets -- -D warnings`, F = none, docx, pptx | 0 each |
| `cargo test --all-features` of `xml-minifier` (189), `litchi-opc` (1,189), `litchi-docx` (2,401), `litchi-pptx` (1,803; consumer, unchanged) | 0 each |
| `cargo test -p litchi-ops --all-features`: `--lib` (922), `--doc` (210), the 138 test targets in seven groups of up to 20 (1,253), including `query_corpus_differential`, `corpus_fidelity` and `checklist_fixes` | 0 each |
| `cargo test -p litchi-ops --no-default-features --features` (every feature but `encryption`) `--test capabilities --test registry` (13; not re-blessed) | 0 |
| `cargo test -p litchi --all-features` (230) | 0 |
| `cargo test -p litchi-cli --all-features --test officecli_corpus` (not re-blessed) | 0 |
| `RUSTDOCFLAGS=-Dwarnings cargo doc --no-deps --all-features -p xml-minifier -p litchi-opc -p litchi-docx -p litchi-pptx` | 0 |
| `python3 tools/check_crate_boundaries.py` | 0 |
| `python3 tools/non_iwork_gate.py verify` | 0 |
| `python3 tools/test_no_3rdparty_references.py` | 0 |
| `RUSTC_BOOTSTRAP=1 python3 tools/check_api_glossary.py --target-dir target/glossary` (40 profiles, 2,306 bans, 228 families) | 0 |
| `crates/pyo3-litchi/tests/run_tests.sh` (279 tests) | 0 |
| `cd tools/perf-baseline && cargo check --locked --all-targets` | 0 |

No capability, registry or OfficeCLI-visible behaviour changed, so neither
the delta table nor the OfficeCLI corpus was regenerated.

The new public items are low-level audit and OPC API outside the unified
vocabulary, and no glossary row changes:
- `xml_minifier::audit::IndexDepth` (`CHILDREN`, `GRANDCHILDREN`, `get`);
- `VerifiedSource::index_depth`;
- the `depth` parameters of `verify_indexed` and
  `verify_indexed_replacement`;
- `litchi_opc::PreparedPartView::xml_audit_proof`;
- `SourceBackedPackage::xml_audit_proof`.

### 6.1 Trunk merge

Trunk `cbbdcd1fde` (record 0314) touches only iWork crates, the iWork
adapter and iWork benches, none of this record's files or measured paths.
After the merge, `cargo fmt --all --check` and
`cargo check -p litchi-ops --all-features --all-targets` exit 0.

## 7. Not claimed

- That a PPTX handle committed twice is proved edit by edit (§5.2).
- Any change to the PPTX edit cost outside the write audit (§5.4).
- That the default-owner add is unchanged: it is +4.8 % in instructions
  (§5.1).
- Any change of a frozen route name. `pptx_large_bench` is a new example,
  and `docx_large_bench` is unchanged.
