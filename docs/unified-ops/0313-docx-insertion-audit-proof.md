# 0313 — DOCX insertion audit proof: successor proofs from recorded content boundaries

- Status: implemented on `uo2/docx-insert`, cut from `feat/unified-ops-api`
  at `12724da2d5`; trunk `1100472b75` (record 0312) merged (`9c79d04592`;
  the merge touches no file this record changes and no measured path).
  Commits: `0c354bc8ef` (xml-minifier), `40dbab6d5d` (litchi-opc),
  `fd964e4103` (bench legs), `2eacd2e922` (this record). Trunk `75e5da7150`
  (rustfmt of a 0312 test, `PROGRESS.md`) was merged last (`6f23858452`).
- Date: 2026-10-16
- Input: record 0311 §4.1. A DOCX paragraph `add` on the default
  source-backed owner carried no audit proof, so the source overlay audited
  the whole replacement main part, 9.1 M of 9.98 M instructions per call.
  That left the add at 1.05–1.22× the eager owner on the 1,000-paragraph
  `docx_large_bench`.
- Targets: the default-owner add within 1.05× the eager owner and ≤ 1.10×
  the leaf; every 0311 DOCX leg re-measured; output byte-identical.
- Scope: `xml-minifier` (`audit.rs`, new `audit/successor.rs`, one helper in
  `audit/namespaces.rs`, new `tests/successor.rs`) and `litchi-opc`
  (`source_backed.rs`, `source_backed/revision.rs`,
  `tests/prepared_topology.rs`), and two new legs in `litchi-ops`
  `examples/docx_large_bench.rs` (`fresh_auto_add`, `fresh_eager_add`: open
  a handle and add, per call). No `litchi-docx` or `litchi-ops` source
  changed.
- `performance_claim`: §3. `profiling` builds (release with debug info,
  stripped for timing) of
  `cargo build --profile profiling -p litchi-ops --features docx,doc,odp,xls,rtf,odt,pptx,ppt --example docx_large_bench`,
  pinned with `taskset -c 30`, processes B A A B. The base build is this
  branch with `crates/{xml-minifier,litchi-opc}/src` reverted to
  `12724da2d5`, so it has the new legs. Instruction and cache counts
  come from callgrind with `--cache-sim=yes --I1=32768,8,64 --D1=32768,8,64`
  (`perf` is blocked, `perf_event_paranoid=4`). Evidence:
  [results/0313/](results/0313/): `measure.sh`, `docx_cg.sh`,
  `summary.txt`, the raw harness outputs, `docx-cachesim.txt` and
  `binaries.sha256`. Host: a shared 32-core EPYC with load 4–6.

## 1. Summary

| Item | Before | After | Target | Status |
|---|---:|---:|---:|---|
| DOCX 1,000 ¶ add, default owner / eager owner's add | 1.037 / 1.029 | **0.177 / 0.174** | ≤ 1.05 | **met** |
| the same add / leaf `insert_paragraph` | 0.399 / 0.398 | **0.067 / 0.066** | ≤ 1.10 | **met** |
| default-owner add, instructions per call | 9.99 M | **0.91 M** | | |
| first add on a freshly opened default owner / eager (new legs) | 1.188 / 1.183 | 0.942 / 1.012 | | improved; 1.029× in instructions (§3.2) |
| add on a handle already committed from the default owner | 454 µs | 269–272 µs | | improved (§4.1) |
| every other 0311 DOCX leg | | within 1 % wall, 0.2 % instructions | | unchanged |

Output stays byte-identical. The query-corpus differential, corpus fidelity,
the checklist goldens, the capability and registry goldens and the
OfficeCLI corpus pass without re-blessing (§5).

## 2. Design

### 2.1 Choice

Record 0311 offered three designs. The decision turned on what each one
must know about the parser state at the insertion point.

- **(a) A proof from an audit the commit already runs.** For an insertion
  there is none. The preserved-markup commit audits nothing, and the
  transaction's snapshot is litchi-docx's own scanner, which cannot mint an
  `xml-minifier` proof. Minting one in the commit costs the same full
  audit, only moved earlier. Rejected.
- **(c) Defer the overlay audit to save.** `apply_xml_bytes` documents that
  it refuses invalid XML. A `PreparedTopology` also hands its effective
  bytes to callers through `with_candidate` and `main_document_part()`,
  and those callers can write them out without passing through a
  litchi-opc writer. A deferred audit would therefore need a new contract
  and a check on every egress route. Missing one route would publish
  unaudited bytes silently. Rejected as the larger and less sound change.
- **(b) An incremental audit in `xml-minifier`.** This was chosen. Its
  soundness argument is the one element windows already rely on: the
  documented invariant at `audit::scan`, which every source-policy check
  must keep. That invariant and the window code (`audit_window`,
  `State::within`) already exist, and debug builds re-derive every window
  proof from a complete audit. What a window over an insertion needs, and
  0311 lacked, is the parser state at the insertion point without a
  second scan of the original. So the audit that first passes the original
  now records that state at its shallow content boundaries.

### 2.2 `xml-minifier`: successor proofs (`0c354bc8ef`)

- `VerifiedSource::verify_indexed(bytes, limits)` runs `verify_source` with
  an observer, `successor::BoundaryRecorder`. Observers only read the
  auditor's state, so the verdict and error are unchanged. The recorder
  keeps each *content boundary* of the document element and of its
  children: every offset between two accepted tokens at which one or two
  elements are open. A boundary stores the following, as `u32` values:
  - its offset;
  - its parent element;
  - the aggregate counters charged up to it.

  Each parent keeps a snapshot of the namespace bindings in scope in its
  content. A child that declares no binding of its own shares the
  document element's snapshot.
- `VerifiedSource::verify_successor(&self, replacement, limits)` returns
  exactly `verify_source(replacement, limits)`'s verdict and error. It
  takes the shortcut only when every condition below holds; anything else
  is audited completely:
  - The proof was made under the same `limits` and has an index.
  - The replacement repeats the original outside a run between two
    boundaries of the same parent: the last boundary at or before the
    common prefix, and the first at or after the common tail. A pure
    insertion is an empty run. The common prefix of an insertion usually
    runs into the next sibling, so the tail is measured again from the
    run's start. When the bytes that took that smaller run's place do not
    begin and end with markup, the first measure is tried too.
  - The bytes that took the run's place are at most half the replacement.
    They pass `audit_window`: they start from the parent's depth and
    bindings, begin with `<`, end with markup, are balanced, and close
    nothing outside themselves.
  - The re-totalled aggregate budgets (`events`, `attributes`,
    `text_bytes`) and `bytes` fit `limits`. The depth limit is checked
    inside the window at absolute depth.

  The result is `ReplacementProof::Splice { original, replacement }`, a new
  variant of the `#[non_exhaustive]` enum. Debug builds confirm every
  splice with a complete audit.
- Why the bytes outside the run need no new scan:
  - Both edges are token boundaries of the original, and the new bytes
    begin with `<` and end with markup. So the last token before the run
    and the first token after it end and start where they did, and text
    never merges across an edge. A pure removal would merge text, so an
    empty window is never a shortcut. A test shows that removal can create
    a `]]>`.
  - The bytes after the run see the same open elements and the same
    bindings, because the window is balanced and closes every binding it
    opens.
  - The index is recorded only by the audit itself, from the proof's own
    bytes. The successor is compared against those bytes. No hint from the
    caller is trusted.
- `VerifiedSource::verify_indexed_replacement(original, replacement,
  limits)` is the pair audit `verify_replacement` with the original's
  indexed proof. It makes one scan of the original, observed by both the
  element-window search and the recorder. The element window is tried
  first and the run second, and the verdict and error are those of
  `verify_source_replacement`.
- Bounds:
  - at most one boundary per 32 input bytes, plus 64, at 20 bytes each;
  - namespace snapshots within the 8 MiB budget of the multi-window path;
  - inputs shorter than `u32::MAX` bytes.

  Past any bound, no index is kept and every successor is audited
  completely. The 1,000-paragraph document has 1,004 boundaries, about
  20 KB.
- `WindowSearch`'s observer is now `#[inline(always)]`. Pairing it with
  the recorder had made LLVM outline it, which cost every pair audit
  about 0.9 % (+155 K instructions on `leaf_set`). With the attribute the
  set legs are back to base (§3.2).

### 2.3 `litchi-opc`: the remembered original keeps its proof (`40dbab6d5d`)

`SourceBackedPackage::audit_physical_replacement_xml` (0311) remembered a
passing original per Part and source revision. On an unmanaged package the
memo now also keeps the original's indexed proof. The proof is built by
the original's first audit, which is the same single scan as before:

- `verify_indexed` when the replacement came with its own proof;
- `verify_indexed_replacement` on the pair route.

An unproven replacement of a remembered original is audited with
`verify_successor`. Before this change it got a full `VerifiedSource::verify`.
The verdict and error are the replacement's complete source audit in every
case. So 0311's contract holds unchanged: the error order of
`apply_xml_bytes`, and an original that fails being audited and reported
every time.

- The proof keeps the original's allocation alive. That allocation is the
  part cache's own payload, and the DOCX owner's snapshot memo pins it
  anyway. At most four remembered originals per package
  (`RETAINED_SOURCE_XML_PROOFS`) keep a proof; an older one keeps only its
  verdict. Managed packages remember nothing, as before.
- `audit_physical_replacement_xml` now takes the original's `PartData`,
  so it can share the allocation.

## 3. Measurements

### 3.1 `docx_large_bench` (µs per call; B A A B)

| leg | base 1 / 2 | after 1 / 2 |
|---|---:|---:|
| leaf text set | 762.1 / 760.8 | 757.7 / 761.0 |
| unified set, default owner | 410.3 / 410.8 | 411.4 / 412.1 |
| unified set, eager owner | 764.3 / 760.9 | 761.8 / 793.5¹ |
| unified set, committed handle | 417.8 / 416.6 | 415.9 / 417.6 |
| leaf `insert_paragraph` | 1,108.0 / 1,108.3 | 1,111.7 / 1,103.9 |
| **unified add, default owner** | 442.2 / 441.3 | **74.3 / 73.4** |
| unified add, eager owner | 426.5 / 429.0 | 420.0 / 422.8 |
| unified add, committed handle | 453.6 / 453.8 | 272.4 / 268.8 |
| open + add, default owner (new) | 1,809.9 / 1,851.4 | 1,454.5 / 1,481.5 |
| open + add, eager owner (new) | 1,523.2 / 1,565.2 | 1,543.7 / 1,463.9 |

¹ One noisy sample; its instruction count equals the base's (§3.2).

### 3.2 Callgrind per call ((15 − 5 loops) / 10, `docx-cachesim.txt`)

| leg | base Ir | after Ir | after L1i misses | after D1 read misses |
|---|---:|---:|---:|---:|
| leaf set | 17.06 M | 17.07 M | 2,567 | 23,047 |
| default-owner set | 9.99 M | 10.00 M | 4,839 | 22,432 |
| committed set | 10.22 M | 10.22 M | 3,447 | 41,682 |
| eager set | 17.07 M | 17.06 M | 2,915 | 23,111 |
| leaf add | 24.08 M | 24.09 M | 2,299 | 21,322 |
| eager add | 8.09 M | 8.09 M | 2,925 | 12,889 |
| **default-owner add** | 9.99 M | **0.91 M** | 5,120 | 16,741 |
| committed add | 10.42 M | 5.88 M | 4,068 | 44,480 |
| open + add, default owner | 40.30 M | 31.59 M | 10,255 | 53,704 |
| open + add, eager owner | 30.72 M | 30.71 M | 7,630 | 37,995 |

The open + add legs take the pair route: the original has not been
audited yet. The base scanned the original and then the whole replacement,
since no element window covers an insertion. After the change, the one
scan of the original also records its boundaries, and the insertion is
proved from them: 8.7 M fewer instructions. The remaining 0.9 M over the
eager owner is that scan of the original, which the eager owner defers to
its writer (0311 §4.1).

In the default-owner add, the successor proof costs about 44 K
instructions per call: the prefix and suffix comparisons, a namespace
snapshot clone, and the audit of the inserted paragraph. The rest is the
commit itself and the `/body` resolution, which the eager owner pays too.
The eager owner's add costs more because it scans its transaction snapshot
on every edit (about 7.9 M), which the source-backed owner serves from the
0305 memo.

## 4. Residuals

### 4.1 Chained edits on a committed handle

A handle committed from the default owner edits its effective revision,
while the remembered proof is the physical original's. The `committed_add`
leg inserts at paragraph 501 into a revision that changed paragraph 2. Its
run spans paragraphs 2–501, which is about half the document and just
under the half-replacement bound, so that half is scanned (5.88 M against
10.42 M). Two edits further apart than half the document fall back to the
complete audit, as before. Closing this needs two more pieces. A successor
proof would have to carry a derived index: the base's entries shifted
past the run, plus the window's own boundaries. And the overlay would have
to keep the effective Part's proof as the base. Both are follow-ups. Sets
take the commit's own window proof (0311) and are unaffected.

### 4.2 Boundaries are recorded only at depth ≤ 2

The recorded boundaries cover DOCX body blocks and XLSX rows (children of
`sheetData`). Edits inside a deeper structure take as the run their
depth-2 ancestor. That is one paragraph or one table in DOCX, but most of a
slide in PPTX (`sld/cSld/spTree/sp`), and a run over half the payload is
audited completely. So nothing gets slower than before, but those formats
gain nothing yet.

### 4.3 The set path still scans the original once

A text set commits through litchi-docx's own pair audit (0311), which scans
the original once per call (9.1 M of the 9.99 M). The remembered indexed
proof could prove that replacement too, but litchi-docx's commit does not
have it. Handing it over is a follow-up that could remove most of the
set's remaining 9.1 M.

## 5. Gates

Every command below ran on the merged tree (`9c79d04592`). The build
environment and the memory-capped cargo shim of the rules were used, one
cargo command ran at a time, and test executables were pruned between
chunks. Every exit status was checked (`gates.sh`, outside the
repository).

| Command | Exit |
|---|---:|
| `cargo fmt --all --check` | 0 after trunk `75e5da7150` was merged (`6f23858452`). On `9c79d04592` it failed only on `crates/litchi-ops/tests/iwork_unified.rs:160`, from record 0312, which trunk then reformatted |
| `cargo clippy -p xml-minifier -p litchi-opc -p litchi-docx -p litchi-ops -p litchi -p litchi-py --all-targets --all-features -- -D warnings` | 0 |
| `cargo clippy -p litchi-ops --no-default-features [--features docx] --all-targets -- -D warnings` (none, docx; again after the `fresh_*` legs) | 0 each |
| `cargo test -p xml-minifier --all-features` (180) | 0 |
| `cargo test -p litchi-opc --all-features` (1,184) | 0 |
| `cargo test -p litchi-docx --all-features` (2,400; consumer, unchanged) | 0 |
| `cargo test -p litchi-ops --all-features`: `--lib` (920), `--doc` (210), the 137 test targets in seven groups of up to 20 (1,252), including `query_corpus_differential`, `corpus_fidelity` and `checklist_fixes` | 0 each |
| `cargo test -p litchi-ops --no-default-features --features` (every feature but `encryption`) `--test capabilities --test registry` (not re-blessed) | 0 |
| `cargo test -p litchi --all-features` (230) | 0 |
| `cargo test -p litchi-cli --all-features --test officecli_corpus` (not re-blessed) | 0 |
| `RUSTDOCFLAGS=-Dwarnings cargo doc --no-deps --all-features -p xml-minifier -p litchi-opc` | 0 |
| `python3 tools/check_crate_boundaries.py` | 0 |
| `python3 tools/non_iwork_gate.py verify` | 0 |
| `python3 tools/test_no_3rdparty_references.py` | 0 |
| `RUSTC_BOOTSTRAP=1 python3 tools/check_api_glossary.py --target-dir target/glossary` (40 profiles, 2,306 bans, 228 families) | 0 |
| `crates/pyo3-litchi/tests/run_tests.sh` (279 tests) | 0 |
| `cd tools/perf-baseline && cargo check --locked --all-targets` | 0 |

No capability, registry or OfficeCLI-visible behaviour changed, so neither
the delta table nor the OfficeCLI corpus was regenerated. The new public
items are `xml_minifier::audit::VerifiedSource::{verify_indexed,
verify_indexed_replacement, verify_successor}` and the
`ReplacementProof::Splice` variant. They are low-level audit API outside
the unified vocabulary, the glossary gate passes, and no glossary row
changes.

## 6. Not claimed

- That a chained add on a committed handle matches the eager owner's add
  (§4.1; it is 0.64–0.65×, but only because its two edits are within half
  the document).
- Any gain for edits whose depth-2 ancestor exceeds half the payload
  (§4.2).
- Any change of a frozen route name. `docx_large_bench` gained the
  `fresh_*` legs only.
