# Evidence packet — change 0631

Relationship hash order decided three OOXML outcomes a caller can see — an XLSX
`Error::PatchConflict` on an exact no-op, a public PPTX `Snapshot::revision()`,
and the DOCX content-control signature staleness token — and all three are now
functions of the document. Record:
[`docs/performance/0631-ooxml-relationship-order-verdict-sites.md`](../../0631-ooxml-relationship-order-verdict-sites.md).
`performance_claim: none`.

## Provenance

| field | value |
| --- | --- |
| base commit | `732cf44ae` (`feat/office-format-completeness`) |
| branch | `perf/0631-ooxml-relationship-order-verdict-sites` |
| worktree | `/home/zhuhe/code/litchi-worktrees/0631` (removed after commit) |
| host | AMD EPYC 9R45, 32 cores, Linux 7.0.0-1012-aws |
| toolchain | rustc 1.95.0 (59807616e 2026-04-14), cargo 1.95.0 (f2d3ce0bd 2026-03-21), valgrind 3.26.0 |
| build | `cargo build --release` for every probe, both legs, same worktree |
| CPU affinity | every measured process pinned to CPU 24 with `taskset` |
| host state | eight agents building and testing concurrently throughout |

The two legs are the same worktree with the three changed source files reverted
(before) or applied (after); nothing else differed, and the added test files are
present on both legs so their before-leg failures are part of the evidence.
Probe binary SHA-256:

| binary | before | after |
| --- | --- | --- |
| `corpus_verdicts` | `94fe298c02f6eb51…0e7caaed` | `43f94c4411c2fa6e…3946a30f5` |
| `open_edit_save_iters` | `ee049e1e69f6ab14…305cf396` | `b93a5af1488b018f…9e32cbedd9` |

## Contents

| path | what it is |
| --- | --- |
| `gates.txt` | the ten gates — `cargo fmt --all --check` plus clippy, doc and test for each of the three crates — each with its tail and exit status, followed by the corpus and callgrind summaries the record cites |
| `differential/xlsx-determinism-before.txt` | `cargo test -p litchi-xlsx --test relationship_order_verdict_determinism` on the unfixed tree: 2 pass, 1 fails with `PatchConflict` on an exact no-op |
| `differential/pptx-determinism-before.txt` | the same for PPTX on the unfixed tree: 2 pass, 2 fail — two distinct public revisions from one graph, and *"ActiveX control patch source is stale"* |
| `differential/pptx-stage-b-binary-capture-sorted.txt` | the PPTX test with only `load_binary`'s array sorted: the revision is stable and the next unsorted use fires — *"ActiveX binary part relationships are stale"* |
| `differential/pptx-stage-c-ensure-binary-sorted.txt` | stage B plus `ensure_binary_part`'s array: the third use fires — *"ActiveX descriptor relationship lifecycle does not match the patch target"* |
| `differential/docx-determinism-before.txt` | the same for DOCX on the unfixed tree: 2 pass, 1 fails with *"package signature topology is stale"* on an exact no-op |
| `differential/determinism-after.txt` | all three test binaries on the fixed tree: 10 of 10 pass |
| `differential/corpus-verdicts-before.txt`, `-after.txt` | `corpus_verdicts` over 321 OOXML fixtures, 8 repeats each: per fixture, an OPC republication SHA-256 plus the set of distinct verdicts and messages for that family's repaired path |
| `differential/corpus-verdicts-verdict-column-diff.txt` | the two runs diffed on the verdict columns: one fixture line (`hello-world-signed-twice.docx`, 2 verdicts → 1) and the two summary lines that count it |
| `tests/suite-results-before-after.txt` | `cargo test -p <crate> --no-fail-fast` on both legs, reduced to one line per test binary and diffed: 189 binaries, 3 lines differ, all three the new determinism binaries |
| `callgrind/{before,after}-{xlsx,pptx,docx}-{1,11}.txt` | valgrind summaries for the N = 1 and N = 11 isolation pairs of `open_edit_save_iters`, one fixture per crate |
| `callgrind/annot-{before,after}-*-11.txt` | `callgrind_annotate --threshold=99.5` self-cost tables for the N = 11 runs |
| `callgrind/symbol-deltas.txt` | their per-symbol difference, which is what the record's attribution reads |
| `probe/` | the complete source of both probe binaries with `Cargo.toml` (the generated `Cargo.lock` is gitignored and not retained). Path dependencies point at the 0631 worktree; retarget them to reproduce. |
| `decision.json` | the machine-readable decision |
| `log-sections.md` | the four log paragraphs for HOTSPOTS, GOAL_AUDIT, REPORT and ADR_COMPLIANCE, for the coordinator to merge |

## Headline numbers

| measurement | before | after | tier |
| --- | --- | --- | --- |
| determinism tests passing (10 in 3 files, 128 repeats per assertion) | 6 / 10 | **10 / 10** | measured |
| corpus fixtures whose repaired-path verdict was unstable over 8 repeats (321 probed) | **1** | **0** | measured |
| published packages byte-identical across the legs (of 312 published, 321 attempted) | — | 312 / 312 | measured |
| crate-suite test binaries whose result changed (189 binaries over 3 crates) | — | 3, all new | measured |
| instructions per open + edit + save, XLSX synthetic closure | 5,125,263.1 | 5,126,735.2 | measured |
| … PPTX `activex_checkbox.pptx` | 12,044,208.5 | 12,052,424.9 | measured |
| … DOCX `NumberedList.docx` | 53,541,460.9 | 53,509,764.0 | measured |
| code-layout noise floor for these counters (3 builds of identical XLSX probe code, N = 11) | 16,429 Ir spread ≈ ±1,600 per round | | measured |

The XLSX and PPTX deltas (+0.029%, +0.068%) sit at or under that layout floor;
the per-symbol tables isolate the genuinely new work as
`insertion_sort_shift_left` plus one `SmallVec::extend`, about 293 Ir per round
on the PPTX leg. The DOCX delta is negative (−0.059%) and the symbols name why:
the `relationship_count != 0` guard skips a second pass over the root
relationships — case-insensitive substring scans and a `PackURI` construction
per internal relationship — for every unsigned package. `performance_claim:
none`: that is reported as a cost that came out negative, on one fixture, not
registered as a claim.

## Reproducing

```sh
# the ten determinism tests
taskset -c 24 cargo test -p litchi-xlsx --test relationship_order_verdict_determinism
taskset -c 24 cargo test -p litchi-pptx --test pptx_activex_relationship_order_determinism
taskset -c 24 cargo test -p litchi-docx --test signature_token_relationship_order

# the probes (retarget probe/Cargo.toml's path dependencies first)
CARGO_TARGET_DIR=/some/disk/path cargo build --release --manifest-path probe/Cargo.toml
cd /path/to/litchi
taskset -c 24 .../corpus_verdicts test-data 8
for N in 1 11; do
  taskset -c 24 valgrind --tool=callgrind --callgrind-out-file=cg-$N.out \
    .../open_edit_save_iters xlsx synthetic $N
  taskset -c 24 valgrind --tool=callgrind --callgrind-out-file=cg-p-$N.out \
    .../open_edit_save_iters pptx test-data/ooxml/pptx/activex/activex_checkbox.pptx $N
  taskset -c 24 valgrind --tool=callgrind --callgrind-out-file=cg-d-$N.out \
    .../open_edit_save_iters docx \
      test-data/libreoffice-core/sw/qa/extras/ooxmlexport/data/NumberedList.docx $N
done
```

For the before leg, revert the three changed source files and rebuild; nothing
else changes. The PPTX staging evidence is reproduced by sorting one call site
at a time, in the order the record's stage table gives.
