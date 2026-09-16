# Change 0649 evidence packet

Attribution of change [0638](../../0638-facade-and-ordinary-save-selectors.md)'s
**133.61 ms** `pptx_real_file_ordinary_save_edit` row: one
`opened_presentation_transaction().set_shape_text(..)` plus
`apply_opened_presentation_commit(..)` on the real 108 KB, 103-member deck
`test-data/libreoffice-core/sd/qa/unit/data/pptx/slide-section-test.pptx`.

Record: [`../../0649-pptx-opened-transaction-real-deck-edit.md`](../../0649-pptx-opened-transaction-real-deck-edit.md).

**No file under `crates/` was modified by this change.** The instrumentation that
produced the deterministic count tables is a measurement-only patch, retained
here and reverted before the gates and the commit.

## Provenance

| | |
| --- | --- |
| base commit | `c7326f68065edf6f2198ca3cb39c38c48cf00ed9` |
| branch | `perf/0649-pptx-opened-transaction-real-deck-edit` |
| host | AMD EPYC 9R45, 32 cores, 123 GiB, Linux 7.0.0-1012-aws |
| toolchain | rustc 1.95.0, cargo 1.95.0, valgrind 3.26.0, perf 7.0.14 |
| pins | `taskset -c 16` for counts, allocations, `perf stat`, the dwarf `perf record` and the harness legs; `-c 17` for the whole-edit `perf record`; `-c 18` for callgrind |
| concurrency | seven other agents built and measured on the other cores throughout |

Binary SHA-256, every one staged outside its Cargo target directory before it ran
(change 0627's lesson):

| binary | sha256 |
| --- | --- |
| `probe0649`, clean leg (`--release`, `debug = 1`) | `9a36d528d2ab7b3f7bd9474972ecb3bbb49558d8b4efd12a63489b74d2b26fdb` |
| `probe0649`, instrumented leg (`--features counts`) | `b23aab8a5b180033cb02395f66b549ba7413eab2487e1bb85583866d37471bdb` |
| `probe0649`, clean leg with the counting allocator | `c364e11478053d4666a038dce604c0cf10931458eea1e29bb2b0eb8dd1101ae6` |
| `probe0649`, instrumented leg with change 0637 applied | `075524837370d6eeabcb97150e8524b0653df27709592338419afa296636bd4e` |
| `litchi-perf-baseline` (base + 0638's harness diff) | `e9b627d76f3f9f0cbdc869f12f4a85a9086b6efecdfaed47521a11c3cf00c1e8` |
| `litchi-perf-baseline-alloc` (base + 0638's harness diff) | `4c1032716cdcbcb875fd7093ad45453b051bffda28e415bf1e14d6a89a13c5a2` |

The base predates change 0638, whose selectors this record ties into, so 0638's
harness-only diff (commit `15363b3a8`, touching only `tools/`) was applied to
build the two harness binaries and reverted afterwards. It touches no file under
`crates/`, so the crate code measured through it is the base's.

## The three decks

| | real | generated | control |
| --- | --- | --- | --- |
| what | the 0638 fixture | the shape of `build_semantic_pptx_corpus(Medium)`, 12 slides × 8 text boxes, rebuilt by the probe | the real deck with every occurrence of the MCE namespace URI replaced by an equal-length URI the codec does not recognize |
| archive bytes | 108,164 | 40,797 (harness corpus 40,788) | 106,536 |
| members | 103 | 61 | 103 |
| uncompressed bytes | 796,725 | 142,110 | 796,725 |
| members mentioning the MCE namespace | 43 (93.0% of bytes) | 0 | 0 |

The control is a control for the code path and **not** a proposal: it is not a
semantically valid transformation of a `.pptx`, it exists only inside the
measurement, and it is not retained as a fixture. `scripts/run.sh` rebuilds it in
six lines of Python.

## Contents

| path | what it is |
| --- | --- |
| `summary.txt` | every derived figure the record quotes, recomputed from the raw files below |
| `decision.json` | the decision, its evidence, its accepted costs and its known gaps |
| `gates.txt` | the five gates and their tails; `cargo test -p litchi-pptx` is 871 tests across 77 suites |
| `log-sections.md` | the four program-log paragraphs the coordinator merges |
| `probe/src/main.rs` | the probe: `shape`, `target`, `phases`, `prefix`, `counts`, `allocations`, `dump` |
| `probe/Cargo.toml.example` | its manifest; substitute `<checkout>` and rename to `Cargo.toml` outside the repository |
| `instrumentation/instrumentation.patch` | the measurement-only counter patch, reverted before the gates |
| `instrumentation/perf0649.rs` | the counter module it adds at `crates/litchi-ooxml-common/src/perf0649.rs` |
| `counts/counts-real-deck.tsv` | deterministic counters per prefix stage, real deck |
| `counts/counts-generated-corpus.tsv` | the same, generated corpus |
| `counts/counts-marker-stripped-control.tsv` | the same, control — structurally identical to the real deck, 0 rewrites |
| `counts/counts-real-deck-with-0637-applied.tsv` | the same, real deck, with change 0637's crate diff applied: every counter identical |
| `counts/mce-trace-one-capture-real-deck.txt` | every MCE pass inside one capture, identified by input length against the archive members: 3 per slide, 5 on `ppt/presentation.xml` |
| `counts/allocations-real-deck.tsv` | allocations, allocated bytes and reallocations per prefix stage |
| `counts/allocations-generated-corpus.tsv` | the same, generated corpus |
| `timing/phases-{real,generated,control}-R{1,2,3}.tsv` | p50/mean/p95/min per documented call, 30 samples, three repeats — the A/A floor |
| `timing/perfstat-isolation-pairs-real-deck.txt` | `perf stat` cycles and instructions at N=2 and N=10 for all six prefix stages |
| `native/perf-capture-self-symbols.txt` | self-cycle shares of the capture stage; MCE codec plus its memmove is 80.69%, hardware SHA-256 is 0.51% |
| `native/perf-edit-self-symbols.txt` | the same for the whole edit; 79.68% and 0.72% |
| `native/callgrind-capture-inclusive.txt` | callgrind inclusive `Ir`; `process_ooxml` 84.17%, `package_fingerprint` 3.24% under *software* SHA-256 |
| `harness/0638-selectors-at-base.json` | 0638's own selectors run at this base, 5 warm-ups and 20 samples |
| `harness/0638-selectors-allocator-binary.json` | the same through `litchi-perf-baseline-alloc`, which reports no allocation metrics for that family |
| `scripts/run.sh` | reproduces every probe leg |
| `scripts/harness-repro.sh` | reproduces the harness leg, including how 0638's diff is applied |

## The result in four numbers

| | |
| --- | ---: |
| the capture plus the commit's recapture, as a share of the edit | **95.9%** wall clock, **96.44%** cycles |
| MCE output produced and thrown away per edit | **28,064,911 bytes** — 259× the source archive |
| removed by the marker-stripped control | **120.301 ms of 128.123 ms — 93.9%** |
| the complete-package revision proof, natively | **0.51%** of the capture's cycles |

## Reproducing

```sh
# 1. counts (instrumented leg)
git -C <checkout> apply instrumentation/instrumentation.patch
cp instrumentation/perf0649.rs <checkout>/crates/litchi-ooxml-common/src/
cargo build --release --features counts --manifest-path <probe>/Cargo.toml

# 2. everything else (clean leg)
git -C <checkout> checkout -- crates/
rm <checkout>/crates/litchi-ooxml-common/src/perf0649.rs
scripts/run.sh <checkout> <probe> <out> 16
```

`scripts/run.sh` carries the full sequence, including the marker-stripped
control's construction and the callgrind isolation pair. `prefix ... capture`
deliberately does not derive an edit target, so the callgrind profile is the
captures and nothing else.
