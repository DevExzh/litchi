# 0649: the 133.61 ms real-deck shape-text edit is six whole-slide MCE rewrites, not the revision proof — a marker-stripped control removes 93.9% of it

Status: retained, attribution and frozen design. **No file under `crates/` was
modified.** `performance_claim: none` — the counts, native cycles and paired
medians below are reported as evidence, not registered as a claim.

OLE2 and OOXML remain the active priority. ODF optimization stays deferred until
that goal completes; iWork is excluded.

Change [0638](0638-facade-and-ordinary-save-selectors.md) measured one
`opened_presentation_transaction().set_shape_text(..)` plus
`apply_opened_presentation_commit(..)` on the real 108 KB, 103-member deck
`test-data/libreoffice-core/sd/qa/unit/data/pptx/slide-section-test.pptx` at
**133.61 ms** — 95.7% of that deck's open-edit-save lifecycle, 490× its complete
serialization, 68.7× the same edit on the generated 61-member corpus — and said
in as many words: *"Nothing in this record explains it."* This record explains
it.

## What was measured

The edit phase 0638 times is four documented public calls, so the phase boundary
is an API boundary and no crate source has to move to decompose it. A retained
scratch probe (`results/change-0649/probe/`) times each call, counts each stage's
work through a retained, reverted instrumentation patch
(`results/change-0649/instrumentation/`), and runs the three decks the
attribution needs:

* **real** — the 0638 fixture, 103 members, 796,725 uncompressed bytes;
* **generated** — the shape of the harness's `build_semantic_pptx_corpus(Medium)`
  (12 slides × 8 text boxes), reproduced by the probe at 61 members and 40,797
  bytes against the harness corpus's 61 and 40,788;
* **control** — the real deck with every occurrence of the MCE namespace URI
  replaced by an equal-length URI the codec does not recognize. Member count,
  element count, attribute count, part-decode count and uncompressed byte total
  are all **identical** to the real deck; the only thing that differs is which
  branch of `process_markup_compatibility` each part takes. Change 0588 used the
  same technique and called it a marker-stripped control.

### The answer, in one line

`litchi-pptx` asks `litchi_ooxml_common::mce` to rewrite **every slide of the
deck, three times per capture, twice per edit** — six whole-slide markup-
compatibility rewrites per edit — and change
[0588](0588-mce-codec-namespace-emission.md)'s frozen namespace emission makes
each of those rewrites produce **16.3× its input**. The complete-package revision
proof — the mechanism this record was asked to rule in or out — is **0.51% of the
capture's cycles**.

## Why it is sound

Nothing under `crates/` changed, so no invariant, error identity, limit, defence
or output byte could move. Every call the probe makes is a documented public
entry point of `litchi-pptx`.

The instrumentation leg is described below and in the packet: it adds
twenty-five `AtomicU64` counters and one `Mutex<Vec<u64>>` trace behind
twenty-six `bump`/`trace` calls in seven files of `litchi-ooxml-common` and
`litchi-pptx`, is used **only** for the deterministic count tables, is not used
for any timing, cycle or instruction number, and was reverted
(`git checkout -- crates/`) before the gates and the commit. The clean-leg binary
`probe0649-base` carries SHA-256
`9a36d528d2ab7b3f7bd9474972ecb3bbb49558d8b4efd12a63489b74d2b26fdb`; the
instrumented binary is `b23aab8a5b180033cb02395f66b549ba7413eab2487e1bb85583866d37471bdb`.

## Measured

Host: AMD EPYC 9R45, 32 cores, 123 GiB, Linux 7.0.0-1012-aws, rustc 1.95.0,
valgrind 3.26.0. Base c7326f680. Every measured process pinned with `taskset`
(CPU 16 unless stated) while seven other agents built and measured on the other
cores. Binaries staged outside every Cargo target directory (0627's lesson).

### Phase decomposition, and the control

p50 of 30 samples, median of three repeats; `real A/A` is the widest p50 spread
across those three repeats in the same window.

| phase | real | generated | control | real A/A | share of the real edit |
| --- | ---: | ---: | ---: | ---: | ---: |
| `opened_presentation()` — the capture | **59.900 ms** | 0.775 ms | 3.375 ms | 4.67% | 46.8% |
| `Snapshot::edit()` — the working clone | 0.020 ms | 0.012 ms | 0.015 ms | 28.63% | 0.0% |
| `set_shape_text` | 4.844 ms | 0.137 ms | 0.547 ms | 2.89% | 3.8% |
| `Transaction::commit` | **62.928 ms** | 0.923 ms | 3.817 ms | 3.96% | 49.1% |
| `apply_opened_presentation_commit` | 0.219 ms | 0.047 ms | 0.076 ms | 8.06% | 0.2% |
| **edit total** | **128.123 ms** | **1.891 ms** | **7.822 ms** | **4.31%** | |

The A/A floor is 4.31% on the edit total and 2.89% to 4.67% on the three phases
above 1 ms. The two sub-millisecond phases have 8.06% and 28.63% floors; they are
reported and **not** relied on. 0638's own A/A spread on this row was 3.02%.
128.123 ms against 0638's 133.61 ms is 4.1% apart, inside both floors: the same
fact on a different driver. The real/generated ratio is **67.7×** against 0638's
68.7×.

**The capture and the commit's recapture are 95.9% of the edit.** Everything
else — the working clone, the one slide `set_shape_text` rewrites, and the
publication — is 4.1%.

**The marker-stripped control removes 120.301 ms of the 128.123 ms — 93.9%.** The
deck it runs on has the same members, the same elements and the same byte total;
it differs only in taking `process_markup_compatibility`'s borrowed fast path
instead of its rewriting path. That is the attribution, established by
construction rather than by profile share.

### Native cycles and instructions (perf stat isolation pairs, CPU 16)

`probe0649 prefix <deck> <stage> N` runs `N × (open + every stage up to
<stage>)`. Profiled at N=2 and N=10 and differenced over the eight extra
iterations.

| stage increment | cycles | share | instructions | share |
| --- | ---: | ---: | ---: | ---: |
| the capture | 260,588,713 | 43.55% | 1,227,624,628 | 46.95% |
| the working clone | 1,425,585 | 0.24% | 3,835,933 | 0.15% |
| `set_shape_text` | 8,292,178 | 1.39% | 88,827,334 | 3.40% |
| `commit` | 316,496,675 | 52.89% | 1,284,705,562 | 49.13% |
| `apply` | 11,580,240 | 1.94% | 9,806,020 | 0.38% |
| **one edit** | **598,383,390** | | **2,614,799,476** | |

The `open` that precedes every iteration is 2,084,063 cycles and 14,888,911
instructions and is outside the edit. Capture plus commit are **96.44% of the
edit's cycles and 96.08% of its instructions**, agreeing with the 95.9% wall
clock.

### Deterministic counts, per edit (`apply` prefix minus `open` prefix)

From the instrumented leg. "MCE pass" is one
`mce::codec::process_markup_compatibility` call; a pass either *borrows* (the
part does not mention the MCE namespace anywhere, so only the namespace scan
runs) or *rewrites* (the whole part is re-serialized).

| counter | real deck | generated corpus | control |
| --- | ---: | ---: | ---: |
| MCE passes | 94 | 91 | 93 |
| — of which **rewriting** | **94** | **0** | **0** |
| MCE input bytes | 1,726,827 | 308,886 | 1,708,797 |
| part decodes (`parts::processed_xml`) | 59 | 55 | 59 |
| part-decode bytes | 1,093,071 | 175,884 | 1,093,071 |
| shape-scene parses | 4 | 4 | 4 |
| shape-scene bytes | 68,012 | 14,014 | 68,012 |
| captures | 2 | 2 | 2 |
| `package_fingerprint` calls | 2 | 2 | 2 |
| **bytes fed to the revision proof** | 1,568,820 | 268,444 | 1,568,820 |
| `packages_equal` | 1 | 1 | 1 |
| allocations | 359,769 | 25,765 | — |
| allocated bytes | 91,012,217 | 1,556,563 | — |
| reallocations | 30,296 | 1,029 | — |

Three things follow immediately. The real deck and the control run **identical**
pass counts, decode counts, scene parses, captures and fingerprint bytes; the
control is 16.4× faster anyway. The generated corpus runs the **same structure**
(91 passes against 94) and **none** of them rewrites. And one edit of a 108 KB
file allocates **91 MB — 841× the source archive**.

### One capture, in full

44 MCE passes over 819,319 bytes — **1.03× the deck's entire uncompressed
size, per capture** — producing:

| | one capture | one edit |
| --- | ---: | ---: |
| MCE input bytes | 819,319 | 1,726,827 |
| MCE **output** bytes | **13,336,354** (16.28×) | **28,064,911** (16.25×) |
| start tags re-emitted | 29,550 | 62,181 |
| **namespace bindings re-declared** | **177,300** | **373,086** |
| — per emitted element | 6.00 | 6.00 |

Every emitted start tag carries every in-scope binding, six of them on this
deck, each ` xmlns:p="…"` escaped character by character by
`mce::codec::esc` into a bounds-checked `BoundedOutput`. That output is what the
cycles buy: 260.6 M cycles per capture is **19.5 cycles per output byte** and 318
cycles per *input* byte. One edit therefore produces and throws away **28.1 MB of
markup to publish one text replacement in a 108 KB file — 259× the source
archive.**

### Which call sites ask for those 44 passes

Instrumented per-site counters, plus callgrind's inclusive `Ir` share of the
capture profile. The three per-slide sites each read **every one of the 13
slides in full**, and the three of them are 94.55% of the capture.

| site | passes | input bytes | callgrind inclusive |
| --- | ---: | ---: | ---: |
| `notes::load_snapshot` → `load_index` → `root_conformance` → `scan_xml` | 15 | 273,892 | 38.95% |
| `SlidePart::name` → `c_sld_name` — reads `p:cSld/@name` | 13 | 269,178 | 27.88% |
| `SlidePart::from_part` → `root_name` — reads the root element name | 13 | 269,178 | 27.72% |
| `PresentationPart::slide_references` (called twice) | 2 | 4,714 | — |
| `PresentationPart::from_part` root check | 1 | 2,357 | — |
| **the complete-package revision proof** | — | 784,445 hashed | **3.24%** |

`root_name` decodes a whole slide to learn that its root is `p:sld`. `c_sld_name`
decodes the same slide again to read one attribute of its second element. The
notes site decodes it a third time, on a deck that contains **no notes slides at
all** — `ppt/notesSlides/` is empty, so `load_index` returns `None` after reading
every slide. And `root_conformance` runs `scan_xml` once per conformance in
`[Transitional, Strict]` until one succeeds, so a Strict deck pays that third
pass twice.

### Native per-symbol, and the callgrind disagreement about SHA-256

`perf record -F 999`, self-cycle shares.

| | the capture stage | the whole edit |
| --- | ---: | ---: |
| `mce::codec` symbols (`extend_from_slice`, `esc`, `start`, `write_start`, `push`, `close`, `Ctx` drop) | 64.74% | 61.30% |
| `__memmove_avx512_unaligned_erms` (resolving under `BoundedOutput::extend_from_slice`) | 15.49% | 18.38% |
| **MCE-attributable** | **80.69%** | **79.68%** |
| `quick_xml` parsing the MCE output | ~9% | ~9.5% |
| `sha2::sha256::x86_sha::compress` — **the revision proof** | **0.51%** | **0.72%** |

Callgrind's inclusive tree on the capture agrees about the mechanism and
disagrees about the proof: `capture_internal` 98.41%,
`mce::codec::process_ooxml` **84.17%**, `write_start` 73.69%, `esc` 67.97%,
`BoundedOutput::extend_from_slice` 51.36% — but `package_fingerprint` **3.24%**,
because callgrind runs `sha2::sha256::soft::unroll::compress` where the hardware
runs SHA-NI. **Callgrind overstates the revision proof 6.4× here**, exactly as
the program's measurement rules warn. Change 0590 measured `package_fingerprint`
at 34.94% of `pptx_eager_batch_edit_save`; that selector's corpus is a
*generated* deck, where every MCE pass borrows and the fingerprint is the only
bulk work left. On a real deck the same function is 3.24% under callgrind and
0.51% natively — about **0.9 ms of the 128 ms**.

The isolation pair itself: 2,534,704,412 `Ir` at N=2 and 7,609,845,677 at N=6,
so 1,268,785,316 `Ir` per (open + capture) against the hardware counter's
1,242,513,539 — the two tools agree to 2.1% on the total while disagreeing 6.4×
on that one symbol.

**This is the one result this record hands to change 0645.** That change owns the
revision-proof format design, and 0590 left part (c) — a cheaper proof — designed
and unimplemented. The proof still runs twice per edit over 1,568,820 bytes here,
unchanged in count and in input from the base, and it is **0.51% of the capture's
native cycles**. Whatever 0645 designs, on this deck it is competing for about
0.9 ms of 128 ms; the 34.94% that made it look like the dominant term in 0590 is
a generated corpus and a software SHA-256 in the same number. This record does
not design a cheaper proof, and does not need to.

### What the generated corpus lacks

| | real deck | generated corpus |
| --- | ---: | ---: |
| archive bytes | 108,164 | 40,788 |
| members | 103 | 61 |
| uncompressed bytes | 796,725 | 142,110 |
| slides / layouts / masters / themes | 13 / 18 / 11 / 11 | 12 / 11 / 1 / 2 |
| largest slide part | 121,498 B | ~3,400 B |
| **members mentioning the MCE namespace** | **43 — 93.0% of the bytes** | **0 — 0.0%** |
| MCE rewrites per capture | 44 | 0 |

The generated corpus is authored by `litchi_pptx::Package::new()`, which never
writes an `mc:` namespace, so all 43 of its capture's passes return
`Cow::Borrowed` after the namespace scan alone. It is **5.6× smaller and takes a
path 13.9× cheaper per byte** — 189.8 MB/s scanning against 13.7 MB/s rewriting
— and 5.57 × 13.9 = 77, against the capture's measured 77.3×. That product is the
whole of 0638's 68.7×.

This is 0587's corpus gap in its sharpest form: **every prior PPTX record measured
on decks that never reach the code that costs 94% of a real deck's edit.** The
control isolates the two halves — at 5.6× the generated corpus's bytes and
none of its rewriting, the control's edit is 4.1× the generated corpus's.

### 0637 does not reach this path

Change [0637](0637-pptx-eager-slide-catalog-memo.md) memoizes the eager slide
catalog on a borrowed `Presentation`, and landed after this record's base. Its
crate diff was applied to the base and the counts re-run: **every counter is
identical**, because `capture_internal` builds one view and makes one query
through it. 0638 measured 133.61 ms with 0637 in the tree and this record
measures 128.12 ms without it; the difference is host load, not 0637. The two
`slide_references` passes per capture survive 0637 for the same reason — the
first is on the `PresentationPart`, before any view exists.

## Correctness evidence

No production code changed, so there is no differential to run: the corpus
differential this record would have cited is the one it did not need. What is
checked instead:

* The probe reproduces 0638's own numbers on 0638's own selectors at this base.
  `pptx_real_file_ordinary_save_edit` and `pptx_ordinary_save_edit` were run from
  the base with 0638's harness-only diff (commit `15363b3a8`, which touches no
  file under `crates/`) applied; retained as
  `results/change-0649/harness/0638-selectors-at-base.json`. Every retained
  sample reproduced 0638's frozen edit outcome and published digest.
* The probe's edit target is derived by 0638's own rule — the first (slide,
  shape) position the documented transaction admits — and lands on the same
  `slide:1/shape:1` for the real deck and `slide:0/shape:0` for the generated
  corpus that 0638 records.
* The marker-stripped control is proved structurally equal to the real deck by
  the count table above, not asserted: identical member count, identical
  uncompressed byte total, identical MCE pass count, part-decode count,
  shape-scene count, capture count and fingerprint bytes.
* The instrumentation leg was reverted before the gates. The gates below run on a
  `crates/` tree that `git status` reports as clean against c7326f680.

**Gates** (`results/change-0649/gates.txt`): `cargo fmt --all --check`;
`cargo clippy -p litchi-pptx -p litchi-ooxml-common --all-targets`;
`cargo doc -p litchi-pptx -p litchi-ooxml-common --no-deps`;
`cargo test -p litchi-ooxml-common`; `cargo test -p litchi-pptx`. All pass.

## The design, frozen

Four candidates remove part of the 28,064,911 bytes of MCE output one edit
produces. Each is sized here in MCE output bytes per edit — A and C save two of
the three per-slide passes in both captures, B saves twelve of the thirteen
slides in the second capture — and each is stopped by something named.

| # | candidate | ceiling | what stops it |
| --- | --- | ---: | --- |
| A | one MCE pass per part per capture instead of three | −17.6 MB (**62.8%**) | needs a per-part memo. `SlidePart` and `PresentationPart` are public `#[derive(Clone, Copy)]` views over `&dyn Part`; a `OnceLock` field removes `Copy` from a public type — the obstacle 0637 recorded for `PresentationPart` and declined for the same reason. Plumbing a cache instead reaches the nine `processed_xml` call sites in `parts/presentation.rs` and `parts/slide.rs` and every public method that owns one, and the cache holds 13.3 MB per capture on a 108 KB file, which is a bounded-resources question ADR 0005 would want answered in its own record. |
| B | the commit's recapture reuses the first capture's per-slide state for the slides the patch did not touch | −12.2 MB (**43.5%**) | moves when a per-slide refusal happens: a malformed *unchanged* slide is validated twice today and would be validated once. |
| C | `root_name`, `c_sld_name` and `root_conformance` read a bounded prefix instead of the whole part | −17.6 MB (**62.8%**) | moves when a refusal happens in the other direction: those three sites refuse a malformed tail today because MCE validates the complete document before they read their one string. |
| D | emit only newly introduced namespace bindings in `mce::codec::write_start` | 0588 measured a further **−80% of the codec** | **already designed, implemented, measured and withdrawn** by change [0588](0588-mce-codec-namespace-emission.md). The per-element re-declaration is load-bearing: `litchi-docx` parses sliced element ranges standalone and two writers publish the processed stream. 0630's next-wave queue carries it as row 4 with 133 classified call sites. Out of this record's scope in any case: it is `litchi-ooxml-common`, shared by every OOXML crate. |

**None is value-identical, semantics-preserving and small, so nothing is
implemented.** The one change that is unambiguously value-identical and in scope
— collapsing the two `slide_references` passes per capture — is worth 2,357 of
819,319 MCE input bytes, 0.29%, about 0.18 ms of 128 ms. That is far below the
4.31% floor, and landing it would be the speculative complexity the decision
rules say to revert.

What this record does change is the *shape* of the opportunity. 0588 asked
whether the codec's output could be made smaller and answered no. This record
asks a question 0588 did not: **how many times does a caller ask for that output
for the same bytes?** Six times per slide per edit, for two short strings and a
notes index that is empty. Candidates A, B and C are all in `litchi-pptx`, all
bounded by the 94.55% the three per-slide sites carry, and all independent of
0588's frozen writer.

## Validation preserved

Every validation named above still runs, in the same order, because nothing ran
differently: the slide-reference resolution, the one-to-one identity checks, the
relationship-type checks, the per-slide name parse, `notes::load_snapshot`'s
topology gate (ADR 0013), the name-index build, and the complete-package
revision proof (ADR 0003). The marker-stripped control is a control for the code
path and **not** a proposal: it is not a semantically valid transformation of a
`.pptx`, it exists only inside the measurement, and it is not retained as a
fixture.

## Limitations

* **One deck.** Every number is `slide-section-test.pptx` on this host at this
  base. 0588 reported the amplification as 14–17× on real producer parts across
  formats; the 16.3× here is one point inside that range, and the 6.00 namespace
  declarations per element are this deck's, not a producer census.
* **No claim is registered**, and no before/after pair exists, because no
  production code changed. The paired structure here is real-versus-control and
  real-versus-generated, not before-versus-after.
* **The ceilings in the design table are MCE output bytes**, which the cycle
  measurements show this path is linear in — but a count of removed work is not
  a saving until the cost it trades into is measured natively, and none of A, B
  or C has been implemented, so none of those ceilings is a measured saving.
* The sub-millisecond phases (`Snapshot::edit`, `apply`) sit on floors of 8% and
  29%; their shares are reported to show they are negligible, not to rank them.
* The per-site table's 38.95% for the notes path and 27.88%/27.72% for the two
  slide-name paths are callgrind inclusive `Ir`, which ranks work, not latency.
  The wall-clock and cycle statements in this record rest on the control and on
  `perf stat`, not on those shares.
* **The third per-slide pass can double.** `notes::codec::root_conformance` runs
  `scan_xml` — and therefore a whole-part MCE pass — once per conformance in
  `[Transitional, Strict]` until one succeeds. This deck is Transitional, so it
  pays one; a Strict deck would pay two, making seven whole-slide rewrites per
  edit instead of six. That is read from the source and **not measured**.
* Neither the DOCX nor the XLSX ordinary-save edit 0638 measured (0.51 ms and
  3.74 ms) was decomposed here. Their editors do not run a whole-package capture,
  which is why they are cheap, but that is an observation about the route rather
  than a measurement of it.
* `litchi-perf-baseline-alloc` reports no allocation metrics for 0638's
  ordinary-save family — the family was registered by a harness-only change that
  did not reach `allocation_metrics.rs`. The allocation counts above come from
  the probe's own counting allocator instead, and the gap is reported here rather
  than fixed.

## Retained evidence

`results/change-0649/README.md`.
