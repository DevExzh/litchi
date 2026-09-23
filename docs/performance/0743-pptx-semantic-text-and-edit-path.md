# 0743: the PPTX semantic read and edit paths stop re-reading bytes they already validated — on the 100 × 100 deck full text −46%, one-percent edit/save −46%, one-shape edit/save −19% and no-op −16%, with every refusal and published byte unchanged; a commit memo worth a further −90% of the one-edit commit is withdrawn pending proposed ADR 0032

Status: retained, implemented in `crates/litchi-pptx` (eight production
commits retained; one withdrawn and proposed as ADR 0032).
`performance_claim: none` — the paired medians, instruction and cycle counts,
allocation counts and per-commit attribution below are reported as evidence
beside an A/A floor measured in the same session, not registered as claims.

OLE2 and OOXML remain the active priority. ODF optimization stays deferred until
that goal completes; iWork is excluded and untouched.

Base `009d515bef`; branch `perf/0743-pptx-semantic-text-and-edit-path`. The
measured head is `ad2c490ee3`: the retained 0743 changes, the revert of the
memo, a compile-time guard, and the separate crash fix of change
[0755](0755-pptx-nested-text-run-panic.md), whose own cost is within ±0.05% of
instructions on every case below.

## Result

The harness's own cases. Both legs are built with the identical command,
`cargo build --release --locked --offline --manifest-path tools/perf-baseline/Cargo.toml --bin litchi-perf-baseline`,
from detached worktrees of equal path length (`0743-src-a` at the base,
`0743-src-b` at the head) into target directories of equal length, and run from
binary paths of equal length; eight processes per arm in four ABBA blocks,
pinned to core 8, paired within each block. The A/A floor is the before binary
against a byte-identical copy, two blocks, in the same window.

| case | deck | before p50 | after p50 | paired p50 change [95% CI] | A/A p50 [95% CI] |
| --- | --- | ---: | ---: | ---: | ---: |
| `pptx_semantic_full_text` | large | 50.688 ms | 27.137 ms | **−46.47%** [−47.48%, −46.18%] | +0.39% [−2.61%, +1.12%] |
| `pptx_semantic_one_percent_edit_save` | large | 269.608 ms | 144.746 ms | **−45.97%** [−46.20%, −45.59%] | +0.40% [+0.02%, +1.70%] |
| `pptx_semantic_one_edit_save` | large | 55.570 ms | 44.547 ms | **−18.72%** [−20.47%, −17.40%] | +0.30% [−1.38%, +0.97%] |
| `pptx_semantic_noop_edit_save` | large | 27.926 ms | 23.171 ms | **−16.42%** [−17.36%, −15.79%] | −0.60% [−1.36%, +2.30%] |
| `pptx_semantic_open` | large | 2.170 ms | 2.093 ms | −0.83% [−5.58%, +2.68%] | −3.27% [−7.96%, +5.04%] |
| `pptx_semantic_full_text` | medium | 0.592 ms | 0.340 ms | **−42.62%** [−42.83%, −41.82%] | −0.07% [−0.48%, +0.03%] |
| `pptx_semantic_one_percent_edit_save` | medium | 1.712 ms | 1.507 ms | **−12.33%** [−12.70%, −10.10%] | +0.63% [−0.49%, +1.57%] |
| `pptx_semantic_one_edit_save` | medium | 1.710 ms | 1.506 ms | **−12.33%** [−12.59%, −9.90%] | +0.80% [−0.61%, +0.94%] |
| `pptx_semantic_noop_edit_save` | medium | 0.851 ms | 0.807 ms | **−5.05%** [−5.85%, −4.13%] | +0.64% [−0.39%, +1.03%] |
| `pptx_semantic_open` | medium | 0.448 ms | 0.453 ms | +1.61% [+0.87%, +1.98%] | +0.93% [+0.00%, +1.78%] |
| `docx_semantic_full_text` (control) | large | 3.197 ms | 3.167 ms | −0.44% [−1.43%, +0.22%] | +0.66% [−0.36%, +3.33%] |
| `docx_semantic_full_text` (control) | medium | 0.065 ms | 0.064 ms | −0.81% [−2.59%, +1.94%] | −0.04% [−0.86%, +0.60%] |

Per operation, from probes built the same way against the same two trees (user
instructions and cycles counted at two iteration counts and differenced, ABBA,
median of four pairs; the edit/save cycles have their per-iteration package
construction subtracted):

| region | deck | before instructions | after instructions | change | before cycles | after cycles | change |
| --- | --- | ---: | ---: | ---: | ---: | ---: | ---: |
| full text | large | 935.06 M | 507.43 M | −45.73% | 222.96 M | 125.01 M | −43.93% |
| full text | medium | 10.41 M | 5.90 M | −43.31% | 2.50 M | 1.49 M | −40.58% |
| no-op edit/save | large | 496.94 M | 412.92 M | −16.91% | 120.14 M | 104.02 M | −13.42% |
| no-op edit/save | medium | 12.09 M | 10.96 M | −9.41% | 3.85 M | 3.64 M | −5.49% |
| one-edit edit/save | large | 1,011.24 M | 821.23 M | −18.79% | 238.70 M | 198.09 M | −17.01% |
| one-edit edit/save | medium | 25.68 M | 21.47 M | −16.40% | 7.64 M | 6.69 M | −12.34% |
| one-percent edit/save | large | 5,134.65 M | 2,732.07 M | −46.79% | 1,174.35 M | 644.30 M | −45.14% |
| open (`Package::from_bytes`, control) | large | 32.15 M | 32.15 M | +0.00% | 9.01 M | 9.03 M | +0.29% |
| open (`Package::from_bytes`, control) | medium | 6.75 M | 6.75 M | −0.04% | 1.73 M | 1.74 M | +0.46% |

The large deck is `pptx-semantic-large` (100 slides × 100 text boxes, archive
215,220 bytes, 3,839,500 bytes of slide XML, 38,395 per slide); medium is
12 × 8. The medium one-percent case makes one edit, like the one-edit case. In
the coordinator's per-slide terms an additional edited slide cost 2.16 ms and
now costs 1.01 ms ((144.75 − 44.55) / 99), and the no-op transaction cost
0.279 ms per slide and now costs 0.232 ms — almost all of it the notes-graph
validation scan of every slide that every capture performs.

The harness phase case (`pptx_semantic_opened_transaction_phases`, one edit,
median of per-process medians) locates the saving:

| phase | large before | large after | change | medium before | medium after | change |
| --- | ---: | ---: | ---: | ---: | ---: | ---: |
| `opened_presentation` (capture) | 27.204 ms | 23.215 ms | −14.66% | 0.760 ms | 0.716 ms | −5.80% |
| `Snapshot::edit` | 0.047 ms | 0.045 ms | −4.37% | 0.013 ms | 0.013 ms | −0.52% |
| `set_shape_text` | 1.193 ms | 0.749 ms | −37.23% | 0.128 ms | 0.084 ms | −34.21% |
| `Transaction::commit` | 25.644 ms | 20.821 ms | −18.81% | 0.651 ms | 0.535 ms | −17.85% |
| `apply_opened_presentation_commit` | 0.201 ms | 0.200 ms | −0.85% | 0.058 ms | 0.058 ms | −0.26% |
| `Package::to_bytes` | 0.346 ms | 0.339 ms | −2.21% | 0.089 ms | 0.088 ms | −0.93% |
| total | 54.601 ms | 45.484 ms | −16.70% | 1.698 ms | 1.491 ms | −12.21% |

Published archives, patch emptiness and snapshot revisions are byte-identical
between base and head for the tiny, medium and large decks under the no-op,
one-edit and one-percent edits, and so is the full text; the no-op output is the
input archive. On the large deck allocation calls fall 48.7% (full text), 84.6%
(no-op), 85.1% (one edit) and 88.8% (one percent).

## Owner-pending: the withdrawn commit-recapture memo (proposed ADR 0032)

`Transaction::commit` recaptures the staged package, and a capture classifies
every slide root with the complete notes-graph scan. The staged package shares
every slide payload allocation the transaction did not rewrite, so for a
one-slide edit the commit rescans 99 slides whose bytes the source snapshot's
capture already classified. That is the 20.8 ms left in the one-edit commit.

Commit `99ce9c5e34` removed it with a `SlideRootMemo` on the snapshot: per slide
payload its capture read, the allocation key, a strong `Arc` and the
classification, admitted only for allocations the snapshot's own digest memo
proves its package owns, projected on a rebind, looked up by the exact raw
observation MCE had just made. Measured in the probe (median of three
processes), it took the one-edit commit from **21.45 to 2.13 ms (−90%)** and the
one-edit edit/save from 46.44 to 27.11 ms (−42%) on the large deck, and the
medium one-edit edit/save from 1.550 to 1.346 ms, with outputs, revisions and
refusals identical over the repository's PPTX corpus. In the first harness run
of this change the one-edit case measured −52.45% with it, against −18.72% now.

It is withdrawn (`07207d2aef`) because it conflicts with an accepted ADR. ADR
0005's 2026-09-16 amendment admits snapshot memos of *digests* over bytes the
owner holds, and requires reservation exhaustion to be a typed resource error;
this memo held notes-root classifications and fell back to an empty table when
it could not reserve. `docs/GOAL.md` forbids implementing behaviour that
conflicts with an accepted record, so the memo is proposed as
[ADR 0032](../adr/0032-snapshot-derived-value-memos.md) — memos of small
derived values that are pure functions of the payload bytes, under exactly the
amendment's conditions, with construction failure as a typed
`Error::Allocation` — for the owner's review. If accepted, `99ce9c5e34`
re-applies unchanged except for that typed-error rule.

## What was changed

| commit | change | files |
| --- | --- | --- |
| `588cc6a317` | **A.** `Scene::scan` stops owning every event and cloning the whole namespace resolver per event | `shape/reader.rs` |
| `8039963c33` | **B.** the notes-graph scanner reads borrowed events and validates names and values in place, owning only reported values | `notes/codec.rs`, `notes/mod.rs` |
| ~~`99ce9c5e34`~~ | **C.** the slide-root memo — **reverted by `07207d2aef`**, proposed as ADR 0032 (above) | — |
| `00234742d8` | **D.** semantic text reads a marker-free slide once instead of twice, keeping raw-scan refusals ahead of semantic ones | `parts/slide.rs` |
| `aff8de9c61` | **E1.** the scene scanner splits each element's qualified name once, not once per classification | `shape/reader.rs` |
| `d10f050c89` | **E2.** commit compaction decides untouched slides by allocation identity and does not re-read a compaction that reproduced the staged bytes | `opened/transaction.rs` |
| `e0e8f35d07` | **F1.** the text-run locator keeps its events borrowed | `opened/xml.rs` |
| `aaed57e900` | **F2.** the raw-span mapper reuses the scene's successful read of the same unmarked bytes instead of a separate offset pass | `tag/shape/codec.rs`, `shape/reader.rs`, `opened/transaction.rs` |
| `e74faad918` | **G.** a transaction records which staged slide payloads its verbs read back as a scene, and compaction does not read those exact allocations again | `opened/transaction.rs` |
| `99e5044fa4` | a compile-time assertion that the processed-text ceiling covers the raw one, which D's single-pass route relies on | `parts/slide.rs` |
| `7cf32442af` | tests only: the corpus oracles visit parts in name order | `notes/codec.rs`, `parts/slide.rs` |
| `b82c81ceef` | the ADR 0032 proposal and its listing under the README's proposed records | `docs/adr/` |

No public function, type, trait, variant, durable format, limit or dependency
is added or removed. One public type's derived `Debug` output changes: `Scene`
(`shape/reader.rs`) gains a private `limits` field for F2, which its derived
`Debug` prints.

### The redundant work each commit removes

**Capture (B).** `opened_presentation` validates the notes graph, which needs
every slide's root conformance, and establishes it with the complete
notes-graph scan (`scan_processed_xml` over the whole slide). That scan was 88%
of capture on the large deck. It read each event into a scratch buffer and
allocated an owned `String` for every element namespace, every local name,
every attribute namespace and every unescaped attribute value, although only
relationship attribute values are ever reported. The scan now reads borrowed
events from the slice reader and validates the same bytes in place. Capture
falls 27.2 → 23.2 ms; the scan itself remains, because it is a validation
obligation (a slide containing CDATA, an unbound prefix or a DTD is still
refused by capture even when the deck has no notes). With the memo withdrawn,
the commit's recapture pays the same scan again, now cheaper by the same
mechanism: 25.6 → 20.8 ms.

**Semantic text (D).** `SlidePart::text` / `Presentation::text` ran a raw-byte
validation scan, then MCE processing, then the semantic pass — two complete
namespace-aware passes over the same bytes whenever MCE changes nothing. For
such a slide the two passes now share one event stream. Each event is validated
by the raw scan first and a raw refusal returns at once; a semantic refusal is
only recorded, the semantic pass stops, and the raw scan continues to the end.
So a raw refusal anywhere still outranks a semantic refusal anywhere, exactly as
when the raw scan ran to completion first, and the semantic pass consumes
exactly the prefix it consumed alone. On this route the semantic pass skips only
its two checks that repeat the raw scan's checks of the same element against the
same bindings (attribute values, attribute-name prefixes); debug builds rerun
them. The route drops the processed-size check, which cannot fire because the
processed ceiling is at least the raw one; `99e5044fa4` makes that a
compile-time assertion. Full text falls 50.7 → 27.1 ms.

MCE processing now runs before the raw scan. It is a pure function of the
bytes, and both routes report failures in the original order, so no value or
refusal moves. One cost does: a malformed slide that carries the MCE namespace
now pays a complete, bounded MCE pass (under the same 64 MiB input and output
ceilings, depth and binding limits) before its raw-scan refusal, where it used
to be refused first. That affects only malformed, marker-bearing input — the
adversarial minority 0652's trade-off 3 lets pay more.

**Edit path (A, E1, F1, F2).** `set_shape_text` read the slide as a scene,
mapped the selected shape to its raw span with two raw passes (offsets, then the
map), located its text runs, rewrote them and read the staged XML back as a
scene. `Scene::scan` owned every event and cloned the namespace resolver (binding
stack and URI buffer) per event (A), and split each element's qualified name
about twenty times (E1); the text-run locator owned every event (F1). The raw
offset pass (F2) exists to feed the MCE offset selection; when the scene just
read these very bytes (same allocation and length) without an MCE rewrite, the
selection would keep every offset, the scene's reader configuration raised no
reader error over them, and its node and input ceilings are no looser than the
raw passes'. In that case every candidate is active and the offset pass is
skipped; the index-agreement, family and active-tree refusals of the mapping
pass all remain, and debug builds run the full route beside it and assert the
same span or the same refusal. `set_shape_text` falls 1.19 → 0.75 ms per call.

**Commit compaction (E2, G).** Commit compacts each rewritten slide: it read the
staged scene, compacted, and read the compacted scene back. Untouched slides were
compared with the source byte by byte although they share its allocation (E2 now
checks identity first). When compaction reproduces the staged bytes the read-back
scene is the one just read, so it is no longer read (E2). And `set_shape_text` /
`set_shape_texts` already read their staged XML back as a scene; the transaction
now records that allocation as a `Weak`, and compaction skips its first read only
for that exact allocation (G). A payload replaced or mutated in place since
(which disassociates the `Weak`) is read as before, so a staged slide the scene
reader refuses — for example one produced by `remove_shape`, which does not read
its output back — is still refused. These carry the one-percent case: 269.6 →
144.7 ms.

## Authority

* **ADR 0003** — `commit()` still validates the complete staged package; no
  validation is skipped for bytes that were not validated in the same
  transaction by the same function. Exact no-ops remain exact (the no-op commit
  returns the source snapshot, and the published archive equals the input);
  revisions, patches and conflicts are unchanged (revisions identical before and
  after).
* **ADR 0005** — no ambient or process-wide state, no new limit, no execution
  context, and no snapshot memo: the one that needed an ADR is withdrawn and
  proposed as ADR 0032. G's scene-read record is transaction-local state that
  holds only `Weak` references, pins nothing, and a miss is the ordinary read.
* **ADR 0006** — validation never mutates and every refusal stays typed; the
  differential oracles below prove the refusals identical, not only present. No
  output byte moves, so the preservation contract is untouched.
* **0652** trade-offs 2 and 3 — correctness first (every shortcut is either an
  identity proof or an equivalence the tests check against the old code), and
  the benign common path is the optimized one: marker-free slides take the
  single-pass and unmarked raw-span routes, marked slides the unchanged routes,
  and only malformed marked slides pay the earlier MCE pass.
* ADRs 0002/0024/0011: one crate, no new dependency, no archive type in any
  signature. ADR 0030: lazy part decode is untouched.

## Evidence that motivated it

The coordinator's sweep and a probe that loops exactly the harness's timed calls
(profiles in [`results/change-0743/profiles`](results/change-0743/profiles)):
capture was 88% notes-root scan (`finish_from_processed::{closure}` →
`scan_processed_xml`), with SHA-256 fingerprinting 7%; the one-edit commit was a
second complete capture; `set_shape_text` was 54% `Scene::read_with` (two reads),
24% raw-span mapping and 12% the rewrite with its text-run locator; in the
one-percent commit profile, scene reads inside compaction were 51% and the
recapture 21% of the process; full text was two complete passes per slide,
with attribute validation run twice on each element.

## Measurements

Host AMD EPYC 9R45, 32 cores, Linux 7.0.0-1012-aws, toolchain 1.95.0 (the
worktree's `rust-toolchain.toml`); every measured process `taskset -c 8`; other
agents built and measured concurrently (the one-minute load average before each
process is in the packet). The harness is unchanged. Binaries:
before `41278898d475e4809ce614fe1efb0452dd034ae83fcdedfbc595e5e301451f67`, after
`7ce633a4a4ff94e807361172a3b0782608eaed8ab5bc5f7945bb98cfb487913c`; the probes
and allocation probes are listed in the packet. Large-deck processes take 15
samples after 3 warmups, medium ones 60 after 10, the phase case 15 after 3.
Each paired ratio is after / before within one ABBA block (positions 0/1 and
3/2); the interval is a percentile bootstrap over the paired processes (10,000
resamples, seed 743). Samples within a process are never treated as
independent. Every harness process also ran under `perf stat` for user
instructions and cycles: over the whole process (which includes corpus
construction and verification) the large group fell from 385.45 G to 224.36 G
instructions (−41.79%) and 90.63 G to 53.55 G cycles, while the A/A pairs moved
−0.00% and +0.18%.

Core 8 was time-shared by other runnable work during parts of the session. The
first A/A attempt shows it: four of its processes ran at almost exactly twice
the wall time of their pairs (for example 55.44 against 27.07 ms) while their
user cycles stayed within 0.38% of every other process. That attempt is kept as
`aa-disturbed/` and the A/A floor above is its rerun on a quiet core. In the
A/B run the same signature appears in the processes of block 2 position 3 and
block 3 positions 2 and 3 (`one_edit_save` and `one_percent_edit_save`, large),
and as a p95 tail in block 2 position 2; they are kept, and the medians and
intervals above include them. Because both processes of a pair were time-shared
in block 3, that pair's ratios still agree with the clean pairs (−15.66% and
−45.00% p50).

A first measurement of this change, with the memo still in and its before leg
built from the base checkout, is kept in the packet as `timing-samecmd/`: full
text −44.28%, no-op −15.08%, one edit −52.45% (with the memo), one percent
−46.50%. The run before that used the prebuilt base harness as the before leg
(`timing/`); the coordinator's note that the prebuilt binary shifts untouched
paths through layout is why both later runs build both legs with one command.

### Per-commit attribution

The retained probe, built at every commit of the first series, in rotating
order, three processes each (median of per-process medians; probe processes
differ by about ±2–3%). The series still contains C, which is now reverted; the
other rows are unchanged by that, since C touched only the commit's recapture:

| after commit | full text L | full text M | no-op L | one edit L (capture / set text / commit) | one percent L (set text / commit) |
| --- | ---: | ---: | ---: | ---: | ---: |
| base | 51.01 | 0.574 | 28.46 | 53.78 (26.97 / 1.189 / 25.21) | 267.8 (114.2 / 110.4) |
| A | 50.18 | 0.572 | 27.33 | 53.33 (26.67 / 1.077 / 25.15) | 244.6 (103.8 / 98.4) |
| B | 50.59 | 0.570 | 23.83 | 46.44 (23.35 / 1.079 / 21.45) | 237.6 (103.5 / 94.9) |
| C (reverted) | 48.86 | 0.562 | 23.68 | 27.11 (23.25 / 1.075 / 2.13) | 236.4 (102.6 / 94.7) |
| D | 27.73 | 0.330 | 24.51 | 28.04 (23.96 / 1.071 / 2.16) | 240.3 (103.5 / 95.5) |
| E1 | 27.59 | 0.327 | 23.89 | 27.01 (23.48 / 0.924 / 2.02) | 208.9 (88.9 / 80.6) |
| E2 | 27.51 | 0.329 | 23.91 | 27.03 (23.80 / 0.897 / 1.75) | 184.5 (87.2 / 57.3) |
| F1 | 28.55 | 0.335 | 23.84 | 26.71 (23.52 / 0.896 / 1.72) | 184.9 (85.8 / 58.9) |
| F2 | 27.94 | 0.330 | 23.57 | 26.15 (23.10 / 0.737 / 1.71) | 167.7 (70.9 / 57.3) |
| G | 27.46 | 0.330 | 23.56 | 25.75 (22.93 / 0.742 / 1.51) | 147.4 (72.0 / 35.5) |

All values in ms. B moved capture (and the recapture inside every commit); D
moved full text; A, E1, E2, F2 and G moved the one-percent case; C moved the
one-edit commit and is the owner-pending item; F1 has no effect resolvable from
the probe's noise on its own.

### Allocations

Exact counts from probes with a counting global allocator (probe-only) around
the same calls, against the base checkout and the head; two repeats agree
exactly:

| region | large before | large after | medium before | medium after |
| --- | ---: | ---: | ---: | ---: |
| `text()` calls / bytes | 250,209 / 18.9 MB | 128,309 / 11.1 MB | 3,498 / 289 KB | 2,118 / 197 KB |
| no-op edit/save calls / bytes | 538,176 / 30.7 MB | 83,135 / 19.5 MB | 11,643 / 3.33 MB | 5,362 / 3.18 MB |
| one-edit edit/save calls / bytes | 1,116,597 / 53.7 MB | 166,350 / 26.4 MB | 26,350 / 4.92 MB | 10,258 / 4.18 MB |
| one-percent edit/save calls / bytes | 5,550,213 / 640 MB | 623,537 / 111 MB | 26,350 / 4.92 MB | 10,258 / 4.18 MB |

### Regression flags

Every paired ratio above 1.05 in p50, p95 or mean is listed in the packet's
`tables.md` files. None is on a p50 of a case this change speeds up. The A/B run
has seven:

* `pptx_semantic_one_edit_save` large p95 +101.94% and
  `pptx_semantic_one_percent_edit_save` large p95 +127.74% — single pairs in
  which one process was time-shared (see Measurements); the same pairs' p50
  ratios are −15.66% and −46.07%.
* `pptx_semantic_open` large p95 +5.40%, +6.78% and +17.28% — a 2 ms region
  timed between verification passes; the clean A/A floor shows the same case
  at +13.33% p95 and −3.27% [−7.96%, +5.04%] p50. Its code is unchanged: the
  probe counts +0.00% instructions per `Package::from_bytes`.
* `docx_semantic_full_text` large p95 +14.26% and medium mean +52.39% — a
  control this change does not reach; the clean A/A floor shows it at +13.28%
  p95.

`pptx_semantic_open` medium moves +1.61% [+0.87%, +1.98%] against an A/A of
+0.93% [+0.00%, +1.78%]: about 5 µs, below the 5% trigger, with identical
instruction counts per call (−0.04%) and +0.46% cycles — a placement effect, not
added work, reported rather than claimed as noise.

## Correctness evidence

* **Semantic text oracle.** The pre-change three-pass implementation is kept
  verbatim as a test oracle and compared with the routed one — value and first
  refusal — over handcrafted precedence cases (raw before semantic, semantic
  before raw, both, CDATA, references, DTD, PI, declarations, roots, unbalanced
  and unbound names, invalid characters and comments, depth, a marked slide), a
  deterministic mutation set, and every XML part of the repository's 78 PPTX
  fixtures and their mutations: 21,404 comparisons, 1,040 of them accepted
  outputs compared byte for byte. A separate test states the precedence without
  the oracle.
* **Notes-scan oracle.** The previous buffered scanner is kept verbatim and
  compared with the borrowed one under five roots, both conformances and a
  spread of raw and processed ceilings, over handcrafted, mutated and 600 corpus
  parts: 650,060 comparisons, 3,306 accepted scans compared value for value.
  Both corpus oracles visit parts in name order, so the counts repeat exactly.
* **Raw span.** Every `set_shape_text` and tag lookup on unmarked XML in the
  crate's suite runs both routes in debug builds and asserts the same span or
  refusal; a focused test keeps the index-disagreement refusal and the marked
  route.
* **Compaction.** Tests cover an untouched slide keeping its allocation, a
  reproduced compaction, a byte-changing compaction read back and compared, a
  staged slide the scene reader refuses still refusing, the skipped read, a
  replaced payload read again, and a record for another allocation not vouching.
* **Outputs.** Probes against the base checkout and the head write each build's
  published archive, revision, patch emptiness and full text for tiny, medium
  and large under the three edits: all 21 artifacts are byte-identical.
* An independent review ran a base-versus-head differential over 66,078 mutated
  packages and about 1.49 million comparisons with no mismatch (with the memo
  still in).

## What is not claimed

* That any real producer deck improves by these amounts. The corpus is the
  harness's generated text-box deck: marker-free, compact, one run per shape,
  edits at shape 0 of a slide. MCE-bearing slides take the unchanged full-text
  and raw-span routes; a slide whose compaction changes bytes still reads the
  compacted scene back.
* Anything for the withdrawn memo beyond its probe measurement: it is not in the
  measured head.
* A general document-open, save, facade (`litchi::Presentation`), source-backed
  PPTX, range-source, cold-cache, concurrency or RSS result. Allocation counts
  are requested calls and bytes on the probe's system allocator, not RSS.
* That capture or full text is at a floor. Capture is still one complete
  notes-root validation scan of every slide (quick-xml's namespace push and
  prefix resolution dominate it) plus SHA-256 over every payload, and every
  commit captures again; publication still audits and deflates every changed
  slide.
* That the small `pptx_semantic_open` shift is noise.

## Verification

Gates run in the worktree at `ad2c490ee3` (the measured code) with its own
target directory, every Cargo command with `--locked --offline` and `TMPDIR`
under the change's scratch directory; commands, exit codes and output tails are
in [`gates.txt`](results/change-0743/gates.txt):

| gate | exit |
| --- | --- |
| `cargo fmt --all --check` | 0 |
| `cargo check -p litchi-pptx --all-targets` | 0 |
| `cargo check -p litchi-pptx --all-targets --all-features` | 0 |
| `cargo check -p litchi --features doc,docx,ppt,pptx,xls,xlsx,xlsb,odt --all-targets` (the facade that depends on this crate) | 0 |
| `cargo clippy -p litchi-pptx --lib --no-deps -- -D warnings` | 0 |
| `cargo clippy -p litchi-pptx --all-targets --no-deps -- -D warnings` | 101 at head **and 101 at base** `009d515bef`: the same three `clippy::err_expect` lints in `opened/tests.rs`, which this change does not touch; not fixed here |
| `cargo test -p litchi-pptx` | 0 — 961 passed, 0 failed, 2 ignored |
| `cargo test -p litchi-pptx --all-features` | 0 — 975 passed, 0 failed, 2 ignored |
| `cargo test -p litchi --features doc,docx,ppt,pptx,xls,xlsx,xlsb,odt` | 0 — 382 passed, 0 failed, 7 ignored |
| `RUSTDOCFLAGS="-D warnings" cargo doc -p litchi-pptx --no-deps` | 0 |
| `python3 tools/check_crate_boundaries.py` | 0 |
| `python3 tools/non_iwork_gate.py verify` | 0 |
| `python3 tools/check_perf_claims.py --registry docs/performance/claim-registry-v1.json --repo-root . --mode structural` | 0 (this record registers no claim) |

The test counts include change 0755's five tests and exclude the memo's ten,
which left with the revert. The harness is unchanged, so its own suite and the
coverage validator were not required; its semantic PPTX cases nevertheless ran
every measured iteration through the harness's reopen verification, and every
measured process exited 0.

## Cleanup

Recorded in [`cleanup.json`](results/change-0743/cleanup.json). The first
series' build and scratch directories were deleted by the coordinator after the
first commit of this record (`dcbf62d252`). For this series, binary digests are
recorded in the packet, the record and packet are committed first, and then the
target directories under `targets/0743`, the scratch contents and the two
detached source worktrees (`0743-src-a`, `0743-src-b`) are removed. The
worktree and branch are kept.

## Retained evidence

[`results/change-0743/README.md`](results/change-0743/README.md).
