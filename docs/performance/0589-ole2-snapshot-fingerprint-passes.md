# 0589: an empty overlay hashes the artifact once, not twice — DOC and PPT source-backed opens lose half their SHA-256 work

Status: retained, partially implemented. `performance_claim: none` — the paired
medians, isolation-pair counters and hash-pass counts below are reported as
evidence, not registered as a claim.

OLE2 and OOXML remain the active priority. ODF optimization stays deferred until
that goal completes; iWork is excluded.

This implements item **DOC-1** of change
[0587](0587-remaining-opportunity-survey.md) (rank 2), and only the part of it
that is value-identical. Because DOC-1 is a fence under ADR 0006 (source
identity) and ADR 0005 (`SourceChanged`), the record opens with a frozen design
that classifies **every** hash pass on the path. Two items in that classification
are designed and deliberately **not** implemented, with their predicted savings
stated.

## Frozen design: what every hash pass is for

### The primitive

One function does all complete-artifact hashing:
`overlay::fingerprints(source, spans)`
(`crates/litchi-cfb/src/overlay.rs:1345`). It reads the artifact once in 1 MiB
chunks and, before this change, drove **two** SHA-256 hashers over that one read:
`source_hasher` over the bytes as read, `target_hasher` over the same buffer
after `apply_spans`. `ValidatedOverlayPlan::write_validated`
(`:918`) is the same shape with a sink in the middle and a 64 KiB chunk.

`apply_spans` (`:1385`) iterates `spans`. **With no span it cannot write a
byte.** So when the span list is empty the two hashers consume the identical
byte sequence and finalize to the identical digest. That is the entire mechanism
this change exploits, and it is value identity by construction, not an
approximation.

The empty span list is not a rare case. It is reached by:

* every `identity_fingerprint` in `litchi-doc`
  (`crates/litchi-doc/src/body_text/source.rs:2089`), which is
  `plan_same_length_stream_splices(Vec::new(), …)`;
* the PPT text-edit snapshot open
  (`crates/litchi-ppt/src/text_edit.rs:422-425`), which *is* one such plan and
  nothing else;
* every **exact byte no-op** edit, because `map_splice`
  (`crates/litchi-cfb/src/splice.rs:629`) and `push_candidate_span`
  (`overlay.rs:1191`) only record a span when the replacement differs from the
  source, so a no-op edit plans zero spans and publishes down the same branch.

### The inventory

Each row says what the pass detects and what removing it would cost.

| # | Site | What it is | Disposition |
| --- | --- | --- | --- |
| **P0** | the target hasher inside `fingerprints` and `write_validated`, when `spans` is empty | a second SHA-256 over the same bytes the first hasher just consumed; `apply_spans` is a total no-op, so the digest is equal by construction | **value-identical duplicate — implemented** |
| P1 | `finish_overlay_plan_with_owner:1033` | the planning fingerprint; captures the identity a plan is bound to | identity capture — keep |
| P2 | `finish_overlay_plan_with_owner:1058`, generic `ReadAt` only | re-reads and re-hashes after the composed-CFB reopen and the owner callback, compared against P1 | **read-twice-compare** bracketing the reopen; a dishonest adapter could otherwise change the bytes the reopen validated — keep both |
| P3 | `preflight_fingerprints:901`, reached from `composed_source`, `write_to` pre/post and `save` pre-temp/pre-rename | closes each distinct staging window: before the first sink byte, after emission, before rename | keep; the ordering rationale is already in the code comments at `:864-890` and `:979-985` |
| P4 | the emission hashing in `write_validated` | proves the bytes handed to the sink are the planned ones, mid-flight | keep (only its P0 half is elided) |
| P5 | `litchi-doc` `open_with_options:526` | `opening_fingerprint`: the identity captured before the CFB index used by the parse is read | identity capture and the opening half of a bracket — keep |
| P6 | `litchi-doc` `finish_open:591` | `final_fingerprint`, compared against P5; closes the bracket around the second index parse, the FIB/DOP/CLX reads and every refusal check | **read-twice-compare over the semantic parse** — keep |
| P7 | `litchi-doc` `finish_open:599` | a third identity, compared against P6 with only `ensure_source_identity` between them | **read-twice-compare pair with P6 — kept**; see "designed, not implemented" below |
| P8/P9 | `litchi-doc` `ensure_current:766` and `:774` | `observed` against the retained identity, then `confirmed` against `observed`; run twice per `resolve` | **read-twice-compare pair — keep both** |
| P10 | `litchi-ppt` `text_edit::SourceSnapshot::open_with_options:423` | the whole open: one empty-splice identity plan, whose own P1/P2 supply the read-twice-compare | identity capture — keep |
| P11 | `litchi-ppt` `publish_source_operation:915`, `:990` | re-plans the identity at commit and compares with the retained one | commit fence — keep |
| I1/I2 | `litchi-doc` `open_with_options:515` and `:531` | two `SharedOleFile::open` calls on the generic `ReadAt` path | **not a duplicate — designed, not implemented** |
| I3 | `finish_overlay_plan_with_owner:1041` | reopens the composed candidate through the ordinary CFB parser, once per identity pass | the reopen proof ADR 0003 requires — keep |

### Designed, not implemented

**P7, the DOC open's third identity pass.** P6 and P7 are a read-twice-compare
pair, and the brief's rule for such pairs is to keep both. The design note worth
recording is that P7 brackets *no work*: between P6 and P7 only
`ensure_source_identity` runs (two `version()` calls and one `len()`), and P6's
own call already contains an internal read-twice-compare (P1/P2). Removing P7
would not remove a class of detection — `ensure_current` re-establishes the same
invariant at every subsequent access — but it *would* move the moment of
detection from `open` to the first read for an adapter that mutates in that
window, and that is a change in when a typed refusal happens. Predicted saving
if a future record decides that relocation is acceptable: **one of the three
identity calls**, i.e. 2 of the 6 remaining SHA-256 passes and 2 of the 6
complete source reads per generic DOC open — about a third of what this change
leaves behind.

**I1/I2, the duplicate CFB index parse on the `ReadAt` path.** The brief lists
this among the candidates to implement. It is not value-identical, and the
source says why: the comment at `source.rs:528-530` states that the reopen
"binds the directory/FAT index used below to the same bytes that supplied the
opening identity even when a custom adapter lies about its token", and
`open_owned_source:632-637` contrasts itself with exactly that. The ordering is
load-bearing. I2's bytes are read **between** P5 and P6, so the bracket
`P5 == P6` covers them; parsing the index once, before P5, would move those bytes
outside every bracket, and an adapter that serves different bytes at index-parse
time than at fingerprint time would be accepted. Eliding I1 instead would need an
index-free fingerprint entry point in `litchi-cfb` and would move the CFB
structural refusal to *after* a complete hash of the artifact, changing when a
refusal happens. Both are therefore designed and not implemented. Predicted
saving, **measured** on the untouched code path
(`results/change-0589/perf/perfstat-extra.jsonl`, `cfb-index-open`): one
`SharedOleFile::open` costs 23,023 cycles on `documentProperties.doc`, 24,341 on
`duplicate-style-names.doc` and 51,326 on `picture.doc` — 6.7%, 2.3% and 0.27%
of those fixtures' *post-change* open. It is a small item behind a real fence.

## What was changed

Three edits, all in `crates/litchi-cfb/src/overlay.rs`.

1. **`fingerprints`** replaces the unconditional second hasher with
   `let mut target_hasher = (!spans.is_empty()).then(Sha256::new);`. When the
   option is `None` the loop skips both `apply_spans` and the target update, and
   the returned target fingerprint is the source fingerprint. The source read,
   the chunking, the two `ensure_length` brackets and every error are untouched.

2. **`write_validated`** does the same for the emission scan. `apply_spans`
   still runs whenever there is a span, because it produces the bytes that are
   written; with no span it is skipped because it cannot change one. The source
   digest is still compared first and the target second, so a divergence yields
   `SourceFingerprintChanged` exactly as before rather than
   `TargetFingerprintChanged`.

3. **`OverlayOperationShape`** is made truthful. It is a content-free
   pass-shape descriptor whose `*_bytes` fields are documented as "logical bytes
   hashed"; those were `source_bytes * 2` per scan unconditionally.
   `OverlayOperationShape::new` now takes `is_noop` and reports
   `source_bytes * 1` for a no-op plan. **Pass counts and chunk counts are
   unchanged**, which is the point: the number of complete source reads did not
   move. `ValidatedOverlayPlan::is_noop()` is already public, so no new
   information is exposed. The only consumers outside `litchi-cfb` are
   `litchi-xls`'s numeric diagnostics and `tools/perf-baseline`; both carry
   effective edits, and neither asserts a `*_bytes` value.

No change was made to `crates/litchi-doc/src/body_text/source.rs` or
`crates/litchi-ppt/src/text_edit.rs` beyond their tests. Every saving on those
two paths comes from the shared primitive.

## Why it is sound

* **Value identity.** `apply_spans(bytes, offset, &[])` iterates
  `spans[partition_point..]` over an empty slice. It cannot write. Therefore
  `target_hasher` in the old code consumed, chunk for chunk, the same bytes as
  `source_hasher` and finalized to the same 32 bytes. Returning the source digest
  for both is the same value, not an approximation of it.
* **Error identity.** The two comparisons in `write_validated` keep their order
  and their variants. With no span the retained `source_fingerprint` and
  `target_fingerprint` are equal (both came from one `fingerprints` call over an
  empty span list), so the first comparison fires first, exactly as before. With
  a span, nothing changed at all.
* **No fence moved.** Complete source reads, read call counts, requested byte
  counts, `len()` calls and `version()` calls per open are **byte-for-byte
  identical** before and after on all 38 admitted fixtures
  (`results/change-0589/counts/`). Every `ensure_length`, `ensure_source_identity`
  and `check_source_version` is where it was.
* **ADR reading.** ADR 0003 says repeatedly that fingerprints are *diagnostic*
  and that "exact byte equality, not the fingerprint, authorizes application"
  (`docs/adr/0003-snapshots-edits-and-patches.md:175-176`, `:214-215`, `:240`).
  ADR 0005 requires that "mutation during a read returns `SourceChanged`"
  (`docs/adr/0005-io-memory-and-performance.md:12`). Both hold: the digest
  values are unchanged, and every mutation that was refused before is refused
  now, with the same typed error.
* **Contracts untouched.** No refusal boundary moved; no limit was relaxed; no
  `unsafe`; no new dependency; no new public item; no change to any `Refusal`,
  `OverlayError` or `OleError` variant; no I/O pattern change; no allocation
  added (the elided hasher removes one stack-resident `Sha256` state on the
  no-op branch); no threading.

## Measured

Host: AMD EPYC 9R45, 32 cores, 123 GiB, Linux 7.0.0-1012-aws, rustc 1.95.0,
valgrind 3.26.0, eight agents building concurrently. Every measured process was
pinned to **CPU 9** with `taskset`. Both legs are `--release --locked` builds of
the same scratch probe against path dependencies — before at `08d968f8e` in a
read-only checkout, after in this branch's worktree. Scenario: a fresh
`FileSource` per operation, then `SourceSnapshot::open`. Corpus: the **8** `.doc`
fixtures the DOC snapshot admits (of 57 in `test-data/`) and **all 30** `.ppt`
fixtures, which the PPT text-edit snapshot admits because it parses nothing.

### Hash passes per open (deterministic — **measured**)

Under callgrind the SHA CPUID bit is masked, so `sha2` runs its software backend
at a constant cost per artifact byte. That turns the `sha2` instruction share
into an exact pass count: 52.07 Ir per byte per complete SHA-256 pass, read off
the PPT leg.

| operation | before | after |
| --- | ---: | ---: |
| DOC `SourceSnapshot::open`, generic `ReadAt` | **12.00** passes (3 identity calls × 2 fingerprint scans × 2 hashers) | **6.00** |
| PPT `text_edit::SourceSnapshot::open` | **4.00** passes (1 × 2 × 2) | **2.00** |

Fitted on three fixtures: `picture.doc` 12.00 → 6.00, `duplicate-style-names.doc`
12.01 → 6.01, `45543.ppt` 4.00 → 2.00
(`results/change-0589/callgrind/isolation-pairs.txt`).

### Complete source reads per open (deterministic — **measured**)

Identical in both legs on every fixture and every mode
(`doc-snapshot-open` ×8, `doc-snapshot-resolve` ×4, `ppt-textedit-open` ×30;
`diff` of `counts-before.jsonl` and `counts-after.jsonl` is empty). A generic
DOC open reads the artifact 6.05–7.47 times over (6 complete identity reads plus
the two index parses and the FIB/CLX ranges); a DOC open plus one paragraph
14.40–16.64 times; a PPT text-edit open 2.02–2.53 times.
**This change removes hashing, not reading**, and that is the evidence.

### Native `perf stat`, isolation pairs (**measured**, primary)

`perf stat -r 5` on N and N+M operations in one process, differenced and divided
by M; `-e cycles,instructions`; pinned to CPU 9. All 38 fixtures, every one of
which improved.

| mode | fixture | bytes | before cyc/op | after cyc/op | Δ | instr Δ |
| --- | --- | ---: | ---: | ---: | ---: | ---: |
| doc-snapshot-open | `picture.doc` | 1,448,448 | 36,846,505 | 19,170,816 | **−48.0%** | −47.6% |
| doc-snapshot-open | `duplicate-style-names.doc` | 64,512 | 1,842,062 | 1,048,606 | **−43.1%** | −37.8% |
| doc-snapshot-open | `table-merged-cells.doc` | 17,408 | 657,035 | 440,202 | −33.0% | −24.3% |
| doc-snapshot-open | `documentProperties.doc` | 9,728 | 466,197 | 344,713 | −26.1% | −17.3% |
| ppt-textedit-open | `cryptoapi-proc2356.ppt` | 1,341,952 | 11,404,643 | 5,958,650 | **−47.8%** | −47.3% |
| ppt-textedit-open | `45543.ppt` | 385,024 | 3,288,139 | 1,721,860 | −47.6% | −46.1% |
| ppt-textedit-open | `SampleShow.ppt` | 125,440 | 1,127,840 | 614,684 | −45.5% | −41.9% |
| ppt-textedit-open | `incorrect_slide_order.ppt` | 9,728 | 156,406 | 116,498 | −25.5% | −17.2% |

Across all 38: cycles **median −36.6%**, range −48.0% to −25.5%; instructions
median −29.0%, range −47.6% to −17.0%. The full table is
`results/change-0589/perf/perfstat-summary.txt`. The saving scales with artifact
size because the fixed per-open work (two CFB index parses, the FIB/CLX reads)
does not shrink: on a 9.7 KB `.doc` the hashing is a quarter of the open, on a
1.4 MB one it is nearly all of it.

**A/A floor, same window, same counters:** a second `before` leg against the
first gives median **+0.1%**, range −0.5% to +1.0% on cycles across the same 38
fixtures. The effect is one to two orders of magnitude above the floor.

### The removed work is exactly SHA-256, and only SHA-256 (**measured**)

If the removed work is `k` complete SHA-256 passes over the artifact, then
`(before − after) / k / bytes` must be one constant: the native per-byte cost of
one pass. With `k = 6` for the DOC generic open and `k = 2` for the PPT
text-edit open it is, across all 38 fixtures and both formats:

| | median | range |
| --- | ---: | --- |
| cycles per byte per pass | **2.048** | 2.029 – 2.115 |
| instructions per byte per pass | **2.579** | 2.561 – 2.632 |

A 4% spread over a 149× range of artifact sizes is the fit; it would not hold if
the delta were anything but whole-artifact hashing. Two figures follow from it.
First, the hashing share of a native open: **51.0%–95.9% of cycles before,
34.3%–92.2% after**, rising with artifact size — so hashing is still the largest
term on large files, just half of what it was. Second, callgrind's software
backend costs 52.07 Ir/byte against 2.579 natively, i.e. it overprices SHA-256
by **20.2×** in instructions; change 0587's survey estimated ~19×, and this is
the independent confirmation. Table:
`results/change-0589/perf/native-sha-cost.txt`.

The DOC readback (open plus the first `paragraph()`, which adds four more
identity passes through `ensure_current`) falls too:
`documentProperties.doc` 1,020,526 → 734,297 cyc (−28.0%), `lists-margins.doc`
1,080,635 → 768,781 (−28.9%), `duplicate-style-names.doc` 4,222,439 → 2,374,368
(−43.8%).

### Paired timing (**measured**)

Order **A1 B1 B2 A2** (before, after, after, before), then two further `before`
legs for the A/A control, all in the same window on CPU 9; per-operation
`Instant` samples, warm page cache.

| scenario | leg | n | p50 µs | mean | p95 | p99 |
| --- | --- | ---: | ---: | ---: | ---: | ---: |
| `picture.doc` (1.4 MB) | before | 80 | 8,267.5 | 8,270.6 | 8,295.2 | 8,324.1 |
| | after | 80 | 4,291.6 | 4,292.9 | 4,306.1 | 4,331.7 |
| `duplicate-style-names.doc` (64 KB) | before | 120 | 411.6 | 412.6 | 416.7 | 432.0 |
| | after | 120 | 232.8 | 234.5 | 239.2 | 254.9 |
| `cryptoapi-proc2356.ppt` (1.34 MB) | before | 80 | 2,555.7 | 2,556.2 | 2,562.5 | 2,566.4 |
| | after | 80 | 1,326.8 | 1,327.4 | 1,333.8 | 1,336.0 |
| `45543.ppt` (385 KB) | before | 120 | 738.4 | 738.1 | 742.5 | 746.6 |
| | after | 120 | 384.4 | 385.8 | 391.1 | 410.9 |

Both directions, at p50: after is **−48.1%**, **−43.4%**, **−48.1%**, **−47.9%**
against before; before is **+92.6%**, **+76.8%**, **+92.6%**, **+92.1%** against
after. At p99: −48.0%, −41.0%, −47.9%, −45.0%.

**A/A floor in the same window:** p50 +0.1%, −0.1%, −0.1%, +0.0%; p99 +0.7%,
+2.2%, −1.8%, +2.5%. That is well inside the host's stated floor (p50 ≈ 4%,
p99 ≈ 14%) and far below the effect, so the medians above are reportable.
**No scenario regressed.**

### Callgrind (**measured**, secondary)

Isolation pairs, `--cache-sim=no --branch-sim=no`. Callgrind prices SHA-256 in
software, so these totals overstate hashing by roughly five times its native
share and rank work rather than latency; they are reported because they yield the
exact pass count above.

| fixture | before Ir/op | after Ir/op |
| --- | ---: | ---: |
| `picture.doc` | 912,550,812 | 460,069,709 |
| `duplicate-style-names.doc` | 41,106,009 | 20,932,549 |
| `45543.ppt` | 81,167,244 | 41,069,470 |

The before figures reproduce change 0587's survey to within 0.003%
(912,551,293 / 41,107,039 / 81,168,004), which is the provenance link between the
two records.

## Correctness evidence

### A corpus differential over every OLE2 artifact

Change 0582's differential harness **does not apply**: it is a `soapberry-zip`
harness that builds only that crate and drives the four public entry points
reaching `strict_layout_for` in the ZIP reader. It never touches `litchi-cfb`,
`litchi-doc` or `litchi-ppt`, and nothing in its corpus is a compound file.

A differential in the same spirit was built instead and is retained in
`results/change-0589/differential/`. For each of the **87** `.doc` and `.ppt`
artifacts under `test-data/`, both legs print one line carrying:

* the empty-splice identity plan's source and target fingerprints, its `is_noop`
  flag, the SHA-256 of the bytes it publishes, and the publish report's two
  fingerprints;
* an **exact byte no-op** splice plan over the first non-empty stream: both
  fingerprints, the span count, and the SHA-256 of what it publishes;
* an **effective** one-byte splice over the same range: both fingerprints (which
  differ from each other), the span count, and the published SHA-256;
* the DOC snapshot fingerprint, or the exact typed refusal;
* the PPT text-edit snapshot fingerprint, or the exact typed refusal.

**The two reports are byte-identical.** Every digest, every span count, every
published artifact and every typed refusal string is unchanged across 87
artifacts and all three span shapes. The lines also show `identity_src ==
identity_tgt == identity_out`, i.e. the digest a no-op plan reports is SHA-256
over the artifact's own bytes.

### Tests added

`crates/litchi-cfb/src/overlay_tests.rs` (six new tests, and one extension of
`operation_shape_matches_generic_and_owned_overlay_policy`):

* `noop_plan_reports_the_source_digest_as_both_identities` — a no-op plan's two
  fingerprints both equal SHA-256 of the artifact, computed independently in the
  test; an effective plan over the same artifact keeps two distinct digests and
  the same source digest.
* `noop_plan_retains_every_complete_source_scan` — a no-op plan performs the same
  number of complete-artifact reads as an effective one, reports the same
  fingerprint scan and chunk counts, and exactly half the logical bytes hashed.
* `noop_plan_catches_a_stable_token_mutation_between_planning_scans` — the
  boundary read is discovered at run time from the request sizes, then a byte of
  an *unselected* stream payload is flipped after it so the composed reopen and
  the precondition both still pass and only the second fingerprint can object;
  the result is `SourceFingerprintChanged`.
* `noop_direct_write_catches_a_late_stable_token_mutation` — a mutation after the
  final emission read still yields
  `IncompleteOutput { CompleteUnflushed, SourceFingerprintChanged }` with the
  full artifact in the sink and the same read count as before.
* `noop_direct_write_catches_a_mutation_inside_the_emission_scan` — the single
  emission hasher still diverges from the retained identity mid-flight.
* `noop_plan_publishes_the_exact_source_bytes` — a no-op publication emits the
  source byte for byte and reports one digest for both identities.

The existing hostile adapter `MutableSource` gained a `mutate_offset` field so a
test can flip a byte that is not in the CFB header/FAT region; its default is the
byte 700 every existing test uses, so no existing test changed behaviour.

`crates/litchi-doc/src/body_text/source.rs` (three new tests and a
`ScheduledMutationSource` adapter that fires after an exact read ordinal):

* `a_stable_token_mutation_after_every_open_read_is_still_refused` — **sweeps
  every read ordinal of a clean open**. For every ordinal at or after the first
  complete identity pass the run must end in `SourceFingerprintChanged`, from the
  open or from the mandatory readback that follows it; for the earlier ordinals —
  where the mutation precedes identity capture and so has nothing to disagree
  with — the snapshot must be bound to exactly the bytes that now exist, checked
  by reopening them. The refusal count is asserted exactly, so a lost fence point
  fails the test rather than silently shrinking the count.
* `a_stable_token_mutation_after_every_readback_read_is_still_refused` — the same
  sweep across `ensure_current`'s bracket around a paragraph read. Every ordinal
  but the last must be refused; the last is the one no finite number of passes
  can bracket.
* `the_retained_identity_is_the_complete_source_artifact_digest` — the owned path
  (one identity pass) and the generic path (three) retain the same digest, and a
  no-op commit publishes the source byte for byte.

`crates/litchi-ppt/src/text_edit.rs` (two new tests and the same adapter):

* `a_stable_token_mutation_after_every_open_read_is_still_refused` — the same
  exhaustive sweep. PPT's `ensure_current` checks the version token and length
  only, so where the open survives, the test drives the **commit** fence and
  requires `SourceFingerprintChanged` there.
* `the_identity_plan_reports_one_digest_for_both_identities` — the empty-splice
  plan's two fingerprints are equal and match the digest retained by both the
  owned and the generic snapshot open.

### Gates

Run in the worktree at `perf/0589-ole2-snapshot-fingerprint-passes`; tails in
`results/change-0589/gates.txt`.

| gate | result |
| --- | --- |
| `cargo fmt --all --check` | pass |
| `cargo clippy -p litchi-cfb -p litchi-doc -p litchi-ppt --all-targets` | pass, no warning (workspace lints are `deny`) |
| `cargo clippy -p litchi-xls -p litchi-ole-common --all-targets` | pass — the two crates outside the change that consume `OverlayOperationShape` and the overlay |
| `cargo test -p litchi-cfb -p litchi-doc -p litchi-ppt` | pass — 77 test binaries, 0 failures; litchi-cfb 322, litchi-doc 995, litchi-ppt 1,077 library tests |
| `cargo test -p litchi-xls -p litchi-ole-common` | pass |
| `cargo doc -p litchi-cfb -p litchi-doc -p litchi-ppt --no-deps` | pass (rustdoc lints are `deny`) |
| 87-artifact fingerprint differential | byte-identical between legs |

No pre-existing failure was encountered.

## Validation preserved

Nothing about validation changed. `plan_same_length_stream_splices` still runs
its full precondition comparison; `finish_overlay_plan_with_owner` still reopens
the composed candidate through the ordinary CFB parser and still re-brackets a
generic source with a second complete fingerprint; `write_to` and `save` keep
every preflight and the post-emission scan; the DOC open keeps all three identity
passes, both CFB index parses and every `reject_fib_policy` /
`reject_container_components` check; `ensure_current` keeps both of its passes.
Validation still never mutates, and an exact no-op is still exact — the
differential proves the published bytes are unchanged on all 87 artifacts.

## Limitations — what is not claimed

* **No claim is registered.** `performance_claim: none`.
* The saving is in **CPU only**. Reads, bytes read, syscalls, allocations and
  peak RSS are unmeasured except for the read counts, which are proven
  unchanged. No allocation, RSS, cold-cache, physical-device or range-source
  measurement was taken.
* Scoped to this host, this build (`rustc 1.95.0`, `--release --locked`, default
  release profile), a warm page cache, a single pinned core, and the two
  scenarios named — `SourceSnapshot::open` on the 8 admitted `.doc` fixtures and
  `text_edit::SourceSnapshot::open` on all 30 `.ppt` fixtures, plus the DOC
  readback on the 4 `.doc` that admit a paragraph read. Machines without SHA-NI
  would see a *larger* relative saving; none was measured.
* No registered `perf-baseline` selector opens either path (change 0587's
  measurement blocker), so nothing here moves a harness scenario. The commit and
  save paths gain the same halving on a no-op publication, which is **not**
  measured here — only the CFB unit tests cover it.
* The largest fixture either path admits is 1.45 MB. Behaviour at DIFAT scale is
  unmeasured.
* Callgrind's figures price SHA-256 in software and are used only to count
  passes, never to claim latency.
* **P7 and I1/I2 are designed, not implemented.** Their predicted savings above
  are a model plus one measurement of the untouched index parse; neither has been
  demonstrated end to end.
* The A/A floor stated here (p50 ≈ 0.1%, p99 ≤ 2.5%) is this window's floor on a
  pinned core, tighter than the host's standing figure; it is not a claim about
  the host in general.
* **One correction to an earlier record.** Change 0587's DOC-1 entry says "this
  is why change 0586 measured exactly zero: a read hint cannot register against
  41 M instructions of hashing". That reading is wrong about 0586: its zero was
  *structural*, not masked — the DOC paragraph hint removed 0 of 3,991 chain
  links on 8 of 8 fixtures, so there was nothing for hashing to hide. The
  hashing term does explain why any *future* DOC read-path change is hard to
  attribute, and that is the sense in which this change helps: the hashing share
  of a native open falls from 51.0%–95.9% of cycles to 34.3%–92.2%. It does not
  revive 0586's hint, which remains correctly rejected.

## Retained evidence

[`results/change-0589/README.md`](results/change-0589/README.md) — the probe
source, both fixture censuses, the deterministic count reports for both legs, the
`perf stat` reports and their summary, the paired-timing samples, the callgrind
extraction, the 87-artifact differential reports, every script, `decision.json`,
`gates.txt` and `log-sections.md`.
