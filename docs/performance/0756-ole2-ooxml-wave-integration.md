# 0756 — the OLE2/OOXML wave of 2026-09-22/23: fifteen records integrated, each changed by its review

Status: integration record, `performance_claim: none`. Each implementing record
owns its measurements, corpora, floors and limitations. This record lists what
landed, what the independent reviews changed, what was withdrawn, and what the
owner is asked to decide. The wave-wide sweep below is descriptive (9-sample
processes, one core, a host shared with other agents), not a claim.

OLE2 and OOXML remain the active priority. ODF optimization stays deferred until
that goal completes; iWork is excluded and untouched.

Wave base `009d515bef`; integration tip `db07e4d53f` (code; this record's commit follows it). Records 0742–0755 and 0757
were implemented by Opus subagents in separate worktrees, reviewed by an
independent adversarial reviewer before merge (most needed a second round), and
cherry-picked onto `feat/office-format-completeness` as linear history. Every
implementing branch `perf/07xx-*` is kept with its original commits. Record 0743
was merged as its net diff (production, 0755's fix and documents as three
commits) because its withdrawn memo commit and that commit's revert cancel out.

The untracked `docs/performance/results/change-0741/` packet (an unfinished
experiment by a previous agent) and `docs/UNIFIED_OPS_API_DESIGN.md` are
preserved untouched and uncommitted. 0742 implements the optimization that
packet was preparing; the packet itself is left for its owner.

## How the wave was chosen

A coordinator sweep of about 100 OLE2/OOXML harness cases on the base
([base-sweep](results/change-0756/base-sweep/)), four read-only surveys and a
timed-region profile ([profile-r2/REPORT.md](results/change-0756/profile-r2/REPORT.md))
ranked the work. The sweep exposed three large costs the queue did not list:
eager XLSX dense-sheet reads and commits (29 ms for the first cell, 152 ms for a
one-cell commit+save on a 384 KB workbook), the PPTX semantic read/edit path on a
100 × 100 deck (50 ms full text, 269 ms for a 1% edit), and owned PPTX media
cross-copy at 22 times the source-backed route's cost on the same content.

## What landed

Headlines are each record's own paired measurement against its own base (see
the record for corpus, samples, core, binaries and every regression flag).

| Record | Change | Headline |
|---|---|---|
| [0742](0742-pptx-owned-cross-copy-media-transfer.md) | Owned PPTX cross-copy publishes copied images from the source's verified compressed bytes; durable format `LPCP0004` (older versions refused by name) | `pptx_cross_copy_media_rich_lifecycle` 410.1 → 183.0 ms (paired 0.444) |
| [0743](0743-pptx-semantic-text-and-edit-path.md) | PPTX semantic read/edit paths stop re-reading validated bytes (slide-root memo withdrawn, see below) | 100 × 100 deck: full text 50.69 → 27.14 ms (−46.5%), 1% edit/save 269.6 → 144.8 ms (−46.0%), one edit −18.7%, no-op −16.4% |
| [0744](0744-xlsx-eager-workbook-cell-path.md) | Strict byte lane for simple XLSX `sheetData` bodies; dense reduced readback | dense-wide: first cell 27.87 → 7.82 ms (−71.9%), one-cell commit+save 149.2 → 60.5 ms (−59.8%) |
| [0745](0745-ppt-lazy-artifact-digests.md) | PPT slide-order artifact digests deferred to durable serialization; editor streams read once | `45543.ppt` slide removal 1,079 → 654 µs (−39.5%); no-op commit −85.9% |
| [0746](0746-xls-validation-only-parse.md) | XLS edit owners validate through a validation-only mode of the complete reader; the validated render is published | `54016.xls` open 10.46 → 4.87 ms, generic commit 23.24 → 10.12 ms |
| [0747](0747-xlsx-publication-audit-reuse.md) | A replaced XLSX part's replacement audit is proved from its original's audit (window proof) | `xlsx_source_backed_cell_values_one_edit_save` −6.6% (medium), −6.3% (dense-sparse) |
| [0748](0748-cfb-overlay-fingerprint-reuse.md) | Sealed owned CFB overlay plans hash their artifact once | `xls_visibility_eager_edit_save` 25.33 → 5.47 ms (0.216), measured on a base without 0746 |
| [0749](0749-cfb-reuse-plan-validation.md) | CFB Reuse-plan validation compares streams in place, linear in mini streams | Reuse `write_to`: `45543.ppt` −40.9%, `FloatingPictures.doc` −32.1%; corpus median −12.5% instructions |
| [0750](0750-xml-audit-well-formedness-gaps.md) | The source XML audit refuses ADR 0006 well-formedness gaps (correctness) | DOCX one-edit save +2.67% (+43 µs); no false refusals against expat/libxml2 |
| [0751](0751-pptx-cross-copy-apply-digest-reuse.md) | Owned cross-copy stops re-hashing proven bytes; the retained archive is shared | commit phase 88.9 → 25.8 ms (0.29) on top of 0742 |
| [0752](0752-streaming-writer-small-write-batching.md) | Budgets charged without per-level handles; scoped reservations; CRC-32 staging; DOCX escaping by runs | `docx_streaming_create` large 91.47 → 48.60 ms (−46.7%) |
| [0753](0753-legacy-fresh-writer-text-paths.md) | DOC/PPT/XLS fresh writers encode text once into final buffers and hash shared strings once | payload-heavy: DOC 5.68 → 1.63 ms, PPT 4.50 → 0.86 ms, XLS 6.23 → 2.32 ms |
| [0754](0754-docx-semantic-edit-and-text-path.md) | Cheaper DOCX admission scan; one audit per preserving publish via `VerifiedSource`; bounded namespace lookups | large: no-op −60.6%, one-edit −47.5%, full text −36.7%; a 30,600-declaration part 436 → 14 ms |
| [0755](0755-pptx-nested-text-run-panic.md) | A nested `a:t` is refused instead of panicking (correctness) | the process no longer aborts |
| [0757](0757-xls-fresh-writer-sst-determinism.md) | Deterministic XLS shared-string order; typed refusal of strings BIFF8 cannot hold (correctness) | 8 byte streams across 8 processes become 1; ~12% on string-heavy sheets is the price of the sort |

## What the reviews changed

- **0742.** Review 1 found redo-after-undo refused (ADR 0003) and a
  reader-accepted member failing the whole copy. Review 2 found the capture
  charge refusing copies recompression performs and caller-defined parts
  silently changing behaviour. Transfer now depends only on recorded revisions
  and published bytes.
- **0743.** A snapshot memo of notes-root classifications (commit `99ce9c5e34`;
  one-edit commit 21.4 → 2.1 ms in its probe) was **withdrawn**: ADR 0005's
  2026-09-16 amendment admits only digest memos and typed exhaustion errors.
  Proposed [ADR 0032](../adr/0032-snapshot-derived-value-memos.md) asks the owner
  to widen it. The review also found the pre-existing panic fixed as 0755.
- **0744.** Lane admission on a foreign-namespace `sheetData`, a changed first
  error above 16 MiB and a dropped omitted-cell collision check were fixed; a
  HashMap-ordered shared-formula traversal was made deterministic.
- **0745.** A test that passed for the wrong reason was corrected; bytes and
  digest memo became inseparable by construction.
- **0746.** A sparse-band bitmap that made one crafted layout 1.7 times slower
  became a per-occupied-row map (2.2 times faster than base on that layout).
- **0747.** The window proof's unstated invariant is documented, and every window
  proof is re-derived by a debug cross-check.
- **0748.** The emission-hash decision now comes from the plan's seal and fails
  closed in release builds.
- **0749.** The first version made mini-stream validation quadratic (up to 14
  times slower); the fix is linear and faster than base on every measured case.
  The coordinator narrowed its new decline to the two plan-derived limits
  (`d61a4d7c79`).
- **0750.** A differential against expat and libxml2 found no false refusals, but
  one CPU denial of service in the new duplicate-attribute check (0.02 s → 182 s
  on a crafted part); namespace names now carry hashed identities.
- **0751.** Adopting a snapshot memo could pin a caller-defined part's payload;
  every adoption now re-projects onto the facade's own allocations.
- **0752** and **0753.** Nits only (stronger tests, `#[must_use]`, record
  wording); 0753's review sharpened the severity of the truncation 0757 fixes.
- **0754.** The prefix index ignored default bindings (an adversarial document
  50% slower than base); it now counts every binding and builds lazily. The
  review also found a pre-existing gap — the OPC writer audited `blob()` but
  published `blob_arc()` — fixed in the same branch.
- **0757.** The review confirmed determinism: 16 processes × 8 threads gave
  one output, and the packed sort key cannot collide. It also confirmed the
  limits exact at N−1, N and N+1 in UTF-16 units. It found number-format and
  defined-name refusals raised only at write time, leaving the writer unable to
  write; they are now refused at registration without mutating the writer. It
  also found further pre-existing length-field bugs, which are queued below.

Four implementers and one profiler stalled for hours at their final cleanup
step (apparently waiting on a background command's notification); the
coordinator stopped them, finished their commits and deletions, and resumed
three of them for review fixes. The briefing now requires long commands in the
foreground and commits before final deletions.

## Correctness fixes found on the way

0750 (ADR 0006 well-formedness gaps in the publication audit), 0755 (a nested
`a:t` panic that aborted the process), 0754's audited-handle gap, 0757
(nondeterministic XLS shared-string order; silent string truncation, including
surrogate pairs split so that litchi's own reader refused the workbook),
deterministic error ordering in 0744 and 0746, and 0749's decline of a
plan-derived limit that used to fail a save outright.

## Owner decisions requested

1. **ADR 0032 (proposed):** admit snapshot memos of small derived values of
   payload bytes under the digest-memo amendment's conditions (evidence: 0743).
2. **0751's reading of ADR 0005's memo amendment (coordinator ruling):** a
   facade may fill an empty memo slot from a capture of its current graph,
   provided every kept entry is re-projected onto the facade's own allocations.
   The ADR text is unchanged; please confirm or reject the reading.
3. **0646's gate G5:** carrying the retained archive's digest with the plan
   instead of recomputing it at application would save about 15 ms per
   media-rich commit (0751).
4. **New deterministic output bytes for creation from scratch:** zlib-rs output
   depends on input chunking (0752), so coalescing writes or changing the
   per-member deflate strategy (PPTX streaming writes 16,421 tiny members)
   changes bytes. Accepting new, still deterministic bytes would unlock about
   10 ms on large DOCX streaming and more on PPTX streaming.
5. **Save durability** (0651 row 6, still open): an opt-in policy to skip or
   weaken fsync on atomic saves; about 5 ms of a 5.5 ms small-DOCX save.
6. **Shared-budget leases** (0752): charging budgets per chunk rather than per
   write changes what sibling holders of a shared budget may observe.

## Follow-up queue

- quick-xml's own duplicate-attribute check uses an unkeyed hash, a pre-existing
  hash-flooding exposure in every parser (0750 review).
- BOM-prefixed parts: XLSX scanner spans and PPTX slide spans are three bytes
  early, so edits are refused rather than corrupted (0744, 0755).
- The CFB structural reparse clears a table-sized map per stream, the remaining
  super-linear term in validation (0749).
- XLSX commit+save is now 60% deflate and 12.6% audit (0744): parallel deflate of
  large regenerated members under an explicit execution context.
- DOCX: the remaining full source audit at commit (~2 ms), multi-window proofs
  for 1% edits, and the source-backed route's duplicate original audit (0754).
- PPTX: the per-capture notes-graph validation scan (~0.24 ms per slide, 0743)
  and per-member deflate in streaming creation.
- XLS: the validation-only walk's store-forwarding stall and per-string UTF-16
  temporaries (0746).
- XLS fresh writer (0757 and its review):
  - Internal-hyperlink lengths are counted in chars but written as UTF-16. A
    sheet named `R😀` makes litchi's reader drop the link and that sheet's cells.
  - AutoFilter string lengths are UTF-8 bytes.
  - Font names are silently cut to 31 units.
  - A font added through `add_cell_style` gets index 4, which the reader refuses.
  - More than 210 custom number formats exceed the reader's cap.
  - A bad formula is refused only at write time.
  - Invalid sheet-name characters are not refused.
  - PivotTable lengths count chars.
  - XFEXT, STYLEEXT and CRN lengths may wrap (unverified).
  These are correctness items for the next wave.
- DOC fresh writes are now 37% CFB output zero-fill (0753).
- `verify_authored` sites and `SourceXmlPart` replacements (0750).
- `opc_mutated_save` on incompressible payloads (60 ms; not profiled).

## Wave-wide sweep (descriptive)

[`results/change-0756/sweep/`](results/change-0756/sweep/README.md) compares the
base and the tip `ddf788eb80` (records 0742–0755; 0757 is a correctness change
measured in its own record). Both harness binaries were built with the same
command, from checkouts and target directories of equal path length. There are
four processes per arm in A B B A A B B A order, with 9 samples each, on one core.
Ratios are final/base, from the median of the process p50s.

| Format (rows) | Geomean ratio | Largest moves |
|---|---:|---|
| DOCX (14) | 0.675 | no-op edit/save 3.82 → 1.51 ms; streaming create 91.5 → 47.4 ms; full text 3.13 → 1.91 ms |
| PPTX (17) | 0.813 | media-rich owned cross-copy 416.7 → 121.8 ms; 1% edit/save 272.1 → 146.0 ms; full text 50.6 → 27.4 ms |
| XLSX (18) | 0.551 | dense first cell 28.4 → 7.8 ms; one-cell commit+save 152.7 → 60.9 ms; eager cell edit 38.2 → 15.3 ms |
| DOC (9) | 0.864 | payload-heavy fresh write 5.92 → 1.90 ms |
| PPT (9) | 0.813 | payload-heavy fresh write 4.87 → 2.35 ms; one-edit save 0.230 → 0.183 ms |
| XLS (12) | 0.502 | visibility edit 24.95 → 2.79 ms; numeric edit 2.38 → 0.29 ms; one-edit save 3.34 → 0.76 ms |
| OLE2 common (8) | 0.991 | unchanged |

The media-rich cross-copy writes different bytes by design, because 0742
transfers source-compressed images. The unweighted geometric means are across
the rows listed in the sweep summary only.

Flags above 5% are disclosed, not averaged away:

- `pptx_source_backed_cross_copy_media_rich_lifecycle` 1.27: each process
  settles into one page-fault mode. Every mode seen in both arms gives the same
  time in both, and a separate run of the case alone shows no arm difference.
- `doc_semantic_open` (large) 1.074: a coordinator follow-up measured identical
  instructions per iteration (28.32 M against 28.31 M) and equal p50s in a fresh
  ABBA run ([docopen-check](results/change-0756/sweep/docopen-check/docopen-large.txt)).
- `ppt_fresh_write_to` (payload-heavy) is 0.483 here but 0.191 in 0753's
  isolated processes: in the sweep it shares a process with the DOC and XLS
  writer cases, and 0757 observed the same case follow heap state rather than
  code (0.79 at one sample count, 0.997 timed alone). The direction is
  consistent; the magnitude depends on process heap history.
- `pptx_semantic_open` (large) 1.055 and `xls_semantic_full_cell_scan` (large)
  1.020: 0743's own re-measurement shows identical instruction counts for open.
  These are code-layout sensitivities, which this host shows at up to ±40%
  (0746) and ±15% (0745).

## Validation

All sixteen gates of change 0675's runner pass on `60faefaa5f`, the tip
before 0757's review follow-up ([logs](results/change-0756/gates/)):

| Gate | Result |
|---|---|
| fmt, check, clippy (`--lib`, `-D warnings`), rustdoc (`-D warnings`) | pass |
| tests (14 crates) | 11,352 passed, 0 failed, 75 ignored |
| facade | 382 passed, 7 ignored |
| harness | 553 passed, 1 ignored (759.6 s) |
| allocator, facade-polyglot | 5 and 104 passed |
| claims (strict and structural), gate-tests, report, coverage, non-iWork, boundaries | pass |

After 0757's registration-time fix (`db07e4d53f`), the affected gates were
re-run ([logs](results/change-0756/gates/after-0757-followup/)): fmt, check,
clippy, facade (382) and rustdoc pass, as do `litchi-xls` (1,539 passed) and a
harness `--all-targets` check. The implementer's own gates also cover
`litchi-doc` (1,206) and `litchi-ppt` (1,235; 1,246 with encryption).

Two earlier integration runs over partial merges passed their gates too:
12,160 tests after 0744–0747, and 12,322 after 0742–0753 and 0755. The only
lint failures under `--all-targets` are four that predate the wave:
`xls_query_index_cache.rs:546` and three `err_expect` in `litchi-pptx`
`opened/tests.rs`.

## Cleanup

[cleanup.json](results/change-0756/cleanup.json) lists every removal:
- the wave's worktrees, build directories and scratch, after the cited evidence
  was copied into this packet;
- obsolete scratchpads and build directories found at wave start;
- about 288 GB of stale incremental caches and example binaries in the main
  checkout's build directory, removed mid-wave to keep the disk from filling.

Implementers, reviewers and testers deleted their own build directories, as
their packets record. Kept: the `perf/07xx-*` branches and the main checkout's
build directory. Untouched: `docs/UNIFIED_OPS_API_DESIGN.md`, the unfinished
`change-0741` packet and its build directory, and other agents' worktrees and
`/tmp` directories.
