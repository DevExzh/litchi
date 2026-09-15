# 0591: the ordinary DOCX edit scans the main part twice, not four times

Status: retained. `performance_claim: none` — this record carries deterministic
`Snapshot::from_xml` call counts and callgrind instruction counts on both legs,
plus paired medians reported beside the A/A floor measured in the same window.
They are reported as evidence, not registered as claims.

OLE2 and OOXML remain the active priority. ODF optimization stays deferred until
that goal completes; iWork is excluded.

This implements parts (a) and (b) of item **DOCX-1** in the change
[0587](0587-remaining-opportunity-survey.md) queue (rank 4). Part (c), compacting
only the replacement fragment, is **not** implemented: it changes output bytes
and needs a frozen design record and an owner decision, both stated below.

## The mechanism

0587 measured that a one-paragraph edit and save scans the whole main document
**four times**, and that `Snapshot::from_xml` is about 64% of the timed
instructions of `docx_semantic_one_edit_save`. Reading the route at
`08d968f8e`, the four scans are:

1. `Package::document_snapshot` (`package/package/document.rs:33`) copies
   `main.blob()` and `Snapshot::from_xml` scans it — the snapshot the edit is
   staged against.
2. `Edit::replace_paragraph_text` (`document/transaction.rs:1885`) splices the
   rewritten paragraph into a fresh `Vec<u8>` and calls `with_rewritten_xml`,
   which rescans the whole part to rebuild the layout.
3. `Edit::commit` (`:4429`) runs `compact_changed_document_xml` over the whole
   part and calls `with_rewritten_xml` again.
4. `Package::apply_document_patch` (`package/package/document.rs:50`) calls
   `document_snapshot` a second time — another copy and another scan — and
   hands the result to `Patch::apply`, whose only use for it is
   `Snapshot::same_source`, a byte comparison. The snapshot is dropped in the
   same statement.

Scan 4 exists to compare bytes. Scan 2 exists to rebuild a layout that the
splice already determines. This change removes both. Scan 3 belongs to
compaction and is part (c); scan 1 is the snapshot the caller asked for.

## What was changed

### (a) The patch's exact-source proof reads the bytes, not a rebuilt snapshot

`Patch` gains `target_for_exact_unmanaged_source(&self, bytes: &[u8]) ->
Option<Snapshot>` (`document/transaction.rs`). It reproduces exactly what
`Patch::apply` decides when its `source` argument is the snapshot
`document_snapshot` would have built: `Snapshot::same_source` reduces to equal
bytes plus `None` on both identity sides, and `document_snapshot` always yields
`XmlStorage::Owned`, whose identity is `None`. So the proof is
`self.before.retains_exact_unmanaged_xml(bytes)` — no source identity, equal
length, then pointer identity or a byte comparison, the idiom
`reuse_if_source_xml_matches` already uses on the source-backed path. On a hit
the patch hands back `after` when it changed the bytes and `before` when it did
not; `before` is byte-equal and identity-equal to the snapshot `apply` would
have cloned, so the two are the same value and the caller cannot tell them
apart, except that the returned snapshot now shares an existing allocation
instead of owning a fresh copy of the part.

`Package::apply_document_patch` asks for that proof through
`document_patch_source`, which runs `ensure_story_opc_current` and
`main_document_part` first, in the order `document_snapshot` runs them, and
falls back to `document_snapshot` plus `Patch::apply` whenever the byte proof
does not hold. The fallback is why no refusal moved: a patch from a foreign
lineage, a package whose main part was replaced, and a package whose main part
no longer parses all reach the same code and the same typed error they reached
before. `Arc::ptr_eq` is not available at this boundary — `Part::blob` yields a
borrowed slice — so the pointer half of the proof compares the slices' data
pointers. On the eager route it never hits today, because `document_snapshot`
copies the blob; the comparison it guards is a `memcmp`, not a scan.

### (b) A same-structure paragraph rewrite derives its layout

`Snapshot::with_spliced_paragraphs(xml, splices)` replaces
`with_rewritten_xml` at the two direct-body paragraph rewrite sites — 
`replace_paragraph_text` and `replace_body_paragraph_texts`. It derives the new
layout from the retained one and falls back to `with_rewritten_xml` — the
unchanged full rescan — whenever the derivation is not provably exact.

A splice is accepted only when its byte range is *exactly* one entry of the
retained paragraph layout and `preserves_body_child_shape` holds between the
fragment it replaces and the fragment replacing it. The new layout is then the
old one with the rewritten paragraphs resized, every later direct-body range
shifted by the running length delta, and `content_end` shifted by the total: a
direct-body paragraph can never be enclosed by a table or a block content
control, because `scan_document` records all three only at `body_depth + 1`,
and every direct body child ends at or before the body-final `w:sectPr` or
`</w:body>`. Table and block-control lists that no splice reaches keep their
`Arc` rather than being rebuilt, which is the common case — the harness corpora
and most fixtures have none.

## Why it is sound

**The layout is derived, not guessed.** `scan_document` produces five values.
`conformance` comes from the `w:body` start tag, which no splice touches.
`paragraphs`, `tables` and `block_controls` are the byte ranges of the body's
direct children, in document order, disjoint; replacing one child's bytes with
`n` bytes moves every later offset by exactly the same delta and changes only
that child's length. `content_end` is the offset of the body-final `w:sectPr`
or of `</w:body>`, both of which follow every direct child.

**The scan is also a validator, so the fragment carries the proof.**
`scan_document` reaches four kinds of conclusion about the inside of a body
child: how the child is classified, from the resolved namespace and local name
of its root element; the running element count against `MAX_DOCUMENT_NODES`;
the nesting depth against `MAX_DOCUMENT_DEPTH`; and the refusals for DTDs,
processing instructions and unbalanced nesting. `preserves_body_child_shape`
parses only the two fragments — paragraph-sized, not document-sized — and
requires that both are exactly one balanced element spanning all their bytes,
that the replacement has no more elements and no more nesting, and that the two
root start tags are **byte-identical**. The bytes around the splice are
unchanged, so the namespace bindings in scope at that root are unchanged, so an
identical start tag resolves identically and the child keeps its
classification; a total that cannot grow cannot cross a ceiling it was already
under; and a DTD or processing instruction in the replacement is refused
outright. Every verdict the scanner would reach on the spliced document is
therefore the verdict it reached before, which is what the differential tests
check empirically on every fixture and generated case.

**Everything the proof does not settle rescans.** `spliced_layout` returns
`None` — never an error — for a splice that is not exactly a paragraph range,
for an out-of-order or overlapping splice list, for any offset arithmetic that
does not fit, for a document over `MAX_DOCUMENT_XML_BYTES` (so the typed
`Limit` error is still raised by `from_xml`, with its identity), and for any
shape the fragment check declines. `None` routes the caller into
`with_rewritten_xml` unchanged.

**Error identity is preserved on both parts.** Part (a) keeps the original
snapshot route for every case the byte proof does not settle, so a stale patch,
a foreign-lineage patch and an unparsable main part each keep the error they
had. Part (b) keeps the original rescan for every case the shape proof does not
settle.

**ADR reading.** ADR 0005's mandatory validation is intact: no check is removed,
weakened, reordered or made conditional. The semantic readback after every
paragraph rewrite still runs against the produced snapshot, and `same_source`
still decides every patch application — part (a) changes *what it is compared
against*, from a copy of the part to the part, not *whether* it is compared.
ADR 0003 is intact: `commit()` still returns the same `Commit`, the snapshot is
still immutable and cheap to share, and the patch is still exact-source-checked
and reversible. ADR 0006 preservation is intact because no output byte changes;
compaction is untouched, which is exactly why part (c) is not here. No new
`unsafe`, no weakened limit, no new ambient I/O, no global pool, no public
leakage of archive types, locks or executors. Managed and source-backed routes
are untouched: both rewrite entry points dispatch to their managed variants
before reaching the spliced path, and `apply_document_patch`'s proof requires
the absence of a source identity, which every managed or source-backed snapshot
carries.

## Measured

Host: AMD EPYC 9R45, 32 cores, 123 GiB, Linux 7.0.0-1012-aws, rustc 1.95.0
(workspace pin), valgrind 3.26.0. Both legs `--release --locked`; every measured
process pinned to CPU 11 with `taskset`; seven other agents were building and
measuring on the same host.

### Deterministic counts — callgrind isolation pair on the harness

`litchi-perf-baseline --warmup 0 --samples N --case <case>` at N=1 and N=3,
differenced and halved, so corpus construction and process start-up cancel.
Each case runs all three shapes, so a "lifecycle" below is one sample of the
24-, the 200- and the 10,000-paragraph document together.

| case | `Snapshot::from_xml` calls | its inclusive Ir | `apply_document_patch` inclusive Ir | whole iteration Ir |
| --- | --- | --- | --- | --- |
| `docx_semantic_one_edit_save` | 12 → **6** | 280,511,387 → **140,478,384** | 72,914,944 → **1,657,943** | 713,306,638 → **573,646,143** (−19.58%) |
| `docx_semantic_one_percent_edit_save` | 12 → **6** | 280,660,423 → **140,272,200** | 72,883,259 → **1,646,556** | 722,550,900 → **582,822,411** (−19.34%) |
| `docx_semantic_noop_edit_save` | 6 → **3** | 140,489,592 → **70,183,341** | 71,803,608 → **439,390** | 436,296,030 → **365,352,877** (−16.26%) |

Per document, that is four scans → two for an edit and two → one for an exact
no-op. `Package::document_snapshot` calls fall 6 → 3 in every case. The new
fragment check costs 18,126 Ir over 6 calls on the one-edit case and 620,948 Ir
over 206 calls on the one-percent case — 0.4% of the 140 M it removes.

### Deterministic counts — the timed region alone

The harness's callgrind figures are whole-child: they include the untimed open,
reopen and verification. A retained probe (`results/change-0591/probe/`) opens a
constant four packages and then runs `edit_document`..`to_stream` — exactly the
harness's timed region — on `measured` of them, so an isolation pair at
`measured = 1` and `3` differences two timed regions and nothing else.

| shape | `noop` | `one` | `one-percent` |
| --- | --- | --- | --- |
| 24 paragraphs | 437,673 → 245,609 (**−43.88%**) | 2,655,864 → 2,284,984 (**−13.96%**) | 2,657,114 → 2,283,234 (**−14.07%**) |
| 200 paragraphs | 2,903,929 → 1,484,140 (**−48.89%**) | 9,896,921 → 7,088,639 (**−28.38%**) | 9,945,769 → 7,146,772 (**−28.14%**) |
| 10,000 paragraphs | 139,954,800 → 70,350,648 (**−49.73%**) | 414,170,045 → 276,371,183 (**−33.27%**) | 418,281,707 → 281,287,465 (**−32.75%**) |

The before column sums to 423.4 M over the three shapes of the one-edit case,
which is the ≈440 M 0587 modelled for the same region.

### Paired timing

Order A1 B1 B2 A2 in one window, `--warmup 5 --samples 50` per run, 100 samples
per pooled leg, pinned to CPU 11. Deltas are `(after − before)/before`; the
inverse column is the same comparison stated the other way.

| case | shape | before p50 | after p50 | p50 Δ | inverse | p95 Δ | p99 Δ | mean Δ | A/A p50 | B/B p50 |
| --- | --- | --- | --- | --- | --- | --- | --- | --- | --- | --- |
| noop | large | 7,739,054 ns | 3,873,201 ns | −49.95% | +99.81% | −49.87% | −48.15% | −50.00% | −0.52% | +1.20% |
| noop | medium | 161,881 ns | 82,680 ns | −48.93% | +95.79% | −47.30% | −46.98% | −48.85% | −1.04% | +0.07% |
| noop | tiny | 26,470 ns | 15,380 ns | −41.90% | +72.11% | −41.78% | −31.59% | −41.53% | −2.06% | +0.46% |
| one edit | large | 21,103,396 ns | 13,299,460 ns | −36.98% | +58.68% | −37.10% | −37.25% | −37.03% | +1.69% | +0.49% |
| one edit | medium | 490,343 ns | 331,092 ns | −32.48% | +48.10% | −31.67% | −30.36% | −32.27% | +1.29% | +0.53% |
| one edit | tiny | 114,190 ns | 92,910 ns | −18.64% | +22.90% | −22.23% | −26.97% | −19.54% | +1.22% | +1.67% |
| one percent | large | 21,327,298 ns | 13,576,161 ns | −36.34% | +57.09% | −36.51% | −36.44% | −36.34% | +0.44% | +0.03% |
| one percent | medium | 495,123 ns | 336,641 ns | −32.01% | +47.08% | −31.22% | −31.75% | −31.94% | +0.59% | +0.48% |
| one percent | tiny | 112,570 ns | 91,441 ns | −18.77% | +23.11% | −13.91% | −15.73% | −18.29% | +1.42% | +0.73% |

**The floor in this window.** The two before runs differ by at most 2.06% at
p50 and the two after runs by at most 1.67%, against the roughly 4% p50 the
program records for this host. Every scenario moved well outside it, every
scenario in the same direction, and none regressed. These are the nine
scenarios run; none is omitted.

The instruction and latency columns agree on ordering and differ modestly in
size on the edit cases (−33.27% Ir against −36.98% p50 on the large one-edit),
which is what removing a pointer-chasing XML pass rather than straight-line work
looks like.

## Correctness evidence

**Differential: the derived layout against a full rescan.** A test-only
`layout_matches_rescan` compares a snapshot's retained paragraph, table and
block-control ranges, its `content_end` and its conformance against
`scan_document` re-run on the snapshot's own bytes.

- *The DOCX fixture corpus.* All 62 `.docx` files under `test-data/` are walked;
  54 open as a package and 53 yield a document snapshot. Up to 16 paragraph positions per fixture are rewritten
  with a short, an empty and a long replacement: **297 rewrites are accepted and
  all 297 both match a rescan and take the incremental route**, plus 4 batch
  rewrites through `replace_body_paragraph_texts`, which also match. The test
  asserts these coverage counts, so it cannot silently become vacuous.
- *The generated corpora.* The harness's own 24-, 200- and 10,000-paragraph
  documents, rebuilt through `Package::new` and a ZIP round trip: three
  positions × three replacement texts per shape plus the 1% batch, each matched
  against a rescan, each required to take the incremental route, and the
  committed snapshot matched too.
- *Tables, controls and the final section.* A document with a leading table, a
  block content control between the paragraphs and a trailing table checks that
  sibling ranges and `content_end` move correctly in both directions.

**The fallback is exercised, not assumed.** A replacement with one extra
element, one extra level of nesting, a different root element, two roots, a
processing instruction, unbalanced nesting, or one byte of padding at either end
is refused by `preserves_body_child_shape` and declined by `spliced_layout`; a
splice one byte short or long of a paragraph range is declined; and a declined
splice still produces a snapshot byte-identical and layout-identical to
`Snapshot::from_xml` of the same bytes. `body_child_shape` is unit-tested for
node count, depth and root-tag length, and for refusing a DOCTYPE, an XML
declaration, unbalanced nesting, empty input and bare text.

**Patch application.** A new package test asserts that an exact no-op publishes
the retained source and leaves the main part's payload `Arc` pointer-identical;
that a patch whose source no longer matches returns `TransactionError::StaleSource`
and leaves the package bytes untouched; that its inverse still reaches the exact
original bytes through the reuse proof; and that when the main part is replaced
with unparsable XML the error is still `TransactionError::Document`, not a
stale-source refusal — the case that pins the fallback. The pre-existing
`document_patch_publishes_atomically_and_reopens` still passes unchanged.

**Gates.** `cargo fmt --all --check`, `cargo clippy -p litchi-docx --all-targets`,
`cargo test -p litchi-docx` (**1,457 passed, 0 failed, 31 ignored across 51 test
binaries**, 9 of them new) and `cargo doc -p litchi-docx --no-deps` all pass in
the change worktree; tails in
[`results/change-0591/gates.txt`](results/change-0591/gates.txt). The harness
verifies every sample it times: each of the four timing runs reopens the saved
package and re-reads the edited paragraphs, and exits non-zero otherwise.

## Validation preserved

Nothing was removed or made conditional. The semantic readback after each
paragraph rewrite runs against the incrementally built snapshot exactly as it
ran against the rescanned one. `Patch::apply`'s `same_source` check still gates
every application. `compact_changed_document_xml` and its rescan are untouched.
The package writer's authored-XML audit at publication is untouched.
`MAX_DOCUMENT_XML_BYTES`, `MAX_DOCUMENT_NODES` and `MAX_DOCUMENT_DEPTH` keep
their values, their error identities and their enforcement points: the first
through the fallback, the other two through the fragment proof that no total
can grow. The managed execution context, its work charges and its admissions
are not reached by either path.

## Limitations

- **No claim is registered.** The numbers above are scoped to these nine
  scenarios, this synthetic corpus, this host, these two `--release --locked`
  builds, and the metrics named. They do not describe real-producer documents,
  cold caches, other paragraph operations, or any other machine.
- **Part (c) is not implemented, and it has an open owner question.** Compaction
  still rewrites the whole main part on every changed commit, and its rescan is
  the second of the two remaining scans — 62.5 M Ir per one-edit lifecycle,
  unchanged by this record. Compacting only the replacement fragment would
  remove it, but `compact_changed_document_xml` today also strips compactable
  whitespace from paragraphs the edit never touched. Under ADR 0006 that is a
  preservation question, not a performance one: whether a one-paragraph edit is
  *allowed* to rewrite untouched markup is the owner's decision, and 0587 logged
  it as an observation rather than a defect. Part (c) needs that answer, then a
  frozen design record and preservation tests, before any code.
- **Only two rewrite sites take the incremental route.** The other eleven
  `Edit` call sites of `with_rewritten_xml` — hyperlink, run, simple- and
  complex-field, revision, content-control, cell and nested-cell text,
  paragraph insertion and removal, section properties — still rescan. Several
  rewrite a paragraph *inside* a table or a block content control, which the
  derivation declines outright today: it accepts a splice only when the range
  is exactly one direct-body paragraph, so the enclosing-range growth those
  sites need is not implemented. They are outside this record's measured
  scenarios and were left alone.
- **`document_snapshot` still copies the main part.** It builds its snapshot
  from `main.blob().to_vec()`, so the pointer half of part (a)'s proof never
  hits on the eager route and the byte comparison runs in full. Handing the
  part's `Arc<Vec<u8>>` to `Snapshot::from_shared_xml` instead would remove both
  the copy and the comparison, and `Part::set_blob` replaces the `Arc` rather
  than mutating it, so the sharing is safe on its face. It changes what an
  ordinary snapshot aliases, so it is named here as the next measurement rather
  than folded into this one.
- **DOCX-2 is untouched.** `DocumentPart::from_part` still builds its paragraph
  index eagerly on every `document()` call; that is a separate item in the 0587
  queue and a separate route from the transaction snapshot.
- Instruction counts rank work; they are not latency. The paired medians are
  reported beside the floor measured in the same window and are not registered
  as a speedup.

## Retained evidence

[`results/change-0591/README.md`](results/change-0591/README.md) — the two legs'
callgrind annotations and their per-symbol summaries, the timed-region probe
with its source and outputs, the four paired-timing reports and their analysis,
the gate tails, `decision.json`, and the log paragraphs for the four program
logs.
