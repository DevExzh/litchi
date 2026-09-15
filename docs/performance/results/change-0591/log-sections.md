# Log paragraphs for change 0591

The coordinator merges these into the four logs; this batch did not edit them.
Each is written in the style of the newest section of its target file.

## For `docs/performance/HOTSPOTS.md`

## 0591 — the ordinary DOCX edit scans the main part twice

Item DOCX-1(a)/(b) of the 0587 queue is implemented. An ordinary one-paragraph
edit and save scanned the whole main document four times — the snapshot, the
rewrite's rescan, the rescan after whole-part compaction, and a second
`document_snapshot` built only to feed `Patch::apply`'s byte comparison — and an
exact no-op scanned it twice. It now scans twice and once. `Patch` gains
`target_for_exact_unmanaged_source`, which answers that byte comparison from
`main.blob()` and hands back its own retained target, and
`Package::apply_document_patch` falls back to the unchanged snapshot route
whenever the byte proof does not hold, so every refusal and error identity stays
put. `Snapshot::with_spliced_paragraphs` replaces `with_rewritten_xml` at the two
direct-body paragraph rewrite sites: it resizes the rewritten paragraphs, shifts
every later direct-body range and `content_end` by the running delta, keeps the
table and block-control `Arc`s no splice reaches, and accepts the derivation only
against a paragraph-sized proof that the replacement fragment is one balanced
element with no more nodes, no more depth and a byte-identical root start tag —
the four things `scan_document` concludes about the inside of a body child.
**Measured** on the isolation pair (`--samples 3` minus `--samples 1`, callgrind,
CPU 11): `Snapshot::from_xml` 12 → 6 calls per `docx_semantic_one_edit_save`
lifecycle, 280,511,387 → 140,478,384 Ir, with `apply_document_patch` inclusive
72,914,944 → 1,657,943; 6 → 3 calls on the no-op. A retained probe isolates the
harness's timed region and measures −33.27% Ir on the 10,000-paragraph one-edit
and −49.73% on its no-op. Paired timing, 100 samples per leg, order A1 B1 B2 A2:
p50 −36.98% large, −32.48% medium, −18.64% tiny on the one-edit selector, the
same shape on one-percent, and −49.95%/−48.93%/−41.90% on the no-op, against an
A/A floor of 2.06% p50 in the same window; no scenario regressed. 1,457
`litchi-docx` tests pass, nine of them new, including a rescan differential over
297 accepted rewrites on 53 DOCX fixtures and the 24/200/10,000-paragraph
corpora. Item (c) — compact only the replacement fragment — changes output bytes
for untouched paragraphs and remains a frozen-design prerequisite with an open
owner question. `performance_claim: none`.
[Change and limitations](0591-docx-edit-single-scan.md);
[evidence](results/change-0591/README.md). OLE2/OOXML remain active; ODF is
deferred until completion and iWork excluded.

## For `docs/performance/GOAL_AUDIT.md`

## 0591 — the ordinary DOCX edit scans the main part twice

`docs/GOAL.md`'s first optimization step — eliminate unnecessary work — applied
to the DOCX opened-document CRUD route that 0587 ranked fourth. Two of the four
whole-part scans per edit are removed because neither produced information the
caller did not already hold: one existed to feed a byte comparison that needs
only the bytes, and one rebuilt a layout that the splice determines exactly.
Both reuses are conditional and fall back to the original route, so no refusal,
no output byte, no limit and no validation moved; the semantic readback after
every paragraph rewrite and the `same_source` check on every patch application
both still run, which is what ADR 0005's mandatory validation requires, and ADR
0006 preservation is untouched because compaction — the part that would change
untouched bytes — was deliberately left alone. **Measured** with exact before and
after call counts (four scans and two `document_snapshot` builds per edit
lifecycle before, two and one after) and −33.27% instructions in the isolated
timed region of the 10,000-paragraph one-edit, matched by −36.98% p50 against a
2.06% A/A floor. What stays open: item DOCX-1(c) removes the last removable scan
but rewrites whitespace in paragraphs the edit never touched, so it needs the
owner's answer to 0587's ADR 0006 observation before a design record exists;
eleven other `with_rewritten_xml` sites still rescan, several of them rewriting a
paragraph inside a table or a block control, a case the derivation declines
outright; `document_snapshot` still copies the main part instead of sharing its
`Arc`, so the reuse proof still pays a full byte comparison; and DOCX-2's eager
paragraph index is untouched. No speedup, RSS, allocation, cold-cache or
real-producer claim follows.

## For `docs/performance/REPORT.md`

## Change 0591: the ordinary DOCX edit scans the main part twice

Change 0591 removes two of the four whole-part scans in the `litchi-docx`
ordinary edit-and-save lifecycle. `Patch::target_for_exact_unmanaged_source`
(`document/transaction.rs`) reproduces what `Patch::apply` decides when its
source is the snapshot `document_snapshot` would have built — no source
identity, equal length, then pointer identity or a byte comparison — and returns
the patch's own retained target, so `Package::apply_document_patch`
(`package/package/document.rs`) no longer copies and rescans the main part to
answer a `memcmp`; `document_patch_source` checks package staleness and main-part
resolution first, in the original order, and routes everything the byte proof
does not settle back through `document_snapshot` and `Patch::apply` unchanged.
`Snapshot::with_spliced_paragraphs` derives the post-rewrite layout for
`replace_paragraph_text` and `replace_body_paragraph_texts` instead of rescanning,
accepting a splice only when its range is exactly one direct-body paragraph and
`preserves_body_child_shape` proves the replacement fragment keeps every
whole-document verdict the scanner reaches, and falling back to the full rescan
otherwise. Validation passed `litchi-docx` `1,457/1,457` across 51 test binaries
with nine new tests: a rescan differential over 297 accepted paragraph rewrites
on 53 DOCX fixtures and over the generated 24/200/10,000-paragraph corpora, both
asserting their own coverage; eight shape-changing replacements and three
off-boundary splices declining the derivation; a declined splice still producing
a byte- and layout-identical snapshot; and an exact no-op keeping the main part's
payload `Arc`, a mismatched source still returning `StaleSource`, and an
unparsable main part still returning `TransactionError::Document`. **Measured**,
callgrind isolation pair on CPU 11: `Snapshot::from_xml` 12 → 6 calls and
280.51 M → 140.48 M Ir per `docx_semantic_one_edit_save` lifecycle, and
−33.27% Ir in the isolated timed region of its largest shape. The paired native
timings — p50 −18.64% to −49.95% over all nine scenarios — are reported beside
this host's A/A floor of 2.06% p50 in the same window and are not a speedup
claim. See [Change 0591](0591-docx-edit-single-scan.md);
`performance_claim: none`.

## For `docs/performance/ADR_COMPLIANCE.md`

## Change 0591 compliance update

Change 0591 keeps ADR 0003's transaction boundary exactly: `commit()` still
returns the same `Commit` with the same snapshot, reversible patch and
diagnostics; snapshots remain immutable and cheap to share; and a patch is still
applied only against the exact source bytes it was produced against. Part (a)
changes what that exact-source proof is compared against — the opened package's
retained main part rather than a throwaway copy of it — not whether it is
compared, and the snapshot it returns is byte-equal and identity-equal to the one
`Patch::apply` would have cloned. ADR 0005's mandatory validation is intact: no
check is removed, weakened, reordered or made conditional; the semantic readback
after every paragraph rewrite runs against the incrementally built snapshot
exactly as it ran against the rescanned one; and where a whole-document rescan
previously re-established `MAX_DOCUMENT_NODES`, `MAX_DOCUMENT_DEPTH`,
well-formedness and the DTD and processing-instruction refusals, a
paragraph-sized fragment proof now establishes that none of those verdicts can
have changed, with a full rescan as the fallback for everything it does not
settle. The `MAX_DOCUMENT_XML_BYTES` limit keeps its identity because oversized
input is routed to `Snapshot::from_xml`, which raises it. ADR 0006 preservation
is intact because no output byte changes: `compact_changed_document_xml` and its
rescan are untouched, and the 0587 observation that ordinary commit compaction
rewrites whitespace in untouched paragraphs is left exactly where it was, as an
owner question that item DOCX-1(c) depends on. No `unsafe`, no weakened limit or
malformed-input defence, no new ambient I/O, no global pool, and no public
leakage of archive types, locks or executors; the managed and source-backed
routes are not reached by either reuse. See
[Change 0591](0591-docx-edit-single-scan.md); `performance_claim: none`.
