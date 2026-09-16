# 0652: the owner's decisions for the third wave: ten blocked rows unblocked, two ADRs accepted, three standing trade-offs recorded

Status: retained, decision record. `performance_claim: none` — this record
carries no measurement; it records decisions and is the authority the records
that implement them cite.

OLE2 and OOXML remain the active priority. ODF optimization stays deferred until
that goal completes; iWork is excluded.

## What this record is

Change [0651](0651-queue-refresh-after-the-second-wave.md) closed the second
wave with a queue whose first ten rows each waited on a human decision: an
owner's answer on a contract, an ADR acceptance, or an ADR clarification. On
2026-09-16 the owner gave those decisions, together with three general
trade-offs that bind every later wave. This record writes them down verbatim,
states how the programme reads each one, flips proposed ADRs 0030 and 0031 to
Accepted, and names what each implementing record must still prove. Until now
every change that moved a contract stopped at a frozen design with a price and
admission gates (0644, 0645, 0646, 0610, 0624, 0602, 0597, 0591, 0617, 0528);
from this record on, those designs are authorized.

## The standing trade-offs

The owner's words, then the programme's reading.

1. **"Since the library is in early alpha, breaking changes are totally
   acceptable."** A design no longer stops at "this needs a breaking public API
   change". It makes the change, states it in the record under a heading
   `Breaking changes`, updates the crate documentation, and keeps the change
   the smallest that serves the decision. Nothing else about the API rules
   moves: the facade stays concise and panic-free (ADR 0001), no archive type,
   raw lock or executor leaks (ADR 0005).
2. **"Correctness and safety is the primary consideration; performance is
   secondary."** Unchanged rules: no `unsafe`, no weakened malformed-input
   defence, typed refusals, exact no-ops stay exact, validation does not
   mutate, preservation by default. When a faster path and a safer path
   conflict, the safer wins, and the record says which conflict it resolved
   that way.
3. **"Assume most of the files are benign; only a small part of the inputs are
   malicious."** The common path is the benign one. Work that exists only to
   refuse the malicious minority may move off the common path (deferred,
   lazy, or taken only when a precondition fires) provided the minority is
   still refused with the same typed error before any partial result reaches
   a caller. A defence is never removed on this ground; its cost is moved onto
   the inputs that need it.

## The decisions

Each row quotes the owner, names the queue row and the record that priced it,
and states what the implementing record must still prove. "Row" numbers are
0651's refreshed queue.

| # | decision (the owner's words) | row; priced by | what it authorizes | what the implementing record must still prove |
| ---: | --- | --- | --- | --- |
| 1 | "MCE namespace re-declaration: accept the public API changes." | row 1; [0588](0588-mce-codec-namespace-emission.md), [0649](0649-pptx-opened-transaction-real-deck-edit.md) | the namespace-emission rewrite 0588 implemented and withdrew: the codec stops re-declaring namespaces on every element, and the slice consumers (`Paragraph::extensions`, `Shape::xml`, the five XLSX raw accessors, the two publishing writers) change their contract to match | the new contract of each consumer, stated; every pristine part still copied byte for byte; every regenerated part's bytes enumerated where they change, with the reason; the 0588 corpus and the 0649 real deck measured on both sides of the codec |
| 2 | "Original byte contract: loose the audit, accept non-compact XMLs." | row 2; [0528](changes/0528-xlsx-publication-attribution.md), [0602](0602-xlsx-real-producer-admission-design.md), [0613](0613-opc-original-audit-memo.md) | the publication audit of original part bytes stops refusing an original part because its XML is not compact (whitespace, declarations, line endings, attribute order); the 94 of 95 real packages it refused become publishable | which checks remain (identity of the original bytes with the planned bytes stays; structural checks that protect the archive stay) and which refusals are removed, each with a witness; the 95-package corpus republished byte-identically for untouched parts; ADR 0006's wording amended where it states compactness |
| 3 | "OPC lazy decoding: accept." | row 3; [0610](0610-opc-lazy-part-decode-design.md), [0581](0581-opc-package-retention.md) | ADR 0030 is Accepted as of this record; candidate C2′ and the migration of the 259 sites 0610 enumerated | 0610's gates; every accessor's error identity preserved or its movement stated; the retention 0610 predicted, measured |
| 4 | "PPTX memoized revision: invalidate the old patches and make the path faster." | row 4; [0645](0645-pptx-memoized-revision-proof-design.md), [0590](0590-pptx-opened-transaction-revision-reuse.md) | 0645's design in full: the per-part memo on `Snapshot`, the `LPRM0001 → LPRM0002` and `LPCP0002 → LPCP0003` bumps, the typed refusal for a patch serialized under the old format, and the facade-carried memo | 0645's nine gates; a patch under the old magic refused by name before any package is read; equal packages still compare equal and different ones different, over the corpus; the `pptx_slide_remove_boundary_save` regression re-measured and disclosed |
| 5 | "Cross-package copy: add a budget for that." | row 5; [0646](0646-pptx-cross-copy-candidate-retention-design.md), [0598](0598-pptx-cross-copy-revision-cache.md) | 0646's design with a retained-candidate budget as a public `opened::Limits` field (a breaking constructor change is acceptable); a candidate over the budget falls back to today's rebuild rather than refusing, because rebuilding is always available | 0646's fourteen gates; the default budget stated with its reason; the fallback exercised by a test; the retained bytes charged and released as 0646 measured |
| 6 | "Parallel deflate: accept ADR 0031." | row 6; [0624](0624-parallel-changed-member-deflate-design.md), [0615](0615-execution-context-completeness-design.md) | ADR 0031 is Accepted as of this record; parallel deflate of changed members under the execution context's budgets, with the measured balance rule of 0624 | no hidden global pool (ADR 0005); the balance rule measured, not assumed; determinism of the published bytes (0625, 0631) unchanged; the documented ordinary save's publication-bound nature (0638) stated so the saving is claimed only where it exists |
| 7 | "Admission-surface widening: accept 0602's D4 — treat other elements or attributes as unfamiliar ones and just copy them; replace the naive whitelist with the D4's dependency rule." | row 7; [0602](0602-xlsx-real-producer-admission-design.md), [0603](0603-xlsx-fused-traversal-marker-admission.md), [0601](0601-perf-harness-real-producer-shape.md) | the value-only editor's admission surface widens: unfamiliar elements and attributes are copied through unchanged, and D4's dependency rule decides which constructs an edit must understand | the dependency rule stated and tested per construct it names; the real-producer corpora of 0601 admitted and republished with the edited cell changed and nothing else; every construct still refused, listed with its reason |
| 8 | "XLSX ineligible read: accept the movement of the timing of error returns." | row 8; [0597](0597-xlsx-selected-cell-ineligibility-gate.md) | 0597's gate: an ineligible selected-cell read may surface its refusal at a different point in the read than today | the refusal keeps its type and text and still precedes any value returned to the caller; the −63% priced on an ineligible read measured on the landed change; the eligible path unchanged |
| 9 | "DOC snapshot's source-identity fence: add a new public entrypoint for it." | row 9; [0644](0644-ole2-snapshot-fence-design.md), [0589](0589-ole2-snapshot-fingerprint-passes.md) | the second public `litchi-cfb` entry point 0644's variant B2 needs, and B2 + C (the DOC open from six complete scans to three); B1 only if its wider trailing window is shown by witness not to admit a mutation B2 catches | 0644's four admission gates; every witness in 0644's sweep re-run on the landed change with the same verdicts; the error identity 0644 mapped preserved by name |
| 10 | "DOCX compaction: add a new policy setting field but default to not touching the unmodified parts. OLE2 side: also controlled by policies and default to reusing, fallback to appending — do less work and making files smaller if possible." | row 10; [0591](0591-docx-edit-single-scan.md), [0617](0617-cfb-copy-through-writer-design.md) | DOCX: a public save-policy field for compaction whose default leaves unmodified paragraphs and parts untouched, with whole-document compaction opt-in; the main-part copy in `document_snapshot` removed with it. OLE2: a public sector-layout policy on the writer whose default reuses the source's sector layout in place and falls back to appending when a stream outgrows its allocation, reclaiming space where it is cheap | DOCX: the default policy publishes byte-identical output to the fragment-only route on every fixture, and the opt-in route is 0591's compaction; OLE2: the default policy's output re-opens with every stream identical, the fallback is exercised by a length-changing edit, 0617's 13-30% measured on the landed change, and `OleWriter`'s deterministic order (0625) kept |

## ADR status changes made by this record

- [ADR 0030](../adr/0030-lazy-opc-part-decode.md): Proposed → **Accepted**
  (decision 3). Its row moves from the proposed table to the accepted table of
  the ADR index.
- [ADR 0031](../adr/0031-execution-context-budgets.md): Proposed → **Accepted**
  (decision 6). Likewise.
- ADR 0006's statements about original part bytes and ADR 0005's about
  retained state are to be amended by the records that implement decisions 2,
  4 and 5, each amendment naming this record; this record changes no other ADR
  text.

## What this record does not decide

The order and sizing of the wave, the exact shape of each public API change,
default values, and anything about ODF or iWork. Each implementing record
keeps the programme's rules: before measurements, the smallest coherent
change, correctness, preservation and adversarial oracles over the corpus,
after measurements beside a floor, every claim scoped, and no claim-registry
entry unless a paired result earns one.

## Retained evidence

[`results/change-0652/`](results/change-0652/README.md): the decision record's
`decision.json` and its four log paragraphs. The owner's decisions were given
in the session that produced this record and are quoted verbatim above.
