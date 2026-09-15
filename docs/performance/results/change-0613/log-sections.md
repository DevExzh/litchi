# Log sections for change 0613

Four paragraphs for the coordinator to merge, one per log, each in the style of
that log's newest section. Nothing here is added to `HOTSPOTS.md`,
`GOAL_AUDIT.md`, `REPORT.md` or `ADR_COMPLIANCE.md` by this change.

## For `HOTSPOTS.md`

### 0613 — the original-bytes publication audit, priced and its memo declined

A callgrind isolation pair on `xlsx_source_backed_cell_values_one_edit_save`
splits change 0528's combined publication-audit figure into halves: the audit of
the **original** Part bytes is 51,844,180 Ir, **27.59%** of source-backed
publication, the audit of the replacement 51,852,530 Ir (27.60%), and the pair
55.19% — an independent reproduction of 0528's 56.2257% on a different base and
corpus. The memo 0587 item XLSX-3 proposed was implemented and declined: it
fires **zero** times, because every publication entry point on
`SourceBackedPackage` takes `self` by value (so one package publishes at most
once and a lineage-keyed memo dies with it) and duplicate part names are refused
with `OpcError::DuplicatePartName` before any audit inside one publication. Its
only measurable effect was its cost, +753.5 Ir per iteration at audit scope
(+0.0007%); paired timing moved −2.02% (`medium`) and −0.60% (`dense-sparse`) at
p50 against A/A floors of 2.18% and 3.03% in the same window, so nothing is
claimed. The 27.59% is the size of change 0602's **D0** prize, not of the memo's:
on real producer packages the first original audit refuses 94 of 95 fixtures and
there is no second publication to serve. No production change and no speedup is
claimed. [Change and limitations](0613-opc-original-audit-memo.md);
[evidence](results/change-0613/README.md).

## For `GOAL_AUDIT.md`

### 0613 — the original-bytes publication audit, priced and its memo declined

Item XLSX-3 of the 0587 queue is answered for its memo half and closed without a
production change. The audit of the original bytes of a replaced Part is
measured at 27.59% of source-backed XLSX publication instructions (0587 modelled
"at most about 28%"), so the opportunity is real, but the mechanism proposed for
it is not reachable: publication consumes its package, so a memo held inside
`SourceBackedPackage` has no second reader, and measurement confirms zero hits.
Reaching it requires either a non-consuming publication door or a caller-owned
cross-lineage memo with a digest key — both contract changes needing their own
records — and both sit behind change 0602's D0, which this change does not
touch. The P1 row "Finish source-backed CRUD adoption across formats" is
unaffected; the P0 publication-intersection rows gain one priced owner. 737
`litchi-opc`/`xml-minifier` tests and 3,629 `litchi-xlsx`/`litchi-docx`/
`litchi-pptx` tests pass on the retained candidate patch; the committed tree
changes no file under `crates/`. `performance_claim: none`. OLE2/OOXML
optimization remains active; ODF is deferred until completion and iWork
excluded. [Change and limitations](0613-opc-original-audit-memo.md);
[evidence](results/change-0613/README.md).

## For `REPORT.md`

### 0613 — the original-bytes publication audit, priced and its memo declined

Source-backed XLSX publication audits two XML Parts per save, each twice — once
as the immutable original and once as the authored replacement — for eight
`verify_authored` calls per harness iteration. Splitting the two halves by call
site prices the original at 51,844,180 Ir per iteration, 27.59% of publication,
and the replacement at 27.60%. The memo that would reuse the original verdict
across publications was built, gated and reverted: every publication entry point
takes `self` by value, so no second publication from one package exists, and the
call count is unchanged at eight in both legs. Six timing legs ordered
A1 B1 B2 A2 A3 A4, 30 samples each on a pinned CPU, put the candidate inside the
A/A floor on both corpus shapes in both directions, with one output SHA-256 per
shape across all six legs. This is instruction attribution and a declined
mechanism, not a latency, allocation, RSS or cold-cache result, and the shares
are measured only on litchi-written corpora, where the original audit accepts.
See [Change 0613](0613-opc-original-audit-memo.md); `performance_claim: none`.

## For `ADR_COMPLIANCE.md`

### 0613 — the original-bytes publication audit, priced and its memo declined

Nothing under `crates/` changes, so no accepted ADR boundary moves. The retained
candidate patch was written to the same boundaries and is recorded here for the
door that might adopt it: both publication audits stay (change 0528), the
replacement is audited in full every time, the original audit keeps its auditor,
its `Limits::default()`, its call site and its `OpcError::XmlPublication { part,
source }` construction so a refusal surfaces at the same point with the same
identity and message, the memo never observes a replacement payload, freshness
is unchanged because `ensure_current` refuses `OpcError::SourceChanged` before
the memo is reached, the retained set is bounded by one entry per admitted Part
with an allocation failure degrading to a fresh audit rather than to a
publication failure, and no archive type, raw lock or executor reaches the
public surface. What change 0602 calls D0 — that the audit of *original* bytes
refuses 94 of 95 real producer packages — is untouched; this change does not
move where a refusal happens. The two doors that would make the memo reachable
are both contract changes and are named as requiring their own records. No
`unsafe` is added, no limit weakened and no malformed-input defence relaxed. See
[Change 0613](0613-opc-original-audit-memo.md); `performance_claim: none`.
