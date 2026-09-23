# Change 0744 log sections

Ready-to-paste sections; newest-first order places each at the top of its log.

## For `HOTSPOTS.md`

## 0744 — eager XLSX sheetData lane and dense reduced readback

[0744](0744-xlsx-eager-workbook-cell-path.md) removes the eager `Workbook`'s
repeated generic XML work on dense sheets. Base profiles showed four of six
complete passes over the touched sheet (base parse, edit scan, compaction,
verification parse; 68% of one-cell commit+save) spent in namespace-resolving
reader machinery over benign `row`/`c`/`v` markup. A strict `<sheetData>`
recognizer replays the reader's exact events into all three passes and declines
everything else; dense value edits above the handoff bound verify through
0525's reduced readback. Against a same-command before leg on CPU 12, dense-wide
first cell falls 71.91% (27.87 → 7.82 ms) and one-cell commit+save 59.53%
(149.2 → 60.5 ms); outputs are byte-identical. Deflate (60%) and the publication
audit (12.6%) now dominate; inline strings, Excel MCE/x14ac pre-passes and
per-row scanner slots remain. Review fixes make the reduced-readback admission
exact (worksheet's own SpreadsheetML body, 16 MiB bound, 0525's collision
refusal); re-measured one-cell −59.79%. `performance_claim: none`;
[evidence](results/change-0744/README.md).

## For `REPORT.md`

## 0744 — eager XLSX sheetData lane and dense reduced readback retained

[0744](0744-xlsx-eager-workbook-cell-path.md) adds a strict recognizer for benign
`<sheetData>` bodies to the eager worksheet parser, edit scanner and compactor,
compact borrowed cell records, and 0525's reduced readback for dense value-only
commits. ABBA against a same-command before leg (eight processes per case, CPU
12): dense-wide first cell −71.91%, full scan −71.81%, one-cell commit+save
−59.53%, one-percent −59.03%; medium −58% to −69%; eager control −52.85%;
source-backed control and open unchanged. Allocation calls fall 88.69% and
region peak 44.82% on dense one-cell. Two no-op pair flags (+7.77%, +15.79% on
16 µs and 0.6 µs operations that run no changed code) do not repeat in a
2,000-sample confirmation. All 18 probe outputs are byte-identical. After
adversarial review, three commits bind the reduced readback's admission to the
worksheet's own SpreadsheetML body, bound it by the 16 MiB web limit, restore
0525's collision refusal, decline byte-order-marked parts, test the two writers
directly and order shared-formula groups deterministically; a rebuilt dense
re-measurement gives one-cell −59.79% and one-percent −59.35%. All gates pass.
`performance_claim: none`; [evidence](results/change-0744/README.md).

## For `GOAL_AUDIT.md`

## 0744 — eager XLSX sheetData lane and dense reduced readback

[0744](0744-xlsx-eager-workbook-cell-path.md) advances workstreams D (fewer
passes, no namespace resolution where a strict local grammar proves it
unnecessary) and C (borrowed value text, compact admitted-cell records, inline
payload spans) for eager XLSX reads and edit/commit/save. Every refusal, limit,
MCE path and output byte is kept; anything outside the benign subset takes the
unchanged reader route. Measured scope: synthetic dense-wide and medium corpora
and the cell-CRUD controls on one host. Not established: RSS, cold cache,
Excel-produced sheets' MCE and x14ac pre-passes, inline-string bodies, and the
publication audit and deflate that now dominate. A pre-existing defect found in
review remains open: eager edits of byte-order-marked worksheets fail on both
routes because the edit scanner's spans exclude the mark. The program goal
remains open.
[Evidence](results/change-0744/README.md).
