# Log paragraphs for change 0601

Four paragraphs, one per merged document, in the style of each file's newest
section. The coordinator merges these; this change does not edit
`HOTSPOTS.md`, `GOAL_AUDIT.md`, `REPORT.md` or `ADR_COMPLIANCE.md` itself.

---

## For `HOTSPOTS.md`

## 0601 — the harness can finally see the real-producer path

Retained harness and corpus work against 0587's finding 1, the first
prerequisite of its top-ranked items. Every generated worksheet this program has
measured since change 0032 carries no markup-compatibility root, no
`x14ac:dyDescent`, no `<cols>`, no shared-string part and no worksheet
relationships, so items XML-1, XML-2, XML-3 and XLSX-2 had no corpus to be
priced on. `tools/perf-baseline/src/producer_shape.rs` adds one, in three XLSX
variants per shape plus a DOCX Word-2013 shape and an `mc:AlternateContent`
PPTX shape, by rewriting only the parts under test in a package the production
writers produced. The first baseline, 20 warm-ups and 100 samples per selector
with an A/A floor of 3.0% at p50 at worst and 1.09% median, prices the producer
signature at **7.04× (medium) and 6.79× (dense)** on a selected cell and
**18.61× and 19.03×** on value-only planning, against the byte-comparable
marker-free control captured in the same run. The planning ratio is what the
namespace *declarations* and the `<cols>` block cost on their own: they are
enough to send the whole part through the MCE codec's rewrite and to take
0546's fused traversal off the table, and the ratio is flat across a 16× change
in worksheet size, so it is a per-byte constant. A real Excel worksheet —
`Excel_file_with_trash_item.xlsx`, the 209,931-byte sheet the survey profiled —
costs 53.63 ms for one cell through the same public API against 122.9 µs to
open the workbook, and the generated shape trips exactly the same three
`source_stream_eligible` conditions. Nothing was optimized; the ranked items
now have a corpus and a number to beat.

---

## For `GOAL_AUDIT.md`

## 0601 — reproducible baselines on the corpus class the audit was missing

Retained against the audit's standing "corpus classes including multi-producer"
and DEFINITION OF DONE clause (a) rows. The audit has carried "most real
DOC/PPT and a chunk of real XLS/XLSX/DOCX/PPTX fixtures can't reach the measured
path" as a gap since the 0587 entry; this change closes the generator half of
it for OOXML. The new corpora are deterministic — two independent release-mode
processes produce byte-identical archives, and both the census diff and the
corpus-identity diff are empty — and they are shown to trip the same library
gates a real Excel file trips, part for part, through the same code path. The
`--real-file PATH` opt-in adds the other half for XLSX reads: the only input in
this harness whose bytes come from outside the process, bounded to 32 MiB,
absent from the default matrix, with the file's path, size and SHA-256 bound
into the corpus identity. `--producer-evidence PATH` makes the marker and
refusal census a first-class harness output, which is the audit's evidence-gap
3 ("a tracked refusal census instead of narrative repetition per record").
**What this batch does not discharge**: the selector gap is untouched for DOCX
and PPTX edit/save, the eager XLSX paths, file-backed and range sources and
XLSB; the dense shape is 128 × 128 rather than 256 × 256 because the codec makes
the larger size seconds per sample; the CI smoke job still compares against
itself; and no instruction-level attribution was taken, so the numbers rank
latency on one host, not work. One standing failure was confirmed on the way and
not fixed: `xls_source_backed_lifecycle_selectors_are_matched_and_local` asserts
`open_reads_zero_worksheet_payload == [true]` and gets `[false]`, identically on
the untouched checkout at `f8cf7d2a1`, and its panic poisons the shared
allocation-metrics mutex so six later tests fail as cascades.

---

## For `REPORT.md`

## 0601 — a producer-shaped corpus, and what it costs to read one

`tools/perf-baseline` gains a generator that writes what Excel, Word and
PowerPoint write by default and sixteen opt-in selectors over it, none of them
in `Case::DEFAULT` and none of them changing an existing corpus identity or the
checked default catalog SHA-256. The XLSX family needs three variants per shape
because the source-backed value-only editor refuses five separate parts of the
producer signature — `mc:Ignorable` on the worksheet, `pageMargins`, the
shared-string workbook relationship, any worksheet relationship, and
`mc:Ignorable` on the workbook — so its reach on Excel output is zero rather
than reduced, and the split lets each refusal be *proven*, untimed, at corpus
construction rather than asserted. Against a marker-free control of the same
grid captured in the same run, one selected cell costs 13.31 ms versus 1.89 ms
(7.04×) at 32 × 32 and 195.54 ms versus 28.80 ms (6.79×) at 128 × 128; value-only
planning costs 11.54 ms versus 0.62 ms (18.61×) and 168.97 ms versus 8.88 ms
(19.03×). A source-backed open is unaffected at roughly 100 µs on every variant,
which is the control that says the cost is worksheet work rather than package
work. `performance_claim: none`: nothing here is an optimization, and the A/A
floor (3.0% p50 worst case, 1.09% median, 20 warm-ups and 100 samples per
selector) is reported so the ratios can be read as measurements rather than
noise.

---

## For `ADR_COMPLIANCE.md`

## 0601 — no boundary moved, and five that are now documented

No file under `crates/` changed, so no ADR boundary moved: ADR 0005's mandatory
validation, ADR 0006's preservation and typed-refusal contracts and ADR 0003's
readback requirement are exactly where change 0595 left them. What this change
adds to the matrix is evidence about them. The five typed refusals the
source-backed value-only editor draws from an ordinary Excel package are now
recorded verbatim, each witnessed on the smallest archive carrying only the one
producer fact that triggers it, with the enforcing site cited
(`cell_values/validation.rs:420` for the element whitelist, `:478` for every
qualified attribute, `cell_values/snapshot.rs:388,444` for worksheet
relationships and `:1729` for workbook relationships). They are 0362's
"deliberately narrow" scope working as designed; the compliance observation is
only that the narrowness is total on real producer output, which is the fact
survey items XLSX-2 and XML-3 will have to argue against under a frozen design
record. The one deliberate deviation from Excel's bytes is named in the record:
the generated parts are compact, with no `\r\n` after the XML declaration,
because `validate_authored_xml` (`litchi-opc/src/pkgwriter.rs:988`) refuses that
whitespace for the whole save — 0587's defect 1, unchanged and unaddressed here.
The `--real-file` opt-in is the only ambient read in the harness; it is bounded,
opt-in, absent from the default matrix and self-identifying, and a source-policy
test holds each of those properties.
