# Log paragraphs for change 0609

The coordinator merges these into the four shared logs. Nothing here edits those
files directly.

## HOTSPOTS.md

**The facade's `.doc` ingress, sized and closed (item CORE-1 of change 0587).**
Change [0609](../../0609-facade-doc-source-route-design.md) is a frozen design
record with no production change. `litchi::Document::open(path)` for `.doc`
reads the whole file once (`read_path_source_bytes`,
`crates/litchi/src/detection_smart/detected.rs:1198`) and parses it eagerly;
CORE-1 proposed routing it to `litchi_doc::body_text::source::SourceSnapshot`.
It is falsified on both halves of 0587's own test. **Measured** over all 57
`.doc` fixtures: the snapshot admits 8 and the facade 42, but the two
populations are not nested — **four artifacts the snapshot admits are refused by
the facade** with `CorruptedFile("invalid stylesheet: style names and aliases
must be unique")`, a validation the snapshot never performs, so a source-first
route would widen the facade's admitted population. **Measured** cost, native
`perf stat` isolation pairs pinned to CPU 29: the snapshot open is 342,797 to
435,287 cycles against the facade's 67,643 to 205,021 on the four fixtures both
admit (2.12× to 5.07×), 3.46 M to 5.90 M Ir against 0.22 M to 0.74 M (7.9× to
15.7×), 30 `read_at` calls against 2 and 153 `statx` against 11, and 6.30× to
9.90× the p50 latency of the one query both answer identically, against an A/A
floor of p50 ≤0.9% and p99 ≤1.9%. The mechanism is that `identity_fingerprint`
(`crates/litchi-doc/src/body_text/source.rs:2089`) reaches
`finish_overlay_plan_with_owner`, which fingerprints **twice** for a generic
`ReadAt`, and `SourceSnapshot::open` takes three identity passes: **six complete
artifact reads per open**, 6.05× to 7.47× the file. Change 0589 halved the
hashers driven over those reads, not the reads. The snapshot's cost therefore
fits `217,274 + 12.90 × file_bytes` cycles (+1.5% and +0.4% on two held-out
fixtures) while the eager open's tracks text units and formatting entries
(change 0596), so the source-backed route is worst precisely where a positional
reader should win: `picture.doc` (1.45 MB) costs 18.9 M cycles to open, 6.5×
what the facade spends opening the larger 1.62 MB kwsymphony form, and then
refuses `paragraph(0)` with `StructuralContent`. The snapshot wins on one axis
only: allocator peak 47.9% to 83.8% lower, retained bytes 80.1% to 97.3% lower.
Two 0587 entries are corrected: CORE-1's modelled "whole-file read retained for
the document's lifetime" is wrong — the slurped `Vec` drops with the package, so
the 1.62 MB fixture retains 1,090,299 bytes — and the DOC/PPT area's "routing
the facade to `SourceSnapshot` ... is slower than the eager open today" now has
the numbers. What remains open in this area is not this route: it is the
snapshot's own per-byte term, which would have to disappear rather than shrink.

## GOAL_AUDIT.md

**A source-backed route is not automatically the cheaper route, and this one is
6.3× to 9.9× worse at p50.** `docs/GOAL.md` hypothesis 1 ranks whole-input
ingestion as the first thing to remove, and the facade's `.doc` slurp is the
last such ingress on a priority format. Change 0609 measured it against the
alternative the repository already owns and found the slurp cheaper on every
axis but heap peak: 2 `pread64` against 30, 11 `statx` against 153, one complete
file read against six, 2.12× to 5.07× fewer native cycles. The reason is a
structural one this audit should carry forward — the eager reader's cost is
proportional to *document content* and the source-backed reader's to *artifact
bytes*, because ADR 0006's complete-artifact identity fence is levied on the
file's length before any paragraph is resolved. Removing whole-input ingestion
is only a saving when what replaces it reads less, and a fence that hashes the
artifact three times (six reads) reads more. The second finding is an admission
one: the two readers' refusal sets are **not nested**, so "try the cheap reader,
fall back to the strict one" is not a safe pattern in this repository whenever
the cheap reader skips a validation the strict one performs — here it would have
admitted four `.doc` fixtures the library calls corrupt. The third is a cost of
the fallback shape itself: a source-backed probe that is *expected* to be
refused still costs two complete artifact passes, measured at +25.8%, +192.7%
and +245.4% on top of the eager open for three real fixtures, growing with file
size. Evidence gap 5 of change 0587 is unchanged by this batch: there is still
no `perf-baseline` selector that opens a `.doc` or `.ppt` through the facade, so
every figure here comes from a retained scratch probe rather than the harness.

## REPORT.md

- [`0609-facade-doc-source-route-design.md`](0609-facade-doc-source-route-design.md)
  — the facade's `.doc` route, sized against the source-backed DOC reader and
  frozen as a design. Retained, design only, `performance_claim: none`, **no
  file under `crates/` changed**. Survey item CORE-1 is closed as falsified on
  both halves of its own test. Over all 57 `.doc` fixtures the eager facade
  route admits 42 and `SourceSnapshot::open` admits 8, with four admitted by the
  snapshot and refused by the facade, so the routes' refusal sets are not
  nested and a source-first route would admit four artifacts the library calls
  corrupt. On the four fixtures both admit the snapshot open costs 2.12× to
  5.07× the native cycles, 7.9× to 15.7× the instructions, 15× the `read_at`
  calls and about 7× the bytes; on the one query both answer identically it is
  6.30× and 9.90× slower at p50 against an A/A floor of p50 ≤0.9% and p99
  ≤1.9%. It wins only on allocator peak (−47.9% to −83.8%) and retained bytes
  (−80.1% to −97.3%). The record also states why the capability does not exist
  — the snapshot serves one of the DOC arm's eight queries, on 2 of 57 fixtures,
  and reaches 31 of one fixture's 134 text bytes — and prices the fallback shape
  at +25.8% to +245.4%. The oracle is an admission census and a value comparison
  over all 57 fixtures, both retained.

## ADR_COMPLIANCE.md

**ADR 0006 (typed refusals) and ADR 0005 (positional source), facade DOC
route.** Change 0609 changes no code, and its finding is about what a change
here would have had to give up. ADR 0006 makes a reader's refusals part of its
contract, so a facade route may not admit what the facade's own reader refuses.
Measured over all 57 `.doc` fixtures, `SourceSnapshot::open` admits four —
`footnote.doc`, `lists-margins.doc`, `duplicate-style-names.doc` and
`picture.doc` — that `litchi::Document::open` refuses with
`CorruptedFile("invalid stylesheet: style names and aliases must be unique")`,
because the source-backed owner never reads the stylesheet: a same-width
paragraph splice does not need it. The gap is structural, not corpus-dependent,
and it is the same 15-refusal facade population change 0596's differential
digest independently covers. Error identity differs on eleven more fixtures the
two readers both refuse (`InvalidFormat("Word 6.0 documents (nFib 0x0065) are
not supported")` against `Refused::AmbiguousTopology`,
`InvalidFormat("DOC password required")` against `Refused::Encrypted`); under a
fallback shape the eager error is the one the caller sees, so identity survives
there and only the four admissions cannot be reconciled. Two ADR 0005 notes.
First, the facade's `.doc` open pins its bytes once and can return
`SourceChanged` only at open; retaining a snapshot past open would add a
`SourceChanged` outcome to `paragraph_text` on an unchanged signature, and the
fallback itself opens a window between the snapshot's six reads and the eager
reader's seventh, which is a relocation of a typed boundary rather than a
routing detail. Second, ADR 0005's "no source generics on a document" is *not* a
blocker for either design: `SourceSnapshot` already erases its source behind
`Arc<dyn ReadAt>` and a `DocumentImpl::DocSource` variant would be
crate-internal exactly as `DocxSource` and `OdtSource` are. Record 0105's
admission contract stands unchanged, and this record adds the measured statement
that it is an edit owner rather than a reader: its whole read surface is
`paragraph(Position)`, which serves 2 of 57 fixtures through the facade.
