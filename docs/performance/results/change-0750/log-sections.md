# Ready-to-paste log sections for change 0750

## HOTSPOTS.md

## 0750 — the source XML audit's well-formedness checks and their cost

[0750](0750-xml-audit-well-formedness-gaps.md) is a correctness change, but it
moves a hot path: every `verify_source` / `verify_source_replacement` now also
checks legal characters (one branch-free pass in 64-byte blocks), names,
references, `]]>`, comments, PI targets, the XML declaration and Namespaces in
XML 1.0. On real parts a source audit executes 14–24% more instructions and
takes 8–18% more time (timed on the first candidate; 7,213 accepted `test-data`
members: 385 → 340 MB/s), plus about 200 ns fixed per audit. Every new check is
linear in the input in expectation. The review-found expanded-name path, which
read ancestors' namespace names on every tag (180 s on one 14.8 MB input), now
compares name identities (0.038 s). Re-measured on the fix, the DOCX one-edit
save is +2.67% (CI [+2.15%, +3.90%], +43 µs): it audits its main document
completely twice per commit, so it pays the new checks in full. Those two audits
are `litchi-docx`'s `publication_accepts_preserved_xml` gate in `Edit::commit`
and the original half of `litchi-opc`'s replacement pair, which on a first edit
are the same bytes. Carrying the gate's verdict into the pair would remove one
complete audit per DOCX commit and is the next opportunity here. The XLSX
dense-sparse one-edit save is +1.91% (about 210 µs of audit, the rest planning
code this change does not touch). The character pass's byte-by-byte tail is
about half of the fixed cost.

## REPORT.md

## 0750 — XML audit well-formedness gaps

[0750](0750-xml-audit-well-formedness-gaps.md) is retained with
`performance_claim: none`. Base `3174242282`; commits `a07d680852` (fixtures),
`de69fb407d` (auditor), and after an independent review `0bbcc9bf94` (a CPU
denial-of-service fix) and `900eb6de11` (a citation). Under `Policy::SOURCE`
only, the publication XML audit now refuses what XML 1.0 (Fifth Edition) and
Namespaces in XML 1.0 refuse, each with `Error::Malformed` at the offending
byte. That covers the reviewer's seven inputs from 0747, undeclared entities and
prefixes, illegal characters and character references, `<` in values, and
expanded-name duplicates. It also covers reserved-prefix misuse, `--` in
comments, PI targets and the XML declaration's place and grammar: 96 of 150
probed inputs are newly refused and none newly accepted. `verify`,
`verify_authored` and the streaming readers are unchanged, because authored
fragments and the ODF generators depend on them; that is recorded as remaining
work. 0747's window proof replays the new position state, namespace-name
identities included. A 24-million-case release campaign, run on both candidates,
found no difference from two complete audits; it caught a `]]>` straddling a
64-byte block in an earlier candidate. The review (differential against expat
and libxml2) found no false refusal and no missed malformation. It did find
that the first candidate's expanded-name check compared ancestor-declared
namespace names byte by byte on every tag: 180 s on an input the base audits in
0.017 s. The fix interns names at declaration and audits that input in 0.038 s.
Census over 7,222 real OOXML XML members and 5,504 evidence-packet members:
three loose sample parts newly refused, all genuinely malformed; no package
member changes verdict. Measured cost: source audits +14–24% in instructions;
DOCX one-edit save +2.67%, XLSX dense-sparse one-edit save +1.91% (on the fix);
PPTX and medium XLSX one-edit saves within noise (on the first candidate); every
output digest identical.

## GOAL_AUDIT.md

## 0750 — ADR 0006 well-formedness in the publication audit

[0750](0750-xml-audit-well-formedness-gaps.md) closes an ADR 0006 gap: the
source audit that publication runs (the eager writer's `audit_published_xml`
over every payload, `litchi-opc`'s source-backed replacement pairs and
`litchi-docx`'s preserved-publication gate) accepted documents that are not
well-formed. Examples are `<r>&;</r>`, `<r>]]></r>`, an inner `<?xml …?>` and
undeclared prefixes, because its checks stopped at quick-xml's tokenizer. No
record tolerated them on purpose, and none of 7,222 real members needs them.
They are now refused with typed errors; no public API changes. Every new check
costs time linear in the input in expectation. The review found one check that
did not: the expanded-name check read ancestors' namespace names on every tag.
It was fixed before merge, and tests now bound the name bytes each audit reads.
Declared non-UTF-8 encodings are refused under OPC rule M1.17. Five dependent
tests had fixtures that were not well-formed and now declare their prefixes, or
put deliberately malformed story bytes past the writer. Still open: `litchi-opc`
sites that audit complete authored documents with `verify_authored` keep the
tokenizer-level checks, because that policy also serves fragments. A
`SourceXmlPart` replacement relies on `validate_source_xml`. Namespace names are
not checked as URIs.
