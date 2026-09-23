# Change 0754 differential: base versus branch

`scripts/make_inputs.py` wrote the inputs listed in `docx-inputs.txt` (719) and
`pptx-inputs.txt` (78): every DOCX, DOTX and DOCM fixture under `test-data`
(63), the harness's three semantic DOCX corpora, and ten deterministic
mutations (seed `0x0754`) of each DOCX main part, three for parts over 400 KB —
byte damage, range deletion, truncation, and insertions after `<w:body>` or
between paragraphs (shadowed and undeclared prefixes, `]]>`, a bad
`xml:space`, other prefixes bound to the Word namespace, tabs and breaks,
tables, content controls, section properties, a second body, a processing
instruction, a comment, an unsupported entity, a bad `xml` binding). The same
719 main parts were also written as raw XML for the snapshot scanner, and
`scripts/make_dos.py` wrote five namespace worst-case witnesses.

The probe built against each leg printed one JSON line per input
(`outputs.tar.gz`: `before-*.jsonl`, `after-*.jsonl`):

* `docx`: open; `Document::text`; `write_text_to`; every paragraph's text;
  `edit_document`'s snapshot counts and paragraph texts; a no-op edit/save; a
  one-paragraph edit at the first, middle and last paragraph under each
  compaction policy, each followed by the inverse patch and a second edit,
  with the SHA-256 of every saved archive and committed snapshot; a composite
  edit of every tenth paragraph; a source-backed edit of the middle paragraph.
  Every error is recorded with its `Display` and `Debug`.
* `snapxml`: `Snapshot::from_xml` on the raw main part: counts and paragraph
  texts, or the error.
* `pptx`: `Presentation::text`.

`compare.txt` / `summary.json`: 719 + 719 + 78 rows, **0 differing**; the
outcome counts show the routes reached. The five witnesses' rows are
identical too.
