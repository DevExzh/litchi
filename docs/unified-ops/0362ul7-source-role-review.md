# 0362ul7 — source-role review of the retained API inventory

- Status: prepared for independent review; not final source closure.
- Date: 2026-10-24.
- Source reviewed: `0dcc767c68967f765412379f0234ed92c21fbc1a`.
- Retained diagnostic input: `/tmp/uo2-result-contract-inventory-final.tsv`,
  156 rows, SHA-256
  `a0166f99f9c1c30d30967bd5ad14d7d6ef56dbad5363788821d62d0ab90db840`.
- Current-source receipt:
  [`source-role-evidence.json`](results/0362ul7/source-role-evidence.json), 30,158
  bytes, SHA-256
  `95995c8693a5952a6b4dc2b0d273404571642a430050ab98128c54ff0c975d6b`.
- Independent source-role review: READY at receipt tip `fc2997619c`; the
  receipt content and the original source/inventory identities are unchanged
  by the subsequent rebase onto the gate-tooling baseline.
- Validator receipt: `tools/validate_api_disposition.py` now registers exactly
  49 current rows across six source-role decisions (38 raw, 4 composition,
  1 archive, 1 wire, 4 OPC source-plan and 1 OPC splice), while rejecting all
  six absent raw aliases and all near-match identities.
- Tool test receipt: `python3 -m unittest tools.test_validate_api_disposition -v`
  — 12 tests passed.

## Result

The old 156-row capture is not a final-source closure. The selected `mut_draft`
and `edit_type` rows divide as follows:

| Source role | Rows | Disposition in this record |
|---|---:|---|
| Current raw `Reader` methods | 38 | Exact source proof under 0351 §2.1.6 and 0359a §4; no package edit lifecycle. |
| Old raw paths absent at the source commit | 6 | Kept as stale diagnostic names, not allowlisted. The current canonical methods are recorded beside each row. |
| Composition/descriptor values | 11 | Exact definition and consumer proof; proposed exact-row categories remain pending review. |

The six absent names are:

* `litchi_doc::raw::Reader::{signatures_with,vba_signature_with,vba_with}`;
* `litchi_ppt::raw::Reader::signatures_with`;
* `litchi_xls::raw::Reader::{signatures_with,vba_with}`.

They must not be registered as raw-reader semantic rows. The receipt records
their nearest current method only so the old capture cannot accidentally
approve compatibility aliases.

The 11 descriptor rows are:

* composition: `litchi_core::SubEdit`,
  `litchi_doc::body_text::SubEdit`, `litchi_odb::SubEdit`, and
  `litchi_odg::edit::SubEdit`;
* raw descriptors: `litchi_iwa_archive::package::EntryEdit` and
  `litchi_iwa_common::wire::NestedFieldEdit`;
* OPC source-plan descriptors:
  `ContentTypeEdit`, `ContentTypeFragmentEdit`, `RelationshipEdit`, and
  `RelationshipFragmentEdit`;
* OPC XML splice descriptor: `OwnedElementEdit`.

The receipt gives each exact definition, consumer, receiver/field shape and
absence of a Package/Edit/Commit lifecycle. The validator registers only the
exact nominal categories above; it does not introduce a wildcard exception.

## Basis and pending work

The raw-reader evidence follows 0351 §2.1.6 and 0359a §4. The ODB/ODG
composition aliases follow 0358a §3; `NestedFieldEdit` and `EntryEdit` follow
the explicit descriptor reviews in 0361b and 0361c. No source-backed family,
Package owner, or unified `File` lifecycle is treated as a descriptor.

Pending before final acceptance:

1. A fresh all-68 diagnostic capture against the final assembled source, not
   reuse of the old 156-row capture.
2. Final coordinator acceptance of the capture and source-role receipt; no
   broad `*Edit`/`Reader` or substrate waiver is permitted.

No production source or `PROGRESS.md` was changed. No Cargo or capture command
was run; the Python tool test above is the only new execution receipt.

## Coordinator integration

The reviewed implementation `22aed11db8` merges at `14161bfd63`. Its older
tooling base overlaps the prior AST/wire registration, so the coordinator
resolves the two conflicts by retaining both sets: 34 prior exact identities
and 49 new ones, for 83 total. The six removed raw-reader aliases are absent
from every semantic category. Independent review verifies this union and
the unchanged one-to-one, duplicate, evidence/owner, scaffold and deferral
guards before the merge is committed.

The four composed tooling modules pass 91 tests on that resolved tree:

```text
python3 -m unittest tools.test_check_api_glossary tools.test_api_debt_inventory \
  tools.test_validate_api_disposition tools.test_final_api_inventory
Ran 91 tests in 1.298s — OK
```

The run precedes the merge commit on the same resolved source; no test log
was separately retained. The earlier 87-test run on `76c21c49cf`, before
this registration, has a retained external log and a reduced receipt at
[`composed-tooling-before-merge.json`](results/0362ul7/composed-tooling-before-merge.json).
These results validate tooling behavior. They do not replace final-source
rustdoc capture, per-row disposition review, or the full workspace gate.
