# 0362ul7 — source-role review of the retained API inventory

- Status: prepared for independent review; not final source closure.
- Date: 2026-10-24.
- Source reviewed: `0dcc767c68967f765412379f0234ed92c21fbc1a`.
- Retained diagnostic input: `/tmp/uo2-result-contract-inventory-final.tsv`,
  156 rows, SHA-256
  `a0166f99f9c1c30d30967bd5ad14d7d6ef56dbad5363788821d62d0ab90db840`.
- Current-source receipt:
  [`source-role-evidence.json`](0362ul7/source-role-evidence.json), 30,158
  bytes, SHA-256
  `95995c8693a5952a6b4dc2b0d273404571642a430050ab98128c54ff0c975d6b`.

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
absence of a Package/Edit/Commit lifecycle. It recommends exact nominal
categories only; it does not change the validator or introduce a wildcard
exception. The body-text and OPC categories still require independent review
because the current semantic checker registers only the exact AST/WIRE rows.

## Basis and pending work

The raw-reader evidence follows 0351 §2.1.6 and 0359a §4. The ODB/ODG
composition aliases follow 0358a §3; `NestedFieldEdit` and `EntryEdit` follow
the explicit descriptor reviews in 0361b and 0361c. No source-backed family,
Package owner, or unified `File` lifecycle is treated as a descriptor.

Pending before final acceptance:

1. Independent review of the 55-row receipt, including the exact nominal
   category predicates and the six stale-path separations.
2. A fresh all-68 diagnostic capture against the final assembled source, not
   reuse of the old 156-row capture.
3. Any checker registration must be exact-row and source-role based; no broad
   `*Edit`/`Reader` or substrate waiver is permitted.

No production source, validator, or `PROGRESS.md` was changed by this record.
No Cargo or capture command was run.
