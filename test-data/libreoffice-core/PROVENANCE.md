# LibreOffice core fixture provenance (record 0257)

Files under this directory keep their path inside the LibreOffice core
source tree (<https://git.libreoffice.org/core>). Metafiles are described in
`METAFILE_PROVENANCE.md`. The files below were copied byte-for-byte on
2026-10-13 from the workspace's local, untracked LibreOffice core checkout,
so that tests and tools no longer read that checkout. The checkout carries no
git metadata, so each copy is pinned by SHA-256 instead of by upstream
commit.

License: MPL-2.0 (the LibreOffice core `COPYING.MPL`); the retained license
text is `test-data/odf/native-resave/source/LICENSE-MPL-2.0.txt`. The filter
registry fragments also carry their own MPL-2.0 / Apache-2.0 headers.

## Test documents

| Path | SHA-256 | Used by |
|---|---|---|
| `sw/qa/extras/ooxmlexport/data/Encrypted_LO_Standard_abc.docx` | `d6b55065e5c8e6debd4a72a957275c22e1f5373a1806e40a1a2bd34334cce073` | `crates/litchi-crypto/tests/native_ooxml_encryption.rs` |
| `sw/qa/extras/ooxmlexport/data/Encrypted_MSO2007_abc.docx` | `ea105a1eb01653c904320ef9ab426686cc468b09503f617725c41cee2f6f549f` | same |
| `sw/qa/extras/ooxmlexport/data/Encrypted_MSO2010_abc.docx` | `7a51bb14cafbc1904779fa176701702251b55c2f9a5b064ba237a92eb3ce73ed` | same |
| `sw/qa/extras/ooxmlexport/data/Encrypted_MSO2013_abc.docx` | `2db365c48e5f3b0302b2f6c0f56b298b15f819e95c6448073566baad910023d3` | same |
| `sw/qa/extras/ooxmlexport/data/multi-page-toc.docx` | `88dbf98cd74204fe364cd426b5a6bc1c7b4c13906e41f876c57f213e0a7ab624` | `crates/litchi-ops/tests/refresh.rs` (record 0279; copied 2026-10-14) |
| `sw/qa/extras/ooxmlexport/data/testTOCFlag_u.docx` | `a27a420de11920520b3e5974bfe9f8fe2323b2ef0b0337fe032ca0b80f8b82e4` | same |
| `sw/qa/extras/ooxmlexport/data/custom-styles-TOC-comma.docx` | `07759e2c9d57a90f5d82be3eeb2292951894ccb605a8d462967181cece11035c` | same |
| `oox/qa/unit/data/testTdf132557_footerCustomShapes.pptx` | `2eb077834e80191fcb8d1540beaac49288537364f19e60c3ce68e79fc9cc9e61` | same |
| `sc/qa/unit/data/xlsb/tdf94627.xlsb` | `b7fed651b265af036afd36a294fb1d8fcb6229f4e5213732610d146a970386a6` | `crates/litchi-xlsb/src/calculation_chain/tests.rs` |
| `sc/qa/unit/data/xlsb/pivottable_error_item_filter.xlsb` | `8e6f1e9e571d92c2f50ee63a7ad23a382b990ccf3efc36fa18f19ca2d652be58` | same |
| `sc/qa/unit/data/xlsb/shared_formula.xlsb` | `e06ae79ed345495eee0bf5e7fd7a32638af0b646c5819f7d83a4f6fb97a151a3` | same |
| `sc/qa/unit/data/xlsx/tdf169496_hidden_graphic.xlsx` | `0b647da300a085f39914fdfae961463ae9e54ffe772b2e0eb9860a841ab93f72` | `crates/litchi-xlsx/tests/drawing_svg_read.rs` |

## Filter registry fragments

`filter/source/config/fragments/` holds, unchanged:

- `types/*.xcu`: every type fragment (188 files), so the "no ODI type is
  registered" check in `tools/test_native_odf_resave.py` still scans the
  whole type registry;
- `filters/*.xcu`: only the 17 filter fragments that
  `tools/native_odf_resave.py` names or that `tools/test_native_odf_resave.py`
  inspects (the other 230 filter fragments are not vendored).
