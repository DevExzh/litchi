# Existing XLSX form-control property fixtures

`native-corpus.json` records seven existing LibreOffice test-corpus packages
with ten `formControlPr` parts. Each entry keeps the original corpus `path`
for provenance and adds a `retained_path` for the tracked byte-for-byte copy
used by the Rust test. The package and part SHA-256 hashes and byte counts are
part of the manifest contract.

| Retained package | Original corpus path | Bytes | SHA-256 |
| --- | --- | ---: | --- |
| `tdf134769.xlsx` | `3rdparty/libreoffice-core/sc/qa/unit/data/xlsx/tdf134769.xlsx` | 16,892 | `50d4f7d5d0ffe17e2251b21cfb66bd5018e3b1791c7f5c27592a47a9d9e004b1` |
| `tdf161365.xlsx` | `3rdparty/libreoffice-core/sc/qa/unit/data/xlsx/tdf161365.xlsx` | 12,161 | `5c5ec45c1dd5db0a911e74d0e7e6cde5bf713b5ef38fcf84483903ddad583957` |
| `button-form-control.xlsx` | `3rdparty/libreoffice-core/sc/qa/unit/data/xlsx/button-form-control.xlsx` | 11,190 | `624321843560b792805592d604e7a645040f19f0ae1f661fec024b9aeec9613f` |
| `tdf120301_xmlSpaceParsing.xlsx` | `3rdparty/libreoffice-core/sc/qa/unit/data/xlsx/tdf120301_xmlSpaceParsing.xlsx` | 11,538 | `73709ecaa9270324eaf3f0c4fae8ec5e733ab639daa7b407f08d42538bac60bf` |
| `tdf60673.xlsx` | `3rdparty/libreoffice-core/sc/qa/unit/data/xlsx/tdf60673.xlsx` | 13,249 | `d9abab27b6cb7ae4eda4729473d8d48175f63fa66457e64a8134f9fa59689bc8` |
| `singlecontrol.xlsx` | `3rdparty/libreoffice-core/sc/qa/unit/data/xlsx/singlecontrol.xlsx` | 17,915 | `a13fd1a411a499f4b474bbe553c7950959ab8895131c40687522d14379885cde` |
| `checkbox-form-control.xlsx` | `3rdparty/libreoffice-core/sc/qa/unit/data/xlsx/checkbox-form-control.xlsx` | 11,394 | `8b602b21777bc06ca574f61783594caa5c37f68bdf66c66e58ef598758367038` |

The retained inputs include CheckBox, Button, and Radio objects. All ten
property parts have
`application/vnd.ms-excel.controlproperties+xml`, the Office 2006 control-
properties relationship type, and the Office 2010 form-control-properties
namespace. Each has one incoming internal relationship. The test verifies the
package and part hashes before parsing, then checks exact source-preserving
leaf no-ops.

These are actual property-part fixtures, unlike ActiveX persistence parts or
worksheet anchor `controlPr` elements. In `singlecontrol.xlsx`, the worksheet
owner is inside nested canonical MCE `Choice Requires="x14"` branches and uses
an `officeDocument/2006/relationships/ctrlProp` edge. A direct-child-only
scanner would miss that owner. This owner-selection evidence is retained for
future package lifecycle tests; the current test covers only the leaf property
part. The corpus is read-only evidence and makes no native package-write,
package relationship CRUD, Office acceptance, rendering, or control-execution
claim.

Fixture discovery inspected the member directories of 643 XLSX archives in
`3rdparty/libreoffice-core/sc/qa/unit/data/xlsx` and
`3rdparty/poi/test-data/spreadsheet`, looking for `ctrlProp` member names.
That bounded search found these seven packages; it is not an exhaustive claim
about all producer corpora. The original paths remain in the manifest as
provenance; test execution uses the retained copies under
`crates/litchi-xlsx/tests/fixtures/form_control_properties/` and therefore
does not require the `3rdparty` symlink.
