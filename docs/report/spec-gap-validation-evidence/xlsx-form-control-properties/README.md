# Existing XLSX form-control property fixtures

`native-corpus.json` records seven existing LibreOffice test-corpus packages
with ten `formControlPr` parts. Package and part SHA-256 hashes bind the
inputs; the manifest also records root attributes, declared content types,
and incoming relationship metadata. It contains no generated output and
establishes no Office acceptance, rendering, or control-execution claim.

The inputs include CheckBox, Button, and Radio objects. All ten parts have
`application/vnd.ms-excel.controlproperties+xml` and one incoming internal
relationship. Their root namespace is
`http://schemas.microsoft.com/office/spreadsheetml/2009/9/main`.

These are actual property-part fixtures, unlike ActiveX persistence parts or
worksheet anchor `controlPr` elements. In `singlecontrol.xlsx`, the worksheet
owner is inside nested canonical MCE `Choice Requires="x14"` branches and uses
an `officeDocument/2006/relationships/ctrlProp` edge. A direct-child-only
scanner would miss that owner. Tests must prove effective owner selection and
preserve the complete worksheet, VML, drawing, and unrelated package bytes.

Fixture discovery inspected the member directories of 643 XLSX archives in
`3rdparty/libreoffice-core/sc/qa/unit/data/xlsx` and
`3rdparty/poi/test-data/spreadsheet`, looking for `ctrlProp` member names.
That bounded search found these seven packages; it is not an exhaustive claim
about all producer corpora. The source files remain in their original corpus
locations and are not duplicated in this directory.
