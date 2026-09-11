# XLSB Theme-family host integration requirements

This batch connects the previously implemented shared `themeFamily` fragment
model to an existing complete DrawingML Theme part and the XLSB Theme owner.
It does not claim completion of the broader specification audit.

## Normative and producer evidence

- `[MS-XLSB]` 2.1.7.52 assigns the workbook Theme relationship and part to XLSB.
- `[MS-ODRAWXML]` 2.2.8 extends the direct `a:theme/a:extLst/a:ext` owner.
  Its table gives `http://schemas.microsoft.com/office/thememl/2012/main`
  as the extension URI for the `themeFamily` child. ECMA `CT_OfficeArtExtension`
  declares `uri` as `xsd:token`; supported identifiers must be compared after
  XML token whitespace normalization while preserving source spelling.
- `[MS-ODRAWXML]` 2.4 and appendix 5.17 define the family namespace and the
  required unqualified `name`, `id`, and `vid` attributes. The shared fragment
  implementation owns those values, limits, and nested extension grammar.
- Native fixtures `test-data/ooxml/xlsb/date.xlsb`, `hyperlink.xlsb`,
  `bug66682.xlsb`, and `cond_format.xlsb` use outer extension URI
  `{05A4C25C-085E-4340-85A3-A5531E510DB2}`. This is a specifically supported
  producer profile, not a license to infer owners under arbitrary URIs.
  `date.xlsb` and `hyperlink.xlsb` have identical 8,390-byte `xl/theme/theme1.xml`
  payloads, SHA-256
  `50b662d8ff0e562157ff80aa19c85ff214cd21357d66457661d070f50c6fdc59`.

## Required behavior

1. Read only an unambiguous, namespace-resolved direct family owner under the
   two supported outer extension identifiers. Foreign ancestry and unprocessed
   MCE branches must not become direct typed owners through descendant scans.
   Selective mutations must refuse relevant MCE-wrapped lists/extensions even
   when empty, and MCE children directly inside admitted extensions. Foreign
   ancestry and unrecognized URI subtrees remain opaque; see
   `../theme-family-mce-refusal/`.
2. Support optional read, add, scalar update, and removal through the existing
   isolated XLSB Theme transaction, including edits combined with base-theme
   values. Reopen the resulting XML before publishing a commit.
3. Preserve unrelated bytes, unknown children/attributes, extension identifiers,
   and package graph state. Removing the selected family is not permission to
   discard its enclosing extension's unrelated content. Remove the selected
   `ext`, then its root `extLst`, when only XML whitespace remains. Comments,
   unrelated content, and foreign attributes retain their container. This
   later closure requirement supersedes the original wrapper-retention policy;
   see `../theme-family-removal-closure/` for fresh validation.
4. Resolve inherited namespaces safely and avoid prefix collisions when
   inserting XML. Retain unchanged lexical XML and exact semantic no-op source
   allocations where required by the existing Theme contract.
5. Preserve source-checked publication and exact inverse restoration. Reject
   stale source, ambiguous supported owners, malformed recognized content,
   exhausted limits, and signature-policy violations atomically. Direct owned
   extension containers must respect their element-only content model. A
   standalone fragment byte-order mark must not become text during embedding.
6. Keep the XML grammar in DrawingML and workbook graph/publication in XLSB.
   A source-backed view retains managed `PartData`; accessing family metadata
   must not load unrelated media or add ambient I/O.
7. Keep resource limits finite and bound insertion output before constructing
   it. New APIs must not introduce a second persistent copy of complete Theme
   XML merely to retain family metadata.

## Evidence scope

Independent tests and review cover ownership, preservation, source checks,
resource boundaries, and publication. Offline XSD validation covers complete
Theme XML plus recognized family subtrees; this does not prove that Office
applications accept newly authored workbooks. Performance measurements are
scoped to the recorded instrumented workloads, toolchain, machine, and corpus.
Existing owner tests remain applicable to package topology and whole-part
lifecycle. Any unimplemented lifecycle surface is recorded explicitly.
