# Ordinary-save artifact review

Review of the fresh export packet at source revision
`0b4799d64973d33fdcd149d664f8e93c27cdc2d8`. The source census,
qualification identities, export manifest, and oracle report are bound by these
digests:

- source census: `8ce041173da709112a32cc4d0dbb43af32d2a33dd5753ba3394bb58fd9845f3a`
- qualification identities: `c34d823321ba1379088d1af682cda33a0ce46e80a5ba0f2a6979d5c87cd7fc7c`
- export manifest: `f194ff5cd9d95ed4b88c5737e6bdfd9ac29f6ae84c466cd3603250e882b0c5c1`
- oracle report: `6b3dea75746b5ecf207c5cfe58cd0fbfd042483e471e65695828af978447d6e0`

The packet covers seven corpora and 28 qualification rows. The independent
oracle passed, and the admission receipt binds the report to the manifest,
qualification identities, source census, build, and runner hashes. The three
relationship/stream controls and all eight content-type/workbook negative
controls rejected their deliberately invalid mutations.

The manifest contains 35 output archives: four durability policies plus the
sequential stream for each corpus. On-disk bytes and SHA-256 values agree with
all manifest records. Within every corpus the five outputs have one shared
published digest, `matches_reference` and `source_unchanged` are true, and
`reopen_admitted` is true. The exporter sets that declaration only after
reopening each output through its format owner; an independent ZIP read also
opened all 35 archives. The refused `alt-chunk-header` case is source-byte
identical on all five routes.

For `real-000-docx` (`NumberedList.docx`),
`word/_rels/document.xml.rels` has the same 11-edge graph in the source and
all five outputs, including relationship IDs, types, targets, and internal
target modes. The `word/footnotes.xml` and `word/endnotes.xml` payloads are
byte-identical in every output (source sizes 1,536 and 1,530 bytes; source
SHA-256 prefixes `c630c3ba2c6b4430` and `810a35859563fe6a`). The ordinary-save
marker is appended exactly once in each output. The Strict coverage is scoped
to these note relationship URIs: the tests inject Strict footnote/endnote
edges into an otherwise Transitional base package and verify preservation and
authored replacement. This does not establish full Strict OOXML dialect
conformance.

For `real-002-xlsx` (`ConditionalFormattingSamples.xlsx`), each of the five
outputs has 131 members versus 132 in the source. The only changed retained
members are `[Content_Types].xml`, `xl/_rels/workbook.xml.rels`,
`xl/workbook.xml`, and `xl/worksheets/sheet1.xml`; the only removed member is
`xl/calcChain.xml`. The content-type table is the exact allowed normalization:
remove the `.bin` printer-settings default and the calcChain override, then
add one printer-settings override for each of the 15 retained printer-settings
members. Effective content types for every retained member are unchanged. The
workbook relationship graph is the source graph minus the single internal
calcChain edge; workbook structure outside `calcPr` is unchanged, and the
dirty-calculation attributes match the contract. `Home!A1` contains the
frozen marker, and no other decoded cell value changed. Non-target worksheet
members remained byte-identical; styles and lexical markup within the edited
part are outside this oracle scope.

No artifact or binding blocker was found. This review does not certify timing
or allocation conclusions; those remain dependent on the native/allocation
lanes and their final packet seal.
