# Position-site survey at base `1d1044e3ac` (condensed)

Four read-only surveys classified every `buffer_position`/`error_position` hit
(`sites-at-base.txt`, 430 grep hits) in the OOXML crates and their XML-reading
neighbours, one per group: litchi-xlsx; litchi-pptx; litchi-docx; and
litchi-opc, litchi-ooxml-common, litchi-drawingml, litchi-xlsb,
litchi-spreadsheet-drawing, litchi-sign, litchi-crypto, litchi-formula,
litchi-xldm, litchi-ole-common and xml-minifier. For each site they recorded
the reader's input (raw part bytes, a stripped input, a generated buffer, a
fragment that starts at an element, a decoded string), how the position is used
(slice, splice, public span, diagnostic, relative arithmetic, bound, absolute
compare, seek) and a verdict. Their findings were read from code; the tests of
change 0765 (`../base-failures/`) are what confirm them.

No production code in these crates calls quick-xml's `error_position()`; the 23
DOCX `error_position` hits are a local paragraph index. No site uses another
quick-xml span API (`read_to_end` spans, `stream().offset()`).

## Verdicts at base

| group | hits | wrong splice | wrong slice | wrong public span | wrong diagnostic | already correct (stripped or +3) | no mark possible (fragment/generated) | relative only | test/comment |
| --- | ---: | ---: | ---: | ---: | ---: | ---: | ---: | ---: | ---: |
| litchi-xlsx | 114 | 47 | 17 | 11 | 0 | 1 | 16 | 3 | 0 |
| litchi-pptx | 105 | 46 | 17 | 19 | 0 | 5 | 16 | 0 | 2 |
| litchi-docx | 97 (+23 identifiers) | 30 | 7 | 4 | 0 | 8 | 40 | 1 | 7 |
| OPC, common, DrawingML, XLSB and neighbours | 91 | 39 | 14 | 7 | 2 | 4 | 13 | 7 | 5 |

The XLSX survey also counted 12 "coupled" hits whose ranges were matched
against litchi-opc's own shifted reader positions (`OwnedXmlPart`), and 7
lane-entry hits that declined marked parts.

## Silent corruption found by the surveys

A splice at a span three bytes early drops three bytes before the element and
leaves the element's last three bytes behind as character data. In an indented
part those are whitespace and a tail such as `me>` or `"/>`, so the output is
well-formed and passes an audit that checks well-formedness only.

* DrawingML theme scheme replacement (`theme/codec.rs` `direct_scheme_range`),
  reached from `litchi_pptx::shape::theme::{put_colors, put_fonts}`. **Confirmed
  at base** (`../base-failures/theme-silent-corruption.txt`).
* XLSX: `Workbook::edit()` cell edits on an indented marked worksheet (edit
  scanner spans), tab selection, sheet catalog edits, page margins/setup/print
  options, page breaks, conditional formatting, data validation, auto filter,
  data-type icons, smart tags, ActiveX controls, web extensions, slicer and
  timeline integration, protection removal.
* PPTX: caption tracks (`vatrue="1"` on any layout), placeholder text with an
  empty first run, `remove_slide_layout` and `add_slide_layout` on indented
  masters, collaboration extension removal, custom shows, content parts with
  MCE, the presentation font list, source-backed `clear_transition`, legacy
  comments with a processing instruction.
* DOCX: settings patches (attached template, document variables, mail merge,
  font embedding, document protection), external-hyperlink detachment and
  redaction, web-settings edits, standalone variables/mail-merge/settings
  codecs. The repository fixture `alt-chunk-header.docx` (Word-marked,
  pretty-printed `word/settings.xml`) reproduces the settings case.
* Shared helpers: `OwnedXmlPart::{update_elements, replace_element,
  remove_element}` with ranges from shifted callers; `custom_data`
  extension-list rewrite.

Silently wrong reads (no output written): DOCX `Document::paragraphs()` text
of a marked main part without MCE (confirmed at base), `OpaqueBlock::xml_bytes`,
`SourceDrawing` ranges, content-control spans; PPTX `Shape::span`/`xml` (confirmed
at base), modern-comment extension bytes, timing trees; DrawingML ink spans and
`Trace::data`; XLSX `Theme::format_scheme_xml` (with a panic on crafted input
when the shifted slice split a UTF-8 sequence), survey and slicer extension
bytes.

## Disposition at `2a2c4a9791`

Every position read that addresses bytes, forms a span or names an error offset
is converted with `litchi_core::xml::ReaderOrigin` of the reader's exact input
(`sites-after.tsv`: 364 converted reads; the rest are tests, comments and the 11
intentional reads below). Helpers that took a reader now take its origin, so the
compiler located every caller. Fragments that start at an element have origin
zero; they are converted anyway so every reader position in a crate is in the
same frame (several consumers compared their own positions with spans from
another reader).

Intentional raw reads: the XLSX lane entries (compaction, worksheet parser, edit
scanner) hand reader positions to `lane::Entry::locate`, which applies the
input's origin itself; `wire::position` feeds the edit scanner, whose `shift`
includes the origin; the crypto label parser's "declaration first" check stays in
reader coordinates; xml-minifier's audit keeps change 0677's marker handling.
litchi-xldm and litchi-formula have no litchi-core dependency and apply the same
three-byte rule locally, with their own tests.
