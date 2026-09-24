#![allow(
    clippy::expect_used,
    clippy::unwrap_used,
    clippy::too_many_lines,
    reason = "golden fixtures favor explicit inputs and panic-driven assertions"
)]

//! Byte-exact goldens for the fresh DOC writer's text paths.
//!
//! Every fixture below was first written by the writer at `6d989cad63`, before
//! the single-pass UTF-16 text encoding of change 0753, and its SHA-256
//! recorded here. The spec-gap merge re-pinned each digest to the output of
//! the incoming writer at `a67a38abf2`, which changes the fresh DOC bytes but
//! does not contain the 0753 encoding: that writer and the merged writer
//! produce identical bytes for all nine fixtures, so the merged 0753 text
//! paths still reproduce a writer without them exactly. The same inputs must
//! keep producing the same bytes: ASCII, Latin-1,
//! CJK and supplementary-plane text, empty runs, runs longer than 64 KiB,
//! field characters next to surrogate pairs, and every story that carries
//! text (main, table, header/footer, footnote, endnote, comment, main and
//! header text boxes, glossary and attached glossary).

use litchi_core::validation::EvidenceDigest;
use litchi_doc::writer::{
    CharacterFormatting, CommentEntry, FloatingPosition, FootnoteEntry, Kind as DrawingKind,
    ParagraphFormatting, Shape as DrawingShape,
};
use litchi_doc::{
    GlossaryItem, GlossaryItemKind, GlossaryMetadata, GlossaryStyle, HeaderFooterParagraph,
    HeaderKind, Writer,
};
use std::io::Cursor;

const FIELD_BEGIN: char = '\u{13}';
const FIELD_SEPARATOR: char = '\u{14}';
const FIELD_END: char = '\u{15}';

fn repeat_to(seed: &str, bytes: usize) -> String {
    let mut text = String::with_capacity(bytes + seed.len());
    while text.len() < bytes {
        text.push_str(seed);
    }
    text
}

fn bold() -> CharacterFormatting {
    CharacterFormatting {
        bold: Some(true),
        ..CharacterFormatting::default()
    }
}

fn special() -> CharacterFormatting {
    CharacterFormatting {
        special: Some(true),
        ..CharacterFormatting::default()
    }
}

fn written(mut writer: Writer) -> Vec<u8> {
    let mut output = Cursor::new(Vec::new());
    writer.write_to(&mut output).unwrap();
    output.into_inner()
}

fn ascii_paragraphs() -> Vec<u8> {
    let mut writer = Writer::new();
    writer.add_paragraph("").unwrap();
    writer.add_paragraph("a").unwrap();
    writer
        .add_paragraph("litchi deterministic ASCII paragraph")
        .unwrap();
    writer
        .add_paragraph(&repeat_to("ascii payload 0123456789 ", 20_000))
        .unwrap();
    writer
        .add_paragraph(&repeat_to("long run beyond 64 KiB. ", 70_001))
        .unwrap();
    writer
        .add_paragraph("\u{7f} control-adjacent \u{1} \u{8} \t tab")
        .unwrap();
    written(writer)
}

fn unicode_paragraphs() -> Vec<u8> {
    let mut writer = Writer::new();
    writer
        .add_paragraph("Latin-1: àéîõüçñ ÀÉÎÕÜÇÑ ß ÿ")
        .unwrap();
    writer.add_paragraph("CJK: 漢字かなカナ한국어").unwrap();
    writer.add_paragraph("Astral: 😀𝄞🦀").unwrap();
    writer.add_paragraph("😀").unwrap();
    writer
        .add_paragraph("\u{80}\u{7ff}\u{800}\u{ffff}\u{10000}\u{10ffff}")
        .unwrap();
    writer
        .add_paragraph(&repeat_to("混合 mixed ✓ 😀 text ", 90_000))
        .unwrap();
    writer
        .add_paragraph_runs(
            vec![
                (String::new(), bold()),
                ("A😀".to_string(), CharacterFormatting::default()),
                (String::new(), CharacterFormatting::default()),
                ("𝄞B".to_string(), bold()),
                ("ascii tail".to_string(), CharacterFormatting::default()),
                ("é".to_string(), bold()),
            ],
            ParagraphFormatting::default(),
        )
        .unwrap();
    writer
        .add_paragraph_runs(Vec::new(), ParagraphFormatting::default())
        .unwrap();
    written(writer)
}

fn main_story_fields() -> Vec<u8> {
    let mut writer = Writer::new();
    writer.add_paragraph("Before the fields").unwrap();
    writer
        .add_paragraph(&format!(
            "{FIELD_BEGIN} PAGE {FIELD_SEPARATOR}1{FIELD_END} ascii field"
        ))
        .unwrap();
    writer
        .add_paragraph(&format!(
            "😀{FIELD_BEGIN} HYPERLINK \"https://example.test\" {FIELD_SEPARATOR}link 🦀{FIELD_END}"
        ))
        .unwrap();
    writer
        .add_paragraph_runs(
            vec![
                (format!("𝄞{FIELD_BEGIN}"), special()),
                (" DATE ".to_string(), CharacterFormatting::default()),
                (format!("{FIELD_SEPARATOR}"), special()),
                ("héllo".to_string(), CharacterFormatting::default()),
                (format!("{FIELD_END}"), special()),
                ("after".to_string(), CharacterFormatting::default()),
            ],
            ParagraphFormatting::default(),
        )
        .unwrap();
    written(writer)
}

fn tables() -> Vec<u8> {
    let mut writer = Writer::new();
    writer.add_paragraph("Table follows").unwrap();
    let table = writer.add_table(2, 3).unwrap();
    writer
        .set_table_cell_text(table, 0, 0, "ascii cell")
        .unwrap();
    writer.set_table_cell_text(table, 0, 1, "").unwrap();
    writer.set_table_cell_text(table, 0, 2, "漢字 😀").unwrap();
    writer
        .set_table_cell_text(
            table,
            1,
            0,
            &format!("{FIELD_BEGIN} PAGE {FIELD_SEPARATOR}2{FIELD_END}"),
        )
        .unwrap();
    writer
        .set_table_cell_text(
            table,
            1,
            1,
            &format!("𝄞{FIELD_BEGIN} NUMPAGES {FIELD_SEPARATOR}9{FIELD_END}é"),
        )
        .unwrap();
    writer
        .set_table_cell_text(table, 1, 2, &repeat_to("wide cell ✓ ", 5_000))
        .unwrap();
    written(writer)
}

fn headers_and_footers() -> Vec<u8> {
    let mut writer = Writer::new();
    writer.add_paragraph("Body").unwrap();
    writer
        .set_odd_header_paragraphs(vec![
            HeaderFooterParagraph::plain("Odd header ascii"),
            HeaderFooterParagraph::field(" PAGE ", "1", CharacterFormatting::default()).unwrap(),
            HeaderFooterParagraph::from_runs(
                vec![
                    ("😀 ".to_string(), bold()),
                    (String::new(), CharacterFormatting::default()),
                    ("漢字".to_string(), CharacterFormatting::default()),
                ],
                ParagraphFormatting::default(),
            ),
        ])
        .unwrap();
    writer
        .set_even_header_paragraphs(vec![HeaderFooterParagraph::plain("Even 𝄞 header")])
        .unwrap();
    writer
        .set_odd_footer_paragraphs(vec![
            HeaderFooterParagraph::field(" NUMPAGES ", "héllo 🦀", bold()).unwrap(),
        ])
        .unwrap();
    writer
        .set_first_footer_paragraphs(vec![HeaderFooterParagraph::plain(repeat_to(
            "first footer ",
            3_000,
        ))])
        .unwrap();
    written(writer)
}

fn notes_and_comments() -> Vec<u8> {
    let mut writer = Writer::new();
    writer
        .add_paragraph("Reference text for notes and comments")
        .unwrap();
    writer
        .add_paragraph("Second paragraph with ünïcode 😀")
        .unwrap();
    writer.add_footnote(FootnoteEntry::new(1, "Footnote ascii", 1));
    writer.add_footnote(FootnoteEntry::new(3, "Footnote 🦀 astral", 1));
    writer.add_footnote(FootnoteEntry::new(5, "", 1));
    writer.add_endnote(FootnoteEntry::new(2, "Endnote 漢字", 1));
    writer.add_endnote(FootnoteEntry::new(40, repeat_to("endnote body ", 4_000), 1));
    writer.add_comment(CommentEntry::new(4, "Review 🦀", "Alice 😀", "A😀"));
    writer.add_comment(CommentEntry::new(6, "plain comment", "Bob", "BB"));
    writer.add_comment(CommentEntry::new(8, "", "Carol", "C"));
    written(writer)
}

fn text_boxes() -> Vec<u8> {
    let mut writer = Writer::new();
    writer.add_paragraph("Body with text boxes").unwrap();
    writer
        .set_odd_header_paragraphs(vec![HeaderFooterParagraph::plain("Header")])
        .unwrap();
    writer
        .insert_floating_text_box(
            DrawingShape::new(DrawingKind::Rectangle, 2000, 1000).unwrap(),
            FloatingPosition::new(1440, 1440),
            "Main box\nsecond line\r\nthird 😀\rfourth\r\r\nlast",
        )
        .unwrap();
    writer
        .insert_floating_text_box(
            DrawingShape::new(DrawingKind::Rectangle, 1000, 500).unwrap(),
            FloatingPosition::new(720, 720),
            "",
        )
        .unwrap();
    writer
        .insert_header_text_box(
            HeaderKind::Odd,
            DrawingShape::new(DrawingKind::Rectangle, 4000, 2000).unwrap(),
            FloatingPosition::new(1000, 500),
            "Watermark 漢字\nDraft 🦀",
        )
        .unwrap();
    writer.add_paragraph("After the boxes").unwrap();
    written(writer)
}

fn glossary_metadata() -> GlossaryMetadata {
    GlossaryMetadata::try_new(
        vec![
            GlossaryItem::try_new("greeting", GlossaryItemKind::NamedAutoText, Some(0), 0, 9)
                .unwrap(),
            GlossaryItem::try_new("teh", GlossaryItemKind::FormattedAutoCorrect, None, 9, 15)
                .unwrap(),
        ],
        vec![GlossaryStyle::try_new("Normal", 1).unwrap()],
        15,
        16,
        17,
    )
    .unwrap()
}

fn glossary_writer() -> Writer {
    let mut writer = Writer::new();
    writer.add_paragraph("Greeting").unwrap();
    writer.add_paragraph("World").unwrap();
    writer.add_paragraph("").unwrap();
    writer.add_paragraph("").unwrap();
    writer.set_glossary_metadata(glossary_metadata());
    writer
}

fn glossary_document() -> Vec<u8> {
    written(glossary_writer())
}

fn attached_glossary_template() -> Vec<u8> {
    let mut template = Writer::new();
    template.add_paragraph("Template body 😀").unwrap();
    template.set_attached_glossary(glossary_writer()).unwrap();
    written(template)
}

#[test]
fn fresh_doc_writer_text_paths_match_the_pre_0753_goldens() {
    let fixtures: [(&str, fn() -> Vec<u8>, &str); 9] = [
        (
            "ascii_paragraphs",
            ascii_paragraphs,
            "ac3ec26bad3fec491f1e7e5371bffa90928546afe4a9cf4fe0f9ee81090906c0",
        ),
        (
            "unicode_paragraphs",
            unicode_paragraphs,
            "7058630276283726732af3036e90b216a5da38834685013f97eb59ab212eb2ae",
        ),
        (
            "main_story_fields",
            main_story_fields,
            "94e9e068236543274d3bace63ea67bb9f533d158c5d063d5fd142044ee9206ce",
        ),
        (
            "tables",
            tables,
            "b132acb223e7168e37a529136c8d90eab67f6552808ce062a650022a9aa0ad89",
        ),
        (
            "headers_and_footers",
            headers_and_footers,
            "7182aece8dc42c91177239d319d84bf5168bd33a3840a5597d707bf072d25e51",
        ),
        (
            "notes_and_comments",
            notes_and_comments,
            "e8ce076ac5c04225e8e2f63bfdb2eea3eed81a157f1b9c6ebdbcf5cd3544ff75",
        ),
        (
            "text_boxes",
            text_boxes,
            "482aeac3d553aa7b6fa7a94284ea357f54ac9452841b690bf29352fb94f4f0d5",
        ),
        (
            "glossary_document",
            glossary_document,
            "4901ae815515a73a07e3703abad209e07fc57c8c31d09c9c27439011800086b2",
        ),
        (
            "attached_glossary_template",
            attached_glossary_template,
            "444e956843618b7b8dcc1a3da56460a72df4e57174e745863711eb533fa853bd",
        ),
    ];
    let mut mismatches = Vec::new();
    for (name, build, expected) in fixtures {
        let first = build();
        assert_eq!(first, build(), "{name} is not deterministic");
        let actual = EvidenceDigest::of(&first).to_string();
        if actual != expected {
            mismatches.push(format!("{name}: {actual} ({} bytes)", first.len()));
        }
    }
    assert!(
        mismatches.is_empty(),
        "digest mismatches:\n{}",
        mismatches.join("\n")
    );
}

/// Writing the same writer twice gives the same bytes: the text stream is
/// built per write and nothing is left behind in the writer.
#[test]
fn a_writer_written_twice_writes_the_same_bytes() {
    let mut writer = Writer::new();
    writer.add_paragraph("once 😀").unwrap();
    writer
        .add_paragraph(&repeat_to("twice payload ", 40_000))
        .unwrap();
    let mut first = Cursor::new(Vec::new());
    writer.write_to(&mut first).unwrap();
    let mut second = Cursor::new(Vec::new());
    writer.write_to(&mut second).unwrap();
    assert_eq!(first.into_inner(), second.into_inner());
}
