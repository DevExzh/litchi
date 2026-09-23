#![allow(
    clippy::expect_used,
    clippy::unwrap_used,
    clippy::too_many_lines,
    reason = "golden fixtures favor explicit inputs and panic-driven assertions"
)]

//! Byte-exact goldens for the fresh DOC writer's text paths.
//!
//! Every fixture below was written by the writer at `6d989cad63`, before the
//! single-pass UTF-16 text encoding of change 0753, and its SHA-256 recorded
//! here. The same inputs must keep producing the same bytes: ASCII, Latin-1,
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
            "fd88d8bff13d61755347f5ec2d5659c240ed7ac0a70f9c95cbc3bd65af38871c",
        ),
        (
            "unicode_paragraphs",
            unicode_paragraphs,
            "334de3f380c1172fa5554904f0cb9450c8dbea61ed4328248daeda11f6df0d8b",
        ),
        (
            "main_story_fields",
            main_story_fields,
            "2bdc6278ad103485746f453240ec3530301fafb3e291cbcaee6b6368b6899480",
        ),
        (
            "tables",
            tables,
            "8665e185f5c59a5eb503cc64efa989514ab5b6fe168c3eb6cf38512fe913d498",
        ),
        (
            "headers_and_footers",
            headers_and_footers,
            "969a91b09342b97891c1b2145f57cd7d4b8e834f15c690e53fa77fe32331e466",
        ),
        (
            "notes_and_comments",
            notes_and_comments,
            "20447c54a9718e5e0343e3037028948edc85b1938eb9543c44cb9897ae2956f9",
        ),
        (
            "text_boxes",
            text_boxes,
            "cdf5835f9b699cf2c36c09f673062aa3455b361c963d946694562f1b1e836ae7",
        ),
        (
            "glossary_document",
            glossary_document,
            "5f59fad273a72d9fe90e86bbd8e2486a5d9b484df21fa3f24fe40847634ed223",
        ),
        (
            "attached_glossary_template",
            attached_glossary_template,
            "7e9790745194f8d9b94e45edc9f904886dbf50627fbdbfc6573f0d69c2bcdda3",
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
