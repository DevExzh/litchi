#![allow(
    clippy::expect_used,
    clippy::shadow_reuse,
    clippy::shadow_same,
    clippy::shadow_unrelated,
    clippy::unwrap_used,
    reason = "test assertions panic on failure by design and rebind fixture names across steps"
)]

mod repair_support;

use litchi_rtf::raw::Tree;
use litchi_rtf::{SectionRendering, write::Writer};

fn write(document: &Tree<'_>) -> Vec<u8> {
    let mut output = Vec::new();
    Writer::new(&mut output).write_document(document).unwrap();
    output
}

#[test]
fn parses_section_rendering_and_column_balance_and_round_trips() {
    let document = Tree::parse(
        r"{\rtf1\ansi\sectd\vertsect\nocolbal\cols2\pgnstarts6 Body}",
        litchi_rtf::read::OpenOptions::default(),
    )
    .unwrap();
    assert_eq!(document.text(), "Body");
    let section = &document.sections()[0];
    assert_eq!(
        section.properties.rendering,
        Some(SectionRendering::Vertical)
    );
    assert!(!section.properties.balance_columns);

    let output = write(&document);
    let serialized = String::from_utf8(output.clone()).unwrap();
    assert!(serialized.contains("\\vertsect"));
    assert!(serialized.contains("\\nocolbal"));
    let reparsed = Tree::parse_bytes(&output, litchi_rtf::read::OpenOptions::default()).unwrap();
    assert_eq!(reparsed.text(), document.text());
    assert_eq!(
        reparsed.sections()[0].properties.rendering,
        Some(SectionRendering::Vertical)
    );
    assert!(!reparsed.sections()[0].properties.balance_columns);
}

#[test]
fn parses_horizontal_rendering_and_balanced_default() {
    let document = Tree::parse(
        r"{\rtf1\ansi\sectd\horzsect Body}",
        litchi_rtf::read::OpenOptions::default(),
    )
    .unwrap();
    let section = &document.sections()[0];
    assert_eq!(
        section.properties.rendering,
        Some(SectionRendering::Horizontal)
    );
    assert!(section.properties.balance_columns);

    let output = write(&document);
    let serialized = String::from_utf8(output.clone()).unwrap();
    assert!(serialized.contains("\\horzsect"));
    assert!(!serialized.contains("nocolbal"));
    let reparsed = Tree::parse_bytes(&output, litchi_rtf::read::OpenOptions::default()).unwrap();
    assert_eq!(
        reparsed.sections()[0].properties.rendering,
        Some(SectionRendering::Horizontal)
    );

    let implicit = Tree::parse(
        r"{\rtf1\ansi\sectd Body}",
        litchi_rtf::read::OpenOptions::default(),
    )
    .unwrap();
    assert!(implicit.sections()[0].properties.rendering.is_none());
    assert!(implicit.sections()[0].properties.balance_columns);
}

#[test]
fn parses_paragraph_suppression_flags_and_round_trips() {
    let document = Tree::parse(
        r"{\rtf1\ansi\pard\noline\notabind Suppressed\par\pard Normal\par}",
        litchi_rtf::read::OpenOptions::default(),
    )
    .unwrap();
    let suppressed = document
        .blocks()
        .iter()
        .find(|block| block.text.contains("Suppressed"))
        .unwrap();
    assert!(suppressed.paragraph.no_line_numbering);
    assert!(suppressed.paragraph.no_auto_tab_indent);
    let normal = document
        .blocks()
        .iter()
        .find(|block| block.text.contains("Normal"))
        .unwrap();
    assert!(!normal.paragraph.no_line_numbering);
    assert!(!normal.paragraph.no_auto_tab_indent);

    let output = write(&document);
    let serialized = String::from_utf8(output.clone()).unwrap();
    assert!(serialized.contains("\\noline"));
    assert!(serialized.contains("\\notabind"));
    let reparsed = Tree::parse_bytes(&output, litchi_rtf::read::OpenOptions::default()).unwrap();
    let reparsed_block = reparsed
        .blocks()
        .iter()
        .find(|block| block.text.contains("Suppressed"))
        .unwrap();
    assert!(reparsed_block.paragraph.no_line_numbering);
    assert!(reparsed_block.paragraph.no_auto_tab_indent);
}

#[test]
fn rejects_parameterized_section_and_paragraph_flags() {
    for rtf in [
        r"{\rtf1\sectd\vertsect1 Body}",
        r"{\rtf1\sectd\horzsect0 Body}",
        r"{\rtf1\sectd\nocolbal1 Body}",
        r"{\rtf1\pard\noline1 Text\par}",
        r"{\rtf1\pard\notabind0 Text\par}",
    ] {
        assert!(
            repair_support::rejected_or_parameter_repaired(&Tree::parse(
                rtf,
                litchi_rtf::read::OpenOptions::default()
            )),
            "accepted malformed {rtf}"
        );
    }
}
