#![allow(
    clippy::expect_used,
    clippy::shadow_reuse,
    clippy::shadow_same,
    clippy::shadow_unrelated,
    clippy::unwrap_used,
    reason = "test assertions panic on failure by design and rebind fixture names across steps"
)]

use litchi_rtf::raw::Tree;
use litchi_rtf::{StyleType, write::Writer};
use std::fs;

fn write_stylesheet(document: &Tree<'_>) -> Vec<u8> {
    let mut output = Vec::new();
    let mut writer = Writer::new(&mut output);
    writer.write_document_header().unwrap();
    writer.write_stylesheet(document.stylesheet()).unwrap();
    writer.write_str("}").unwrap();
    output
}

#[test]
fn parses_implicit_normal_all_types_list_metadata_and_round_trips() {
    let document = Tree::parse(
        concat!(
            r#"{\rtf1\ansi\ansicpg1250\uc2{\stylesheet"#,
            r#"{\b\fs22\snext0 Normal;}"#,
            r#"{\s1\i\sbasedon0\snext1\slink2\spriority9\styrsid42 Heading \u268?x;}"#,
            r#"{\*\cs2\additive\sbasedon0\slink1 Emphasis;}"#,
            r#"{\*\ds3 Section;}"#,
            r#"{\*\ts4 Table;}"#,
            r#"{\s5\ls4\ilvl2 List;}"#,
            r#"}Body}"#,
        ),
        litchi_rtf::read::OpenOptions::default(),
    )
    .unwrap();
    let sheet = document.stylesheet();
    assert_eq!(sheet.styles().len(), 6);
    assert_eq!(
        sheet.get_typed(StyleType::Paragraph, 0).unwrap().name,
        "Normal"
    );
    assert!(
        sheet
            .get_typed(StyleType::Paragraph, 0)
            .unwrap()
            .formatting
            .bold
    );
    assert_eq!(
        sheet.get_typed(StyleType::Paragraph, 1).unwrap().name,
        "Heading \u{10c}"
    );
    assert_eq!(
        sheet
            .get_typed(StyleType::Paragraph, 5)
            .unwrap()
            .paragraph
            .unwrap()
            .list_override,
        Some(4)
    );
    assert_eq!(
        sheet
            .get_typed(StyleType::Paragraph, 5)
            .unwrap()
            .paragraph
            .unwrap()
            .list_level,
        Some(2)
    );
    assert!(sheet.get_typed(StyleType::Character, 2).is_some());
    assert!(sheet.get_typed(StyleType::Section, 3).is_some());
    assert!(sheet.get_typed(StyleType::Table, 4).is_some());
    let chain = sheet.inheritance_chain(StyleType::Paragraph, 1).unwrap();
    assert_eq!(
        chain.iter().map(|style| style.id).collect::<Vec<_>>(),
        vec![0, 1]
    );

    let reparsed = Tree::parse_bytes(
        &write_stylesheet(&document),
        litchi_rtf::read::OpenOptions::default(),
    )
    .unwrap();
    assert_eq!(reparsed.stylesheet(), sheet);
}

#[test]
fn rejects_malformed_scope_order_duplicates_cycles_and_bounds() {
    for source in [
        r"{\rtf1{\stylesheet{\s0 A;}}{\stylesheet{\s1 B;}}}",
        r"{\rtf1{\b{\stylesheet{\s0 A;}}}}",
        r"{\rtf1 Body{\stylesheet{\s0 A;}}}",
        r"{\rtf1{\stylesheet{\cs1 Character;}}}",
        r"{\rtf1{\stylesheet{\b\s1 Late;}}}",
        r"{\rtf1{\stylesheet{\s1\s2 Duplicate;}}}",
        r"{\rtf1{\stylesheet{\s1 No terminator}}}",
        r"{\rtf1{\stylesheet{\s1 A;}{\s1 B;}}}",
        r"{\rtf1{\stylesheet{\s1\sbasedon2 A;}{\s2\sbasedon1 B;}}}",
        r"{\rtf1{\stylesheet{\s1\sbasedon0\sbasedon0 A;}}}",
        r"{\rtf1{\stylesheet{\s1\shidden\shidden0 A;}}}",
        r"{\rtf1{\stylesheet{\s1\spriority1\spriority2 A;}}}",
        r"{\rtf1{\stylesheet{\s1\spriority100 A;}}}",
        r"{\rtf1{\stylesheet{\s1\styrsid-1 A;}}}",
    ] {
        assert!(
            Tree::parse(source, litchi_rtf::read::OpenOptions::default()).is_err(),
            "accepted {source}"
        );
    }
}

#[test]
fn parses_real_libreoffice_stylesheet_types_and_implicit_normal() {
    const FIXTURES: &[&str] = &[
        "sw/qa/core/data/rtf/pass/tdf116851.rtf",
        "sw/qa/extras/rtfexport/data/fdo55504-1-min.rtf",
        "sw/qa/writerfilter/filters-test/data/pass/TCI-TN65GP-DDRHDLL-partial.rtf",
    ];
    let root = concat!(
        env!("CARGO_MANIFEST_DIR"),
        "/../../test-data/libreoffice-core"
    );
    for fixture in FIXTURES {
        let bytes = fs::read(format!("{root}/{fixture}")).unwrap();
        let document = Tree::parse_bytes(&bytes, litchi_rtf::read::OpenOptions::default())
            .unwrap_or_else(|error| panic!("failed to parse {fixture}: {error}"));
        let sheet = document.stylesheet();
        assert_eq!(
            sheet.get_typed(StyleType::Paragraph, 0).unwrap().name,
            "Normal"
        );
        assert!(
            sheet
                .styles()
                .iter()
                .any(|style| style.style_type == StyleType::Character)
        );
        assert!(
            sheet
                .styles()
                .iter()
                .any(|style| style.style_type == StyleType::Table)
        );
        let reparsed = Tree::parse_bytes(
            &write_stylesheet(&document),
            litchi_rtf::read::OpenOptions::default(),
        )
        .unwrap();
        assert_eq!(reparsed.stylesheet(), sheet);
    }
}
