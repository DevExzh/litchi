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
use litchi_rtf::{StyleType, write::Writer};

const TABLE_STYLE_SOURCE: &str = concat!(
    r#"{\rtf1\ansi{\stylesheet{\ql Normal;}"#,
    r#"{\*\ts16\tsrowd\b \tscfirstrow\tsclastrow\tscfirstcol\tsclastcol"#,
    r#"\tscbandhorzodd\tscbandverteven\tscbandsh2\tscbandsv3 Table List;}}Body}"#,
);

fn write(document: &Tree<'_>) -> Vec<u8> {
    let mut output = Vec::new();
    Writer::new(&mut output).write_document(document).unwrap();
    output
}

#[test]
fn parses_table_style_conditional_formatting_and_round_trips() {
    let document =
        Tree::parse(TABLE_STYLE_SOURCE, litchi_rtf::read::OpenOptions::default()).unwrap();
    assert_eq!(document.text(), "Body");
    let style = document
        .stylesheet()
        .styles()
        .iter()
        .find(|style| style.style_type == StyleType::Table)
        .expect("table style");
    let conditional = style.table_conditional;
    assert!(conditional.row_defaults_marker);
    assert!(conditional.first_row);
    assert!(conditional.last_row);
    assert!(conditional.first_column);
    assert!(conditional.last_column);
    assert!(conditional.band_horizontal_odd);
    assert!(!conditional.band_horizontal_even);
    assert!(!conditional.band_vertical_odd);
    assert!(conditional.band_vertical_even);
    assert_eq!(conditional.horizontal_band_size, Some(2));
    assert_eq!(conditional.vertical_band_size, Some(3));

    let output = write(&document);
    let serialized = String::from_utf8(output.clone()).unwrap();
    for control in [
        "\\tsrowd",
        "\\tscfirstrow",
        "\\tsclastrow",
        "\\tscfirstcol",
        "\\tsclastcol",
        "\\tscbandhorzodd",
        "\\tscbandverteven",
        "\\tscbandsh2",
        "\\tscbandsv3",
    ] {
        assert!(
            serialized.contains(control),
            "missing {control} in {serialized}"
        );
    }

    let reparsed = Tree::parse_bytes(&output, litchi_rtf::read::OpenOptions::default()).unwrap();
    let reparsed_style = reparsed
        .stylesheet()
        .styles()
        .iter()
        .find(|style| style.style_type == StyleType::Table)
        .unwrap();
    assert_eq!(reparsed_style.table_conditional, conditional);
}

#[test]
fn table_style_without_conditional_metadata_stays_empty() {
    let document = Tree::parse(
        r"{\rtf1\ansi{\stylesheet{\ql Normal;}{\*\ts16\b Table List;}}Body}",
        litchi_rtf::read::OpenOptions::default(),
    )
    .unwrap();
    let style = document
        .stylesheet()
        .styles()
        .iter()
        .find(|style| style.style_type == StyleType::Table)
        .unwrap();
    assert!(style.table_conditional.is_empty());

    let reparsed =
        Tree::parse_bytes(&write(&document), litchi_rtf::read::OpenOptions::default()).unwrap();
    assert!(
        reparsed.stylesheet().styles()[1]
            .table_conditional
            .is_empty()
    );
}

#[test]
fn repeated_table_style_id_is_distinguished_by_conditional_scope() {
    let document = Tree::parse(
        concat!(
            r"{\rtf1{\stylesheet",
            r"{\*\ts16\tscfirstrow\b Example;}",
            r"{\*\ts16\i Example;}",
            r"{\*\ts16\tsclastrow\ul Example;}}Body}"
        ),
        litchi_rtf::read::OpenOptions::default(),
    )
    .unwrap();
    assert_eq!(document.stylesheet().styles().len(), 3);
    let base = document
        .stylesheet()
        .get_typed(StyleType::Table, 16)
        .unwrap();
    assert!(base.table_conditional.is_empty());
    assert!(base.formatting.italic);
    let reparsed =
        Tree::parse_bytes(&write(&document), litchi_rtf::read::OpenOptions::default()).unwrap();
    assert_eq!(reparsed.stylesheet(), document.stylesheet());
}

#[test]
fn repeated_table_scope_is_not_a_new_style_when_only_band_size_changes() {
    for source in [
        r"{\rtf1{\stylesheet{\*\ts1\tscfirstrow One;}{\*\ts1\tscfirstrow Two;}}}",
        r"{\rtf1{\stylesheet{\*\ts1 One;}{\*\ts1\tsrowd Two;}}}",
        r"{\rtf1{\stylesheet{\*\ts1\tscbandsh1 One;}{\*\ts1\tscbandsh2 Two;}}}",
        r"{\rtf1{\stylesheet{\s1 One;}{\s1 Two;}}}",
        r"{\rtf1{\stylesheet{\*\cs1 One;}{\*\cs1 Two;}}}",
    ] {
        assert!(
            Tree::parse(source, litchi_rtf::read::OpenOptions::default()).is_err(),
            "{source}"
        );
    }
}

#[test]
fn transfer_copies_all_conditional_definitions_and_checks_scoped_conflicts() {
    use litchi_rtf::{Package, edit::TransferPlan};
    let source = Package::parse(concat!(
        r"{\rtf1{\stylesheet{\*\ts16 Base;}",
        r"{\*\ts16\tscfirstrow\b Base;}",
        r"{\*\ts16\tsclastrow\i Base;}}Source}"
    ))
    .unwrap();
    let target = Package::parse(r"{\rtf1{\stylesheet{\*\ts16 Base;}}Target}").unwrap();
    let copied = TransferPlan::style(&source, StyleType::Table, 16, &target)
        .unwrap()
        .commit()
        .unwrap()
        .into_package();
    let reopened = Package::from_bytes(&copied.to_bytes().unwrap()).unwrap();
    assert_eq!(reopened.raw().stylesheet(), source.raw().stylesheet());
    assert_eq!(reopened.text(), "Target");
    assert_eq!(target.raw().stylesheet().styles().len(), 1);

    let conflict = Package::parse(concat!(
        r"{\rtf1{\stylesheet{\*\ts16 Base;}",
        r"{\*\ts16\tscfirstrow\i Base;}}Target}"
    ))
    .unwrap();
    assert!(TransferPlan::style(&source, StyleType::Table, 16, &conflict).is_err());
    assert_eq!(conflict.raw().stylesheet().styles().len(), 2);
}

#[test]
fn rejects_misplaced_or_duplicate_table_style_controls() {
    let cases = [
        // Conditional controls in a paragraph style.
        r"{\rtf1{\stylesheet{\s1\tscfirstrow Para;}}}",
        r"{\rtf1{\stylesheet{\s1\tsrowd Para;}}}",
        r"{\rtf1{\stylesheet{\s1\tscbandsh2 Para;}}}",
        // Conditional controls in a character style.
        r"{\rtf1{\stylesheet{\*\cs1\tscfirstcol Char;}}}",
        // Duplicate controls.
        r"{\rtf1{\stylesheet{\*\ts16\tscfirstrow\tscfirstrow T;}}}",
        r"{\rtf1{\stylesheet{\*\ts16\tsrowd\tsrowd T;}}}",
        r"{\rtf1{\stylesheet{\*\ts16\tscbandsh2\tscbandsh3 T;}}}",
        // Parameterized flag controls.
        r"{\rtf1{\stylesheet{\*\ts16\tscfirstrow1 T;}}}",
        r"{\rtf1{\stylesheet{\*\ts16\tsrowd0 T;}}}",
        // Missing band-size parameter.
        r"{\rtf1{\stylesheet{\*\ts16\tscbandsh T;}}}",
        // Out-of-range band size.
        r"{\rtf1{\stylesheet{\*\ts16\tscbandsv65536 T;}}}",
    ];
    for rtf in cases {
        assert!(
            repair_support::rejected_or_parameter_repaired(&Tree::parse(
                rtf,
                litchi_rtf::read::OpenOptions::default()
            )),
            "accepted malformed {rtf}"
        );
    }
}
