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
use litchi_rtf::{DocumentEastAsianCompatibility, UnderlineStyle, write::Writer};

const CONTROLS: [&str; 6] = [
    "dntblnsbdb",
    "expshrtn",
    "nospaceforul",
    "noultrlspc",
    "noxlattoyen",
    "lnbrkrule",
];

fn write(doc: &Tree<'_>) -> String {
    let mut bytes = Vec::new();
    Writer::new(&mut bytes).write_document(doc).unwrap();
    String::from_utf8(bytes).unwrap()
}

#[test]
fn parses_complete_cluster_as_passive_metadata() {
    let rtf = concat!(
        r"{\rtf1\ansi\dntblnsbdb\expshrtn\nospaceforul",
        r"\noultrlspc\noxlattoyen\lnbrkrule\ul Body text }"
    );
    let doc = Tree::parse(rtf, litchi_rtf::read::OpenOptions::default()).unwrap();
    let compatibility = doc.east_asian_compatibility();
    assert!(compatibility.do_not_balance_sbcs_dbcs);
    assert!(compatibility.expand_spacing_at_shift_return);
    assert!(compatibility.do_not_add_space_for_underline);
    assert!(compatibility.do_not_underline_trailing_spaces);
    assert!(compatibility.do_not_translate_backslash_to_yen);
    assert!(compatibility.use_legacy_line_breaking_rules);
    assert_eq!(doc.text(), "Body text ");
    assert_ne!(doc.runs()[0].formatting.underline, UnderlineStyle::None);
}

#[test]
fn omission_is_empty_and_serializes_no_compatibility_flags() {
    let doc = Tree::parse(
        r"{\rtf1\ansi Body}",
        litchi_rtf::read::OpenOptions::default(),
    )
    .unwrap();
    assert!(doc.east_asian_compatibility().is_empty());
    let serialized = write(&doc);
    for name in CONTROLS {
        assert!(!serialized.contains(&format!("\\{name}")));
    }
}

#[test]
fn typed_api_round_trips_in_specification_order_and_clears_passively() {
    let mut doc = Tree::parse(
        r"{\rtf1\ansi Body}",
        litchi_rtf::read::OpenOptions::default(),
    )
    .unwrap();
    doc.set_east_asian_compatibility(DocumentEastAsianCompatibility {
        do_not_balance_sbcs_dbcs: true,
        expand_spacing_at_shift_return: true,
        do_not_add_space_for_underline: true,
        do_not_underline_trailing_spaces: true,
        do_not_translate_backslash_to_yen: true,
        use_legacy_line_breaking_rules: true,
    });

    let serialized = write(&doc);
    let positions: Vec<_> = CONTROLS
        .iter()
        .map(|name| serialized.find(&format!("\\{name}")).unwrap())
        .collect();
    assert!(positions.windows(2).all(|pair| pair[0] < pair[1]));
    let reparsed = Tree::parse(&serialized, litchi_rtf::read::OpenOptions::default()).unwrap();
    assert_eq!(
        *reparsed.east_asian_compatibility(),
        *doc.east_asian_compatibility()
    );
    assert_eq!(reparsed.text(), "Body");

    doc.clear_east_asian_compatibility();
    assert!(doc.east_asian_compatibility().is_empty());
    assert_eq!(doc.text(), "Body");
}

#[test]
fn rejects_parameters_duplicates_starred_grouped_and_late_flags() {
    for name in CONTROLS {
        let hostile = [
            format!(r"{{\rtf1\ansi\{name}0 Body}}"),
            format!(r"{{\rtf1\ansi\{name}\{name} Body}}"),
            format!(r"{{\rtf1\ansi{{\*\{name}}}Body}}"),
            format!(r"{{\rtf1\ansi{{\{name}}}Body}}"),
            format!(r"{{\rtf1\ansi Body\{name}}}"),
        ];
        for input in hostile {
            assert!(
                repair_support::rejected_or_parameter_repaired(&Tree::parse(
                    &input,
                    litchi_rtf::read::OpenOptions::default()
                )),
                "accepted {input}"
            );
        }
    }
    assert!(repair_support::rejected_or_parameter_repaired(
        &Tree::parse(
            r"{\rtf1\ansi\dntblnsbdb999999999999 Body}",
            litchi_rtf::read::OpenOptions::default()
        )
    ));
}

#[test]
fn parses_bundled_libreoffice_word6_japanese_compatibility_fixture() {
    let path = concat!(
        env!("CARGO_MANIFEST_DIR"),
        "/../../test-data/libreoffice-core/sw/qa/extras/layout/data/A020-min.rtf"
    );
    let bytes = std::fs::read(path).unwrap();
    let doc = Tree::parse_bytes(&bytes, litchi_rtf::read::OpenOptions::default()).unwrap();
    let compatibility = doc.east_asian_compatibility();
    assert!(compatibility.do_not_balance_sbcs_dbcs);
    assert!(compatibility.expand_spacing_at_shift_return);
    assert!(compatibility.do_not_add_space_for_underline);
    assert!(compatibility.do_not_underline_trailing_spaces);
    assert!(compatibility.do_not_translate_backslash_to_yen);
    assert!(!compatibility.use_legacy_line_breaking_rules);
}
