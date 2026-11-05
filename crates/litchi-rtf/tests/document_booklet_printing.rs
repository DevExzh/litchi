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
use litchi_rtf::{DocumentBookletPrinting, write::Writer};

fn write(document: &Tree<'_>) -> std::io::Result<Vec<u8>> {
    let mut output = Vec::new();
    Writer::new(&mut output).write_document(document)?;
    Ok(output)
}

#[test]
fn parses_complete_producer_style_booklet_metadata_passively() {
    let document = Tree::parse(
        r"{\rtf1\bookfold\bookfoldrev\bookfoldsheets16 Body}",
        litchi_rtf::read::OpenOptions::default(),
    )
    .unwrap();
    assert_eq!(
        *document.booklet_printing(),
        DocumentBookletPrinting {
            book_fold: true,
            reverse_book_fold: true,
            sheets_per_booklet: Some(16),
        }
    );
    assert_eq!(document.text(), "Body");
}

#[test]
fn preserves_producer_observed_zero_separately_from_omission() {
    let explicit = Tree::parse(
        r"{\rtf1\bookfoldsheets0 Body}",
        litchi_rtf::read::OpenOptions::default(),
    )
    .unwrap();
    assert_eq!(explicit.booklet_printing().sheets_per_booklet, Some(0));
    assert_eq!(explicit.text(), "Body");

    let omitted = Tree::parse(r"{\rtf1 Body}", litchi_rtf::read::OpenOptions::default()).unwrap();
    assert!(omitted.booklet_printing().is_empty());
    let serialized = String::from_utf8(write(&omitted).unwrap()).unwrap();
    assert!(!serialized.contains("bookfold"));
}

#[test]
fn accepts_independent_flags_without_synthesizing_layout_or_printing() {
    for (source, book_fold, reverse) in [
        (r"{\rtf1\bookfold Body}", true, false),
        (r"{\rtf1\bookfoldrev Body}", false, true),
    ] {
        let document = Tree::parse(source, litchi_rtf::read::OpenOptions::default()).unwrap();
        assert_eq!(document.booklet_printing().book_fold, book_fold);
        assert_eq!(document.booklet_printing().reverse_book_fold, reverse);
        assert_eq!(document.text(), "Body");
        let serialized = String::from_utf8(write(&document).unwrap()).unwrap();
        assert!(!serialized.contains("\\landscape"));
        assert!(!serialized.contains("\\twoonone"));
    }
}

#[test]
fn typed_api_round_trips_boundaries_in_canonical_order_and_clears() {
    for sheets in [0, 4, 16, 2_147_483_644] {
        let mut document =
            Tree::parse(r"{\rtf1 Body}", litchi_rtf::read::OpenOptions::default()).unwrap();
        document.set_booklet_printing(DocumentBookletPrinting {
            book_fold: true,
            reverse_book_fold: true,
            sheets_per_booklet: Some(sheets),
        });
        let output = write(&document).unwrap();
        let serialized = String::from_utf8(output.clone()).unwrap();
        let controls = ["\\bookfold\\", "\\bookfoldrev\\", "\\bookfoldsheets"];
        for pair in controls.windows(2) {
            assert!(serialized.find(pair[0]).unwrap() < serialized.find(pair[1]).unwrap());
        }
        let reparsed =
            Tree::parse_bytes(&output, litchi_rtf::read::OpenOptions::default()).unwrap();
        assert_eq!(reparsed.booklet_printing(), document.booklet_printing());
        assert_eq!(reparsed.text(), "Body");
    }

    let mut document = Tree::parse(
        r"{\rtf1\bookfold Body}",
        litchi_rtf::read::OpenOptions::default(),
    )
    .unwrap();
    document.clear_booklet_printing();
    assert!(document.booklet_printing().is_empty());
    assert_eq!(document.text(), "Body");
}

#[test]
fn coexists_independently_with_adjacent_print_layout_metadata() {
    let document = Tree::parse(
        r"{\rtf1\gutterprl\twoonone\bookfold\bookfoldsheets8\bookfoldrev Body}",
        litchi_rtf::read::OpenOptions::default(),
    )
    .unwrap();
    let output = write(&document).unwrap();
    let reparsed = Tree::parse_bytes(&output, litchi_rtf::read::OpenOptions::default()).unwrap();
    assert_eq!(reparsed.booklet_printing(), document.booklet_printing());
    assert_eq!(
        reparsed.print_layout_settings(),
        document.print_layout_settings()
    );
    assert_eq!(reparsed.text(), "Body");
}

#[test]
fn rejects_invalid_values_parameters_duplicates_starred_grouped_and_late_forms() {
    for source in [
        r"{\rtf1\bookfoldsheets Body}",
        r"{\rtf1\bookfoldsheets-4 Body}",
        r"{\rtf1\bookfoldsheets1 Body}",
        r"{\rtf1\bookfoldsheets2 Body}",
        r"{\rtf1\bookfoldsheets5 Body}",
        r"{\rtf1\bookfoldsheets2147483647 Body}",
        r"{\rtf1\bookfoldsheets99999999999 Body}",
    ] {
        assert!(
            repair_support::rejected_or_parameter_repaired(&Tree::parse(
                source,
                litchi_rtf::read::OpenOptions::default()
            )),
            "accepted malformed {source}"
        );
    }
    for name in ["bookfold", "bookfoldrev"] {
        for suffix in ["0", "1", "2147483647", "99999999999"] {
            let source = format!(r"{{\rtf1\{name}{suffix} Body}}");
            assert!(
                repair_support::rejected_or_parameter_repaired(&Tree::parse(
                    &source,
                    litchi_rtf::read::OpenOptions::default()
                )),
                "accepted malformed {source}"
            );
        }
    }
    for control in ["bookfold", "bookfoldrev", "bookfoldsheets4"] {
        let name = control.trim_end_matches(|character: char| character.is_ascii_digit());
        for source in [
            format!(r"{{\rtf1\{control}\{control} Body}}"),
            format!(r"{{\rtf1{{\*\{control}}}Body}}"),
            format!(r"{{\rtf1{{\{control}}}Body}}"),
            format!(r"{{\rtf1 Body\{control}}}"),
        ] {
            assert!(
                repair_support::rejected_or_parameter_repaired(&Tree::parse(
                    &source,
                    litchi_rtf::read::OpenOptions::default()
                )),
                "accepted malformed {name}: {source}"
            );
        }
    }

    for invalid in [2, i32::MAX as u32 + 1] {
        let mut document =
            Tree::parse(r"{\rtf1 Body}", litchi_rtf::read::OpenOptions::default()).unwrap();
        document.set_booklet_printing(DocumentBookletPrinting {
            sheets_per_booklet: Some(invalid),
            ..DocumentBookletPrinting::default()
        });
        assert!(
            write(&document).is_err(),
            "serialized invalid value {invalid}"
        );
    }
}
