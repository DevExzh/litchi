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
use litchi_rtf::{DocumentReviewDisplay, write::Writer};

fn write(document: &Tree<'_>) -> Vec<u8> {
    let mut output = Vec::new();
    Writer::new(&mut output).write_document(document).unwrap();
    output
}

#[test]
fn parses_all_passive_review_display_flags_and_round_trips_in_stable_order() {
    let document = Tree::parse(
        r"{\rtf1\donotshowinsdel\donotshowcomments\donotshowmarkup Body}",
        litchi_rtf::read::OpenOptions::default(),
    )
    .unwrap();
    assert_eq!(
        *document.review_display(),
        DocumentReviewDisplay {
            hide_markup: true,
            hide_comments: true,
            hide_insertions_and_deletions: true,
        }
    );
    assert_eq!(document.text(), "Body");

    let output = write(&document);
    let serialized = String::from_utf8(output.clone()).unwrap();
    assert!(
        serialized.find("\\donotshowmarkup").unwrap()
            < serialized.find("\\donotshowcomments").unwrap()
    );
    assert!(
        serialized.find("\\donotshowcomments").unwrap()
            < serialized.find("\\donotshowinsdel").unwrap()
    );
    let reparsed = Tree::parse_bytes(&output, litchi_rtf::read::OpenOptions::default()).unwrap();
    assert_eq!(reparsed.review_display(), document.review_display());
    assert_eq!(reparsed.text(), document.text());
}

#[test]
fn typed_api_mutates_and_clears_fixed_size_metadata() {
    let mut document =
        Tree::parse(r"{\rtf1 Body}", litchi_rtf::read::OpenOptions::default()).unwrap();
    assert!(document.review_display().is_empty());
    document.set_review_display(DocumentReviewDisplay {
        hide_markup: false,
        hide_comments: true,
        hide_insertions_and_deletions: false,
    });
    let reparsed =
        Tree::parse_bytes(&write(&document), litchi_rtf::read::OpenOptions::default()).unwrap();
    assert_eq!(reparsed.review_display(), document.review_display());
    document.clear_review_display();
    assert!(document.review_display().is_empty());
}

#[test]
fn rejects_parameters_duplicates_starred_nested_and_late_flags() {
    for source in [
        r"{\rtf1\donotshowmarkup0 Body}",
        r"{\rtf1\donotshowcomments1 Body}",
        r"{\rtf1\donotshowinsdel-1 Body}",
        r"{\rtf1\donotshowmarkup\donotshowmarkup Body}",
        r"{\rtf1\donotshowcomments\donotshowcomments Body}",
        r"{\rtf1\donotshowinsdel\donotshowinsdel Body}",
        r"{\rtf1{\*\donotshowmarkup}Body}",
        r"{\rtf1{\*\donotshowcomments}Body}",
        r"{\rtf1{\donotshowinsdel nested}Body}",
        r"{\rtf1 Body\donotshowmarkup}",
    ] {
        assert!(
            repair_support::rejected_or_parameter_repaired(&Tree::parse(
                source,
                litchi_rtf::read::OpenOptions::default()
            )),
            "accepted malformed {source}"
        );
    }
}

#[test]
fn display_preferences_do_not_remove_or_accept_review_content() {
    let document = Tree::parse(
        concat!(
            r#"{\rtf1\donotshowmarkup\donotshowcomments\donotshowinsdel"#,
            r#"{\*\revtbl{Ada;}}{\revised\revauth0 kept}"#,
            r#"{\*\atnid C}{\*\atnauthor Ada}\chatn{\*\annotation note}}"#,
        ),
        litchi_rtf::read::OpenOptions::default(),
    )
    .unwrap();
    assert_eq!(document.text(), "kept");
    assert_eq!(document.revisions().len(), 1);
    assert_eq!(document.annotations().len(), 1);
}
