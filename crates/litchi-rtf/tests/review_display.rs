#![allow(
    clippy::expect_used,
    clippy::shadow_reuse,
    clippy::shadow_same,
    clippy::shadow_unrelated,
    clippy::unwrap_used,
    reason = "test assertions panic on failure by design and rebind fixture names across steps"
)]

use litchi_rtf::{DocumentReviewDisplay, RtfDocument, RtfWriter};

fn write(document: &RtfDocument<'_>) -> Vec<u8> {
    let mut output = Vec::new();
    RtfWriter::new(&mut output)
        .write_document(document)
        .unwrap();
    output
}

#[test]
fn parses_all_passive_review_display_flags_and_round_trips_in_stable_order() {
    let document =
        RtfDocument::parse(r"{\rtf1\donotshowinsdel\donotshowcomments\donotshowmarkup Body}")
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
    let reparsed = RtfDocument::parse_bytes(&output).unwrap();
    assert_eq!(reparsed.review_display(), document.review_display());
    assert_eq!(reparsed.text(), document.text());
}

#[test]
fn typed_api_mutates_and_clears_fixed_size_metadata() {
    let mut document = RtfDocument::parse(r"{\rtf1 Body}").unwrap();
    assert!(document.review_display().is_empty());
    document.set_review_display(DocumentReviewDisplay {
        hide_markup: false,
        hide_comments: true,
        hide_insertions_and_deletions: false,
    });
    let reparsed = RtfDocument::parse_bytes(&write(&document)).unwrap();
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
            RtfDocument::parse(source).is_err(),
            "accepted malformed {source}"
        );
    }
}

#[test]
fn display_preferences_do_not_remove_or_accept_review_content() {
    let document = RtfDocument::parse(concat!(
        r#"{\rtf1\donotshowmarkup\donotshowcomments\donotshowinsdel"#,
        r#"{\*\revtbl{Ada;}}{\revised\revauth0 kept}"#,
        r#"{\*\atnid C}{\*\atnauthor Ada}\chatn{\*\annotation note}}"#,
    ))
    .unwrap();
    assert_eq!(document.text(), "kept");
    assert_eq!(document.revisions().len(), 1);
    assert_eq!(document.annotations().len(), 1);
}
