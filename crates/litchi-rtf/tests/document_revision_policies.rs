#![allow(
    clippy::expect_used,
    clippy::shadow_reuse,
    clippy::shadow_same,
    clippy::shadow_unrelated,
    clippy::unwrap_used,
    reason = "test assertions panic on failure by design and rebind fixture names across steps"
)]

use litchi_rtf::{DocumentRevisionPolicies, RtfDocument, RtfWriter};

fn write(document: &RtfDocument<'_>) -> Vec<u8> {
    let mut output = Vec::new();
    RtfWriter::new(&mut output)
        .write_document(document)
        .unwrap();
    output
}

#[test]
fn parses_all_explicit_policy_combinations_without_enabling_tracking() {
    for (moves, formatting) in [(false, false), (false, true), (true, false), (true, true)] {
        let source = format!(
            r"{{\rtf1\trackmoves{}\trackformatting{} Body}}",
            i32::from(moves),
            i32::from(formatting)
        );
        let document = RtfDocument::parse(&source).unwrap();
        assert_eq!(
            *document.revision_policies(),
            DocumentRevisionPolicies {
                track_moves: Some(moves),
                track_formatting: Some(formatting),
            }
        );
        assert_eq!(document.text(), "Body");
        assert!(document.revisions().is_empty());
    }
}

#[test]
fn omission_remains_unspecified_and_is_not_serialized() {
    let document = RtfDocument::parse(r"{\rtf1 Body}").unwrap();
    assert!(document.revision_policies().is_empty());
    let serialized = String::from_utf8(write(&document)).unwrap();
    assert!(!serialized.contains("trackmoves"));
    assert!(!serialized.contains("trackformatting"));
}

#[test]
fn typed_api_round_trips_in_stable_order_and_clears_without_creating_revisions() {
    let mut document = RtfDocument::parse(r"{\rtf1 Body}").unwrap();
    document.set_revision_policies(DocumentRevisionPolicies {
        track_moves: Some(true),
        track_formatting: Some(false),
    });
    let output = write(&document);
    let serialized = String::from_utf8(output.clone()).unwrap();
    assert!(
        serialized.find("\\trackmoves1").unwrap() < serialized.find("\\trackformatting0").unwrap()
    );
    let reparsed = RtfDocument::parse_bytes(&output).unwrap();
    assert_eq!(reparsed.revision_policies(), document.revision_policies());
    assert_eq!(reparsed.text(), "Body");
    assert!(reparsed.revisions().is_empty());

    document.clear_revision_policies();
    assert!(document.revision_policies().is_empty());
    assert_eq!(document.text(), "Body");
}

#[test]
fn coexists_with_embedding_and_xml_policies_as_independent_metadata() {
    let document = RtfDocument::parse(concat!(
        r#"{\rtf1\trackformatting1\validatexml0\donotembedsysfont1"#,
        r#"\trackmoves0\showxmlerrors1 Body}"#,
    ))
    .unwrap();
    let output = write(&document);
    let reparsed = RtfDocument::parse_bytes(&output).unwrap();
    assert_eq!(reparsed.revision_policies(), document.revision_policies());
    assert_eq!(reparsed.embedding_policies(), document.embedding_policies());
    assert_eq!(reparsed.xml_policies(), document.xml_policies());
    assert_eq!(reparsed.text(), "Body");
}

#[test]
fn rejects_missing_non_boolean_overflow_and_duplicate_values() {
    for name in ["trackmoves", "trackformatting"] {
        for suffix in ["", "-1", "2", "32767", "99999999999"] {
            let source = format!(r"{{\rtf1\{name}{suffix} Body}}");
            assert!(
                RtfDocument::parse(&source).is_err(),
                "accepted malformed {source}"
            );
        }
        let source = format!(r"{{\rtf1\{name}0\{name}1 Body}}");
        assert!(
            RtfDocument::parse(&source).is_err(),
            "accepted duplicate {source}"
        );
    }
}

#[test]
fn rejects_every_starred_grouped_and_late_revision_policy() {
    for control in [r"\trackmoves0", r"\trackformatting1"] {
        for source in [
            format!(r"{{\rtf1{{\*{control}}}Body}}"),
            format!(r"{{\rtf1{{{control}}}Body}}"),
            format!(r"{{\rtf1 Body{control}}}"),
        ] {
            assert!(
                RtfDocument::parse(&source).is_err(),
                "accepted malformed {source}"
            );
        }
    }
}
