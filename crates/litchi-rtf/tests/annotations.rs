use litchi_rtf::{Annotation, NavigationEntry, RtfDocument, RtfWriter};
use std::borrow::Cow;

#[test]
fn parses_full_comment_metadata_range_unicode_and_hidden_body() {
    let source = r#"{\rtf1\ansi A{\*\atrfstart 7}\u20320?{\*\atrfend 7}{\*\atnid AM}{\*\atnauthor Ada \u20320?}\chatn{\*\annotation{\*\atnref 7}{\*\atndate -12}{\*\atnparent root}{\*\atnicn 2}{\*\atntime 42}Review \u-10179?\u-8704?}Z}"#;
    let document = RtfDocument::parse(source).unwrap();
    assert_eq!(document.text(), "A你Z");
    let annotation = &document.annotations()[0];
    assert!(annotation.has_reference);
    assert_eq!(annotation.id, 7);
    assert_eq!(annotation.position, 1);
    assert_eq!(annotation.range_end, "A你".len());
    assert_eq!(annotation.author, "Ada 你");
    assert_eq!(annotation.initials, "AM");
    assert_eq!(annotation.date.as_deref(), Some("-12"));
    assert_eq!(annotation.parent_id.as_deref(), Some("root"));
    assert_eq!(annotation.icon.as_deref(), Some("2"));
    assert_eq!(annotation.time.as_deref(), Some("42"));
    assert_eq!(annotation.text, "Review 😀");
}

#[test]
fn parses_bundled_libreoffice_comment_fixtures() {
    let root = std::path::Path::new(env!("CARGO_MANIFEST_DIR"))
        .join("../../3rdparty/libreoffice-core/sw/qa/extras/rtfexport/data");
    if !root.exists() {
        return;
    }
    for fixture in [
        "text-with-comment.rtf",
        "tdf136445-1-min.rtf",
        "fdo38244.rtf",
        "tdf94377.rtf",
        "fdo63428.rtf",
    ] {
        let source = std::fs::read_to_string(root.join(fixture)).unwrap();
        let document = RtfDocument::parse(&source)
            .unwrap_or_else(|error| panic!("failed to parse {fixture}: {error}"));
        assert!(!document.annotations().is_empty(), "{fixture}");
    }
}

#[test]
fn point_and_range_comments_round_trip_without_visible_text_duplication() {
    let source = r#"{\rtf1 A{\*\atnid }{\*\atnauthor Point}\chatn{\*\annotation point}B{\*\atrfstart 9}\u20320?{\*\atrfend 9}{\*\atnid R}{\*\atnauthor Range}\chatn{\*\annotation{\*\atnref 9}range}C}"#;
    let document = RtfDocument::parse(source).unwrap();
    assert!(!document.annotations()[0].has_reference);
    let mut output = Vec::new();
    RtfWriter::new(&mut output)
        .write_document(&document)
        .unwrap();
    let output = String::from_utf8(output).unwrap();
    let reparsed = RtfDocument::parse(&output).unwrap();
    assert_eq!(reparsed.text(), document.text());
    assert_eq!(reparsed.annotations(), document.annotations());
}

#[test]
fn typed_mutation_validates_positions_identity_and_coexists_with_navigation_marks() {
    let mut document = RtfDocument::parse(r#"{\rtf1 Body{\tc\v Heading}}"#).unwrap();
    assert!(matches!(
        document.navigation_entries()[0],
        NavigationEntry::TableOfContents(_)
    ));
    let mut annotation = Annotation::comment(4, Cow::Borrowed("Ada"), Cow::Borrowed("Review"));
    annotation.position = 0;
    annotation.range_end = 4;
    document.push_annotation(annotation.clone()).unwrap();
    assert!(document.push_annotation(annotation).is_err());
    document.clear_annotations();
    assert!(document.annotations().is_empty());
}

#[test]
fn rejects_conflicts_orphans_active_data_and_invalid_metadata() {
    for source in [
        r#"{\rtf1{\*\atrfstart nope}x}"#,
        r#"{\rtf1{\*\atrfend 1}x}"#,
        r#"{\rtf1{\*\atrfstart 1}{\*\atrfstart 1}x}"#,
        r#"{\rtf1{\*\atrfstart 1}x}"#,
        r#"{\rtf1{\*\atnauthor A}{\*\atnauthor B}\chatn{\*\annotation x}}"#,
        r#"{\rtf1{\*\atnauthor A}{\*\atnid X}{\*\annotation x}}"#,
        r#"{\rtf1\chatn{\*\annotation{\*\atnref nope}x}}"#,
        r#"{\rtf1\chatn{\*\annotation{\*\atndate 1}{\*\atndate 2}x}}"#,
        r#"{\rtf1\chatn{\*\annotation{\field danger}}}"#,
        r#"{\rtf1\chatn{\*\annotation{\object danger}}}"#,
        "{\\rtf1\\chatn{\\*\\annotation\\bin4 abcd}}",
    ] {
        assert!(RtfDocument::parse(source).is_err(), "{source}");
    }
}
