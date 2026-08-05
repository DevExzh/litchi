//! Regression tests for the index-mark owner.

use super::*;

mod parser {
    use super::*;

    const TEXT: &str = "urn:oasis:names:tc:opendocument:xmlns:text:1.0";

    #[test]
    fn parses_point_range_and_bibliography_marks_in_document_order() {
        let xml = format!(
            r#"<x:text xmlns:x="{TEXT}" xmlns:u="urn:vendor"><x:p><x:toc-mark-start x:id="same" x:outline-level="2" u:flag="yes"/>Alpha<x:span>&amp;</x:span><x:toc-mark x:string-value="Manual &amp; Entry"/><x:alphabetical-index-mark-start x:id="same" x:key1="K" x:main-entry="1"/>Beta<![CDATA[!]]><x:alphabetical-index-mark-end x:id="same"/></x:p><x:p>Next<x:toc-mark-end x:id="same"/><x:bibliography-mark x:bibliography-type="book" x:author="A &amp; B" x:title="T">[1] &amp; more</x:bibliography-mark><x:user-index-mark x:index-name="Custom" x:string-value="User" x:outline-level="3"/></x:p></x:text>"#
        );
        let marks = parse_text_index_marks(&xml).unwrap();
        assert_eq!(marks.len(), 5);
        assert_eq!(marks[0].kind(), TextIndexMarkKind::TableOfContents);
        assert!(marks[0].is_range());
        assert_eq!(marks[0].id(), Some("same"));
        assert_eq!(marks[0].value(), "Alpha&Beta!\nNext");
        assert_eq!(marks[0].attribute(Some(TEXT), "outline-level"), Some("2"));
        assert_eq!(marks[0].attribute(Some("urn:vendor"), "flag"), Some("yes"));

        assert_eq!(marks[1].kind(), TextIndexMarkKind::TableOfContents);
        assert!(!marks[1].is_range());
        assert_eq!(marks[1].value(), "Manual & Entry");

        assert_eq!(marks[2].kind(), TextIndexMarkKind::Alphabetical);
        assert_eq!(marks[2].id(), Some("same"));
        assert_eq!(marks[2].value(), "Beta!");
        assert_eq!(marks[2].attribute(Some(TEXT), "key1"), Some("K"));

        assert_eq!(marks[3].kind(), TextIndexMarkKind::Bibliography);
        assert_eq!(marks[3].value(), "[1] & more");
        assert_eq!(marks[3].attribute(Some(TEXT), "author"), Some("A & B"));
        assert_eq!(
            marks[3].attribute(Some(TEXT), "bibliography-type"),
            Some("book")
        );

        assert_eq!(marks[4].kind(), TextIndexMarkKind::User);
        assert_eq!(marks[4].value(), "User");
        assert_eq!(marks[4].attribute(Some(TEXT), "index-name"), Some("Custom"));
    }

    #[test]
    fn index_marks_reject_missing_ambiguous_and_unmatched_metadata() {
        let missing = format!(r#"<x:toc-mark xmlns:x="{TEXT}"/>"#);
        assert!(parse_text_index_marks(&missing).is_err());
        let unmatched = format!(r#"<x:toc-mark-end xmlns:x="{TEXT}" x:id="a"/>"#);
        assert!(parse_text_index_marks(&unmatched).is_err());
        let unclosed = format!(r#"<x:toc-mark-start xmlns:x="{TEXT}" x:id="a"/>"#);
        assert!(parse_text_index_marks(&unclosed).is_err());
        let duplicate = format!(
            r#"<x:p xmlns:x="{TEXT}"><x:toc-mark-start x:id="a"/><x:toc-mark-start x:id="a"/></x:p>"#
        );
        assert!(parse_text_index_marks(&duplicate).is_err());
        let aliases = format!(
            r#"<x:toc-mark xmlns:x="{TEXT}" xmlns:y="{TEXT}" x:string-value="A" y:string-value="B"/>"#
        );
        assert!(parse_text_index_marks(&aliases).is_err());
        let invalid_level = format!(
            r#"<x:user-index-mark xmlns:x="{TEXT}" x:index-name="I" x:string-value="V" x:outline-level="0"/>"#
        );
        assert!(parse_text_index_marks(&invalid_level).is_err());
        let invalid_boolean = format!(
            r#"<x:alphabetical-index-mark xmlns:x="{TEXT}" x:string-value="A" x:main-entry="yes"/>"#
        );
        assert!(parse_text_index_marks(&invalid_boolean).is_err());
        let invalid_bibliography_type = format!(
            r#"<x:bibliography-mark xmlns:x="{TEXT}" x:bibliography-type="novel">bad</x:bibliography-mark>"#
        );
        assert!(parse_text_index_marks(&invalid_bibliography_type).is_err());
        let bibliography_child = format!(
            r#"<x:bibliography-mark xmlns:x="{TEXT}" x:bibliography-type="book"><x:span>bad</x:span></x:bibliography-mark>"#
        );
        assert!(parse_text_index_marks(&bibliography_child).is_err());
        assert!(parse_text_index_marks("<x:toc-mark>").is_err());

        let empty_strings =
            format!(r#"<x:user-index-mark xmlns:x="{TEXT}" x:index-name="" x:string-value=""/>"#);
        assert_eq!(
            parse_text_index_marks(&empty_strings).unwrap()[0].value(),
            ""
        );
    }

    #[test]
    fn index_marks_enforce_nesting_bound() {
        let mut xml = format!(r#"<x:p xmlns:x="{TEXT}">"#);
        for _ in 0..MAX_MARK_DEPTH {
            xml.push_str("<x:span>");
        }
        for _ in 0..MAX_MARK_DEPTH {
            xml.push_str("</x:span>");
        }
        xml.push_str("</x:p>");
        assert!(parse_text_index_marks(&xml).is_err());
    }
}

mod package {
    use super::super::package::{TEXT, validated_marks};
    use super::*;
    use crate::{BibliographyField, TextBibliographyType};

    const OFFICE: &str = "urn:oasis:names:tc:opendocument:xmlns:office:1.0";
    fn document(body: &str) -> String {
        format!(
            "<office:document-content xmlns:office=\"{OFFICE}\" xmlns:text=\"{TEXT}\"><office:body><office:text>{body}</office:text></office:body></office:document-content>"
        )
    }

    #[test]
    fn canonical_point_and_range_fragments_round_trip() {
        let toc = TextIndexMark::toc_point("Manual & entry", Some(2)).unwrap();
        assert_eq!(
            toc.to_xml_fragments().unwrap(),
            TextIndexMarkFragments::Point(format!(
                "<text:toc-mark xmlns:text=\"{TEXT}\" text:outline-level=\"2\" text:string-value=\"Manual &amp; entry\"/>"
            ))
        );
        let user = TextIndexMark::user_range("r1", "Custom", Some(1)).unwrap();
        let TextIndexMarkFragments::Range { start, end } = user.to_xml_fragments().unwrap() else {
            panic!()
        };
        let xml = document(&format!("<text:p>{start}Visible{end}</text:p>"));
        let parsed = validated_marks(&xml).unwrap();
        assert_eq!(parsed[0].value(), "Visible");
        assert_eq!(
            parsed[0].attribute(Some(TEXT), "index-name"),
            Some("Custom")
        );
    }

    #[test]
    fn insertion_replacement_and_removal_preserve_unrelated_bytes() {
        let xml = document("<text:p><!--keep-->alpha<text:span>beta</text:span></text:p><text:p/>");
        let range = TextIndexMark::alphabetical_range(
            "a1",
            TextAlphabeticalMarkMetadata {
                key1: Some("A".to_string()),
                ..Default::default()
            },
        )
        .unwrap();
        let xml = insert_text_index_mark_xml(&xml, 0, &range).unwrap();
        assert!(xml.contains("<!--keep-->alpha<text:span>beta</text:span>"));
        let replacement = TextIndexMark::toc_range("t1", Some(1)).unwrap();
        let xml = replace_text_index_mark_xml(&xml, 0, &replacement).unwrap();
        let xml = remove_text_index_mark_xml(&xml, 0).unwrap();
        assert!(xml.contains("<text:p><!--keep-->alpha<text:span>beta</text:span></text:p>"));
        let bibliography = TextIndexMark::bibliography_point(
            TextBibliographyType::Www,
            "[Test]",
            vec![(BibliographyField::Identifier, "Test".to_string())],
        )
        .unwrap();
        let xml = insert_text_index_mark_xml(&xml, 1, &bibliography).unwrap();
        let xml = remove_text_index_mark_xml(&xml, 0).unwrap();
        assert!(xml.contains("<text:p>[Test]</text:p>"));
    }

    #[test]
    fn hostile_metadata_and_identity_are_rejected() {
        assert!(TextIndexMark::toc_range("", None).is_err());
        assert!(TextIndexMark::toc_point("x", Some(0)).is_err());
        assert!(
            TextIndexMark::bibliography_point(
                TextBibliographyType::Book,
                "x",
                vec![
                    (BibliographyField::Title, "a".to_string()),
                    (BibliographyField::Title, "b".to_string())
                ]
            )
            .is_err()
        );
        let spoofed = document(
            "<text:p xmlns:u=\"urn:bad\"><text:toc-mark text:string-value=\"x\" u:outline-level=\"1\"/></text:p>",
        );
        assert!(validated_marks(&spoofed).is_err());
        let crossed = document(
            "<text:p><text:toc-mark-start text:id=\"a\"/><text:toc-mark-end text:id=\"b\"/></text:p>",
        );
        assert!(validated_marks(&crossed).is_err());
    }

    #[test]
    fn libreoffice_point_range_and_bibliography_marks_round_trip() {
        let root = std::path::Path::new(env!("CARGO_MANIFEST_DIR")).join("../..");
        for relative in [
            "test-data/libreoffice-core/sw/qa/extras/ooxmlexport/data/toxmarkhyperlink.fodt",
            "test-data/libreoffice-core/sw/qa/extras/layout/data/tdf112256-diacritic-index-mark.fodt",
            "test-data/libreoffice-core/sw/qa/uibase/shells/data/protectedLinkCopy.fodt",
        ] {
            let xml = std::fs::read_to_string(root.join(relative)).unwrap();
            let marks = validated_marks(&xml).unwrap();
            assert!(!marks.is_empty());
            for mark in marks {
                mark.to_xml_fragments().unwrap();
            }
        }
    }
}
