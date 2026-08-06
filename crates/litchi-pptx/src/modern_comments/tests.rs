//! Regression coverage for the modern-comments model, codec, and package graph.

use super::*;
use litchi_opc::{BlobPart, OpcPackage, PackURI};

#[test]
fn semantic_facade_uses_contextual_names() {
    fn assert_type<T>() {}

    assert_type::<Change>();
    assert_type::<Changes>();
    assert_type::<Metadata>();
    assert_type::<Entry>();
    assert_type::<OpaqueXml>();
    assert_type::<Payload>();
    assert_type::<Instance>();
    assert_type::<Reaction>();
    assert_type::<Action>();
    assert_type::<Assign>();
    assert_type::<Details>();
    assert_type::<Event>();
    assert_type::<History>();
    assert_type::<Schedule>();
    assert_type::<Title>();
    assert_type::<Undo>();
    assert_type::<User>();

    assert_type::<semantic::changes::Reply>();
    assert_type::<semantic::changes::Replies>();
    assert_type::<semantic::extensions::List>();
    assert_type::<semantic::monikers::Kind>();
    assert_type::<semantic::monikers::List>();
    assert_type::<semantic::monikers::Node>();
    assert_type::<semantic::reactions::List>();
}

mod comment_tests {
    use super::*;
    use crate::modern_comments::{AC, MAX_BYTES, P188, PC};
    use litchi_opc::Part as _;
    use std::mem::size_of;

    const AUTHOR: &str = "{CD37207E-7903-4ED4-8AE8-017538D2DF7E}";
    const COMMENT: &str = "{62A8A96D-E5A8-4BFC-B993-A6EAE3907CAD}";
    const REPLY: &str = "{E524A04C-CF22-45D7-A60D-09322EA5A80D}";

    fn sdk_xml() -> Vec<u8> {
        format!(r#"<p188:cmLst xmlns:p188="{P188}" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><p188:cm id="{COMMENT}" authorId="{AUTHOR}" created="2024-12-30T20:26:06.503"><p188:txBody><a:bodyPr/><a:lstStyle/><a:p><a:r><a:t>Needs more cowbell</a:t></a:r></a:p></p188:txBody></p188:cm></p188:cmLst>"#).into_bytes()
    }

    fn package() -> OpcPackage {
        let mut package = OpcPackage::new();
        package.rels_mut().add_relationship(
            "http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument"
                .into(),
            "ppt/presentation.xml".into(),
            "rId1".into(),
            false,
        );
        let mut presentation = BlobPart::new(
            PackURI::new("/ppt/presentation.xml").unwrap(),
            "application/vnd.openxmlformats-officedocument.presentationml.presentation.main+xml"
                .into(),
            Vec::new(),
        );
        presentation.rels_mut().add_relationship(
            "http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide".into(),
            "slides/slide1.xml".into(),
            "rId1".into(),
            false,
        );
        package.add_part(Box::new(presentation));
        package.add_part(Box::new(BlobPart::new(
            PackURI::new("/ppt/slides/slide1.xml").unwrap(),
            SLIDE_CONTENT_TYPE.into(),
            Vec::new(),
        )));
        package
    }

    fn value() -> Part {
        Part {
            slide_part_name: "/ppt/slides/slide1.xml".into(),
            relationship_id: "rId9".into(),
            part_name: "/ppt/comments/modernComment1.xml".into(),
            comments: List::parse(&sdk_xml()).unwrap(),
        }
    }

    #[test]
    fn loads_microsoft_open_xml_sdk_documentation_specimen() {
        let parsed = List::parse(&sdk_xml()).unwrap();
        assert_eq!(parsed.comments.len(), 1);
        assert_eq!(parsed.comments[0].id, COMMENT);
        assert!(
            std::str::from_utf8(parsed.comments[0].text_body_xml.as_ref().unwrap())
                .unwrap()
                .contains("Needs more cowbell")
        );
        assert_eq!(List::parse(&parsed.to_xml().unwrap()).unwrap(), parsed);
    }

    #[test]
    fn package_round_trip_keeps_monikers_replies_and_extensions_inert() {
        let xml = format!(
            r#"<p188:cmLst xmlns:p188="{P188}" xmlns:pc="{PC}" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:x="urn:payload"><p188:cm id="{COMMENT}" authorId="{AUTHOR}" status="resolved" created="2026-07-19T12:00:00+08:00" assignedTo="{AUTHOR}" complete="50%" title="Review"><pc:sldMkLst><pc:sldMk/></pc:sldMkLst><p188:pos x="10" y="-20"/><p188:replyLst><p188:reply id="{REPLY}" authorId="{AUTHOR}" created="2026-07-19T12:01:00+08:00"><p188:txBody><a:bodyPr/><a:lstStyle/><a:p/></p188:txBody><p188:extLst><p:ext uri="{{A}}"><x:data relationship="rId999"/></p:ext></p188:extLst></p188:reply></p188:replyLst><p188:extLst><p:ext uri="{{B}}"><x:payload r:id="rId666" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"/></p:ext></p188:extLst></p188:cm></p188:cmLst>"#
        );
        let expected = List::parse(xml.as_bytes()).unwrap();
        assert_eq!(
            expected.comments[0].complete,
            Some(Progress::new(50).unwrap())
        );
        let mut package = package();
        let mut part = value();
        part.comments = expected.clone();
        store_modern_comment(&mut package, &part).unwrap();
        let loaded = load_modern_comments(&package).unwrap();
        assert_eq!(loaded.len(), 1);
        assert_eq!(loaded[0].comments, expected);
        assert!(
            loaded[0].comments.comments[0]
                .extension_xml
                .as_ref()
                .unwrap()
                .windows(6)
                .any(|window| window == b"rId666")
        );
        assert!(
            package
                .get_part(&PackURI::new("/ppt/comments/modernComment1.xml").unwrap())
                .unwrap()
                .rels()
                .is_empty()
        );
    }

    #[test]
    fn progress_is_bounded_typed_and_written_in_office_units() {
        assert_eq!(size_of::<Option<Progress>>(), size_of::<u32>());
        assert_eq!(Progress::ZERO.thousandths(), 0);
        assert_eq!(Progress::FULL.thousandths(), 100_000);
        assert_eq!(Progress::new(25).unwrap().thousandths(), 25_000);
        assert_eq!(
            Progress::from_thousandths(50_250).unwrap().to_string(),
            "50250"
        );
        assert!(Progress::new(101).is_err());
        assert!(Progress::from_thousandths(100_001).is_err());

        let xml = format!(
            r#"<p188:cmLst xmlns:p188="{P188}"><p188:cm id="{COMMENT}" authorId="{AUTHOR}" created="2024-12-30T20:26:06.503" complete="50.25%"/></p188:cmLst>"#
        );
        let parsed = List::parse(xml.as_bytes()).unwrap();
        assert_eq!(
            parsed.comments[0].complete,
            Some(Progress::from_thousandths(50_250).unwrap())
        );
        let serialized = parsed.to_xml().unwrap();
        assert!(
            serialized
                .windows(b"complete=\"50250\"".len())
                .any(|window| window == b"complete=\"50250\"")
        );
        assert_eq!(List::parse(&serialized).unwrap(), parsed);

        for complete in ["-1%", "100.01%", "50.123%", "1e2%", "100001", ""] {
            let xml = format!(
                r#"<p188:cmLst xmlns:p188="{P188}"><p188:cm id="{COMMENT}" authorId="{AUTHOR}" created="2024-12-30T20:26:06.503" complete="{complete}"/></p188:cmLst>"#
            );
            assert!(
                List::parse(xml.as_bytes()).is_err(),
                "accepted invalid progress {complete:?}"
            );
        }
    }

    #[test]
    fn rejects_hostile_or_schema_invalid_comment_xml() {
        let cases = [
            format!(r#"<!DOCTYPE x><p188:cmLst xmlns:p188="{P188}"/>"#),
            r#"<x:cmLst xmlns:x="urn:wrong"/>"#.to_string(),
            format!(
                r#"<p188:cmLst xmlns:p188="{P188}"><p188:cm id="bad" authorId="{AUTHOR}" created="2024-12-30T20:26:06.503"/></p188:cmLst>"#
            ),
            format!(
                r#"<p188:cmLst xmlns:p188="{P188}"><p188:cm id="{COMMENT}" authorId="{AUTHOR}" status="pending" created="2024-12-30T20:26:06.503"/></p188:cmLst>"#
            ),
            format!(
                r#"<p188:cmLst xmlns:p188="{P188}" xmlns:pc="{PC}" xmlns:ac="{AC}"><p188:cm id="{COMMENT}" authorId="{AUTHOR}" created="2024-12-30T20:26:06.503"><pc:sldMkLst/><ac:deMkLst/></p188:cm></p188:cmLst>"#
            ),
            format!(
                r#"<p188:cmLst xmlns:p188="{P188}"><p188:cm id="{COMMENT}" authorId="{AUTHOR}" created="2024-12-30T20:26:06.503"><p188:txBody/><p188:replyLst/></p188:cm></p188:cmLst>"#
            ),
        ];
        for xml in cases {
            assert!(List::parse(xml.as_bytes()).is_err(), "accepted {xml}");
        }
        assert!(List::parse(&vec![b' '; MAX_BYTES + 1]).is_err());
    }

    #[test]
    fn rejects_invalid_package_graphs_and_failed_store_is_atomic() {
        let mut external = package();
        external
            .get_part_mut(&PackURI::new("/ppt/slides/slide1.xml").unwrap())
            .unwrap()
            .rels_mut()
            .add_relationship(
                MODERN_COMMENT_RELATIONSHIP_TYPE.into(),
                "https://invalid.example/comments.xml".into(),
                "rId9".into(),
                true,
            );
        assert!(load_modern_comments(&external).is_err());

        let mut wrong_source = package();
        wrong_source
            .get_part_mut(&PackURI::new("/ppt/presentation.xml").unwrap())
            .unwrap()
            .rels_mut()
            .add_relationship(
                MODERN_COMMENT_RELATIONSHIP_TYPE.into(),
                "comments/modern.xml".into(),
                "rId9".into(),
                false,
            );
        assert!(load_modern_comments(&wrong_source).is_err());

        let mut orphan = package();
        orphan.add_part(Box::new(BlobPart::new(
            PackURI::new("/ppt/comments/orphan.xml").unwrap(),
            MODERN_COMMENT_CONTENT_TYPE.into(),
            sdk_xml(),
        )));
        assert!(load_modern_comments(&orphan).is_err());

        let mut atomic = package();
        let mut invalid_value = value();
        invalid_value.comments.comments[0].id = "not-a-guid".into();
        let before_parts = atomic.iter_parts().count();
        let before_rels = atomic
            .get_part(&PackURI::new("/ppt/slides/slide1.xml").unwrap())
            .unwrap()
            .rels()
            .len();
        assert!(store_modern_comment(&mut atomic, &invalid_value).is_err());
        assert_eq!(atomic.iter_parts().count(), before_parts);
        assert_eq!(
            atomic
                .get_part(&PackURI::new("/ppt/slides/slide1.xml").unwrap())
                .unwrap()
                .rels()
                .len(),
            before_rels
        );
    }
}

mod author_tests {
    use super::*;
    use crate::modern_comments::store_modern_comment;
    use crate::modern_comments::{MAX_BYTES, P188};
    use litchi_opc::Part as _;

    const AUTHOR: &str = "{CD37207E-7903-4ED4-8AE8-017538D2DF7E}";
    const OTHER: &str = "{0B2043D4-0908-4C42-8A79-51EA2CC309F7}";
    const COMMENT: &str = "{62A8A96D-E5A8-4BFC-B993-A6EAE3907CAD}";

    fn sdk_author_xml() -> Vec<u8> {
        format!(r#"<p188:authorLst xmlns:p188="{P188}"><p188:author id="{AUTHOR}" name="Ada Lovelace" initials="AL" userId="ada@example.com::4b640067-2830-4c10-9c4f-5879bb2e41d1" providerId=""/></p188:authorLst>"#).into_bytes()
    }

    fn package() -> OpcPackage {
        let mut package = OpcPackage::new();
        package.rels_mut().add_relationship(
            "http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument"
                .into(),
            "ppt/presentation.xml".into(),
            "rId1".into(),
            false,
        );
        let mut presentation = BlobPart::new(
            PackURI::new("/ppt/presentation.xml").unwrap(),
            "application/vnd.openxmlformats-officedocument.presentationml.presentation.main+xml"
                .into(),
            Vec::new(),
        );
        presentation.rels_mut().add_relationship(
            "http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide".into(),
            "slides/slide1.xml".into(),
            "rId1".into(),
            false,
        );
        package.add_part(Box::new(presentation));
        package.add_part(Box::new(BlobPart::new(
            PackURI::new("/ppt/slides/slide1.xml").unwrap(),
            "application/vnd.openxmlformats-officedocument.presentationml.slide+xml".into(),
            Vec::new(),
        )));
        package
    }

    fn author_part() -> AuthorPart {
        AuthorPart {
            relationship_id: "rId8".into(),
            part_name: "/ppt/authors/author1.xml".into(),
            authors: Authors::parse(&sdk_author_xml()).unwrap(),
        }
    }

    fn comment_part(author: &str) -> Part {
        let xml = format!(
            r#"<p188:cmLst xmlns:p188="{P188}"><p188:cm id="{COMMENT}" authorId="{author}" created="2024-12-30T20:26:06.503" assignedTo="{author}"/></p188:cmLst>"#
        );
        Part {
            slide_part_name: "/ppt/slides/slide1.xml".into(),
            relationship_id: "rId9".into(),
            part_name: "/ppt/comments/modernComment1.xml".into(),
            comments: List::parse(xml.as_bytes()).unwrap(),
        }
    }

    #[test]
    fn loads_open_xml_sdk_shaped_author_specimen() {
        let parsed = Authors::parse(&sdk_author_xml()).unwrap();
        assert_eq!(parsed.authors.len(), 1);
        assert_eq!(parsed.authors[0].name, "Ada Lovelace");
        assert_eq!(parsed.authors[0].provider_id, "");
        assert_eq!(Authors::parse(&parsed.to_xml().unwrap()).unwrap(), parsed);
    }

    #[test]
    fn author_and_comment_package_graph_round_trip_and_resolve() {
        let extension = format!(r#"<p188:extLst xmlns:p188="{P188}" xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:x="urn:payload"><p:ext uri="{{A}}"><x:data authorId="{OTHER}" relationship="rId999"/></p:ext></p188:extLst>"#).into_bytes();
        let mut authors = author_part();
        authors.authors.authors[0].extension_xml = Some(extension.clone());
        let mut package = package();
        store_modern_comment_authors(&mut package, &authors).unwrap();
        store_modern_comment(&mut package, &comment_part(AUTHOR)).unwrap();
        let graph = load_modern_comment_graph(&package).unwrap();
        assert_eq!(
            graph.authors.unwrap().authors.authors[0].extension_xml,
            Some(extension)
        );
        assert_eq!(graph.comments.len(), 1);
    }

    #[test]
    fn rejects_hostile_author_grammar_and_unresolved_modeled_references() {
        let cases = [
            format!(r#"<!DOCTYPE x><p188:authorLst xmlns:p188="{P188}"/>"#),
            "<x:authorLst xmlns:x=\"urn:wrong\"/>".into(),
            format!(
                r#"<p188:authorLst xmlns:p188="{P188}"><p188:author id="bad" name="A" userId="u" providerId="p"/></p188:authorLst>"#
            ),
            format!(
                r#"<p188:authorLst xmlns:p188="{P188}"><p188:author id="{AUTHOR}" name="A" userId="u"/></p188:authorLst>"#
            ),
            format!(
                r#"<p188:authorLst xmlns:p188="{P188}"><p188:author id="{AUTHOR}" name="A" userId="u" providerId="p"><p188:extLst/><p188:extLst/></p188:author></p188:authorLst>"#
            ),
            format!(
                r#"<p188:authorLst xmlns:p188="{P188}"><p188:author id="{AUTHOR}" name="A" userId="u" providerId="p"/><p188:author id="{AUTHOR}" name="B" userId="v" providerId="p"/></p188:authorLst>"#
            ),
        ];
        for xml in cases {
            assert!(Authors::parse(xml.as_bytes()).is_err(), "accepted {xml}");
        }
        assert!(Authors::parse(&vec![b' '; MAX_BYTES + 1]).is_err());

        let authors = author_part();
        assert!(
            validate_modern_comment_author_references(Some(&authors), &[comment_part(OTHER)])
                .is_err()
        );
        assert!(validate_modern_comment_author_references(None, &[comment_part(AUTHOR)]).is_err());
    }

    #[test]
    fn rejects_author_package_graphs_and_failed_store_is_atomic() {
        let mut external = package();
        external
            .get_part_mut(&PackURI::new("/ppt/presentation.xml").unwrap())
            .unwrap()
            .rels_mut()
            .add_relationship(
                MODERN_COMMENT_AUTHOR_RELATIONSHIP_TYPE.into(),
                "https://invalid.example/authors.xml".into(),
                "rId8".into(),
                true,
            );
        assert!(load_modern_comment_authors(&external).is_err());

        let mut orphan = package();
        orphan.add_part(Box::new(BlobPart::new(
            PackURI::new("/ppt/authors/orphan.xml").unwrap(),
            MODERN_COMMENT_AUTHOR_CONTENT_TYPE.into(),
            sdk_author_xml(),
        )));
        assert!(load_modern_comment_authors(&orphan).is_err());

        let mut outbound = package();
        store_modern_comment_authors(&mut outbound, &author_part()).unwrap();
        outbound
            .get_part_mut(&PackURI::new("/ppt/authors/author1.xml").unwrap())
            .unwrap()
            .rels_mut()
            .add_relationship(
                "urn:forbidden".into(),
                "other.xml".into(),
                "rId1".into(),
                false,
            );
        assert!(load_modern_comment_authors(&outbound).is_err());

        let mut atomic = package();
        store_modern_comment(&mut atomic, &comment_part(OTHER)).unwrap();
        let before_parts = atomic.iter_parts().count();
        let before_rels = atomic
            .get_part(&PackURI::new("/ppt/presentation.xml").unwrap())
            .unwrap()
            .rels()
            .len();
        assert!(store_modern_comment_authors(&mut atomic, &author_part()).is_err());
        assert_eq!(atomic.iter_parts().count(), before_parts);
        assert_eq!(
            atomic
                .get_part(&PackURI::new("/ppt/presentation.xml").unwrap())
                .unwrap()
                .rels()
                .len(),
            before_rels
        );
    }
}
