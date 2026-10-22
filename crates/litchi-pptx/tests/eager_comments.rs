#![allow(
    clippy::expect_used,
    clippy::unwrap_used,
    reason = "integration assertions intentionally panic on malformed fixtures"
)]

use std::io::Cursor;

use litchi_opc::constants::{content_type as ct, relationship_type as rt};
use litchi_opc::part::BlobPart;
use litchi_opc::{OpcPackage, PackURI};
use litchi_pptx::presentation::related::raw;
use litchi_pptx::{Package, SourceBackedPresentation};

const PRESENTATION_XML: &[u8] = concat!(
    r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>"#,
    r#"<p:presentation xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">"#,
    r#"<p:sldIdLst><p:sldId id="256" r:id="rIdSlideOne"/></p:sldIdLst>"#,
    r#"</p:presentation>"#,
)
.as_bytes();
const SLIDE_XML: &[u8] = concat!(
    r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>"#,
    r#"<p:sld xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"><p:cSld><p:spTree/></p:cSld></p:sld>"#,
)
.as_bytes();
const AUTHORS_XML: &[u8] = concat!(
    r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>"#,
    r#"<p:cmAuthorLst xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"><p:cmAuthor id="7" name="Ada Lovelace" initials="AL" lastIdx="1" clrIdx="0"/></p:cmAuthorLst>"#,
)
.as_bytes();
const COMMENTS_XML: &[u8] = concat!(
    r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>"#,
    r#"<p:cmLst xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"><p:cm authorId="7" idx="3" dt="2026-01-02T03:04:05Z"><p:pos x="120" y="240"/><p:text>Review this slide</p:text></p:cm></p:cmLst>"#,
)
.as_bytes();

#[test]
fn eager_and_source_comment_snapshots_share_author_and_native_metadata() {
    let package = package_with_comments();
    let bytes = package.to_bytes().unwrap();

    let eager = {
        let presentation = package.presentation().unwrap();
        let slide = presentation.slide(0).unwrap().unwrap();
        slide.comment_snapshot().unwrap().unwrap()
    };
    let eager_entry = &eager.entries()[0];
    assert_eq!(eager_entry.text(), "Review this slide");
    assert_eq!(eager_entry.date_time(), Some("2026-01-02T03:04:05Z"));
    assert_eq!(eager_entry.location(), (120, 240));
    assert_eq!(eager_entry.author_name(), Some("Ada Lovelace"));
    assert_eq!(eager_entry.initials(), Some("AL"));
    assert_eq!(raw::comment_author_id(eager_entry), 7);
    assert_eq!(raw::comment_index(eager_entry), 3);
    assert_eq!(
        raw::author_part_uri(&eager).unwrap().as_str(),
        "/ppt/commentAuthors.xml"
    );

    let source = SourceBackedPresentation::from_reader(Cursor::new(bytes)).unwrap();
    let source_comments = source.slide(0).unwrap().comments().unwrap().unwrap();
    let source_entry = &source_comments.entries()[0];
    assert_eq!(source_entry.text(), eager_entry.text());
    assert_eq!(source_entry.author_name(), eager_entry.author_name());
    assert_eq!(source_entry.initials(), eager_entry.initials());
    assert_eq!(raw::comment_author_id(source_entry), 7);
    assert_eq!(raw::comment_index(source_entry), 3);
}

#[test]
fn eager_comment_snapshot_requires_an_author_relationship() {
    let package = package_with_comments();
    let bytes = package.to_bytes().unwrap();
    let mut opc = OpcPackage::from_bytes(&bytes, litchi_opc::SourceLimits::default()).unwrap();
    opc.get_part_mut(&PackURI::new("/ppt/presentation.xml").unwrap())
        .unwrap()
        .rels_mut()
        .remove("rIdCommentAuthors");
    let package = Package::from_opc_package(opc).unwrap();
    let presentation = package.presentation().unwrap();
    let result = presentation
        .slide(0)
        .unwrap()
        .unwrap()
        .comment_snapshot()
        .unwrap_err();
    assert!(format!("{result:?}").contains("comment-author"));
}

#[test]
fn eager_comment_snapshot_rejects_dangling_author_reference() {
    let package = rewrite_package(package_with_comments(), |opc| {
        let comments = opc
            .get_part_mut(&PackURI::new("/ppt/comments/comment1.xml").unwrap())
            .unwrap();
        comments.set_blob(
            String::from_utf8(comments.blob().to_vec())
                .unwrap()
                .replace("authorId=\"7\"", "authorId=\"99\"")
                .into_bytes(),
        );
    });
    let presentation = package.presentation().unwrap();
    assert!(
        presentation
            .slide(0)
            .unwrap()
            .unwrap()
            .comment_snapshot()
            .is_err()
    );
}

#[test]
fn eager_comment_snapshot_rejects_duplicate_authors() {
    let package = rewrite_package(package_with_comments(), |opc| {
        let authors = opc
            .get_part_mut(&PackURI::new("/ppt/commentAuthors.xml").unwrap())
            .unwrap();
        let xml = String::from_utf8(authors.blob().to_vec()).unwrap();
        authors.set_blob(
            xml.replacen(
                "</p:cmAuthorLst>",
                r#"<p:cmAuthor id="7" name="Grace Hopper" initials="GH" lastIdx="1" clrIdx="1"/></p:cmAuthorLst>"#,
                1,
            )
            .into_bytes(),
        );
    });
    let presentation = package.presentation().unwrap();
    assert!(
        presentation
            .slide(0)
            .unwrap()
            .unwrap()
            .comment_snapshot()
            .is_err()
    );
}

#[test]
fn eager_comment_snapshot_rejects_external_author_relationship() {
    let package = rewrite_package(package_with_comments(), |opc| {
        let presentation = opc
            .get_part_mut(&PackURI::new("/ppt/presentation.xml").unwrap())
            .unwrap();
        presentation.rels_mut().remove("rIdCommentAuthors");
        presentation.rels_mut().add_relationship(
            rt::COMMENT_AUTHORS.to_string(),
            "https://example.invalid/commentAuthors.xml".to_string(),
            "rIdCommentAuthors".to_string(),
            true,
        );
    });
    let presentation = package.presentation().unwrap();
    assert!(
        presentation
            .slide(0)
            .unwrap()
            .unwrap()
            .comment_snapshot()
            .is_err()
    );
}

fn package_with_comments() -> Package {
    let package = Package::new().unwrap();
    let presentation_name = PackURI::new("/ppt/presentation.xml").unwrap();
    let slide_name = PackURI::new("/ppt/slides/slide1.xml").unwrap();
    let authors_name = PackURI::new("/ppt/commentAuthors.xml").unwrap();
    let comments_name = PackURI::new("/ppt/comments/comment1.xml").unwrap();

    let bytes = package.to_bytes().unwrap();
    let mut package = OpcPackage::from_bytes(&bytes, litchi_opc::SourceLimits::default()).unwrap();
    {
        let presentation = package.get_part_mut(&presentation_name).unwrap();
        presentation.set_blob(PRESENTATION_XML.to_vec());
        presentation.rels_mut().add_relationship(
            rt::SLIDE.to_string(),
            "slides/slide1.xml".to_string(),
            "rIdSlideOne".to_string(),
            false,
        );
        presentation.rels_mut().add_relationship(
            rt::COMMENT_AUTHORS.to_string(),
            "commentAuthors.xml".to_string(),
            "rIdCommentAuthors".to_string(),
            false,
        );
    }
    package.add_part(Box::new(BlobPart::new(
        slide_name.clone(),
        ct::PML_SLIDE.to_string(),
        SLIDE_XML.to_vec(),
    )));
    package.add_part(Box::new(BlobPart::new(
        authors_name,
        ct::PML_COMMENT_AUTHORS.to_string(),
        AUTHORS_XML.to_vec(),
    )));
    package.add_part(Box::new(BlobPart::new(
        comments_name,
        ct::PML_COMMENTS.to_string(),
        COMMENTS_XML.to_vec(),
    )));
    package
        .get_part_mut(&slide_name)
        .unwrap()
        .rels_mut()
        .add_relationship(
            rt::COMMENTS.to_string(),
            "../comments/comment1.xml".to_string(),
            "rIdComments".to_string(),
            false,
        );
    Package::from_opc_package(package).unwrap()
}

fn rewrite_package(package: Package, edit: impl FnOnce(&mut OpcPackage)) -> Package {
    let bytes = package.to_bytes().unwrap();
    let mut opc = OpcPackage::from_bytes(&bytes, litchi_opc::SourceLimits::default()).unwrap();
    edit(&mut opc);
    Package::from_opc_package(opc).unwrap()
}
