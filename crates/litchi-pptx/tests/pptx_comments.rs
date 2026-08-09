#![allow(
    clippy::unwrap_used,
    clippy::expect_used,
    reason = "test assertions panic on failure by design"
)]

use litchi_opc::constants::{content_type as ct, relationship_type as rt};
use litchi_opc::part::BlobPart;
use litchi_opc::{OpcPackage, PackURI};
use litchi_pptx::comments::load_presentation_comments;
use litchi_pptx::{Error, Package};

// These package-writer fixtures are intentionally compact. The external
// relationship below is the hostile condition under test; formatting
// whitespace is unrelated and must not bypass production XML publication.
const PRESENTATION_XML: &[u8] = concat!(
    r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>"#,
    r#"<p:presentation xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">"#,
    r#"<p:sldIdLst><p:sldId id="256" r:id="rIdSlideOne"/></p:sldIdLst>"#,
    r#"</p:presentation>"#,
)
.as_bytes();
const SLIDE_XML: &[u8] = concat!(
    r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>"#,
    r#"<p:sld xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main">"#,
    r#"<p:cSld><p:spTree/></p:cSld></p:sld>"#,
)
.as_bytes();
const AUTHORS_XML: &[u8] = concat!(
    r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>"#,
    r#"<p:cmAuthorLst xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main">"#,
    r#"<p:cmAuthor id="0" name="Ada Lovelace" initials="AL" lastIdx="1" clrIdx="0"/>"#,
    r#"</p:cmAuthorLst>"#,
)
.as_bytes();
const COMMENTS_XML: &[u8] = concat!(
    r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>"#,
    r#"<p:cmLst xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main">"#,
    r#"<p:cm authorId="0" idx="1"><p:pos x="120" y="240"/>"#,
    r#"<p:text>Review this slide</p:text></p:cm></p:cmLst>"#,
)
.as_bytes();

#[test]
fn presentation_comments_are_typed_and_legacy_adapters_preserve_data() {
    let package = package_with_comments();
    let comments = load_presentation_comments(package.opc().unwrap())
        .unwrap()
        .unwrap();
    assert_eq!(comments.author_relationship_id, "rIdCommentAuthors");
    assert_eq!(comments.author_part_name, "/ppt/commentAuthors.xml");
    assert_eq!(comments.authors.len(), 1);
    assert_eq!(comments.authors[0].id, 0);
    assert_eq!(comments.authors[0].name, "Ada Lovelace");
    assert_eq!(comments.authors[0].initials, "AL");
    assert_eq!(comments.slides.len(), 1);
    assert_eq!(comments.slides[0].slide_part_name, "/ppt/slides/slide1.xml");
    assert_eq!(comments.slides[0].relationship_id, "rIdComments");
    assert_eq!(comments.slides[0].part_name, "/ppt/comments/comment1.xml");
    assert_eq!(comments.slides[0].comments.len(), 1);
    assert_eq!(comments.slides[0].comments[0].text, "Review this slide");

    assert_eq!(comments.authors[0].last_index, 1);
}

#[test]
fn legacy_comment_adapter_rejects_external_comment_relationships() {
    let mut package = package_with_comments();
    let slide_name = PackURI::new("/ppt/slides/slide1.xml").unwrap();
    package = edit_package(package, |opc| {
        let slide = opc.get_part_mut(&slide_name).unwrap();
        slide.rels_mut().remove("rIdComments");
        slide.rels_mut().add_relationship(
            rt::COMMENTS.to_string(),
            "https://example.invalid/comments.xml".to_string(),
            "rIdComments".to_string(),
            true,
        );
    });

    assert!(matches!(
        load_presentation_comments(package.opc().unwrap()),
        Err(Error::Invalid(message))
            if message.contains("cannot be external")
    ));
}

fn package_with_comments() -> Package {
    let mut package = Package::new().unwrap();
    let presentation_name = PackURI::new("/ppt/presentation.xml").unwrap();
    let slide_name = PackURI::new("/ppt/slides/slide1.xml").unwrap();
    let authors_name = PackURI::new("/ppt/commentAuthors.xml").unwrap();
    let comments_name = PackURI::new("/ppt/comments/comment1.xml").unwrap();

    let bytes = package.to_bytes().unwrap();
    let mut package = OpcPackage::from_bytes(&bytes).unwrap();
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
    }
    Package::from_opc_package(package).unwrap()
}

fn edit_package(mut package: Package, edit: impl FnOnce(&mut OpcPackage)) -> Package {
    let bytes = package.to_bytes().unwrap();
    let mut opc = OpcPackage::from_bytes(&bytes).unwrap();
    edit(&mut opc);
    Package::from_opc_package(opc).unwrap()
}
