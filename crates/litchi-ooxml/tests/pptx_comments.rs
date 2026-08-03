use litchi_ooxml::pptx::Package;
use litchi_ooxml::{OoxmlError, PackURI};
use litchi_opc::constants::{content_type as ct, relationship_type as rt};
use litchi_opc::part::BlobPart;

const PRESENTATION_XML: &[u8] =
    include_bytes!("../../../test-data/ooxml/pptx/comments/presentation.xml");
const SLIDE_XML: &[u8] = include_bytes!("../../../test-data/ooxml/pptx/comments/slide.xml");
const AUTHORS_XML: &[u8] =
    include_bytes!("../../../test-data/ooxml/pptx/comments/comment-authors.xml");
const COMMENTS_XML: &[u8] = include_bytes!("../../../test-data/ooxml/pptx/comments/comments.xml");

#[test]
fn presentation_comments_are_typed_and_legacy_adapters_preserve_data() {
    let package = package_with_comments();
    let presentation = package.presentation().unwrap();

    let comments = presentation.comments().unwrap().unwrap();
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

    let legacy_comments = presentation.get_comments().unwrap();
    assert_eq!(legacy_comments.len(), 1);
    assert_eq!(legacy_comments[0].0, 0);
    assert_eq!(legacy_comments[0].1.author_id, 0);
    assert_eq!(legacy_comments[0].1.text, "Review this slide");
    assert_eq!(legacy_comments[0].1.index, Some(1));

    let legacy_authors = presentation.get_comment_authors().unwrap();
    assert_eq!(legacy_authors.len(), 1);
    assert_eq!(legacy_authors[0].id, 0);
    assert_eq!(legacy_authors[0].name, "Ada Lovelace");
    assert_eq!(legacy_authors[0].initials, "AL");
}

#[test]
fn legacy_comment_adapter_rejects_external_comment_relationships() {
    let mut package = package_with_comments();
    let slide_name = PackURI::new("/ppt/slides/slide1.xml").unwrap();
    package
        .edit_opc(|opc| {
            let slide = opc.get_part_mut(&slide_name).unwrap();
            slide.rels_mut().remove("rIdComments");
            slide.rels_mut().add_relationship(
                rt::COMMENTS.to_string(),
                "https://example.invalid/comments.xml".to_string(),
                "rIdComments".to_string(),
                true,
            );
            Ok(())
        })
        .unwrap();

    assert!(matches!(
        package.presentation().unwrap().get_comments(),
        Err(OoxmlError::InvalidFormat(message))
            if message.contains("cannot be external")
    ));
}

fn package_with_comments() -> Package {
    let mut package = Package::new().unwrap();
    let presentation_name = PackURI::new("/ppt/presentation.xml").unwrap();
    let slide_name = PackURI::new("/ppt/slides/slide1.xml").unwrap();
    let authors_name = PackURI::new("/ppt/commentAuthors.xml").unwrap();
    let comments_name = PackURI::new("/ppt/comments/comment1.xml").unwrap();

    package
        .edit_opc(|opc| {
            let presentation = opc.get_part_mut(&presentation_name).unwrap();
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
            opc.add_part(Box::new(BlobPart::new(
                slide_name.clone(),
                ct::PML_SLIDE.to_string(),
                SLIDE_XML.to_vec(),
            )));
            opc.add_part(Box::new(BlobPart::new(
                authors_name,
                ct::PML_COMMENT_AUTHORS.to_string(),
                AUTHORS_XML.to_vec(),
            )));
            opc.add_part(Box::new(BlobPart::new(
                comments_name,
                ct::PML_COMMENTS.to_string(),
                COMMENTS_XML.to_vec(),
            )));
            opc.get_part_mut(&slide_name)
                .unwrap()
                .rels_mut()
                .add_relationship(
                    rt::COMMENTS.to_string(),
                    "../comments/comment1.xml".to_string(),
                    "rIdComments".to_string(),
                    false,
                );
            Ok(())
        })
        .unwrap();
    package
}
